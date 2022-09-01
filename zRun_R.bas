Attribute VB_Name = "zRun_R"
Option Explicit
'===================================================================
'Portable module of subs for interfacing with R in VBA
'Author: Ulchc
'TODO: Verify working on MacOS
'===================================================================
'-------------------------------------------------------------------
' Temp_RScript(ScriptContents As String)
'
'   Creates a random named temporary folder, creates an .R script
'   titled Temp.R containing {ScriptContents}, returns Temp.R path
'
'-------------------------------------------------------------------
' Get_DesktopPath()
'
'   Returns Windows desktop directory (even if hosted on OneDrive)
'
'-------------------------------------------------------------------
' Easy_RScript(ByVal RScriptPath)
'
'   Takes a string or cell reference {RScriptPath} & runs it on the
'   latest version of R on the OS
'
'-------------------------------------------------------------------
' Run_RScript(ByVal RLocation As String, _
'             ByVal ScriptLocation As String, _
'             Optional ByVal Visibility As String)
'
'   Uses the RScript.exe pointed to by {RLocation} to run the script
'   found at {ScriptLocation}. Rscript.exe window displayed by default,
'   but {Visibility}:= "VeryHidden" or "Minimized" can be used
'
'-------------------------------------------------------------------
' Get_RExePath() As String
'
'   Returns the path to the latest version of Rscript.exe
'
'-------------------------------------------------------------------
' Get_LatestRVersion(ByVal RVersions As Variant)
'
'   Returns the latest version of R currently installed
'
'-------------------------------------------------------------------
' Get_RVersions(ByVal RFolderPath As String)
'
'   Returns an array of the R versions currently installed
'
'-------------------------------------------------------------------
' Get_RFolder() As String
'
'   Returns the parent R folder path which houses the installed
'   versions of R on the OS from which the sub is called
'
'-------------------------------------------------------------------
' MyOS()
'
'   'Returns "Windows",  "Mac", or "Neither Windows or Mac"
'
'-------------------------------------------------------------------

Sub R_Example()

    Call Easy_RScript("...Something.R")
    
End Sub

Sub R_CellExample()

Dim ScriptLocation As String
    'Paste script in selection that is long enough to keep RScript window open
    ScriptLocation = Temp_RScript(Selection.Cells(1).Value)

    Call Easy_RScript(ScriptLocation)
    
End Sub

Function Temp_RScript(ScriptContents As String)

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
Dim TempFolder As String
    
    'Create temporary folder to house temp.R
    TempFolder = Get_DesktopPath & "\Temp" & Left(Format(Now(), "ss") * Rnd() * 1000, 3)
    Call MkDir(TempFolder)
    
    'Write temp.R into {TempFolder}
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile(TempFolder & "\" & "Temp.R", True, False)
        Fileout.Write ScriptContents
        Fileout.Close
            
            Temp_RScript = TempFolder & "\" & "Temp.R"

End Function

Function Get_DesktopPath()

If MyOS = "Windows" Then
    
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim UserFolder As Object: Set UserFolder = FSO.GetFolder("C:\Users")
    
    Dim Item As Object, _
        UserPath As String, _
        DesktopPath As String
    
        
        For Each Item In UserFolder.SubFolders
            'TODO verify this works on other windows OS
            If InStr(Item.Path, "All Users") = 0 And _
               InStr(Item.Path, "Default") = 0 And _
               InStr(Item.Path, "Public") = 0 Then
               
               UserPath = Item.Path
               Exit For
            
            End If
        Next Item
            
            Dim UserNameFolder As Object: Set UserNameFolder = FSO.GetFolder(UserPath)
            For Each Item In UserNameFolder.SubFolders
                If InStr(Item.Path, "Desktop") > 1 Then
                    DesktopPath = Item.Path
                    Exit For
                End If
            Next Item
                
                'C:\Users\UserName\Desktop will be found in above procedure
                'except in cases that this file is hosted on OneDrive
                If DesktopPath = vbNullString Then
                    DesktopPath = UserPath & "\OneDrive\Desktop"
                End If
                
                    Set UserFolder = Nothing
                    Set UserNameFolder = Nothing
                    Set FSO = Nothing
            
ElseIf MyOS = "Mac" Then
    'TODO
End If

    Get_DesktopPath = DesktopPath
    
End Function

Sub Easy_RScript(ByVal RScriptPath)

If TypeName(RScriptPath) = "Range" Then
    RScriptPath = RScriptPath.Cells(1).Value
End If

Call Run_RScript(RLocation:=Get_RExePath, _
                 ScriptLocation:=RScriptPath, _
                 Visibility:="Visible")
                 
End Sub

Sub Run_RScript(ByVal RLocation As String, ByVal ScriptLocation As String, Optional ByVal Visibility As String)

Dim WaitTillComplete As Boolean: WaitTillComplete = True
Dim Style As Integer: Style = 1

Dim oShell As Object, _
    ErrorCode As Integer, _
    eRExe As String, _
    eRScript As String, _
    RExe_RScript As String
    
    If Visibility = "VeryHidden" Then
        Style = 0
    ElseIf Visibility = "Minimized" Then
        Style = 2
    End If
    
        Set oShell = CreateObject("WScript.Shell")
        
        eRExe = Chr(34) & Replace(RLocation, "\", "\\") & Chr(34)
        eRScript = Chr(34) & Replace(ScriptLocation, "\", "\\") & Chr(34)

            RExe_RScript = eRExe & eRScript
            
                ErrorCode = oShell.Run(RExe_RScript, Style, WaitTillComplete)

End Sub

Function Get_RExePath() As String

Dim RVersions As Variant: RVersions = Get_RVersions(Get_RFolder)
Dim LatestRVersion As String: LatestRVersion = Get_LatestRVersion(RVersions)
    
    Get_RExePath = LatestRVersion & "\bin\Rscript.exe"
             
End Function

Function Get_LatestRVersion(ByVal RVersions As Variant)

Dim i As Integer
    For i = LBound(RVersions) To UBound(RVersions)
        If Get_LatestRVersion < RVersions(i) Then
           Get_LatestRVersion = RVersions(i)
        End If
    Next i
    
End Function

Function Get_RVersions(ByVal RFolderPath As String)

Dim FSO As Object, _
    RProgramFolder As Object, _
    VersionFolder As Object, _
    arrRVersions() As Variant, _
    i As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set RProgramFolder = FSO.GetFolder(RFolderPath)

        For Each VersionFolder In RProgramFolder.SubFolders
            ReDim Preserve arrRVersions(i): i = i + 1
            arrRVersions(UBound(arrRVersions)) = VersionFolder.Path
        Next VersionFolder
    
            Set VersionFolder = Nothing
            Set RProgramFolder = Nothing
            Set FSO = Nothing
                
               Get_RVersions = arrRVersions
End Function

Function Get_RFolder() As String

Dim OS As String: OS = MyOS

    If OS = "Windows" Then
        Get_RFolder = "C:\Program Files\R"
    ElseIf OS = "Mac" Then
        Get_RFolder = "/Library/Frameworks/R.framework/Resources/bin/R"
    End If

End Function

Function MyOS()

    If InStr(1, Application.OperatingSystem, "Windows") <> 0 Then
        MyOS = "Windows"
    ElseIf InStr(1, Application.OperatingSystem, "Mac") <> 0 Then
        MyOS = "Mac"
    Else
        MyOS = "Neither Windows or Mac"
    End If
        
End Function

