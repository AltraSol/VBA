Attribute VB_Name = "zRun_R"
Option Explicit
'===================================================================
'## Module: zRun_R.bas
'Subs and Functions to interface with R in VBA; relies on
'zPortable_Subs and zPortable_Functions from github.com/ulchc/VBA-tools
'===================================================================
'------------------------------------------------------------------- VBA
'  QuickRun_RScript(ByVal ScriptContents As String)
'
''   Writes a temporary .R script containing {ScriptContents}, runs
''   it, prompts for the deletion of the temporary script
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  WriteTemp_RScript(ByVal ScriptContents As String)
'
''   Creates a random named temporary folder on desktop, creates an
''   .R file "Temp.R" containing {ScriptContents}, returns Temp.R path
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  FindAndRun_RScript(ByVal ScriptLocation)
'
''   Takes a string or cell reference {RScriptPath} & runs it on the
''   latest version of R on the OS
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Run_RScript(ByVal RLocation As String, _
'              ByVal ScriptLocation As String, _
'              Optional ByVal Visibility As String)
'
''   Uses the RScript.exe pointed to by {RLocation} to run the script
''   found at {ScriptLocation}. Rscript.exe window displayed by default,
''   but {Visibility}:= "VeryHidden" or "Minimized" can be used
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_RExePath() As String
'
''   Returns the path to the latest version of Rscript.exe
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_LatestRVersion(ByVal RVersions As Variant)
'
''   Returns the latest version of R currently installed
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_RVersions(ByVal RFolderPath As String)
'
''   Returns an array of the R versions currently installed
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_RFolder() As String
'
''   Returns the parent R folder path which houses the installed
''   versions of R on the OS from which the sub is called
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Test_QuickRun_RScript()
'
''   Writes a computationally intensive script to Desktop and asks
''   if you want to run it (to visually verify all zRun_R f(x) worked)
'
'-------------------------------------------------------------------

Sub Test_QuickRun_RScript()

'NOTE: It's best to paste a script into a cell and reading
'it's .Value as opposed to writing it in the VBA editor

Dim PackagesList As String, _
    arrPackages As Variant, _
    i As Integer, _
    HighComputeScript As String
    
    PackagesList = "pdftools, tesseract, stringr, dplyr, qdapRegex, tidyr, stringi, purrr, openxlsx, tidyverse"
    arrPackages = Split(PackagesList, ", ")
    
    For i = LBound(arrPackages) To UBound(arrPackages)

        'Formatting to:
        'if (!require(Package)) install.packages('Package')
        'library(Package)
        
        'Installing & referencing many packages is computationally intensive
        'which allows a chance to verify the script is running on the device
        
        HighComputeScript = HighComputeScript & vbNewLine & _
                            "if (!require(" & arrPackages(i) & _
                            ")) install.packages('" & arrPackages(i) & "')" & _
                            vbNewLine & "library (" & arrPackages(i) & ")"
    Next i
    
        Dim Answer: Answer = MsgBox("The following script is about to be ran in R:" & _
                                    vbNewLine & HighComputeScript & vbNewLine & vbNewLine & _
                                    "Press OK to continue, or Cancel to exit.", vbOKCancel)
                                    
        If Answer = vbOK Then Call QuickRun_RScript(HighComputeScript)
        
End Sub

Sub QuickRun_RScript(ByVal ScriptContents As String)

Dim TempScriptPath As String, _
    TempFolderPath As String, _
    i As Integer, _
    Slash As String, _
    Answer As String
    
        TempScriptPath = WriteTemp_RScript(ScriptContents)
        If MyOS = "Windows" Then Slash = "\" Else Slash = "/"
        Call FindAndRun_RScript(TempScriptPath)
                
                'NOTE: MsgBox question serves as both an option and a workaround for long R procedures
                'which prevent VBA's command line call from deleting Temp.R prior to Rscript.exe unloading Temp.R
                Answer = MsgBox("Temporary script written to desktop and ran in R." & vbNewLine & vbNewLine & _
                                "Would you like to delete the temporary file and it's folder?", vbYesNo, "Delete Temp.R File & Folder?")
            
                If Answer = vbYes Then
                    'Deletion successful
                    If Delete_FileAndFolder(TempScriptPath) = True Then
                        
                        'Initially set {TempFolderPath} to {TempScriptPath} prior to loop
                        TempFolderPath = TempScriptPath
                        
                        'Reduce {TempFolderPath} until it's a directory
                        For i = Len(TempFolderPath) To 1 Step -1
            
                            TempFolderPath = Left(TempFolderPath, Len(TempFolderPath) - 1)
                            If Right(TempFolderPath, 1) = Slash Then
                                TempFolderPath = Left(TempFolderPath, Len(TempFolderPath) - 1)
                                Exit For
                            End If
                            
                        Next i
                        MsgBox "Temp.R and folder directory " & TempFolderPath & " deleted."
                    Else
                        MsgBox "Error deleting " & TempFolderPath
                    End If
                End If
            
End Sub

Function WriteTemp_RScript(ByVal ScriptContents As String)

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
            
            WriteTemp_RScript = TempFolder & "\" & "Temp.R"

End Function

Sub FindAndRun_RScript(ByVal ScriptLocation)

If TypeName(ScriptLocation) = "Range" Then
    ScriptLocation = ScriptLocation.Cells(1).Value
End If

Call Run_RScript(RLocation:=Get_RExePath, _
                 ScriptLocation:=ScriptLocation, _
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


