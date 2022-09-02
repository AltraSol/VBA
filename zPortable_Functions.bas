Attribute VB_Name = "zPortable_Functions"
Option Explicit
'===================================================================
'## Module: zPortable_Functions.bas
'Portable module of functions which can be exported to any workbook
'and are only dependent on one-another
'===================================================================
'------------------------------------------------------------------- VBA
'  Tabs_MatchingCodeName(MatchCodeName As String,
'                        ExcludePerfectMatch As Boolean)
'
''   Returns array of tab names with MatchCodeName found in the CodeName
''   property (useful for detecting copies of a code-named template)
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  WorksheetExists (aName)
'
''   True or False dependent on if tab name {aName} already exists
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  ExtractFirstInt_RightToLeft (aVariable)
'
''   Returns the first integer found in a string when searcing
''   from the right end of the string to the left
'
''   ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  ExtractFirstInt_LeftToRight (aVariable)
'
''   Returns the first integer found in a string when searcing
''   from the left end of the string to the right
'
''   ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Truncate_Before_Int (aString)
'
''   Removes characters before first integer in a sequence of characters
'
''   Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Truncate_After_Int (aString)
'
''   Removes characters after first integer in a sequence of characters
'
''   Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  IsInt_NoTrailingSymbols (aNumeric)
'
''   Checks if supplied value is both numeric, and contains no numeric
''   symbols (different from IsNumeric)
'
''   IsInt_NoTrailingSymbols(9999) = True
''   IsInt_NoTrailingSymbols(9999,) = False
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  MyOS()
'
''   Returns "Windows",  "Mac", or "Neither Windows or Mac"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_WindowsUsername()
'
''   Loops through folders to find paths matching C:\Users\...\AppData
''   then extracts Username from correct path. Superior to reading
''   .FullName of workbook which does not work for OneDrive files
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_MacUsername()
'
''   Reads Activeworkbook.FullName property to get Mac user
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_Username()
'
''   Returns username regardless of Windows or Mac OS
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_DesktopPath()
'
''   Returns Mac or Windows desktop directory (even if on OneDrive)
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Delete_FileAndFolder(ByVal aFilePath As String)
'
''   Read code directly prior to use
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Print_Pad()
'
''   Uses Debug.Print to print a timestamped seperator of "======"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Print_Named(Something, Optional Label)
'
''   Uses Debug.Print to add a space between each {Something} printed,
''   labels each {Something} if {Label} supplied
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Clipboard_Load(ByVal aString As String)
'
''   Stores {aString} in clipboard
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Clipboard_Read(Optional IfRngConcatAllVals As Boolean = True,
'                 Optional Sep As String = ", ")
'
''   Returns text from the copied object (clipboard text or range)
'
''   >> NOT TO BE USED ON-SHEET << creates a sheet each refresh
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  Get_CopiedRangeVals()
'
''   If range copied, returns an array of each non-blank Cell.Value
'
''   >> NOT TO BE USED ON-SHEET << creates a sheet each refresh
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' Clipboard_IsRange()
'
''   Returns True if a range is currently copied; only works in VBA
'
'-------------------------------------------------------------------

Function Clipboard_Load(ByVal aString As String)

On Error GoTo NoLoad
    CreateObject("HTMLFile").ParentWindow.ClipboardData.SetData "text", aString
    Clipboard_Load = True
    Exit Function
    
NoLoad:
Clipboard_Load = False
On Error GoTo -1

End Function

'>> NOT TO BE USED ON-SHEET << creates a sheet each refresh
Function Clipboard_Read(Optional IfRngConcatAllVals As Boolean = True, Optional Sep As String = ", ")
On Error GoTo NoRead

If Clipboard_IsRange() = True Then
    Dim CopiedRangeText As Variant
        CopiedRangeText = Get_CopiedRangeVals()
        
        If IfRngConcatAllVals = False Then
            Clipboard_Read = CopiedRangeText(LBound(CopiedRangeText))
        Else
            Clipboard_Read = Application.WorksheetFunction.TextJoin(Sep, True, CopiedRangeText)
        End If
        
Else
    Clipboard_Read = CreateObject("HTMLFile").ParentWindow.ClipboardData.GetData("text")
End If

Exit Function

NoRead:
Clipboard_Read = False
On Error GoTo -1
End Function

'>> NOT TO BE USED ON-SHEET << creates a sheet each refresh
Function Get_CopiedRangeVals()

If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
            
Dim aCell As Range, _
    arrCellText() As Variant
    
    Sheets.Add
    
    On Error GoTo PasteIssue:
    ActiveSheet.Paste Link:=True
        
        ReDim Preserve arrCellText(0)

        For Each aCell In Selection
            If aCell.Value <> vbNullString Then
                arrCellText(UBound(arrCellText)) = aCell.Value
                ReDim Preserve arrCellText(UBound(arrCellText()) + 1)
            End If
        Next aCell
            
            'To reverse final ReDim after the last aCell added
            ReDim Preserve arrCellText(UBound(arrCellText()) - 1)
                
                Get_CopiedRangeVals = arrCellText()
PasteIssue:
                ActiveSheet.Delete
                       
If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
If Application.ScreenUpdating = False Then Application.ScreenUpdating = True

End Function

Function Clipboard_IsRange()

Clipboard_IsRange = False
Dim aFormat As Variant

    For Each aFormat In Application.ClipboardFormats
        If aFormat = xlClipboardFormatCSV Then
            Clipboard_IsRange = True
        End If
    Next
    
End Function

Function Tabs_MatchingCodeName(ByVal MatchCodeName As String, ExcludePerfectMatch As Boolean)
Dim aSheet As Worksheet, _
    i As Integer, _
    j As Integer, _
    arrTabNames() As Variant

    If ExcludePerfectMatch = False Then
        'Loop through all tabs in this workbook
        For i = 0 To ActiveWorkbook.Sheets.Count - 1
            'If CodeName of sheet includes the aCodeName supplied
            If InStr(1, ActiveWorkbook.Sheets(i + 1).CodeName, MatchCodeName) Then
                    'Then add tab name to the array
                    ReDim Preserve arrTabNames(j): j = j + 1
                    arrTabNames(UBound(arrTabNames)) = ActiveWorkbook.Sheets(i + 1).Name
            End If
        Next i
    ElseIf ExcludePerfectMatch = True Then
        'Loop through all tabs in this workbook
        For i = 0 To ActiveWorkbook.Sheets.Count - 1
            'If CodeName of sheet includes the aCodeName supplied;
            'but is not a perfect match with aCodeName
            If InStr(1, ActiveWorkbook.Sheets(i + 1).CodeName, MatchCodeName) And _
               ActiveWorkbook.Sheets(i + 1).CodeName <> MatchCodeName Then
                    'Then add tab name to the array
                    ReDim Preserve arrTabNames(j): j = j + 1
                    arrTabNames(UBound(arrTabNames)) = ActiveWorkbook.Sheets(i + 1).Name
            End If
        Next i
    End If

Tabs_MatchingCodeName = arrTabNames

End Function

Function WorksheetExists(aName As String, Optional wb As Workbook) As Boolean
Dim aSheet As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
        
        On Error Resume Next
            Set aSheet = wb.Sheets(aName)
        On Error GoTo 0
        
            WorksheetExists = Not aSheet Is Nothing
    
End Function

Function ExtractFirstInt_RightToLeft(ByVal aVariable)
On Error GoTo ExtractNothing:

Dim i As Integer, _
    CheckCharacter As String, _
    CountCharsToRemove As Integer, _
    NewStrLength As Integer

'If range supplied, convert to string
If TypeName(aVariable) = "Range" Then
    aVariable = aVariable.Cells(1).Value
End If

aVariable = Trim(aVariable)

'Return immediately if already integer
If IsInt_NoTrailingSymbols(aVariable) Then
    ExtractFirstInt_RightToLeft = aVariable
    Exit Function
End If

If Len(aVariable) = 0 Then
    GoTo ExtractNothing
End If

'Remove any characters **following an integer
aVariable = Truncate_After_Int(aVariable)
'Pad {aVariable} with **starting character to set up loop
aVariable = "A" & aVariable

    'Remove integers one-by-one until character found
    For i = 1 To Len(aVariable)
        If IsInt_NoTrailingSymbols(Right(aVariable, i)) = False Then
            ExtractFirstInt_RightToLeft = Right(aVariable, i - 1)
            Exit Function
        End If
    Next i
    
ExtractNothing:
ExtractFirstInt_RightToLeft = vbNullString

End Function

Function ExtractFirstInt_LeftToRight(ByVal aVariable)
On Error GoTo ExtractNothing:

Dim i As Integer, _
    CheckCharacter As String, _
    CountCharsToRemove As Integer, _
    NewStrLength As Integer

'If range supplied, convert to string
If TypeName(aVariable) = "Range" Then
    aVariable = aVariable.Cells(1).Value
End If

aVariable = Trim(aVariable)

'Return immediately if already integer
If IsInt_NoTrailingSymbols(aVariable) Then
    ExtractFirstInt_LeftToRight = aVariable
    Exit Function
End If

If Len(aVariable) = 0 Then
    GoTo ExtractNothing
End If

'Remove any characters **leading up to an integer
aVariable = Truncate_Before_Int(aVariable)
'Pad {aVariable} with **ending character to set up loop
aVariable = aVariable & "A"
    
    'Remove integers one-by-one until character found
    For i = 1 To Len(aVariable)
        If IsInt_NoTrailingSymbols(Left(aVariable, i)) = False Then
            ExtractFirstInt_LeftToRight = Left(aVariable, i - 1)
            Exit Function
        End If
    Next i
    
ExtractNothing:
ExtractFirstInt_LeftToRight = vbNullString

End Function

Function Truncate_Before_Int(ByVal aString As String)
On Error GoTo NoInt:

Dim CountCharsToRemove As Integer, _
    CheckCharacter As String, _
    NewStrLength As Integer, _
    i As Integer
    
    aString = Trim(aString)
    
    'Return immediately if already integer
    If IsInt_NoTrailingSymbols(aString) Then
        Truncate_Before_Int = aString
        Exit Function
    End If
        
        CountCharsToRemove = 0
    
        'Loop to determine number of starting characters to remove
        For i = 1 To Len(aString)
        
            'Single character string at point i in {aString}, e.g. "S" or "o"
            CheckCharacter = Right(Left(aString, i), 1)
                
                If IsInt_NoTrailingSymbols(CheckCharacter) = False Then
                    CountCharsToRemove = CountCharsToRemove + 1
                ElseIf IsNumeric(CheckCharacter) = True Then
                    Exit For
                End If
        Next i
                    
                    NewStrLength = Len(aString) - CountCharsToRemove
                    Truncate_Before_Int = Right(aString, NewStrLength)
                    
                    Exit Function
        
NoInt:
Truncate_Before_Int = vbNullString

End Function

Function Truncate_After_Int(ByVal aString As String)
On Error GoTo NoInt:

Dim CountCharsToRemove As Integer, _
    CheckCharacter As String, _
    NewStrLength As Integer, _
    i As Integer
    
    aString = Trim(aString)
    
    'Return immediately if already integer
    If IsInt_NoTrailingSymbols(aString) Then
        Truncate_After_Int = aString
        Exit Function
    End If
        
        CountCharsToRemove = 0
    
        'Loop to determine number of starting characters to remove
        For i = 1 To Len(aString)
        
            'Single character string at point i in {aString}, e.g. "S" or "o"
            CheckCharacter = Left(Right(aString, i), 1)
                
                If IsNumeric(CheckCharacter) = False Then
                    CountCharsToRemove = CountCharsToRemove + 1
                ElseIf IsNumeric(CheckCharacter) = True Then
                    Exit For
                End If
        Next i
            
                    NewStrLength = Len(aString) - CountCharsToRemove
                    Truncate_After_Int = Left(aString, NewStrLength)
                    
                    Exit Function
        
NoInt:
Truncate_After_Int = vbNullString

End Function

Function IsInt_NoTrailingSymbols(ByVal aNumeric)
On Error GoTo NotInt

    'False if {aNumeric} padded with comma or period
    If Len(aNumeric * 1) = Len(aNumeric) Then
        IsInt_NoTrailingSymbols = True
        Exit Function
    End If
    
NotInt:
IsInt_NoTrailingSymbols = False
End Function

'>> WARNING - READ CODE <<
Function Delete_FileAndFolder(ByVal aFilePath As String)
'Make this a private function when it is moved to a new workbook

On Error GoTo NoDelete

Dim Slash As String, _
    ContainerFolder As String, _
    ThisUser As String, _
    i As Integer
    
ThisUser = Get_Username()
If MyOS = "Windows" Then Slash = "\" Else Slash = "/"
            
'Check to verify file path supplied, if not, 2 folders would be deleted so exit
If InStr(1, aFilePath, ".") = 0 Then GoTo NoDelete
        
        For i = Len(aFilePath) To 1 Step -1
            
            'Reduce {aFilePath} until it's Dir ending with {Slash}
            aFilePath = Left(aFilePath, Len(aFilePath) - 1)
            If Right(aFilePath, 1) = Slash Then
                ContainerFolder = aFilePath
                Exit For
            End If
            
        Next i
        
If Dir(ContainerFolder, vbDirectory) = "" Then GoTo NoDelete

If Right(ContainerFolder, Len(Slash & "Desktop" & Slash)) = (Slash & "Desktop" & Slash) Then
    Debug.Print "!!WARNING!! Path supplied to Delete_FileAndFolder() would delete all files in your Desktop folder"
    GoTo NoDelete
End If

If Right(ContainerFolder, Len(Slash & "Documents" & Slash)) = (Slash & "Documents" & Slash) Then
    Debug.Print "!!WARNING!! Path supplied to Delete_FileAndFolder() would delete all files in your Documents folder"
    GoTo NoDelete
End If

If Len(ContainerFolder) - Len(Replace(ContainerFolder, Slash, "")) <= 4 Then
    Debug.Print Len(ContainerFolder) - Len(Replace(ContainerFolder, "/", ""))
    Debug.Print "!!WARNING!! Path supplied to Delete_FileAndFolder() is a high level folder that could delete many files"
    GoTo NoDelete
End If
    
    Kill ContainerFolder & "*.*"
    RmDir ContainerFolder
    Debug.Print ContainerFolder & " and all files within it deleted."

        Delete_FileAndFolder = True
        Exit Function

NoDelete:
Delete_FileAndFolder = False
Exit Function
            
End Function

Function Get_DesktopPath()

If MyOS = "Windows" Then
    
    Dim ThisUserName As String, _
        ExpectedPath As String
        ThisUserName = Get_WindowsUsername()
        
        ExpectedPath = "C:\Users\" & ThisUserName & "\Desktop"
            
        If Dir(ExpectedPath, vbDirectory) = "" Then
            'OneDrive Windows OS
            Get_DesktopPath = "C:\Users\" & ThisUserName & "\OneDrive\Desktop"
            Exit Function
        End If
            
            Get_DesktopPath = ExpectedPath
        
ElseIf MyOS = "Mac" Then
    Get_DesktopPath = "/Users/" & Get_MacUsername
End If
    
End Function

Function Get_Username()

Dim ThisOS As String
    ThisOS = MyOS()

    If ThisOS = "Windows" Then
        Get_Username = Get_WindowsUsername
    ElseIf ThisOS = "Mac" Then
        Get_Username = Get_MacUsername
    Else
        'Windows & Mac only
        Get_Username = vbNullString
    End If

End Function

Function Get_WindowsUsername()

If MyOS() <> "Windows" Then
    Get_WindowsUsername = ""
    Exit Function
End If

Dim FSO As Object, _
    ParentFolder As Object, _
    ParentFolderPath As String, _
    ChildFolder As Object, _
    ChildFolderPath As String, _
    arrPaths() As Variant, _
    i As Integer
    
    'Folder on all Windows OS
    ParentFolderPath = "C:\Users\"
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ParentFolder = FSO.GetFolder(ParentFolderPath)
        
        For Each ChildFolder In ParentFolder.SubFolders
            
            'C:\Users\{ChildFolder}
            ChildFolderPath = ChildFolder.Path
            
            Dim ChildChildFolder As Object: Set ChildChildFolder = FSO.GetFolder(ChildFolderPath)
            Dim ChildChildFolderPath As String
            
            'To handle permission errors
            On Error Resume Next
            For Each ChildChildFolder In ChildFolder.SubFolders
                
                'Returns path with Username between "C:\Users\" & "\AppData"
                '& filters out "C:\Users\All Users\AppData" & "C:\Users\Default\AppData"
                If InStr(1, ChildChildFolder.Path, "AppData") <> 0 And _
                   InStr(1, ChildChildFolder.Path, "All Users") = 0 And _
                   InStr(1, ChildChildFolder.Path, "Default") = 0 Then
                    
                    'C:\Users\{ChildFolder}\{ChildChildFolder}
                    ReDim Preserve arrPaths(i): i = i + 1
                    arrPaths(UBound(arrPaths)) = ChildChildFolder.Path
                    
                    'Extract Username
                    arrPaths(UBound(arrPaths)) = Replace(arrPaths(UBound(arrPaths)), "C:\Users\", vbNullString)
                    arrPaths(UBound(arrPaths)) = Replace(arrPaths(UBound(arrPaths)), "\AppData", vbNullString)
                    
                End If
                
            Next ChildChildFolder
            
        Next ChildFolder
            
            Set FSO = Nothing
            Set ParentFolder = Nothing
            Set ChildFolder = Nothing
            Set ChildChildFolder = Nothing
                
                If UBound(arrPaths) <> 0 Then
                    Get_WindowsUsername = "Error Finding User"
                End If
                
                    Get_WindowsUsername = arrPaths(0)

End Function

Function Get_MacUsername()
'Will not work for cloud hosted files

If MyOS() <> "Mac" Then
    Get_MacUsername = ""
    Exit Function
End If

Dim TempString As String, _
    SlashLocation As Integer
    
    TempString = Replace(ActiveWorkbook.FullName, "/Users/", "")
       SlashLocation = InStr(1, TempString, "/")
           Get_MacUsername = Left(TempString, SlashLocation - 1)

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

Function Print_Pad()
    Debug.Print ("================== " & Format(Now(), "Long Time") & " ==================")
End Function

Function Print_Named(ByVal Something, Optional Label)
On Error GoTo SomethingIsNothing
    
    If IsMissing(Label) = True Then
        Debug.Print (">> " & Something)
    Else
        Debug.Print (Label & ":")
        Debug.Print (">> " & Something)
    End If
        Debug.Print ""
        Exit Function
        
SomethingIsNothing:
On Error GoTo -1
    Debug.Print "Error Printing Value"
    Debug.Print ""
End Function


