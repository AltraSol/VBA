Attribute VB_Name = "SimplifyVBA"
Option Explicit
'
'TODO: finish documenting (ctrl+f ooooooooooooooooooooooooooooooooooooooooo)
'
'===============================================================================================================================================================================================================================================================
'#  SimplifyVBA ¬ github.com/ulchc (10-17-22)
'===============================================================================================================================================================================================================================================================
'
'A collection of code to interface R with VBA, make application building easier, or improve VBA readability.
'
'Prefix: ƒ— denotes a function which has a notable load time or file interactions outside ThisWorkbook. Only use these within the VBA IDE.
'
'===============================================================================================================================================================================================================================================================
'##  Important
'===============================================================================================================================================================================================================================================================
'
'#### If you intend to use the User Interface section, the following sub must be placed within ThisWorkbook:
'
'----------------------------------------------------------------``` VBA
'   Private Sub Workbook_BeforeClose(Cancel As Boolean)
'       Call Remove_TempMenuCommands
'       Call Remove_TempMenuCommandSections
'   End Sub
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
Public GlobalUser As String
'
''   Prevents needless rerunning of the file search component of
''   Get_WindowsUsername() once the local user has been determined.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
Public GlobalTempMenuCommands() As Variant
Public GlobalTempMenuSections() As Variant
'
''    Tracks menu commands or menu sections that have been added using
''    the CreateMenuCommand() or CreateMenuSection() commands with a
''    Temporary:=True property. Allows for the deletion of all user
''    created menus or menu items on the Workbook_BeforeClose() event.
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  Functions
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
'  Get_Username()
'
''   Returns username regardless of Windows or Mac OS.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_DesktopPath()
'
''   Returns Mac or Windows desktop directory (even if on OneDrive).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_DownloadsPath()
'
''   Returns Mac or Windows downloads directory (even if on OneDrive).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_LatestFile( _
'     FromFolder As String, _
'     MatchingString As String, _
'     FileType As String _
' )
'
''   Returns the latest file of the specified {FileType} with a name
''   that includes {MatchingString} from the directory {FromFolder}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ListFiles(FromFolder As String)
'
''   Returns an array of all file paths located in {FromFolder}
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ListFolders(FromFolder As String)
'
''   Returns an array of all file paths located in {FromFolder}
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CopySheets_FromFolder( _
'     FromFolder As String, _
'     Optional Copy_xlsx As Boolean, _
'     Optional Copy_xlsm As Boolean, _
'     Optional Copy_xls As Boolean, _
'     Optional Copy_csv As Boolean _
' )
''   Opens all file types specified by the boolean parameters in the
''   directory {FromFolder}, copies all sheets to ThisWorkbook, then
''   returns an array of the new sheet names.
'
'    Dim CopiedSheets(): CopiedSheets() = CopySheets_FromFolder(...)
'    Sheets(CopiedSheets(1)).Activate
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' PasteSheetVals_FromFolder( _
'     FromFolder As String, _
'     Optional Copy_xlsx As Boolean, _
'     Optional Copy_xlsm As Boolean, _
'     Optional Copy_xls As Boolean, _
'     Optional Copy_csv As Boolean _
' )
'
''   Opens all file types specified by the boolean parameters in the
''   directory {FromFolder}, pastes cell values from each sheet to new
''   tabs in ThisWorkbook, then returns an array of the new sheet names.
'
'    Dim PastedSheets(): PastedSheets() = PasteSheetVals_FromFolder(...)
'    Sheets(PastedSheets(1)).Activate
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Clipboard_IsRange()
'
''   Returns True if a range is currently copied.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Clipboard_Load(ByVal aString As String)
'
''   Stores {aString} in clipboard.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ƒ—Clipboard_Read( _
'     Optional IfRngConcatAllVals As Boolean = True, _
'     Optional Sep As String = ", " _
' )
'
''   Returns text from the copied object (clipboard text or range).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  ƒ—Get_CopiedRangeVals()
'
''   If range copied, checks each Cell.Value in the range and
''   returns an array of each non-blank value.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CopySheets_FromFile(FromFile As String)
'
''   Opens {FromFile}, copies all sheets within it to ThisWorkbook,
''   then returns an array of the new sheet names.
'
'    Dim CopiedSheets(): CopiedSheets() = CopySheets_FromFile(...)
'    Sheets(CopiedSheets(1)).Activate
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' PasteSheetVals_FromFile(FromFile As String)
'
''   Opens {FromFile}, pastes cell values from all sheets within it
''   to ThisWorkbook, then returns an array of the new sheet names.
'
'    Dim PastedSheets(): PastedSheets() = PasteSheetVals_FromFile(...)
'    Sheets(PastedSheets(1)).Activate
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_FilesMatching( _
'     FromFolder As String, _
'     MatchingString As String, _
'     FileType As String _
' )
'
''   Returns an array of file paths located in {FromFolder} which have
''   a file name containing {MatchingString} and a specific {FileType}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' RenameSheet( _
'     CurrentName As String, _
'     NewName As String, _
'     OverrideExisting As Boolean _
' )
'
''   Changes Sheets({CurrentName}).Name to {NewName} if {NewName}
''   is not already in use, otherwise, a bracketed number (n) is added
''   to {NewName}. The final name of the renamed sheet is returned.
'
''   If {OverrideExisting} = True and a sheet with the name {NewName}
''   exists, it will be deleted and Sheets({CurrentName}).Name will
''   always be set to {NewName}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Tabs_MatchingCodeName( _
'     MatchCodeName As String, _
'     ExcludePerfectMatch As Boolean _
' )
'
''   An array of tab names where {MatchCodeName} is within the CodeName
''   property (useful for detecting copies of a code-named template).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' WorksheetExists( _
'     aName As String, _
'     Optional wb As Workbook _
' )
'
''   True or False dependent on if tab name {aName} already exists.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  ƒ—Delete_FileAndFolder(ByVal aFilePath As String) as Boolean
'
''   Use with caution. Deletes the file supplied {aFilePath}, all
''   files in the same folder, and the directory itself.
'
''   Will exit the deletion procedure if {aFilePath} is a file
''   within the Desktop or Documents directory, or if the directory
''   is considered high level (it's within the user directory).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_WindowsUsername()
'
''   Loops through folders to find paths matching C:\Users\...\AppData
''   then extracts the User from correct path. Superior to reading
''   .FullName of workbook which does not work for OneDrive.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_MacUsername()
'
''   Reads ActiveWorkbook.FullName property to get Mac user.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' PlatformFileSep()
'
''   Returns "\" or "/" depending on the operating system.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  MyOS()
'
''   "Windows",  "Mac", or "Neither Windows or Mac".
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'NOTE: Windows only (uses CreateObject("VBScript.RegExp"))
'
' Replace_SpecialChars( _
'     YourString As String, _
'     Replacement As String, _
'     Optional ReplaceAll As Boolean, _
'     Optional TrimWS As Boolean _
' )
'
''   Replaces `!@#$%^&“”*(")-=+{}\/?:;'.,<> from {YourString} with
''   {Replacement}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'NOTE: Windows only (uses CreateObject("VBScript.RegExp"))
'
' Function Replace_Any( _
'     Of_Str As String, _
'     With_Str As String, _
'     Within_Str As String, _
'     Optional TrimWS As Boolean _
' )
'
''   Replaces all characters {Of_Str} in the supplied {Within_Str}.
''   Distinct from VBA's Replace() in that all matched characters
''   are removed instead of perfect matches.
'
'    Debug.Print Replace_Any(" '. ", "_", "Here's an example.")
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  ExtractFirstInt_RightToLeft (aVariable)
'
''   Returns the first integer found in a string when searcing
''   from the right end of the string to the left.
'
'    ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  ExtractFirstInt_LeftToRight (aVariable)
'
''   Returns the first integer found in a string when searcing
''   from the left end of the string to the right.
'
'    ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Truncate_Before_Int (aString)
'
''   Removes characters before first integer in a sequence of characters.
'
'    Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Truncate_After_Int (aString)
'
''   Removes characters after first integer in a sequence of characters.
'
'    Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  IsInt_NoTrailingSymbols (aNumeric)
'
''   Checks if supplied value is both numeric, and contains no numeric
''   symbols (different from IsNumeric).
'
''   IsInt_NoTrailingSymbols(9999) = True
''   IsInt_NoTrailingSymbols(9999,) = False
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'## Subs
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
' ReDim_Add(ByRef aArr() As Variant, ByVal aVal)
'
''    Simplifies the addition of a value to a one dimensional array by
''    handling the initalization & resizing of an array in VBA
'
'     Call ReDim_Add(aArr(), aVal) '-> last element of aArr() now aVal
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ReDim_Rem(ByRef aArr() As Variant)
'
''    Simplifies the sequential removal of the last element of a one
''    dimensional array by handing the resizing of the array as well
''    as the removal of the 0th value
'
'     Call ReDim_Rem(aArr()) '-> last element of aArr() has been removed
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' SaveToDownloads( _
'    SaveTabNamed As String, _
'    AsFileNamed As String, _
'    OpenAfterSave As Boolean, _
'    Optional SaveAsType As String = "xlsx" _
' )
'
''    {SaveTabNamed} is the ActiveSheet.Name property, {AsFileNamed}
''    is a plain string which is automatically combined with the local
''    download folder to create the full path to save to.
'
''    {SaveAsType} can be "xlsx", "xlsm", "xlsb", or "csv". A bracketed
''    (n) will automatically be added to the file name if it is
''    already taken.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' SaveToDownloads_Multiple( _
'    SaveTabsNamed_Array As Variant, _
'    AsFileNamed As String, _
'    OpenAfterSave As Boolean, _
'    Optional SaveAsType As String = "xlsx" _
' )
'
''    Operates the same as SaveToDownloads() but takes an array of
''    tab names.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' MergeAndCombine(MergeRange As Range, Optional SepValsByNewLine = True)
'
''    Concatenates each Cell.Value in a range & merges range as opposed
''    to Merge & Center which only keeps a single value
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' AutoAdjustZoom(rngBegin As Range, rngEnd As Range)
'
''   Adjusts user view to the width of rngBegin to rngEnd
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' LaunchLink (aLink)
'
''   Launches aLink in existing browser with error handling for
''   invalid Links
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' InsertSlicer( _
'     NamedRange As String, _
'     NumCols As Integer, _
'     aHeight As Double, _
'     aWidth As Double _
' )
''   Creates a slicer for the active sheet named range {NamedRange}
''   with {NumCols} buttons per slicer row, and with dimensions
''   {aHeight} by {aWidth}
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' AlterSlicerColumns(SlicerName As String, NumCols)
'
''   Loops through workbook to find {SlicerName} and sets the number
''   of buttons per row to {NumCols}
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' MoveSlicer( _
'     SlicerSelection, _
'     rngPaste As Range, _
'     leftOffset, _
'     IncTop _
' )
''   Takes Selection as {SlicerSelection}, cuts & pastes it to a rough
''   location {rngPaste} to be incrementally adjusted from paste
''   location by {leftOffset} and {IncTop}
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  ToggleDisplayMode()
'
''   Toggles display of ribbon, formula bar, status bar & headings
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Print_Pad()
'
''   Uses Debug.Print to print a timestamped seperator of "======"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Print_Named(Something, Optional Label)
'
''   Uses Debug.Print to add a space between each {Something} printed,
''   labels each {Something} if {Label} supplied.
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  User Interface Additions
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
' ConvertStrCommand( _
'     CommandString As String, _
'     Optional Verbose As Boolean = True _
' )
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ChangeMenuVisibility( _
'     MenuItems_Array As Variant, _
'     VisibleProperty As Boolean _
' )
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ResetCellMenu
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CreateMenuCommand( _
'    MenuCommandName As String, _
'    StrCommand As String, _
'    Optional Temporary As Boolean = True, _
'    Optional MenuFaceID As Long _
' )
'PARAMETERS:
''    {PARAMETERS} =
''    {PARAMETERS} =
''    {PARAMETERS} =
'
'EXPLANATION:
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    Call RemoveMenuCommand(...) to remove
'
'EXAMPLES: '(Ctrl+f to view & run)
'     Sub Try_CreateMenuCommand
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CreateMenuSection( _
'    MenuSectionName As String, _
'    Array_SectionMenuNames As Variant, _
'    Array_StrCommands As Variant, _
'    Optional Temporary As Boolean = True _
' )
'PARAMETERS:
''    {PARAMETERS} =
''    {PARAMETERS} =
''    {PARAMETERS} =
'
'EXPLANATION:
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    Call RemoveMenuSection(...) to remove
'
'EXAMPLES: '(Ctrl+f to view & run)
'     Sub Try_CreateMenuSection
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'NOTE: Popup menus are Windows only
'
' CreatePopupMenu( _
'    PopupMenuName As String, _
'    Array_ItemNames As Variant, _
'    Array_StrCommands As Variant, _
'    Array_ItemFaceIDs As Variant, _
'    Optional Temporary As Boolean = True _
' )
'PARAMETERS:
''    {PARAMETERS} =
''    {PARAMETERS} =
''    {PARAMETERS} =
'
'EXPLANATION:
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    Call RemovePopupMenu(...) to remove
'
'EXAMPLES: '(Ctrl+f to view & run)
'     Sub Try_CreatePopupMenu
'     Sub Try_CreatePopupMenuColorful
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CreateAddInButtons( _
'    ButtonSectionName As String, _
'    ButtonNames_Array As Variant, _
'    ButtonTypes_Array As Variant, _
'    ButtonStrCommands_Array As Variant, _
'    Optional MenuFaceIDs_Array As Variant, _
'    Optional Temporary As Boolean = True _
' )
'
'PARAMETERS:
''    {ButtonSectionName} = Name of the row added to the Add-ins ribbon (visible on hover).
''    {ButtonNames_Array} = Array of names for each command (visible on hover).
''    {ButtonTypes_Array} = Array of types (1, 2 or 3) for the display of command buttons.
''    {ButtonStrCommands_Array} = Array of commands for each button (see ConvertStrCommand).
''    {MenuFaceIDs_Array} = Array of FaceId numbers (only applicable to ButtonTypes 1 and 3).
''    {Temporary} = Specifies whether the Add-ins section will automatically be removed when workbook closes.
'
'EXPLANATION:
''    Creates a row of commands within the "Custom Toolbars" section
''    of the Add-ins ribbon and Debug.Prints the details.
'
''    Adds each command in {ButtonStrCommands_Array}
''    to the section with properties as specified in {ButtonTypes_Array},
''    {MenuFaceIDs_Array} and {ButtonNames_Array}. Each {..._Array}
''    parameter must be of equal length, but the item of {MenuFaceIDs_Array}
''    will be ignored if the corresponding element of {ButtonTypes_Array} is
''    2 given that it's a caption only display type.
'
'     Call RemoveAddInSection(...) to remove
'
'EXAMPLES: '(Ctrl+f to view & run)
'     Sub Try_CreateAddInButtons_Type1
'     Sub Try_CreateAddInButtons_Type2
'     Sub Try_CreateAddInButtons_Type3
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CreateButtonShape( _
'    Optional StrCommand As String, _
'    Optional btnLabel As String = "Blank Button", _
'    Optional btnName As String, _
'    Optional ShapeType As Integer = 5, _
'    Optional btnColor As Long = 6299648, _
'    Optional Lef As Long = 10, _
'    Optional Top As Long = 10, _
'    Optional Wid As Long = 100, _
'    Optional Hei As Long = 20 _
' )
'PARAMETERS:
''    {PARAMETERS} =
''    {PARAMETERS} =
''    {PARAMETERS} =
'
'EXPLANATION:
''    ooooooooooooooooooooooooooooooooooooooooo
'
''    ooooooooooooooooooooooooooooooooooooooooo
'
'EXAMPLES: '(Ctrl+f to view & run)
'     Sub Try_CreateButtonShape
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  RScript
'===============================================================================================================================================================================================================================================================
'
'### TODO: Remove notification of deletion
'
'    All RScript functions are currently Windows OS only.
'
'----------------------------------------------------------------``` VBA
'  QuickRun_RScript(ByVal ScriptContents As String)
'
''   Writes a temporary .R script containing {ScriptContents}, runs
''   it, prompts for the deletion of the temporary script
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  WriteTemp_RScript(ByVal ScriptContents As String)
'
''   Creates a random named temporary folder on desktop, creates an
''   .R file "Temp.R" containing {ScriptContents}, returns Temp.R path
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  LocateRScript_Run(ByVal Script_Path)
'
''   Takes a string or cell reference {RScriptPath} & runs it on the
''   latest version of R on the OS
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Run_RScript( _
'     RScriptExe_Path As String, _
'     Script_Path As String, _
'     Optional Visibility As String, _
'     Optional OnErrorEnd As Boolean = True _
' )
''   Uses the RScript.exe pointed to by {RScriptExe_Path} to run the script
''   found at {Script_Path}. Rscript.exe window displayed by default,
''   but {Visibility}:= "VeryHidden" or "Minimized" can be used.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_RScriptExePath() As String
'
''   Returns the path to the latest version of Rscript.exe
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_LatestRVersion(ByVal RVersions As Variant)
'
''   Returns the latest version of R currently installed
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_RVersions(ByVal RFolderPath As String)
'
''   Returns an array of the R versions currently installed
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_RFolder() As String
'
''   Returns the parent R folder path which houses the installed
''   versions of R on the OS from which the sub is called
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Test_QuickRun_RScript()
'
''   Writes a computationally intensive script to Desktop and asks
''   if you want to run it (to visually verify all zRun_R f(x) worked)
'
'----------------------------------------------------------------```

Function Replace_Any( _
    Of_Str As String, _
    With_Str As String, _
    Within_Str As String, _
    Optional TrimWS As Boolean _
)

Dim Regex As Object: Set Regex = CreateObject("VBScript.RegExp") 'Windows Only

With Regex
    .Global = True
    .MultiLine = True
    .IgnoreCase = False
    .Pattern = "[" & Of_Str & "]"
End With

    Within_Str = Regex.Replace(Within_Str, With_Str)
    
    If TrimWS = True Then
        Within_Str = Application.WorksheetFunction.Trim(Within_Str)
    End If
    
    Replace_Any = Within_Str

End Function

Function Replace_SpecialChars( _
    YourString As String, _
    Replacement As String, _
    Optional ReplaceAll As Boolean, _
    Optional TrimWS As Boolean _
)

Dim Regex As Object: Set Regex = CreateObject("VBScript.RegExp") 'Windows Only

With Regex
    .Global = ReplaceAll
    .MultiLine = True
    .IgnoreCase = False
    .Pattern = "[" & "`!@#$%^&“”*(" & Chr(34) & ")-=+{}\/?:;'.,<>" & "]"
End With

    YourString = Regex.Replace(YourString, Replacement)
    
    If TrimWS = True Then
        YourString = Application.WorksheetFunction.Trim(YourString)
    End If
    
    Replace_SpecialChars = YourString

End Function

Function RenameSheet( _
    CurrentName As String, _
    NewName As String, _
    OverrideExisting As Boolean _
)

Dim wsToRename As Worksheet: Set wsToRename = ThisWorkbook.Sheets(CurrentName)
If wsToRename.Name = NewName Then GoTo ApplyName:

If WorksheetExists(NewName) = False Then
    GoTo ApplyName:
End If

If OverrideExisting = True Then
    Application.DisplayAlerts = False
        ThisWorkbook.Sheets(NewName).Delete
        GoTo ApplyName:
    Application.DisplayAlerts = True
Else
    Dim i As Integer, TryName As String: TryName = NewName
    Do While WorksheetExists(TryName) = True
        i = i + 1
        TryName = NewName & " (" & i & ")"
    Loop
    NewName = TryName
    GoTo ApplyName:
End If
    
ApplyName:
wsToRename.Name = NewName
RenameSheet = NewName
    
End Function

Function PasteSheetVals_FromFile(FromFile As String)

Dim ws As Worksheet, _
    wb As Workbook, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant, _
    ScreenUpdatingState As Boolean

ScreenUpdatingState = Application.ScreenUpdating
Application.ScreenUpdating = False

Application.StatusBar = "Opening " & Right(FromFile, Len(FromFile) - InStrRev(FromFile, PlatformFileSep())) & "..."
Workbooks.Open FileName:=FromFile, ReadOnly:=True
Set wb = ActiveWorkbook
    
    wb.Sheets(1).Select
    SheetCount = wb.Sheets.Count
    
    For Each ws In wb.Worksheets
        Sheet_i = Sheet_i + 1
        Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & wb.Name & "..."
         
        With ws
            .Activate
            .Cells.Select
        End With
        Selection.Copy
        With ThisWorkbook
            .Activate
            .Sheets.Add after:=ActiveSheet
        End With
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Call ReDim_Add(NewSheets(), ActiveSheet.Name)
    Next ws
    
Application.DisplayAlerts = False
wb.Close
Application.DisplayAlerts = True

PasteSheetVals_FromFile = NewSheets()
    
Application.StatusBar = False
Application.ScreenUpdating = ScreenUpdatingState
End Function

Function CopySheets_FromFile(FromFile As String)

Dim ws As Worksheet, _
    wb As Workbook, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant, _
    ScreenUpdatingState As Boolean

ScreenUpdatingState = Application.ScreenUpdating
Application.ScreenUpdating = False

Application.StatusBar = "Opening " & Right(FromFile, Len(FromFile) - InStrRev(FromFile, PlatformFileSep())) & "..."
Workbooks.Open FileName:=FromFile, ReadOnly:=True
Set wb = ActiveWorkbook

    wb.Sheets(1).Select
    SheetCount = wb.Sheets.Count
    
    For Each ws In wb.Worksheets
        Sheet_i = Sheet_i + 1
        Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & wb.Name & "..."
        ws.Copy after:=ThisWorkbook.ActiveSheet
        Call ReDim_Add(NewSheets(), ActiveSheet.Name)
    Next ws
    
wb.Close
    
CopySheets_FromFile = NewSheets()
    
Application.StatusBar = False
Application.ScreenUpdating = True
End Function

Function PasteSheetVals_FromFolder( _
    FromFolder As String, _
    Optional Copy_xlsx As Boolean, _
    Optional Copy_xlsm As Boolean, _
    Optional Copy_xls As Boolean, _
    Optional Copy_csv As Boolean _
)

Dim Sep As String, _
    FileTypes() As Variant, _
    FileType As Variant, _
    ScreenUpdatingState As Boolean

ScreenUpdatingState = Application.ScreenUpdating
Application.ScreenUpdating = False

'Ensures {FromFolder} ends with a PlatformFileSep()
Sep = PlatformFileSep()
FromFolder = Replace(FromFolder & Sep, Sep & Sep, Sep)

If Copy_xlsx = True Then Call ReDim_Add(FileTypes(), "xlsx")
If Copy_xlsm = True Then Call ReDim_Add(FileTypes(), "xlsm")
If Copy_xls = True Then Call ReDim_Add(FileTypes(), "xls")
If Copy_csv = True Then Call ReDim_Add(FileTypes(), "csv")

Dim DirPaths As String, _
    wbName As String, _
    wbFiles() As Variant
    
    For Each FileType In FileTypes()
        DirPaths = Dir(FromFolder & "*." & FileType)
        Do While DirPaths <> vbNullString
            wbName = FromFolder & DirPaths
                'Exclude partial matches (ex: {xls} matches .xls and .xls[x])
                If Right(wbName, Len(FileType)) = FileType Then
                    Call ReDim_Add(wbFiles(), wbName)
                End If
            DirPaths = Dir()
        Loop
    Next FileType
    
Dim ws As Worksheet, _
    wb As Workbook, _
    wbFile As Variant, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant

    For Each wbFile In wbFiles()
    
        Workbooks.Open FileName:=wbFile, ReadOnly:=True
        Set wb = ActiveWorkbook
        wb.Sheets(1).Select
        SheetCount = wb.Sheets.Count
            
            For Each ws In wb.Worksheets
                Sheet_i = Sheet_i + 1
                Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & wb.Name & "..."
                 
                With ws
                    .Activate
                    .Cells.Select
                End With
                Selection.Copy
                With ThisWorkbook
                    .Activate
                    .Sheets.Add after:=ActiveSheet
                End With
                Selection.PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                
                Call ReDim_Add(NewSheets(), ActiveSheet.Name)
            Next ws
    
        Application.DisplayAlerts = False
        wb.Close
        Application.DisplayAlerts = True
        
    Next wbFile
    
PasteSheetVals_FromFolder = NewSheets()
    
Application.StatusBar = False
Application.ScreenUpdating = ScreenUpdatingState
End Function

Function CopySheets_FromFolder( _
    FromFolder As String, _
    Optional Copy_xlsx As Boolean, _
    Optional Copy_xlsm As Boolean, _
    Optional Copy_xls As Boolean, _
    Optional Copy_csv As Boolean _
)

Dim Sep As String, _
    FileTypes() As Variant, _
    FileType As Variant, _
    ScreenUpdatingState As Boolean

ScreenUpdatingState = Application.ScreenUpdating
Application.ScreenUpdating = False

'Ensures {FromFolder} ends with a PlatformFileSep()
Sep = PlatformFileSep()
FromFolder = Replace(FromFolder & Sep, Sep & Sep, Sep)

If Copy_xlsx = True Then Call ReDim_Add(FileTypes(), "xlsx")
If Copy_xlsm = True Then Call ReDim_Add(FileTypes(), "xlsm")
If Copy_xls = True Then Call ReDim_Add(FileTypes(), "xls")
If Copy_csv = True Then Call ReDim_Add(FileTypes(), "csv")

Dim DirPaths As String, _
    wbName As String, _
    wbFiles() As Variant
    
    For Each FileType In FileTypes()
        DirPaths = Dir(FromFolder & "*." & FileType)
        Do While DirPaths <> vbNullString
            wbName = FromFolder & DirPaths
                'Exclude partial matches (ex: {xls} matches .xls and .xls[x])
                If Right(wbName, Len(FileType)) = FileType Then
                    Call ReDim_Add(wbFiles(), wbName)
                End If
            DirPaths = Dir()
        Loop
    Next FileType
    
Dim ws As Worksheet, _
    wb As Workbook, _
    wbFile As Variant, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant

    For Each wbFile In wbFiles()
        Workbooks.Open FileName:=wbFile, ReadOnly:=True
            
            Set wb = ActiveWorkbook
            wb.Sheets(1).Select
            SheetCount = wb.Sheets.Count
            
            For Each ws In wb.Worksheets
                Sheet_i = Sheet_i + 1
                Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & wb.Name & "..."
                ws.Copy after:=ThisWorkbook.ActiveSheet
                Call ReDim_Add(NewSheets(), ActiveSheet.Name)
            Next ws
        
        wb.Close
    Next wbFile
    
CopySheets_FromFolder = NewSheets()
    
Application.StatusBar = False
Application.ScreenUpdating = ScreenUpdatingState
End Function

Function Get_DownloadsPath()

If MyOS = "Windows" Then
    Get_DownloadsPath = "C:\Users\" & Get_WindowsUsername() & "\Downloads"
ElseIf MyOS = "Mac" Then
    Get_DownloadsPath = "/Users/" & Get_MacUsername & "/Downloads"
End If
    
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
    Get_DesktopPath = "/Users/" & Get_MacUsername & "/Desktop"
End If
    
End Function

Function Get_LatestFile( _
    FromFolder As String, _
    MatchingString As String, _
    FileType As String _
)

Dim MatchedFiles() As Variant: MatchedFiles() = Get_FilesMatching(FromFolder, MatchingString, FileType)
Dim MatchedFilesTime() As Variant, _
    KeepIndex As Integer, _
    i As Integer

        For i = LBound(MatchedFiles) To UBound(MatchedFiles)
            Call ReDim_Add(MatchedFilesTime(), FileDateTime(MatchedFiles(i)))
        Next i
        
            For i = LBound(MatchedFilesTime) To UBound(MatchedFilesTime)
                If i = LBound(MatchedFilesTime) Then
                    KeepIndex = LBound(MatchedFilesTime)
                Else
                    If MatchedFilesTime(i) > MatchedFilesTime(KeepIndex) Then
                        KeepIndex = i
                    End If
                End If
            Next i
            
                Get_LatestFile = MatchedFiles(KeepIndex)
End Function

Function Get_FilesMatching( _
    FromFolder As String, _
    MatchingString As String, _
    FileType As String _
)

Dim Sep As String, _
    LenFileType As Integer, _
    DirPaths As String, _
    NextPath As String, _
    ArrMatches() As Variant
    
    Sep = PlatformFileSep()
    'Ensures {FromFolder} ends with a PlatformFileSep()
    FromFolder = Replace(FromFolder & Sep, Sep & Sep, Sep)
    'Ensures {FileType} is the correct format
    FileType = Replace(FileType, ".", vbNullString)
    LenFileType = Len(FileType)
    DirPaths = Dir(FromFolder & "*." & FileType)
    
        Do While DirPaths <> vbNullString
            NextPath = FromFolder & DirPaths
                'Exclude partial matches (ex: {xls} matches .xls and .xls[x])
                If Right(NextPath, LenFileType + 1) = "." & FileType Then
                    If InStr(1, NextPath, MatchingString, vbTextCompare) <> 0 Then
                       Call ReDim_Add(ArrMatches(), NextPath)
                    End If
                End If
            DirPaths = Dir()
        Loop
        
    On Error GoTo NoMatches
    If LBound(ArrMatches()) = 0 Then
        Get_FilesMatching = ArrMatches()
        Exit Function
    End If
    
NoMatches:
Get_FilesMatching = vbNullString
    
End Function

Function ListFiles(FromFolder As String)

Dim Sep As String, _
    DirPaths As String, _
    NextPath As String, _
    ArrMatches() As Variant
    
    Sep = PlatformFileSep()
    'Ensures {FromFolder} ends with a PlatformFileSep()
    FromFolder = Replace(FromFolder & Sep, Sep & Sep, Sep)
    DirPaths = Dir(FromFolder)
    
        Do While DirPaths <> vbNullString
            NextPath = FromFolder & DirPaths
                Call ReDim_Add(ArrMatches(), NextPath)
            DirPaths = Dir()
        Loop
        
    On Error GoTo NoFiles
    If LBound(ArrMatches()) = 0 Then
        ListFiles = ArrMatches()
        Exit Function
    End If
    
NoFiles:
ListFiles = vbNullString
    
End Function

Function ListFolders(FromFolder As String)

Dim Sep As String, _
    DirPaths As String, _
    NextPath As String, _
    ArrMatches() As Variant
    
    Sep = PlatformFileSep()
    'Ensures {FromFolder} ends with a PlatformFileSep()
    FromFolder = Replace(FromFolder & Sep, Sep & Sep, Sep)
    DirPaths = Dir(FromFolder, vbDirectory)
    
        Do While DirPaths <> vbNullString
            NextPath = FromFolder & DirPaths
                Call ReDim_Add(ArrMatches(), NextPath)
            DirPaths = Dir()
        Loop
        
    On Error GoTo NoFiles
    If LBound(ArrMatches()) = 0 Then
        ListFolders = ArrMatches()
        Exit Function
    End If
    
NoFiles:
ListFolders = vbNullString
    
End Function

Function PlatformFileSep()
    If InStr(1, Application.OperatingSystem, "Windows") <> 0 Then
        PlatformFileSep = "\"
    Else
        PlatformFileSep = "/"
    End If
End Function

Function Clipboard_Load(ByVal aString As String) As Boolean

On Error GoTo NoLoad
    CreateObject("HTMLFile").ParentWindow.ClipboardData.SetData "text", aString
    Clipboard_Load = True
    Exit Function
    
NoLoad:
Clipboard_Load = False
On Error GoTo -1

End Function

Function ƒ—Clipboard_Read( _
    Optional IfRngConcatAllVals As Boolean = True, _
    Optional Sep As String = ", " _
)
On Error GoTo NoRead

If Clipboard_IsRange() = True Then
    Dim CopiedRangeText As Variant
        CopiedRangeText = ƒ—Get_CopiedRangeVals()
        
        If IfRngConcatAllVals = False Then
            ƒ—Clipboard_Read = CopiedRangeText(LBound(CopiedRangeText))
        Else
            ƒ—Clipboard_Read = Application.WorksheetFunction.TextJoin(Sep, True, CopiedRangeText)
        End If
        
Else
    ƒ—Clipboard_Read = CreateObject("HTMLFile").ParentWindow.ClipboardData.GetData("text")
End If

Exit Function

NoRead:
ƒ—Clipboard_Read = False
On Error GoTo -1
End Function

Function ƒ—Get_CopiedRangeVals()

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
                
                ƒ—Get_CopiedRangeVals = arrCellText()
PasteIssue:
                ActiveSheet.Delete
                       
If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
If Application.ScreenUpdating = False Then Application.ScreenUpdating = True

End Function

Function Clipboard_IsRange() As Boolean

Clipboard_IsRange = False
Dim aFormat As Variant

    For Each aFormat In Application.ClipboardFormats
        If aFormat = xlClipboardFormatCSV Then
            Clipboard_IsRange = True
        End If
    Next
    
End Function

Function Tabs_MatchingCodeName( _
    MatchCodeName As String, _
    ExcludePerfectMatch As Boolean _
)
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

Function WorksheetExists( _
    aName As String, _
    Optional wb As Workbook _
)
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

Function ƒ—Delete_FileAndFolder(ByVal aFilePath As String) As Boolean

On Error GoTo NoDelete

Dim Slash As String, _
    ContainerFolder As String, _
    ThisUser As String, _
    i As Integer
    
ThisUser = Get_Username()
Slash = PlatformFileSep()
            
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
    Debug.Print "!!WARNING!! Path supplied to ƒ—Delete_FileAndFolder() would delete all files in your Desktop folder"
    GoTo NoDelete
End If

If Right(ContainerFolder, Len(Slash & "Documents" & Slash)) = (Slash & "Documents" & Slash) Then
    Debug.Print "!!WARNING!! Path supplied to ƒ—Delete_FileAndFolder() would delete all files in your Documents folder"
    GoTo NoDelete
End If

If Len(ContainerFolder) - Len(Replace(ContainerFolder, Slash, "")) <= 4 Then
    Debug.Print Len(ContainerFolder) - Len(Replace(ContainerFolder, "/", ""))
    Debug.Print "!!WARNING!! Path supplied to ƒ—Delete_FileAndFolder() is a high level folder that could delete many files"
    GoTo NoDelete
End If
    
    Kill ContainerFolder & "*.*"
    RmDir ContainerFolder
    Debug.Print ContainerFolder & " and all files within it deleted."

        ƒ—Delete_FileAndFolder = True
        Exit Function

NoDelete:
ƒ—Delete_FileAndFolder = False
            
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

If GlobalUser <> vbNullString Then
    Get_WindowsUsername = GlobalUser
    Exit Function
End If

Dim UserDirectory As String, _
    UserFolders As Variant, _
    FilterTerms As Variant, _
    Term As Variant, _
    FilteredFolders As Variant

    UserDirectory = "C:\Users\"
    UserFolders = ListFolders(UserDirectory)
        
        'Non-username folders contained within the UserDirectory
        FilterTerms = Array(".", "Public", "AppData", "Default", "All Users")
        FilteredFolders = UserFolders 'Initalize
        
        For Each Term In FilterTerms
            FilteredFolders = Filter(FilteredFolders, CStr(Term), False)
        Next Term
        
        GlobalUser = Replace(FilteredFolders(0), "C:\Users\", vbNullString)
        Get_WindowsUsername = GlobalUser
        
End Function

Function Get_WindowsUsernameOld()

If GlobalUser <> vbNullString Then
    Get_WindowsUsername = GlobalUser
    Exit Function
End If

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
                
                    GlobalUser = arrPaths(0)
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

'===================================================================
'## SUBS
'===================================================================

Sub SaveToDownloads_Multiple( _
    SaveTabsNamed_Array As Variant, _
    AsFileNamed As String, _
    OpenAfterSave As Boolean, _
    Optional ByVal SaveAsType As String = "xlsx" _
)

Dim IdealFileName As String, _
    TryFileName As String, _
    VisibleProp, _
    wbNew As Workbook, _
    i As Integer, j As Integer

IdealFileName = Get_DownloadsPath() & PlatformFileSep() & AsFileNamed & "." & SaveAsType
TryFileName = IdealFileName

Do While Dir(TryFileName, vbDirectory) <> ""
    i = i + 1
    TryFileName = Replace(IdealFileName, ".", " (" & i & ").")
Loop

    For j = LBound(SaveTabsNamed_Array) To UBound(SaveTabsNamed_Array)
        
        ThisWorkbook.Activate: Application.StatusBar = _
            "Copying sheet " & j + 1 & " of " & UBound(SaveTabsNamed_Array) + 1 & _
            " (" & SaveTabsNamed_Array(j) & ") to " & TryFileName & "..."
            
        VisibleProp = ThisWorkbook.Sheets(SaveTabsNamed_Array(j)).Visible
            
        If j = 0 Then
            With ThisWorkbook.Sheets(SaveTabsNamed_Array(j))
                .Visible = xlSheetVisible
                .Copy
                .Visible = VisibleProp
            End With
            Set wbNew = ActiveWorkbook
        Else
            With ThisWorkbook.Sheets(SaveTabsNamed_Array(j))
                .Visible = xlSheetVisible
                .Copy before:=wbNew.Sheets(1)
                .Visible = VisibleProp
            End With
        End If
    
    Next j

ThisWorkbook.Activate: Application.StatusBar = "Saving sheet to the following location: " & TryFileName & "..."
    Select Case SaveAsType
        Case "xlsx"
            SaveAsType = xlOpenXMLWorkbook
        Case "xlsm"
            SaveAsType = xlOpenXMLWorkbookMacroEnabled
        Case "xlsb"
            SaveAsType = xlExcel12
        Case "csv"
            SaveAsType = xlCSV
    End Select
    With wbNew
        .Activate
        .SaveAs _
            FileName:=TryFileName, _
            FileFormat:=SaveAsType, _
            CreateBackup:=False
    End With
ActiveWindow.Close

Application.StatusBar = False
If OpenAfterSave = True Then
    Workbooks.Open FileName:=TryFileName
End If

End Sub

Sub SaveToDownloads( _
    SaveTabNamed As String, _
    AsFileNamed As String, _
    OpenAfterSave As Boolean, _
    Optional SaveAsType As String = "xlsx" _
)

Dim IdealFileName As String, _
    TryFileName As String, _
    VisibleProp, _
    wbNew As Workbook, _
    i As Integer
    
    Application.StatusBar = "Copying sheet " & Chr(34) & SaveTabNamed & Chr(34) & " to downloads folder..."
    
    VisibleProp = ThisWorkbook.Sheets(SaveTabNamed).Visible
    
    With ThisWorkbook.Sheets(SaveTabNamed)
        .Visible = xlSheetVisible
        .Copy
        .Visible = VisibleProp
    End With
    
    Set wbNew = ActiveWorkbook
    
    IdealFileName = Get_DownloadsPath() & PlatformFileSep() & AsFileNamed & "." & SaveAsType
    TryFileName = IdealFileName
    
    Do While Dir(TryFileName, vbDirectory) <> ""
        i = i + 1
        TryFileName = Replace(IdealFileName, ".", " (" & i & ").")
    Loop

Application.StatusBar = "Saving sheet to the following location: " & TryFileName & "..."
    Select Case SaveAsType
        Case "xlsx"
            SaveAsType = xlOpenXMLWorkbook
        Case "xlsm"
            SaveAsType = xlOpenXMLWorkbookMacroEnabled
        Case "xlsb"
            SaveAsType = xlExcel12
        Case "csv"
            SaveAsType = xlCSV
    End Select
    With wbNew
        .Activate
        .SaveAs _
            FileName:=TryFileName, _
            FileFormat:=SaveAsType, _
            CreateBackup:=False
    End With
ActiveWindow.Close

    Application.StatusBar = False
    If OpenAfterSave = True Then
        Workbooks.Open FileName:=TryFileName
    End If

End Sub

Sub ToggleDisplayMode()

    If Application.DisplayStatusBar = True Then
        Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",False)"
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        ActiveWindow.DisplayHeadings = False
    Else
        Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",True)"
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        ActiveWindow.DisplayHeadings = True
    End If
    
End Sub

Sub MergeAndCombine(MergeRange As Range, Optional SepValsByNewLine = True)

On Error Resume Next 'Unmerge if entire range already merged
If MergeRange.MergeCells = True Then MergeRange.MergeCells = False: On Error GoTo -1: Exit Sub
On Error GoTo -1

Dim i As Integer, _
    CombinedString As String, _
    Seperator As String
    
    Seperator = " "
    If SepValsByNewLine = True Then Seperator = " " & vbNewLine
    
    CombinedString = Trim(MergeRange.Cells(1).Value)
    If MergeRange.Count = 1 Then GoTo SkipConcat
    
        For i = 2 To MergeRange.Count
            If MergeRange.Cells(i).Value <> vbNullString Then
                CombinedString = CombinedString & Seperator & Trim(MergeRange.Cells(i).Value)
            End If
        Next i
        
SkipConcat:

            MergeRange.Cells(1).Value = CombinedString
            
                Application.DisplayAlerts = False
                    MergeRange.Merge
                Application.DisplayAlerts = True
                
            MergeRange.WrapText = True
            MergeRange.VerticalAlignment = xlTop
            
End Sub

Sub ReDim_Add(ByRef aArr() As Variant, ByVal aVal)

On Error GoTo Initalize:
Dim Dummy: Dummy = UBound(aArr())

    ReDim Preserve aArr(UBound(aArr()) + 1)
    aArr(UBound(aArr)) = aVal
    Exit Sub

Initalize:
On Error GoTo -1

ReDim Preserve aArr(0)
aArr(UBound(aArr)) = aVal

End Sub

Sub ReDim_Rem(ByRef aArr() As Variant)

On Error GoTo ZerothElement:
Dim Dummy: Dummy = UBound(aArr())

    ReDim Preserve aArr(UBound(aArr()) - 1)
    Exit Sub

ZerothElement:
On Error GoTo -1
Erase aArr()

End Sub

Sub LaunchLink(aLink)
On Error GoTo InvalidLink

ActiveWorkbook.FollowHyperlink Address:=aLink
Exit Sub

InvalidLink:
MsgBox "Unable to launch link in browser.", vbInformation, "Invalid Link?"

End Sub

Sub AutoAdjustZoom(rngBegin As Range, rngEnd As Range)
On Error Resume Next

Dim rngPrevious As Range
    
    'Only AutoAdjustZoom when maximized window
    If Application.WindowState <> xlMaximized Then
        Exit Sub
    End If
        
        'Zoom window into defined view range
        Set rngPrevious = Selection
            Range(rngBegin, rngEnd).Select
            ActiveWindow.Zoom = True
    
                'Return selection to original selection
                rngPrevious.Select

End Sub

Sub InsertSlicer( _
    NamedRange As String, _
    NumCols As Integer, _
    aHeight As Double, _
    aWidth As Double _
)
On Error Resume Next
DoEvents

Dim tblName As String, colName As String, SlicerName As String

    tblName = ActiveSheet.Range(NamedRange).Cells(1).ListObject 'The active sheet table name
    colName = ActiveSheet.Range(NamedRange).Cells(1).Offset(-1, 0).Value 'The column to filter
    SlicerName = tblName & "_" & colName ' & Format(Right(Now() * 100, 4) + 100, "0000") 'A semi-random name for the Slicer object
        
        'Add a Slicer titled {colName} that filters {colName} on {tblName}, then name the obj {SlicerName}
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects(tblName), colName) _
            .Slicers.Add ActiveSheet, , SlicerName, colName, 204.75, 559.874960629921, 144, 193.75
                DoEvents
                
                'Refer to the Slicer by the {SlicerName}, and set it's height & width
                ActiveSheet.Shapes(SlicerName).Height = aHeight
                ActiveSheet.Shapes(SlicerName).Width = aWidth
                DoEvents
                
                    'Use f(x) AlterSlicerColumns to change {SlicerName}'s number of cols to {NumCols}
                    Call AlterSlicerColumns(SlicerName, NumCols)
End Sub

Sub AlterSlicerColumns(SlicerName As String, NumCols)
On Error Resume Next

Dim i As Integer
    
    'Loop through each Slicer within workbook
    For i = 1 To ActiveWorkbook.SlicerCaches.Count
        'Neccesarily will error for all but one loop, when the correct Slicer
        'called {SlicerName} is found. using Slicers(1) or Slicers(j) does not
        'work consistently
         ActiveWorkbook.SlicerCaches(i).Slicers(SlicerName).NumberOfColumns = NumCols
    Next i
    
    DoEvents

End Sub

Sub MoveSlicer( _
    SlicerSelection, _
    rngPaste As Range, _
    leftOffset, _
    IncTop _
)
On Error Resume Next
    
    DoEvents
    SlicerSelection.Cut 'Cut the slicer current selected, which is {SlicerSelection}
    rngPaste.Select 'Select the range with which we're aligning {SlicerSelection}'s top and left positions with
    ActiveSheet.Paste 'Paste the slicer onto cell {rngPaste}
    DoEvents
    
        'After pasting, {SlicerSelection} is once again the selected object
        
        'Move {SlicerSelection} to the RIGHT of {rngPaste} by {leftOffset}
        ActiveSheet.Shapes(Selection.Name).IncrementLeft leftOffset
        DoEvents
        
        'Move {SlicerSelection} upwards by {IncTop}
        ActiveSheet.Shapes(Selection.Name).IncrementTop IncTop
        DoEvents
        
End Sub

Sub Print_Pad()
    Debug.Print ("================== " & Format(Now(), "Long Time") & " ==================")
End Sub

Sub Print_Named(ByVal Something, Optional Label)
On Error GoTo SomethingIsNothing
    
    If IsMissing(Label) = True Then
        Debug.Print (">> " & Something)
    Else
        Debug.Print (Label & ":")
        Debug.Print (">> " & Something)
    End If
        Debug.Print ""
        Exit Sub
        
SomethingIsNothing:
On Error GoTo -1
    Debug.Print "Error Printing Value"
    Debug.Print ""
End Sub

'===============================================================================================================================================================================================================================================================
'# USER INTERFACE
'===============================================================================================================================================================================================================================================================

'NOTES:

'FaceIds: https://bettersolutions.com/vba/ribbon/face-ids-2003.htm
'FaceId = 1 is blank

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## MISC / FOR EXAMPLES
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function ConvertStrCommand(CommandString As String, Optional Verbose As Boolean = True)
If CommandString = vbNullString Then Exit Function
If Verbose = True Then Call Print_Named(CommandString, "Original CommandString")
    'Replace curly parenthesis with spaces...
    CommandString = Replace(CommandString, "}", " ")
    CommandString = Replace(CommandString, "{", " ")
    '...substitute apostrophes with quotation marks...
    CommandString = Replace(CommandString, "'", Chr(34))
    '...trim white space
    CommandString = Application.WorksheetFunction.Trim(CommandString)
    '...encase with apostrophes and return
    CommandString = "'" & CommandString & "'"
    ConvertStrCommand = CommandString
If Verbose = True Then Call Print_Named(CommandString, "Converted CommandString")
End Function

Sub ExampleSub()
    MsgBox "This a message shown by calling 'ExampleSub'", vbInformation, "ExampleSub"
End Sub

Sub WriteLines(MyText As String, Optional Repeat As Integer = 1)
    MyText = Application.WorksheetFunction.Rept(MyText & vbNewLine, Repeat)
    MsgBox MyText 'Simple sub to call (with parameters)
End Sub

Function ChangeMenuVisibility(MenuItems_Array As Variant, VisibleProperty As Boolean)
Dim MenuItem As Variant
    For Each MenuItem In MenuItems_Array
        If MenuCommandExists(CStr(MenuItem)) Then
            CommandBars("Cell").Controls(CStr(MenuItem)).Visible = VisibleProperty
        ElseIf MenuSectionExists(CStr(MenuItem)) Then
            Application.ShortcutMenus(xlWorksheetCell).MenuItems(CStr(MenuItem)).Visible = VisibleProperty
        End If
    Next MenuItem
End Function

Function ResetCellMenu()
    CommandBars("Cell").Reset
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## CREATE MENU COMMAND
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Try_CreateMenuCommand()

'Open the immediate window to see a print out of the events

'The name that can be used in RemoveMenuCommand() to delete the Menu Command
Dim MyCommandName As String: MyCommandName = "Merge and Combine"

Call CreateMenuCommand( _
    MenuCommandName:=MyCommandName, _
    StrCommand:="MergeAndCombine{Selection}", _
    MenuFaceID:=402, _
    Temporary:=True _
)

Exit Sub 'Comment this and replay the sub to remove the example command
Call RemoveMenuCommand(MyCommandName)
End Sub

Sub CreateMenuCommand( _
    MenuCommandName As String, _
    StrCommand As String, _
    Optional Temporary As Boolean = True, _
    Optional MenuFaceID As Long _
)

'Overwrite the command with new parameters if the user has created a version of it before
Call RemoveMenuCommand(MenuCommandName)

Dim MenuObject As Object
Set MenuObject = CommandBars("Cell").Controls.Add(before:=1)
    
    StrCommand = ConvertStrCommand(StrCommand, Verbose:=False)

    With MenuObject
       .Caption = MenuCommandName
       .OnAction = StrCommand
       .FaceId = MenuFaceID
       .BeginGroup = True
    End With
    
    Debug.Print "Menu: [" & MenuCommandName & "]"
    Debug.Print "Runs: " & StrCommand & vbNewLine
    
        If Temporary = True Then
        
            'Add the {MenuCommandName} to the Public variable GlobalTempMenuCommands()...
            Call ReDim_Add(GlobalTempMenuCommands(), MenuCommandName)
            
            '...filter array to check if more than one instance of {MenuCommandName} is present...
            If UBound(Filter(GlobalTempMenuCommands(), MenuCommandName)) > 0 Then
                
                '...delete the newly added element if it already exists
                Call ReDim_Rem(GlobalTempMenuCommands())

            End If
            
        End If
        
        Set MenuObject = Nothing

End Sub

Sub Remove_TempMenuCommands()
On Error GoTo NoTempMenus
Dim i As Integer
    For i = UBound(GlobalTempMenuCommands()) To LBound(GlobalTempMenuCommands()) Step -1
        'Remove the menu...
        Call RemoveMenuCommand(CStr(GlobalTempMenuCommands(i)))
        '...and the last element from GlobalTempMenuCommands()
        Call ReDim_Rem(GlobalTempMenuCommands())
    Next i
NoTempMenus:
On Error GoTo -1
End Sub

Sub RemoveMenuCommand(MenuCommandName As String)
    If MenuCommandExists(MenuCommandName) Then CommandBars("Cell").Controls(MenuCommandName).Delete
End Sub

Function MenuCommandExists(MenuCommandName As String)
Dim i As Integer

    For i = 1 To CommandBars("Cell").Controls.Count
        If MenuCommandName = CommandBars("Cell").Controls(i).Caption Then
            MenuCommandExists = True
            Exit Function
        End If
    Next i
    
MenuCommandExists = False
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## CREATE MENU COMMAND SECTION
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Try_CreateMenuSection()

'Open the immediate window to see a print out of the events

'The name that can be used in RemoveMenuSection() to delete the Menu Section
Dim MySectionName As String: MySectionName = "Custom Section"

Call CreateMenuSection( _
    MenuSectionName:=MySectionName, _
    Array_SectionMenuNames:=Array("Merge and Combine Selection", "Toggle Display Mode"), _
    Array_StrCommands:=Array("MergeAndCombine{Selection}", "ToggleDisplayMode"), _
    Temporary:=True _
)

Exit Sub 'Comment this and replay the sub to remove the example section
Call RemoveMenuSection(MySectionName)
End Sub

Sub CreateMenuSection( _
    MenuSectionName As String, _
    Array_SectionMenuNames As Variant, _
    Array_StrCommands As Variant, _
    Optional Temporary As Boolean = True _
)

'Overwrite the section with new parameters if the user has created a version of it before
Call RemoveMenuSection(MenuSectionName)

Dim MenuObject As Object
Set MenuObject = Application.ShortcutMenus(xlWorksheetCell).MenuItems.AddMenu( _
    Caption:=MenuSectionName, _
    before:=1 _
)

Debug.Print "=========================================="
Debug.Print "Menu Section Added: [" & MenuSectionName & "]"
Debug.Print "==========================================" & vbNewLine
    
    Dim i As Integer
    For i = LBound(Array_StrCommands) To UBound(Array_StrCommands)
        
        'Convert StrCommands to runable command...
        Array_StrCommands(i) = ConvertStrCommand(CStr(Array_StrCommands(i)), Verbose:=False)
        
        '...then add each sub menu name and command to main menu
        With MenuObject.MenuItems.Add( _
            Caption:=Array_SectionMenuNames(i), _
            OnAction:=Array_StrCommands(i))
        End With
        
        Debug.Print "Sub Menu: [" & Array_SectionMenuNames(i) & "]"
        Debug.Print "    Runs: " & Array_StrCommands(i) & vbNewLine
        
    Next i

        If Temporary = True Then

            'Add the {MenuSectionName} to the Public variable GlobalTempMenuSections()...
            Call ReDim_Add(GlobalTempMenuSections(), MenuSectionName)

            '...filter array to check if more than one instance of {MenuName} is present...
            If UBound(Filter(GlobalTempMenuSections(), MenuSectionName)) > 0 Then

                '...delete the newly added element if it already exists
                Call ReDim_Rem(GlobalTempMenuSections())

            End If

        End If

            Set MenuObject = Nothing

End Sub

Sub Remove_TempMenuCommandSections()
On Error GoTo NoTempMenus
Dim i As Integer
    For i = UBound(GlobalTempMenuSections()) To LBound(GlobalTempMenuSections()) Step -1
        'Remove the menu...
        Call RemoveMenuSection(CStr(GlobalTempMenuSections(i)))
        '...and the last element from GlobalTempMenus()
        Call ReDim_Rem(GlobalTempMenuSections())
    Next i
NoTempMenus:
On Error GoTo -1
End Sub

Sub RemoveMenuSection(MenuSectionName As String)
    If MenuSectionExists(MenuSectionName) Then Application.ShortcutMenus(xlWorksheetCell).MenuItems(MenuSectionName).Delete
End Sub

Function MenuSectionExists(MenuSectionName As String)
Dim i As Integer

    For i = 1 To Application.ShortcutMenus(xlWorksheetCell).MenuItems.Count
        If MenuSectionName = Application.ShortcutMenus(xlWorksheetCell).MenuItems(i).Caption Then
            MenuSectionExists = True
            Exit Function
        End If
    Next i
    
MenuSectionExists = False
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## CREATE POPUP MENU
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Try_CreatePopupMenu()

'Open the immediate window to see a print out of the events

'The name that can be used in RemovePopupMenu() to delete the Menu Section
Dim MyMenuName As String: MyMenuName = "Popup Menu"

'The menu specified will be generated when called
Call CreatePopupMenu( _
    PopupMenuName:=MyMenuName, _
    Array_ItemNames:=Array("Toggle Display Mode", "Change Theme", "Print Sheet"), _
    Array_StrCommands:=Array("ExampleSub", "ExampleSub", "ExampleSub"), _
    Array_ItemFaceIDs:=Array(9378, 508, 4), _
    Temporary:=True _
)

'The popup will be shown any time this command is used
Application.CommandBars(MyMenuName).ShowPopup

'Place the .ShowPopup (or both the create and .ShowPopup) on a Worksheet event to show the menu:
'Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'    Cancel = True
'    Call Try_CreatePopupMenu
'End Sub

Exit Sub 'Comment this and replay the sub to remove the example menu
Call RemovePopupMenu(MyMenuName)
End Sub

Sub Try_CreatePopupMenuColorful()

'Open the immediate window to see a print out of the events

'The name that can be used in RemovePopupMenu() to delete the Menu Section
Dim MyMenuName As String: MyMenuName = "Popup Menu"

'The menu specified will be generated when called
Call CreatePopupMenu( _
    PopupMenuName:=MyMenuName, _
    Array_ItemNames:=Array( _
       "Dicrete Colour Wheel - 417", "Paint Brush - 108", "Continuous Colour Wheel - 7166", _
       "Paint Bucket & Brush - 3061", "Multi-Coloured Bars - 5873", "Multi-Coloured Bars / Shades - 6714", _
       "Coloured Cells - 6862", "Multi-Coloured Butterfly - 9678", _
       "Eraser - 7884", "Eraser With Cell - 2901", "Blank Cell - 410" _
    ), _
    Array_StrCommands:=Array( _
       "ExampleSub", "ExampleSub", "ExampleSub", "ExampleSub", "ExampleSub", "ExampleSub", _
       "ExampleSub", "ExampleSub", "ExampleSub", "ExampleSub", "ExampleSub" _
    ), _
    Array_ItemFaceIDs:=Array( _
        417, 108, 7166, 3061, 5873, 6714, 6862, 9678, 7884, 2901, 410 _
    ), _
    Temporary:=True _
)

'The popup will be shown any time this command is used
Application.CommandBars(MyMenuName).ShowPopup

'Place the .ShowPopup (or both the create and .ShowPopup) on a Worksheet event to show the menu:
'Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'    Cancel = True
'    Call Try_CreatePopupMenu
'End Sub

Exit Sub 'Comment this and replay the sub to remove the example menu
Call RemovePopupMenu(MyMenuName)
End Sub

Sub CreatePopupMenu( _
    PopupMenuName As String, _
    Array_ItemNames As Variant, _
    Array_StrCommands As Variant, _
    Array_ItemFaceIDs As Variant, _
    Optional Temporary As Boolean = True _
)

Call RemovePopupMenu(PopupMenuName)

Dim PopupMenu As CommandBar, _
    PopupMenuItem As CommandBarControl
    
Set PopupMenu = Application.CommandBars.Add( _
    Name:=PopupMenuName, _
    Position:=5, _
    MenuBar:=False, _
    Temporary:=Temporary _
)

Debug.Print "================================================"
Debug.Print "Application.CommandBars(" & Chr(34) & PopupMenuName & Chr(34) & ") added"
Debug.Print "================================================" & vbNewLine
    
    Dim i As Integer
    For i = LBound(Array_StrCommands) To UBound(Array_StrCommands)
    
        Set PopupMenuItem = PopupMenu.Controls.Add
    
            'Convert StrCommands to runable command...
            Array_StrCommands(i) = ConvertStrCommand(CStr(Array_StrCommands(i)), Verbose:=False)
            
            '...then add each sub menu name and command to main menu
            With PopupMenuItem
                .Caption = Array_ItemNames(i)
                .OnAction = Array_StrCommands(i)
                .FaceId = Array_ItemFaceIDs(i)
            End With
            
            Debug.Print "Popup Item: [" & Array_ItemNames(i) & "]"
            Debug.Print "      Runs: " & Array_StrCommands(i) & vbNewLine
            
        Set PopupMenuItem = Nothing
        
    Next i
    
Set PopupMenu = Nothing

End Sub

Sub RemovePopupMenu(MenuName As String)
    If PopupMenuExists(MenuName) Then Application.CommandBars(MenuName).Delete
End Sub

Function PopupMenuExists(MenuName As String)
Dim i As Integer

    For i = 1 To Application.CommandBars.Count
        If MenuName = Application.CommandBars(i).Name Then
            PopupMenuExists = True
            Exit Function
        End If
    Next i
    
PopupMenuExists = False
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## CREATE ADD-INS RIBBON BUTTONS
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Try_CreateAddInButtons_Type1()

'The name that can be used in RemoveAddInSection() to delete the Add-In row
Dim ButtonSectionName As String: ButtonSectionName = "Icons"

Call CreateAddInButtons( _
    ButtonSectionName:=ButtonSectionName, _
    ButtonNames_Array:=Array("Icon 1", "Icon  2", "Icon 3"), _
    ButtonTypes_Array:=Array(1, 1, 1), _
    ButtonStrCommands_Array:=Array("ExampleSub", "ExampleSub", "ExampleSub"), _
    MenuFaceIDs_Array:=Array(483, 482, 484), _
    Temporary:=True _
)

Exit Sub 'Comment this to delete the section added in the example
Call RemoveAddInSection(ButtonSectionName)
End Sub

Sub Try_CreateAddInButtons_Type2()

'The name that can be used in RemoveAddInSection() to delete the Add-In row
Dim ButtonSectionName As String: ButtonSectionName = "Captions"

Call CreateAddInButtons( _
    ButtonSectionName:=ButtonSectionName, _
    ButtonNames_Array:=Array("Plain Text Button"), _
    ButtonTypes_Array:=Array(2), _
    ButtonStrCommands_Array:=Array("ExampleSub"), _
    Temporary:=True _
)

Exit Sub 'Comment this to delete the section added in the example
Call RemoveAddInSection(ButtonSectionName)
End Sub

Sub Try_CreateAddInButtons_Type3()

'Open the immediate window to see a print out of the events

'The name that can be used in RemoveAddInSection() to delete the Add-In row
Dim ButtonSectionName As String: ButtonSectionName = "Caption Icons"

Call CreateAddInButtons( _
    ButtonSectionName:=ButtonSectionName, _
    ButtonNames_Array:=Array("TextIcon 1", "TextIcon 2"), _
    ButtonTypes_Array:=Array(3, 3), _
    ButtonStrCommands_Array:=Array("ExampleSub", "ExampleSub"), _
    MenuFaceIDs_Array:=Array(356, 487), _
    Temporary:=True _
)

Exit Sub 'Comment this to delete the section added in the example
Call RemoveAddInSection(ButtonSectionName)
End Sub

Sub CreateAddInButtons( _
    ButtonSectionName As String, _
    ButtonNames_Array As Variant, _
    ButtonTypes_Array As Variant, _
    ButtonStrCommands_Array As Variant, _
    Optional MenuFaceIDs_Array As Variant, _
    Optional Temporary As Boolean = True _
)

'Overwrite the section with new parameters if the user has created a version of it before
Call RemoveAddInSection(ButtonSectionName)

'Create the Add-In section called {ButtonSectionName}
Dim CustomToolbarsRow As CommandBar
Set CustomToolbarsRow = Application.CommandBars.Add(Temporary:=Temporary)
    
    With CustomToolbarsRow
        .Visible = True
        .Name = ButtonSectionName
    End With
        
        'Begin printing the events
        Debug.Print "==========================================================================================="
        Debug.Print "Added Row: " & Chr(34) & ButtonSectionName & Chr(34) & " to " & Chr(34) & "Custom Toolbars" & Chr(34) & " Section of the Add-ins Ribbon"
        Debug.Print "===========================================================================================" & vbNewLine

Dim ToolbarButton As CommandBarButton
Dim i As Integer
    
    'For each command button specified in the parameter arrays...
    For i = LBound(ButtonStrCommands_Array) To UBound(ButtonStrCommands_Array)

        '...convert StrCommands to a runable command...
        ButtonStrCommands_Array(i) = ConvertStrCommand(CStr(ButtonStrCommands_Array(i)), Verbose:=False)
        
        '...set up the ToolbarButton object...
        Set ToolbarButton = CustomToolbarsRow.Controls.Add(Type:=1)
            
            '...add the ToolbarButton according to the parameter arrays...
            With ToolbarButton
                .Caption = ButtonNames_Array(i)
                .Style = ButtonTypes_Array(i)
                .OnAction = ButtonStrCommands_Array(i)
            End With
            
            '...if the button is a type with a FaceId, set the property...
            If ButtonTypes_Array(i) = 1 Or ButtonTypes_Array(i) = 3 Then
                With ToolbarButton
                    .FaceId = MenuFaceIDs_Array(i)
                End With
            End If
            
            '...print what has happened in a readable way...
            Debug.Print "Button Name: [" & ButtonNames_Array(i) & "]"
            Debug.Print "       Runs: " & ButtonStrCommands_Array(i) & vbNewLine
        
        '...release the ToolbarButton object for the next loop
        Set ToolbarButton = Nothing
        
    Next i

Set CustomToolbarsRow = Nothing
End Sub

Sub RemoveAddInSection(MenuName As String)
    If AddInMenuExists(MenuName) Then Application.CommandBars(MenuName).Delete
End Sub

Function AddInMenuExists(MenuName As String)
Dim i As Integer

    For i = 1 To Application.CommandBars.Count
        If MenuName = Application.CommandBars(i).Name Then
            AddInMenuExists = True
            Exit Function
        End If
    Next i
    
AddInMenuExists = False
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## CREATE BUTTON SHAPE
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Try_CreateButtonShape()
Range("A1").Select

MsgBox "Creating a [Blank Button] shape that does nothing:", vbInformation, "[Blank Button]"
    Call CreateButtonShape

MsgBox "Creating a button, calling it [Button One], and assigning some properties", vbInformation, "[Button One]"
    Call CreateButtonShape( _
        btnLabel:="Button One", _
        StrCommand:="WriteLines{'This is a message'}", _
        btnColor:=5242976, _
        Top:=40 _
    )

MsgBox "Creating a button, calling it [Button Two], assigning some properties, and using it in a sub...", vbInformation, "[Button Two]"
    Dim Button As Object
    Set Button = _
        CreateButtonShape( _
            btnLabel:="Button Two", _
            btnName:="btnTwo", _
            StrCommand:="WriteLines{MyText:='This is a message', Repeat:=4}", _
            btnColor:=5242976, _
            Lef:=120, _
            Hei:=50 _
        )
        
MsgBox "Changing [Button Two]'s fill:", vbInformation, "Modifying Button in a Sub"
    With Button
        .Fill.ForeColor.RGB = RGB(255, 247, 254)
    End With

Application.Calculate 'To register fill on the screen change when running the sub

MsgBox "Changing [Button Two]'s text:", vbInformation, "Modifying Button in a Sub"
    With ActiveSheet.Shapes.Range(Array(Button.Name))
        .TextFrame2.TextRange.Characters.Text = "New Text"
    End With

End Sub

Function CreateButtonShape( _
    Optional StrCommand As String, _
    Optional btnLabel As String = "Blank Button", _
    Optional btnName As String, _
    Optional ShapeType As Integer = 5, _
    Optional btnColor As Long = 6299648, _
    Optional Lef As Long = 10, _
    Optional Top As Long = 10, _
    Optional Wid As Long = 100, _
    Optional Hei As Long = 20 _
)

Dim btn As Object, _
    btnRange As Object, _
    btnTextFrame As Object
        
    Set btn = ActiveSheet.Shapes.AddShape( _
        ShapeType, _
        Lef, Top, Wid, Hei _
    )
    Set btnRange = ActiveSheet.Shapes.Range(Array(btn.Name))
    Set btnTextFrame = btnRange.TextFrame2

        If btnName <> vbNullString Then
            Select Case ShapeExists(btnName)
                Case True
                    Call Err.Raise(Number:=1004, Description:="That shape name is already taken on the active sheet. Try a different one.")
                Case False
                    btn.Name = btnName
            End Select
        End If
        
        StrCommand = ConvertStrCommand(StrCommand, Verbose:=False)
          
        With btn
            .Line.ForeColor.RGB = btnColor
            .Fill.Visible = 0
            .OnAction = StrCommand
        End With
        
        With btnTextFrame
            .VerticalAnchor = 3
        End With
        
        With btnTextFrame.TextRange
            .Font.Name = "Avenir Next LT Pro"
            .Font.Fill.ForeColor.RGB = btnColor
            .Characters.Text = btnLabel
            .ParagraphFormat.Alignment = 2
        End With

        Set CreateButtonShape = btn

Set btn = Nothing
Set btnRange = Nothing
Set btnTextFrame = Nothing
    
End Function

Function ShapeExists(ShapeName As String)
Dim objShape As Object

    For Each objShape In ActiveSheet.Shapes
        If objShape.Name = ShapeName Then
            ShapeExists = True
            Exit Function
        End If
    Next objShape
    
ShapeExists = False
End Function

'===============================================================================================================================================================================================================================================================
'# RSCRIPT
'===============================================================================================================================================================================================================================================================

Sub CheckSystemInfo()
Print_Pad
               Print_Named MyOS(), "MyOS()"
    Print_Named PlatformFileSep(), "PlatformFileSep()"
       Print_Named Get_Username(), "Get_Username()"
    Print_Named Get_DesktopPath(), "Get_DesktopPath()"
  Print_Named Get_DownloadsPath(), "Get_DownloadsPath()"
        Print_Named Get_RFolder(), "Get_RFolder()"
       Print_Named Get_RScriptExePath(), "Get_RScriptExePath()"
Print_Pad
End Sub

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
        Call LocateRScript_Run(TempScriptPath)
                
                'NOTE: MsgBox question serves as both an option and a workaround for long R procedures
                'which prevent VBA's command line call from deleting Temp.R prior to Rscript.exe unloading Temp.R
                Answer = MsgBox("Temporary script written to desktop and ran in R." & vbNewLine & vbNewLine & _
                                "Would you like to delete the temporary file and it's folder?", vbYesNo, "Delete Temp.R File & Folder?")
            
                If Answer = vbYes Then
                    'Deletion successful
                    If ƒ—Delete_FileAndFolder(TempScriptPath) = True Then
                        
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

Set Fileout = Nothing
Set FSO = Nothing
End Function

Sub LocateRScript_Run(Script_Path As String)

    Call Run_RScript( _
        RScriptExe_Path:=Get_RScriptExePath, _
        Script_Path:=Script_Path, _
        Visibility:="Visible" _
    )

End Sub

Sub Run_RScript( _
    RScriptExe_Path As String, _
    Script_Path As String, _
    Optional Visibility As String, _
    Optional OnErrorEnd As Boolean = True _
)

Dim WaitTillComplete As Boolean: WaitTillComplete = True
Dim Style As Integer: Style = 1

Dim oShell As Object, _
    ErrorCode As Integer, _
    Escaped_RScriptExe As String, _
    Escaped_Script As String, _
    RShellCommand As String
    
    If Visibility = "VeryHidden" Then
        Style = 0
    ElseIf Visibility = "Minimized" Then
        Style = 2
    End If
    
Set oShell = CreateObject("WScript.Shell")
        
    Escaped_RScriptExe = Chr(34) & Replace(RScriptExe_Path, "\", "\\") & Chr(34)
    Escaped_Script = Chr(34) & Replace(Script_Path, "\", "\\") & Chr(34)
    RShellCommand = Escaped_RScriptExe & Escaped_Script
    ErrorCode = oShell.Run(RShellCommand, Style, WaitTillComplete)

        If OnErrorEnd = True And ErrorCode <> 0 Then
            MsgBox "Error attempting to run script. Ensure that any potential exceptions are wrapped in try().", vbInformation, "Run Failure"
            End
        End If
        
        Call Print_Named(ErrorCode, "ErrorCode")
        
Set oShell = Nothing
End Sub

Function Get_RScriptExePath() As String
Dim RVersions As Variant: RVersions = Get_RVersions(Get_RFolder)
Dim LatestRVersion As String: LatestRVersion = Get_LatestRVersion(RVersions)
    
    Get_RScriptExePath = LatestRVersion & "\bin\Rscript.exe"
             
End Function

Function Get_LatestRVersion(RVersions As Variant)
Dim i As Integer
    
    For i = LBound(RVersions) To UBound(RVersions)
        If Get_LatestRVersion < RVersions(i) Then
           Get_LatestRVersion = RVersions(i)
        End If
    Next i
    
End Function

Function Get_RVersions(RFolderPath As String)
    'Filter out C:\Program Files\R\.. & C:\Program Files\R\.
    Get_RVersions = Filter(ListFolders(RFolderPath), PlatformFileSep() & ".", False)
End Function

Function Get_RFolder() As String
Dim OS As String: OS = MyOS()

    If OS = "Windows" Then
        Get_RFolder = "C:\Program Files\R"
    ElseIf OS = "Mac" Then
        Get_RFolder = "/Library/Frameworks/R.framework/Resources/bin/R"
    End If
    
End Function

Function WriteScript( _
    TextContents As String, _
    SaveToDir As String, _
    Optional OverWrite As Boolean = False, _
    Optional ScriptName As String = "Script.R" _
)

'Add FileSep to directory string if required
If Right(SaveToDir, 1) <> PlatformFileSep() Then SaveToDir = SaveToDir & PlatformFileSep()

If OverWrite <> True Then
    If Dir(SaveToDir & ScriptName) <> vbNullString Then
        Dim i As Integer, SplitName As Variant, TryName As String
        For i = 1 To 100
            SplitName = Split(ScriptName, ".")
            TryName = SplitName(0) & " (" & i & ")" & "." & SplitName(1)
            If Dir(SaveToDir & TryName) = vbNullString Then
                ScriptName = TryName
                Exit For
            End If
        Next i
    End If
End If

Open SaveToDir & ScriptName For Output As #1
Print #1, TextContents
Close #1

WriteScript = CStr(SaveToDir & ScriptName)
End Function

'===============================================================================================================================================================================================================================================================
'# EXTRAS
'===============================================================================================================================================================================================================================================================

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'## LEGAL SPECIAL CHARACTER REFERENCE
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' ¶ € § Ø µ ª ° ¹ ² ³ · • ¿ ¡ ƒ × ¤ » « ‡ ¦ ± ÷ ¨ ¯ — ¬

'https://homepage.cs.uri.edu/faculty/wolfe/book/Readings/R02%20Ascii/completeASCII.htm


