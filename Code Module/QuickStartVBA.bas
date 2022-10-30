Attribute VB_Name = "QuickStartVBA"
Option Explicit
'===============================================================================================================================================================================================================================================================
'#  QuickStartVBA ¬ github.com/ulchc (10-29-22)
'===============================================================================================================================================================================================================================================================
'===============================================================================================================================================================================================================================================================
'## Overview
'===============================================================================================================================================================================================================================================================
'
'A collection of generic functions and subs for every action I had to implement in VBA more than once.
'
'Prefix ƒ— denotes a function which has a notable load time or file interactions
'outside ThisWorkbook. Since my intended use of the QuickStartVBA module/repo was to quickly
'port in many potentially useful snippets of code, then build onto a secondary module for
'a specific use case, I've opted to use this non-common chracter prefix instead of using
'Private Functions so that functions are available in any module.
'
'#### Sections
'  * [Functions](#functions)
'  * [Subs](#subs)
'  * [Data Transformation](#data-transformation)
'  * [User Interface](#user-interface)
'
'#### See Also
'  * [RscriptVBA](https://github.com/ulchc/RscriptVBA)
'
'===============================================================================================================================================================================================================================================================
'##  Important
'===============================================================================================================================================================================================================================================================
'
'If you intend to use the User Interface section, the following sub must be placed within ThisWorkbook:
'
'----------------------------------------------------------------``` VBA
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'   Call Remove_TempMenuCommands
'   Call Remove_TempMenuCommandSections
'End Sub
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
''   Returns username by reading the environment variable.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_DesktopPath()
'
''   Returns the desktop path regardless of platform with handling
''   for OneDrive hosted desktops.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  Get_DownloadsPath()
'
''   Returns the desktop path regardless of platform.
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
'  Clipboard_Load(ByVal YourString As String)
'
''   Stores {YourString} in clipboard.
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
' ƒ—Delete_FileAndFolder(ByVal aFilePath As String) as Boolean
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
' PlatformFileSep()
'
''   Returns "\" or "/" depending on the operating system.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
'  MyOS()
'
''   Read the system environment OS variable and returns "Windows",
''   "Mac", or the unaltered Environ("OS") output if neither.
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
' ExtractFirstInt_RightToLeft(aVariable)
'
''   Returns the first integer found in a string when searcing
''   from the right end of the string to the left.
'
'    ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ExtractFirstInt_LeftToRight(aVariable)
'
''   Returns the first integer found in a string when searcing
''   from the left end of the string to the right.
'
'    ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Truncate_Before_Int(YourString)
'
''   Removes characters before first integer in a sequence of characters.
'
'    Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Truncate_After_Int(YourString)
'
''   Removes characters after first integer in a sequence of characters.
'
'    Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' IsInt_NoTrailingSymbols(aNumeric)
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
' LaunchLink(aLink)
'
''   Launches aLink in existing browser with error handling for
''   invalid Links
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ToggleDisplayMode()
'
''   Toggles display of ribbon, formula bar, status bar & headings
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CreateSlicer( _
'     tblKeyAddress As String, _
'     ColumnName As String, _
'     HorizAlignAddress As String, _
'     HorizAlignRight As Boolean, _
'     Optional Wb As Workbook, _
'     Optional Ws As Worksheet, _
'     Optional BtnsPerRow As Long = 3, _
'     Optional BtnsPerCol As Long = 2, _
'     Optional BtnsPointWidth As Long = 80 _
' )
'
''   Uses {tblKeyAddress} to determine the ListObject name,
''   creates a slicer for {ColumnName}, and then aligns it with the
''   cell specified by {HorizAlignAddress}.
'
''   Aligns the slicer with the top right corner of the cell when
''   {HorizAlignRight} = True and the top left corner when
''   {HorizAlignRight} = False.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' HorizAlignShape( _
'     ShapeObject As Object, _
'     AlignToRange As Range, _
'     RightAlign As Boolean _
' )
'
''   Written for use in CreateSlicer(), but can be used to
''   skip the calculations involved to right align any shape.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_AvenirSlicerStyle()
'
''   Creates the .TableStyle "AvenirSlicerStyle" for CreateSlicer().
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' TableStyleExists(StyleNamed As String)
'
''   Returns True or False depending on if .TableStyle({StyleNamed})
''   exists.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Create_Comment( _
'     rngComment As Range, _
'     arrTextLines As Variant, _
'     Optional WidthFactor As Single = 1, _
'     Optional HeightFactor As Single = 1, _
'     Optional FontSize As Single = 9, _
'     Optional FontColor As Long = 0, _
'     Optional BoldChoice As Boolean = False, _
'     Optional UseFormatStrings As Boolean = False, _
'     Optional VisibleProperty As Boolean = True, _
'     Optional BorderColor As Long = 6299648, _
'     Optional BorderWeight As Single = 1.3, _
'     Optional FillColor As Long = 16777215, _
'     Optional FillPicturePath As String, _
'     Optional OverrideExisting As Boolean = True _
' )
'
''   Adds a comment to {rngComment} that has a cleaner look than
''   the base Excel comment, with each item of {arrTextLines} written
''   to the comment as a line of text seperated by a new line character,
''   and optional arguments to make changing the comment's properties
''   less convoluted.
'
''   More notably, automatically adjusts the dimensions of image
''   comments to match the aspect ratio of the image, and enables the
''   use of format strings to bolden specific sections of text in
''   {arrTextLines}.
'
''   If {arrTextLines} = Array("#Header#", "Point 1", "Point 2")
'
''   The comment would show as:
'
''   Header (bold)
''   Point 1
''   Point 2
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_AspectRatio(ImgPath As String)
'
''   Written for use in Create_Comment(), but will return the aspect
''   ratio (width / height) for any image.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' PrintEnvironVariables()
'
''   Print the environment variables to the Immediate window.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Print_Named(Something, Optional Label)
'
''   Uses Debug.Print to add a space between each {Something} printed,
''   labels each {Something} if {Label} supplied.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Print_Pad()
'
''   Uses Debug.Print to print a timestamped seperator of "======"
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'## Data Transformation
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
' Filter_By( _
'     rngColumn As Range, _
'     AdvFilterTerm As String, _
'     Optional rngTable As Range _
' )
'
''   Removes filtered terms from either a range or ListObject by
''   copying the .CurrentRegion of {rngColumn} (unless {rngTable}
''   is specified), using .AdvancedFilter on the copy with the
''   {AdvFilterTerm} applied to {rngColumn}, then overwriting the
''   previous range or ListObject with the filtered result.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Order_by( _
'     rngColumn As Range, _
'     Optional Descending As Boolean = True _
' )
'
''   Self-explanatory simplification of .Sort on a table.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Pivot_Wider( _
'     rngTable As Range, _
'     NamesFrom As String, _
'     ValuesFrom As String, _
'     JoinFrom As String, _
'     Optional PerfectMatchOnly As Boolean = True _
' )
'
''   Adds new columns to {rngTable} by seperating each category in
''   column {NamesFrom} into it's own column, with column values
''   obtained from the column named {ValuesFrom}.
'
''   If a column name should be approximately matched, {PerfectMatchOnly}
''   can be set equal to False.
'
''   After pivoting categories into columns, if it is found that there
''   are mutiple rows for a single value of the column named {JoinFrom},
''   the category values are consolidated into a single row and the
''   duplicate rows are removed.
'
''   Works similarily to dplyr pivot_wider() in R, with the original
''   columns removed after the pivot.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ColumnSub( _
'     rngColumn As Range, _
'     strSubstitute As String, _
'     strReplacement As String _
' )
'
''   Subsitutes each occurance of {strSubstitute} with {strReplacement}
''   for all values in {rngColumn}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Drop_Columns( _
'     rngTable As Range, _
'     strMatch As String, _
'     Optional PerfectMatch As Boolean = False _
' )
'
''   Deletes any column with a header matching {strMatch} in {rngTable},
''   with optional parameter {PerfectMatch} to adjust match precision.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Set_MinColWidth( _
'     rngTable As Range, _
'     MinWidth As Single, _
'     Optional OverrideAll As Boolean = False _
' )
'
''   Adjusts columns widths of {rngTable} to at minimum be {MinWidth}
''   wide, with optional parameter {OverrideAll} to reset all widths
''   to {MinWidth}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Split_ColumnValues( _
'     rngColumn As Range, _
'     SplitTerm As String, _
'     SplitKeepIndex As Long _
' )
'
''   Splits the values in {rngColumn} by {SplitTerm} and substitutes
''   column values with the split index specified: {SplitKeepIndex}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Reorder_Columns( _
'     Named As Variant, _
'     FromTable As Range, _
'     Optional PerfectMatch As Boolean = False, _
'     Optional ToLocation As String = "{Start} or {End}" _
' )
'
''   Rearranges a subset of columns specified in the array {Named}
''   by the order they were supplied, either to the "Start" or "End"
''   of the .CurrentRegion of the table, with optional parameter
''   {PerfectMatch} to adjust match precision of column names.
'
''   Note: {ToLocation} is not Optional. The default value is simply
''   a means to make the choice values visible when calling the sub.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Filter_Dupes( _
'     FromColsNamed As Variant, _
'     rngTable As Range _
' )
'
''   Creates a non-ListObject copy of {rngTable}, filter duplicates
''   across all the columns specified in {FromColsNamed}, then
''   overwrites the original table with the filtered result.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Overwrite_Table( _
'     tblCurrent As Range, _
'     tblNew As Range _
' )

''   Overwrites {tblCurrent} with {tblNew} and resizes the ListObject
''   that linked with {tblCurrent} if applicable.
'
''   Note: Column dimensions must be the same (Excel table compatability)
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Fast_Copy( _
'     rngToCopy As Range, _
'     Optional rngOutput As Range _
' )
'
''   Returns {rngOutput} after using .AdvancedFilter to copy
''   {rngToCopy} to the right of itself. If [rngOutput} is specified,
''   the default output location will be overridden.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_FirstRow(rngFrom As Range)
'
''   Returns the first row of a given range.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_LastRow(rngFrom As Range)
'
''   Returns the last row of a given range.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_LastColumn(rngFrom As Range)
'
''   Returns the last column of a given range.
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  User Interface
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
' ConvertStrCommand( _
'     CommandString As String, _
'     Optional Verbose As Boolean = True _
' )
'
''   Automatically applied to all {StrCommand}'s passed to the menu
''   and button creation functions below (prior to linking the
''   macro to the object).
'
''   Changes existing apostrophes in {StrCommand} to quotation marks,
''   encases the full command in apostrophes, and substitutes curly
''   braces for brackets.
'
''   This is to make it easier to supply parameters to a sub or function
''   called by a menu or shape without having to include a long list of values
''   seperated with Chr(34) & "..." & Chr(34).
'
''   Original:   "MySub(Range{'NamedRange'}, 2)"
''   Converted: "'MySub(Range("NamedRange"), 2)'"
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ChangeMenuVisibility( _
'     MenuItems_Array As Variant, _
'     VisibleProperty As Boolean _
' )
'
''   Toggles the visibility of items on the menu shown by right
''   clicking a cell. For situations where the menu is becoming
''   overcrowded with custom commands.
'
''   Menu items can be refered to with the same string as their
''   display names *except* in the case of an underlined letter,
''   in which case, the true name includes an &. For example,
''   "Copy" is actually "&Copy".
'
''   All visibility modifications can be returned to default by
''   calling ResetCellMenu()
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ResetCellMenu()
'
''   Restores CommandBars("Cell") and ShortcutMenus(xlWorksheetCell)
''   to their default states.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CreateMenuCommand( _
'    MenuCommandName As String, _
'    StrCommand As String, _
'    Optional Temporary As Boolean = True, _
'    Optional MenuFaceID As Long _
' )
'PARAMETERS:
''    {MenuCommandName} = The name of the menu that will be created.
''    {StrCommand} = Command to run when clicked (see ConvertStrCommand).
''    {Temporary} = Whether the menu should be deleted on the WorkbookClose event.
''    {MenuFaceID} = The FaceId integer for the menu command.
'
'EXPLANATION:
''    Adds an item to the top of the menu displayed when right clicking
''    a cell on a worksheet. Shows up with the caption {MenuCommandName}
''    and the icon specified with {MenuFaceID}.
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
''    {MenuSectionName} = The name of the menu section that will be created.
''    {Array_SectionMenuNames} = Array of display names for each command.
''    {Array_StrCommands} = Array of commands for each menu item (see ConvertStrCommand).
''    {Temporary} = Whether the menu should be deleted on the WorkbookClose event.
'
'EXPLANATION:
''    Adds a section to the top of the menu displayed when right clicking
''    a cell on a worksheet. Shows up with the caption {MenuSectionName}
''    and no icon.
'
''    When hovering over the menu section, a list of commands
''    specified by {Array_SectionMenuNames} will display, each running
''    it's corresponding macro specified in {Array_StrCommands}.
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
''    {PopupMenuName} = The name of the menu that will be created.
''    {Array_ItemNames} = Array of display names for each command.
''    {Array_StrCommands} = Array of commands for each menu item (see ConvertStrCommand).
''    {Array_ItemFaceIDs} = Array of FaceId integers for each menu item.
''    {Temporary} = Whether the menu should be deleted on the WorkbookClose event.
'
'EXPLANATION:
''    Creates a custom menu named {PopupMenuName} which can be displayed with
''    Application.CommandBars({PopupMenuName}).ShowPopup.
'
''    Each item from {Array_ItemNames} is included in the menu with
''    it's corresponding integer FaceId specified by {Array_ItemFaceIDs}.
''    When an item is clicked, it runs it's respective macro as
''    assigned by {Array_StrCommands}.
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
''    {StrCommand} = Commands to run when clicked (see ConvertStrCommand).
''    {btnLabel} = The display name of the shape.
''    {btnName} = The .Name property of the shape.
''    {ShapeType} = The look of the shape as specified by the integer type.
''    {btnColor} = Color code of the shape
''    {Lef} = .Left property of the shape
''    {Top} = .Top property of the shape
''    {Wid} = .Width property of the shape
''    {Hei} = .Hweight property of the shape
'
'EXPLANATION:
''    Inserts a shape onto the sheet that has the appears of a button
''    and runs {StrCommand} when clicked.
'
'EXAMPLES: '(Ctrl+f to view & run)
'     Sub Try_CreateButtonShape
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' IdentifyMenus(Optional RemoveIndicators As Boolean = False)
'
''   Loops through each CommandBar in the workbook and adds a
''   new indicator command *This is CommandBar(i)* to the top
''   of the menu so that the index of the menu can be identified.
'
''   This is simply to enable the modification of CommandBars other
''   than the worksheet cell menus (ex. ListObject), which aren't
''   often named in an intuitive way.
'
'----------------------------------------------------------------```

'===============================================================================================================================================================================================================================================================
'##  Functions
'===============================================================================================================================================================================================================================================================

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

Dim Ws As Worksheet, _
    Wb As Workbook, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant, _
    ScreenUpdatingState As Boolean

ScreenUpdatingState = Application.ScreenUpdating
Application.ScreenUpdating = False

Application.StatusBar = "Opening " & Right(FromFile, Len(FromFile) - InStrRev(FromFile, PlatformFileSep())) & "..."
Workbooks.Open FileName:=FromFile, ReadOnly:=True
Set Wb = ActiveWorkbook
    
    Wb.Sheets(1).Select
    SheetCount = Wb.Sheets.Count
    
    For Each Ws In Wb.Worksheets
        Sheet_i = Sheet_i + 1
        Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & Wb.Name & "..."
         
        With Ws
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
    Next Ws
    
Application.DisplayAlerts = False
Wb.Close
Application.DisplayAlerts = True

PasteSheetVals_FromFile = NewSheets()
    
Application.StatusBar = False
Application.ScreenUpdating = ScreenUpdatingState
End Function

Function CopySheets_FromFile(FromFile As String)

Dim Ws As Worksheet, _
    Wb As Workbook, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant, _
    ScreenUpdatingState As Boolean

ScreenUpdatingState = Application.ScreenUpdating
Application.ScreenUpdating = False

Application.StatusBar = "Opening " & Right(FromFile, Len(FromFile) - InStrRev(FromFile, PlatformFileSep())) & "..."
Workbooks.Open FileName:=FromFile, ReadOnly:=True
Set Wb = ActiveWorkbook

    Wb.Sheets(1).Select
    SheetCount = Wb.Sheets.Count
    
    For Each Ws In Wb.Worksheets
        Sheet_i = Sheet_i + 1
        Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & Wb.Name & "..."
        Ws.Copy after:=ThisWorkbook.ActiveSheet
        Call ReDim_Add(NewSheets(), ActiveSheet.Name)
    Next Ws
    
Wb.Close
    
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
    
Dim Ws As Worksheet, _
    Wb As Workbook, _
    wbFile As Variant, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant

    For Each wbFile In wbFiles()
    
        Workbooks.Open FileName:=wbFile, ReadOnly:=True
        Set Wb = ActiveWorkbook
        Wb.Sheets(1).Select
        SheetCount = Wb.Sheets.Count
            
            For Each Ws In Wb.Worksheets
                Sheet_i = Sheet_i + 1
                Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & Wb.Name & "..."
                 
                With Ws
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
            Next Ws
    
        Application.DisplayAlerts = False
        Wb.Close
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
    
Dim Ws As Worksheet, _
    Wb As Workbook, _
    wbFile As Variant, _
    SheetCount As Integer, _
    Sheet_i As Integer, _
    NewSheets() As Variant

    For Each wbFile In wbFiles()
        Workbooks.Open FileName:=wbFile, ReadOnly:=True
            
            Set Wb = ActiveWorkbook
            Wb.Sheets(1).Select
            SheetCount = Wb.Sheets.Count
            
            For Each Ws In Wb.Worksheets
                Sheet_i = Sheet_i + 1
                Application.StatusBar = "Adding sheet " & Sheet_i & " of " & SheetCount & " from " & Wb.Name & "..."
                Ws.Copy after:=ThisWorkbook.ActiveSheet
                Call ReDim_Add(NewSheets(), ActiveSheet.Name)
            Next Ws
        
        Wb.Close
    Next wbFile
    
CopySheets_FromFolder = NewSheets()
    
Application.StatusBar = False
Application.ScreenUpdating = ScreenUpdatingState
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

Function ReadLines( _
    TxtFile As String, _
    Optional ToImmediate As Boolean = True, _
    Optional ToClipboard As Boolean = True, _
    Optional Replace_AnyOf As String = "Of String", _
    Optional Replace_With As String = "With String" _
)

Dim SheetFX As Object: Set SheetFX = Application.WorksheetFunction
Dim FileNum As Integer: FileNum = FreeFile
Dim TxtFileLines() As String
    
    Open TxtFile For Input As FileNum
        TxtFileLines = Split(Input$(LOF(FileNum), FileNum), vbNewLine)
    Close FileNum

Dim TxtFileContents As String
    TxtFileContents = SheetFX.TextJoin(vbNewLine, False, TxtFileLines)

'Use UDF Replace_Any() on Windows (Regex not available on Mac)
If MyOS() = "Windows" Then
    'Optional default is meant to indicate how the parameter works...
    If Replace_AnyOf <> "Of String" Then
        '...only proceed with replacement if the default value has been changed
        TxtFileContents = Replace_Any(Replace_AnyOf, Replace_With, TxtFileContents)
    End If
End If

If ToImmediate = True Then
    Debug.Print TxtFileContents
End If

If ToClipboard = True Then
    Call Print_Named( _
        IIf(Clipboard_Load(TxtFileContents) = True, _
        "Output copied to clipboard.", _
        "Output could not be copied to clipboard."), _
        "Clipboard Status" _
    )
End If

ReadLines = TxtFileContents
    
Set SheetFX = Nothing
End Function

Function Clipboard_Load(ByVal YourString As String) As Boolean

On Error GoTo NoLoad
    CreateObject("HTMLFile").ParentWindow.ClipboardData.SetData "text", YourString
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
    Optional Wb As Workbook _
)
Dim aSheet As Worksheet
    
    If Wb Is Nothing Then Set Wb = ThisWorkbook
        
        On Error Resume Next
            Set aSheet = Wb.Sheets(aName)
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

Function Truncate_Before_Int(ByVal YourString As String)
On Error GoTo NoInt:

Dim CountCharsToRemove As Integer, _
    CheckCharacter As String, _
    NewStrLength As Integer, _
    i As Integer
    
    YourString = Trim(YourString)
    
    'Return immediately if already integer
    If IsInt_NoTrailingSymbols(YourString) Then
        Truncate_Before_Int = YourString
        Exit Function
    End If
        
        CountCharsToRemove = 0
    
        'Loop to determine number of starting characters to remove
        For i = 1 To Len(YourString)
        
            'Single character string at point i in {YourString}, e.g. "S" or "o"
            CheckCharacter = Right(Left(YourString, i), 1)
                
                If IsInt_NoTrailingSymbols(CheckCharacter) = False Then
                    CountCharsToRemove = CountCharsToRemove + 1
                ElseIf IsNumeric(CheckCharacter) = True Then
                    Exit For
                End If
        Next i
                    
                    NewStrLength = Len(YourString) - CountCharsToRemove
                    Truncate_Before_Int = Right(YourString, NewStrLength)
                    
                    Exit Function
        
NoInt:
Truncate_Before_Int = vbNullString

End Function

Function Truncate_After_Int(ByVal YourString As String)
On Error GoTo NoInt:

Dim CountCharsToRemove As Integer, _
    CheckCharacter As String, _
    NewStrLength As Integer, _
    i As Integer
    
    YourString = Trim(YourString)
    
    'Return immediately if already integer
    If IsInt_NoTrailingSymbols(YourString) Then
        Truncate_After_Int = YourString
        Exit Function
    End If
        
        CountCharsToRemove = 0
    
        'Loop to determine number of starting characters to remove
        For i = 1 To Len(YourString)
        
            'Single character string at point i in {YourString}, e.g. "S" or "o"
            CheckCharacter = Left(Right(YourString, i), 1)
                
                If IsNumeric(CheckCharacter) = False Then
                    CountCharsToRemove = CountCharsToRemove + 1
                ElseIf IsNumeric(CheckCharacter) = True Then
                    Exit For
                End If
        Next i
            
                    NewStrLength = Len(YourString) - CountCharsToRemove
                    Truncate_After_Int = Left(YourString, NewStrLength)
                    
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
    
ThisUser = Get_Username
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

Function Get_DownloadsPath()
    Get_DownloadsPath = Environ("USERPROFILE") & PlatformFileSep & "Downloads"
End Function

Function Get_DesktopPath()
    If Environ("OneDriveConsumer") <> vbNullString Then
        Get_DesktopPath = Environ("OneDriveConsumer") & PlatformFileSep() & "Desktop"
    Else
        Get_DesktopPath = Get_Username() & PlatformFileSep() & "Desktop"
    End If
End Function

Function Get_Username()
    Get_Username = Environ("USERNAME")
End Function

Function MyOS()
Dim EnvOS As String: EnvOS = Environ("OS")
    If InStr(1, EnvOS, "Windows") <> 0 Then
        MyOS = "Windows"
    ElseIf InStr(1, EnvOS, "Mac") <> 0 Then
        MyOS = "Mac"
    Else
        MyOS = EnvOS
    End If
End Function

'===============================================================================================================================================================================================================================================================
'##  Subs
'===============================================================================================================================================================================================================================================================

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

Sub CreateSlicer( _
    tblKeyAddress As String, _
    ColumnName As String, _
    HorizAlignAddress As String, _
    HorizAlignRight As Boolean, _
    Optional Wb As Workbook, _
    Optional Ws As Worksheet, _
    Optional BtnsPerRow As Long = 3, _
    Optional BtnsPerCol As Long = 2, _
    Optional BtnsPointWidth As Long = 80 _
)

Dim NewCache As SlicerCache, _
    NewSlicer As Slicer, _
    SheetFX As Object, _
    AutoWidth As Long, _
    AutoHeight As Long, _
    tblName As String, _
    PrevSelection As Object
    
Set PrevSelection = Selection
If TypeName(Wb) = "Nothing" Then Set Wb = ActiveWorkbook
If TypeName(Ws) = "Nothing" Then Set Ws = ActiveSheet

'Obtain the list object name
With Wb.Sheets(Ws.Name)
    tblName = .ListObjects(Wb.Sheets(Ws.Name).Range(tblKeyAddress).ListObject.Name)
End With

'Create the new slicer cache
Set NewCache _
  = Wb.SlicerCaches.Add2( _
    Source:=Wb.Sheets(Ws.Name).ListObjects(tblName), _
    SourceField:=CStr(ColumnName) _
)

Set SheetFX = Application.WorksheetFunction
'Default of 35 for the shape border whitespace, increase by {BtnsPerRow}
AutoWidth = 35 + (BtnsPointWidth * BtnsPerRow)
'Default of 30 for the header, increase by minimum of {BtnsPerCol} and items
AutoHeight = 29 + SheetFX.RoundUp(SheetFX.Min(NewCache.SlicerItems.Count / BtnsPerRow, BtnsPerCol), 0) * 22.5

'Create the new slicer
Set NewSlicer _
  = NewCache.Slicers.Add( _
    SlicerDestination:=Wb.Sheets(Ws.Name), _
    Caption:=CStr(ColumnName), _
    Top:=0, _
    Left:=0, _
    Width:=AutoWidth, _
    Height:=AutoHeight _
)

'Modify the look of the slicer
With NewCache.Slicers(NewSlicer.Name)
    .NumberOfColumns = BtnsPerRow
    .Style = Get_AvenirSlicerStyle()
End With
    
'Move but don't size with cells
NewCache.Slicers.Item(NewSlicer.Name).Shape.Placement = xlMove

Call HorizAlignShape( _
    ShapeObject:=NewSlicer, _
    AlignToRange:=Wb.Sheets(Ws.Name).Range(HorizAlignAddress), _
    RightAlign:=HorizAlignRight _
)

Set NewCache = Nothing
Set NewSlicer = Nothing
Set SheetFX = Nothing
    PrevSelection.Select

End Sub

Sub HorizAlignShape( _
    ShapeObject As Object, _
    AlignToRange As Range, _
    RightAlign As Boolean _
)

Dim ShapePointWidth As Single: ShapePointWidth = ActiveSheet.Shapes(ShapeObject.Name).Width
Dim RangePointWidth As Single: RangePointWidth = AlignToRange.Offset(0, 1).Left - AlignToRange.Left

With ShapeObject
    .Left = AlignToRange.Left
    .Top = AlignToRange.Top
End With

If RightAlign = True Then
    With ActiveSheet.Shapes(ShapeObject.Name)
        .IncrementLeft (RangePointWidth - ShapePointWidth)
    End With
End If

End Sub

Function Get_AvenirSlicerStyle()
 
If TableStyleExists("AvenirSlicerStyle") Then
    GoTo GetStyle
End If

Dim AvenirStyle As Object
Set AvenirStyle = ActiveWorkbook.TableStyles("SlicerStyleLight1").Duplicate( _
    NewTableStyleName:="AvenirSlicerStyle" _
)

AvenirStyle.ShowAsAvailableSlicerStyle = True
AvenirStyle.TableStyleElements(xlWholeTable).Clear

With AvenirStyle.TableStyleElements(xlWholeTable)
    .Font.Name = "Avenir Next LT Pro"
    .Font.ThemeFont = xlThemeFontNone
End With

With AvenirStyle.TableStyleElements(xlHeaderRow)
    .Font.FontStyle = "Bold"
    .Font.Color = 6299648
    .Font.Size = 12
    .Borders(xlEdgeBottom).Color = 6299648
    .Borders(xlEdgeBottom).Weight = xlThick
    .Borders(xlEdgeBottom).LineStyle = 9
End With

GetStyle: Get_AvenirSlicerStyle = "AvenirSlicerStyle"

End Function

Function TableStyleExists(StyleNamed As String)
TableStyleExists = False

Dim i As Long
For i = 1 To ActiveWorkbook.TableStyles.Count
    If ActiveWorkbook.TableStyles(i) = StyleNamed Then
        TableStyleExists = True
        Exit Function
    End If
Next i

End Function

Sub Create_Comment( _
    rngComment As Range, _
    arrTextLines As Variant, _
    Optional WidthFactor As Single = 1, _
    Optional HeightFactor As Single = 1, _
    Optional FontSize As Single = 9, _
    Optional FontColor As Long = 0, _
    Optional BoldChoice As Boolean = False, _
    Optional UseFormatStrings As Boolean = False, _
    Optional VisibleProperty As Boolean = True, _
    Optional BorderColor As Long = 6299648, _
    Optional BorderWeight As Single = 1.3, _
    Optional FillColor As Long = 16777215, _
    Optional FillPicturePath As String, _
    Optional OverrideExisting As Boolean = True _
)

Dim txtComment As String, _
    NewComment As Comment

txtComment = Join(arrTextLines, Chr(10))

If TypeName(rngComment.Comment) = "Nothing" Then
    Set NewComment = rngComment.AddComment
Else
    If OverrideExisting = True Then
        rngComment.Comment.Delete
        Set NewComment = rngComment.AddComment
    Else
        Set NewComment = rngComment.Comment
    End If
End If

With NewComment
    .Text txtComment
    .Visible = VisibleProperty
    
    With .Shape
           .Fill.ForeColor.RGB = FillColor
           .Line.ForeColor.RGB = BorderColor
           .Line.Weight = BorderWeight
           .ScaleWidth Factor:=WidthFactor, RelativeToOriginalSize:=0
           .ScaleHeight Factor:=HeightFactor, RelativeToOriginalSize:=0
           .TextFrame.AutoMargins = False
           .TextFrame.MarginLeft = 15
           .TextFrame.MarginRight = 15
           .TextFrame.MarginBottom = 11
           .TextFrame.MarginTop = 11

        With .TextFrame.Characters.Font
               .Color = vbBlack
               .Size = FontSize
               .Name = "Avenir Next LT Pro"
               .Bold = BoldChoice
               
        End With
    
    End With
    
End With

If UseFormatStrings = False Then
    GoTo SkipLoop
End If

Dim LocStart As Long, _
    LocEnd As Long, _
    adjLocStart As Long, _
    adjLocEnd As Long, _
    iBold As Long, _
    adjBoldFactor As Long, _
    i As Long

'Check how many bold format strings #{str}# there are...
iBold = (Len(txtComment) - Len(Replace(txtComment, "#", vbNullString))) / 2

'...iterate over the comment iBold times
If iBold > 0 Then
    For i = 1 To iBold
          
        'Note: LocEnd will initially be 0
         
        'Search for the location to begin formatting after the last iteration's LocEnd + 1
        LocStart = Application.Find("#", txtComment, LocEnd + 1)
        
        'Search for the location to end formatting after the symbol just found, LocStart + 1
        LocEnd = Application.Find("#", txtComment, LocStart + 1)
        
        'For each iteration, NewComment's text will be reduced by two characters (formatting symbols). As such, we need to adjust LocStart and LocEnd to account for the sequential removal of these terms.
        
        'Somehow, NewComment's text property only stores 255 characters, so the full length of a comment's text cannot be read. As such, adjusting based on {txtComment} is neccesary.
        
        adjBoldFactor = (i * 2 - 2)
        adjLocStart = LocStart - adjBoldFactor
        adjLocEnd = LocEnd - adjBoldFactor
           
            With NewComment.Shape.TextFrame
                'Bolden the #{str}# between the symbols...
                .Characters( _
                    Start:=adjLocStart, _
                    Length:=(adjLocEnd - adjLocStart + 1)).Font.Bold = True
                '...remove <#>{str}#...
                 .Characters( _
                    Start:=adjLocStart, _
                    Length:=1).Text = vbNullString
                '...remove {str}<#>...
                .Characters( _
                    Start:=adjLocEnd - 1, _
                    Length:=1).Text = vbNullString
            End With
    Next i
End If

SkipLoop:

If FillPicturePath <> vbNullString Then
    
    Dim AspectRatio As Single: AspectRatio = Get_AspectRatio(FillPicturePath)
        
    With NewComment.Shape
           .Fill.UserPicture FillPicturePath
           .ScaleWidth Factor:=IIf(AspectRatio > 1, 0.6, 0.6 * 1.7), RelativeToOriginalSize:=0
           .ScaleHeight Factor:=IIf(AspectRatio > 1, AspectRatio, AspectRatio * 1.7), RelativeToOriginalSize:=0
    End With

End If

End Sub

Function Get_AspectRatio(ImgPath As String)
Dim WinShell As Object, _
    ImgObject As Object, _
    ImgDir As String, _
    ImgName As String, _
    ImgDim As String, _
    ImgAspect As Single

'Image container folder
ImgDir = Left(ImgPath, InStrRev(ImgPath, "\") - 1)
'Image name from path
ImgName = Right(ImgPath, Len(ImgPath) - InStrRev(ImgPath, "\"))
    
Set WinShell = CreateObject("Shell.Application")
Set ImgObject = WinShell.Namespace(CStr(ImgDir)).ParseName(ImgName)

ImgDim = ImgObject.ExtendedProperty("Dimensions")
'Remove invalid characters on left & right of "Dimensions"
ImgDim = Left(Right(ImgDim, Len(ImgDim) - 1), Len(ImgDim) - 2)

Get_AspectRatio = CSng(Split(ImgDim, " x ")(1) / Split(ImgDim, " x ")(0))

Set WinShell = Nothing
End Function

Sub PrintEnvironVariables()
Dim i As Integer: i = 1

Dim EnvVarItem As Variant, _
    EnvVarName As String, _
    EnvVar As String

    EnvVarItem = Split(Environ(i), "=")
    EnvVarName = CStr(EnvVarItem(0))
    EnvVar = CStr(EnvVarItem(1))

    Do Until Environ(i) = vbNullString
        EnvVarItem = Split(Environ(i), "=")
        EnvVarName = CStr(EnvVarItem(0))
        EnvVar = CStr(EnvVarItem(1))
        Call Print_Named(EnvVar, EnvVarName)
        i = i + 1
    Loop

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

Sub Print_Pad()
    Debug.Print ("================== " & Format(Now(), "Long Time") & " ==================")
End Sub

'===============================================================================================================================================================================================================================================================
'## Data Transformation
'===============================================================================================================================================================================================================================================================

Function Filter_By( _
    rngColumn As Range, _
    AdvFilterTerm As String, _
    Optional rngTable As Range _
)

'Wrap {AdvFilterTerm} to ="{AdvFilterTerm}" so...
AdvFilterTerm = "=" & Chr(34) & AdvFilterTerm & Chr(34)
'...that it is compatible with AdvancedFilter syntax

'Default the table region to the CurrentRegion of {rngColumn}
If TypeName(rngTable) = "Nothing" Then Set rngTable = rngColumn.CurrentRegion

'A temporary copy of the table will be placed to the right of the original range + 1...
Dim rngOutput As Range: Set rngOutput = rngTable.Offset(0, Get_FirstRow(rngTable).Count + 1)
'...so that in the case of {rngTable} being an Excel table, it does not auto-include the copy

    'Create a copy of {rngTable}
    With rngTable.AdvancedFilter( _
        Action:=xlFilterCopy, _
        CopyToRange:=rngOutput _
    ): End With
        
'Obtain the headers of the copied table...
Dim rngHeaders As Range: Set rngHeaders = Get_FirstRow(rngOutput)
'...then obtain the top two rows for the AdvancedFilter criteria section...
Dim rngCriteriaSection As Range: Set rngCriteriaSection = Range(rngHeaders, rngHeaders.Offset(1, 0))
'...and finally, set {rngCriteria} to an area *FAR AWAY* since it becomes an auto-referenced named range for every advanced filter afterwards
Dim rngCriteria As Range: Set rngCriteria = rngCriteriaSection.Offset(rngOutput.Cells(rngOutput.Count).Row * 5, rngHeaders.Column * 5)

    'Create a copy of the criteria section: {rngCriteria}
    With rngCriteriaSection.AdvancedFilter( _
        Action:=xlFilterCopy, _
        CopyToRange:=rngCriteria _
    ): End With
    
        'Clear the first row of {rngCriteria}, as it represents...
        rngCriteria.Resize(1).Offset(1, 0).Clear
        '...an entry from the copied table, not criteria
        
Dim i As Integer
'Since {rngCriteria} is a 2 row table...
For i = 1 To rngCriteria.Count / 2 '...this only includes the headers
    If rngCriteria.Cells(i).Value = rngColumn.Cells(1).Value Then
        'Locate {rngCriteria}'s header that corresponds with {rngColumn}
        Dim rngCriteriaCol As Range: Set rngCriteriaCol = rngCriteria.Cells(i)
        Exit For
    End If
Next i

    'Place {advFilterTerm} into the criteria slot
    rngCriteriaCol.Offset(1, 0).Value = AdvFilterTerm
         
        'Filter the original copy {rngOutput} by {rngCriteria}...
        With rngOutput.AdvancedFilter( _
            Action:=xlFilterCopy, _
            CriteriaRange:=rngCriteria, _
            CopyToRange:=rngTable _
        ): End With '...then override {rngTable} with the result
    
'In the case of {rngTable} being an Excel table object...
If TypeName(rngTable.Cells(1).ListObject) <> "Nothing" Then
    
    '...we'll also need to resize the Excel table object...
    With ActiveSheet.ListObjects(rngTable.ListObject.Name)
        .Resize rngTable.CurrentRegion '...to the new .CurrentRegion
    End With
    
End If

'Once {rngTable} has been overwritten...
rngCriteria.Clear
rngOutput.Clear
'...remove the tables used to apply AdvancedFilter

Set Filter_By = rngTable.CurrentRegion
End Function

Sub Order_by( _
    rngColumn As Range, _
    Optional Descending As Boolean = True _
)

Dim WbWs As Object: Set WbWs = ActiveWorkbook.ActiveSheet
    WbWs.Range(rngColumn.Cells(1).Address).Select 'Required

'Guarantee AutoFilterMode state of the worksheet is initially = False...
If WbWs.AutoFilterMode = True Then WbWs.Range(rngColumn.Cells(1).Address).AutoFilter

    '...then apply AutoFilter and clear criteria
    WbWs.Range(rngColumn.Cells(1).Address).AutoFilter

Dim objSort As Object 'Different depending on Range vs Excel Table
    
    If TypeName(rngColumn.Cells(1).ListObject) <> "Nothing" Then
        Set objSort = WbWs.ListObjects(rngColumn.Cells(1).ListObject.Name)
    Else
        Set objSort = WbWs.AutoFilter
    End If

'Guarantee the SortFields of the object are cleared...
objSort.Sort.SortFields.Clear
    
    '...add the desired criteria...
    With objSort.Sort.SortFields.Add2( _
        Key:=Range(rngColumn.Address), _
        SortOn:=xlSortOnValues, _
        Order:=IIf(Descending = True, xlDescending, xlAscending), _
        DataOption:=xlSortNormal _
    ): End With
        
    '...then apply it
    With objSort.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Set objSort = Nothing
End Sub

Sub Pivot_Wider( _
    rngTable As Range, _
    NamesFrom As String, _
    ValuesFrom As String, _
    JoinFrom As String, _
    Optional PerfectMatchOnly As Boolean = True _
)

Dim SheetFX As Object, _
    rngPrevSelection As Range, _
    rngNamesFrom As Range, _
    rngValuesFrom As Range, _
    rngJoinFrom  As Range, _
    rngDupeID As Range, _
    rngRemoveRows As Range, _
    ColSeperation As Long, _
    iCell As Range, _
    i As Long, j As Long

Set rngPrevSelection = Selection
Set rngNamesFrom = Get_Column(NamesFrom, rngTable, ExcludeHeader:=True)
Set rngValuesFrom = Get_Column(ValuesFrom, rngTable, ExcludeHeader:=True)
Set rngJoinFrom = Get_Column(JoinFrom, rngTable, ExcludeHeader:=True)
    ColSeperation = rngValuesFrom.Column - rngNamesFrom.Column + 1
 
If TypeName(rngNamesFrom) = "Nothing" Or TypeName(rngValuesFrom) = "Nothing" Then
    Exit Sub 'Matching columns not found
End If

Set SheetFX = Application.WorksheetFunction
'Transpose {rngNamesFrom} to a one-dimensional array, then obtain the unique values
Dim NewColumnNames As Variant: NewColumnNames = SheetFX.Unique(SheetFX.Transpose(rngNamesFrom), True)

    For i = LBound(NewColumnNames) To UBound(NewColumnNames)
        
        With rngNamesFrom
            .EntireColumn.Insert Shift:=xlToRight
            .Cells(1).Offset(-1, -1).Value = NewColumnNames(i)
            'If the header name NewColumnNames(i) matches the value of this row's {rngNamesFrom} value...
            .Offset(0, -1).FormulaR1C1 = "=IF(" & "R1C" & "=" & "RC[1]" & "," & "RC[" & ColSeperation & "]" & ","""")"
            '...assign this row's NewColumnNames(i) value equal to the {rngValuesFrom} value...
            .Offset(0, -1).Value = .Offset(0, -1).Value '...then make cell values static
        End With
        
            'Apply the same formatting to the new column
            rngValuesFrom.Cells(1).EntireColumn.Copy
            rngNamesFrom.Cells(1).EntireColumn.Offset(0, -1).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        
    Next i

'Initialize the deletion range to the empty row beneath the table...
Set rngRemoveRows = rngTable.Cells(rngTable.Count).Offset(1, 0)
'...so that it can be used with Union in the loop below

Dim arrDeleteRows() As Variant

'Order by the unique ID range, {rngJoinFrom}...
Call Order_by(rngJoinFrom, Descending:=False)
    
    '...iterate across the values in the ordered range...
    For Each iCell In rngJoinFrom
        
        '...when a duplicate unique ID from {rngJoinFrom} is detected beneath iCell...
        If iCell.Value <> vbNullString And iCell.Value = iCell.Offset(1, 0).Value Then
            
            '...and the cell above is not a duplicate (which would be counted on the first scan)
            If iCell.Value <> iCell.Offset(-1, 0).Value Then
                
                '...then determine the total count of duplicate entries
                j = 0: Do While iCell.Offset(1 + j, 0).Value = iCell.Value
                           j = j + 1
                       Loop
                        
                       'For each pivoted column...
                       For i = 1 To UBound(NewColumnNames)
                        
                            '...set {rngDupeID} equal...
                            Set rngDupeID = Range( _
                                iCell.Offset(0, rngNamesFrom.Column - 1 - i), _
                                iCell.Offset(j, rngNamesFrom.Column - 1 - i) _
                            ) '...to the span of duplicate rows in the pivoted column...
                            
                            '...join all the text across the rows (there will only be one non-blank value)...
                            rngDupeID.Cells(1).Value = Application.TextJoin(", ", True, rngDupeID)
                            '...and set the first cell in {rngDupeID} equal to all the joined text
                            
                       Next i
                       
                'Add all but the first cell of {rngDupeID} to the an array to hold deletion ranges
                Call ReDim_Add(arrDeleteRows, Range(rngDupeID.Cells(2), rngDupeID.Cells(rngDupeID.Count)).Address)
            
            End If
                
        End If

    Next iCell

'Delete duplicate rows, from the bottom up, one at a time...
For i = UBound(arrDeleteRows) To LBound(arrDeleteRows) Step -1
    '...because list objects prevent bulk deletion
    Range(arrDeleteRows(i)).EntireRow.Select
    Range(arrDeleteRows(i)).EntireRow.Delete
Next i

'Delete pivoted columns
rngNamesFrom.EntireColumn.Hidden = False
rngValuesFrom.EntireColumn.Hidden = False
    
    Call Drop_Columns( _
        rngTable:=rngTable.CurrentRegion, _
        strMatch:=NamesFrom, _
        PerfectMatch:=PerfectMatchOnly _
    )
    Call Drop_Columns( _
        rngTable:=rngTable.CurrentRegion, _
        strMatch:=ValuesFrom, _
        PerfectMatch:=PerfectMatchOnly _
    )

rngPrevSelection.Select
End Sub

Sub ColumnSub( _
    rngColumn As Range, _
    strSubstitute As String, _
    strReplacement As String _
)

Dim arrColVals As Variant, _
    strCombined As String
    
'Load column values into one dimensional array...
arrColVals = Application.Transpose( _
    Range(rngColumn.Cells(2), _
    rngColumn.Cells(rngColumn.Count)) _
) '...excluding the column header

    '...combine all values into single string...
    strCombined = Replace( _
        CStr(Join(arrColVals, "::")), _
        strSubstitute, _
        strReplacement _
    ) '...apply one string replacement...
    
    '...split by deliminator term back into array...
    arrColVals = Split(strCombined, "::")

'...override column values (excluding headers)...
Range(rngColumn.Cells(2), _
    rngColumn.Cells(rngColumn.Count) _
) = Application.Transpose(arrColVals) '...with array substitutions

End Sub

Sub Drop_Columns( _
    rngTable As Range, _
    strMatch As String, _
    Optional PerfectMatch As Boolean = False _
)

Dim rngHeaders As Range, _
    rngDrop As Range
Set rngHeaders = Get_FirstRow(rngTable)
    
LookFor:
Set rngDrop = rngHeaders.Find( _
    What:=strMatch, _
    LookIn:=xlValues, _
    LookAt:=IIf(PerfectMatch = True, xlWhole, xlPart), _
    MatchCase:=False, _
    SearchDirection:=xlNext _
)
    If TypeName(rngDrop) <> "Nothing" Then
        rngDrop.EntireColumn.Delete
        GoTo LookFor 'See if there are any other matches
    End If

End Sub

Sub Set_MinColWidth( _
    rngTable As Range, _
    MinWidth As Single, _
    Optional OverrideAll As Boolean = False _
)

Dim rngHeaders As Range, rngCell As Range
Set rngHeaders = Get_FirstRow(rngTable)

Select Case OverrideAll

    Case False: For Each rngCell In rngHeaders
                    If rngCell.ColumnWidth < MinWidth Then
                        rngCell.ColumnWidth = MinWidth
                    End If
                Next rngCell
        
    Case True:  For Each rngCell In rngHeaders
                    rngCell.ColumnWidth = MinWidth
                Next rngCell
End Select

End Sub

Sub Split_ColumnValues( _
    rngColumn As Range, _
    SplitTerm As String, _
    SplitKeepIndex As Long _
)

Dim ArrayActivity() As Variant: ArrayActivity = rngColumn
    Dim i As Long
    
    For i = LBound(ArrayActivity) To UBound(ArrayActivity)
        ArrayActivity(i, 1) = Split(ArrayActivity(i, 1), SplitTerm)(SplitKeepIndex)
    Next i
rngColumn.Value = ArrayActivity

End Sub

Sub Reorder_Columns( _
    Named As Variant, _
    FromTable As Range, _
    Optional PerfectMatch As Boolean = False, _
    Optional ToLocation As String = "{Start} or {End}" _
)

Dim rngHeaders As Range, _
    rngMove As Range, _
    rngBefore As Range, _
    jShifts As Long, _
    i As Long

For i = LBound(Named) To UBound(Named)
    
    '{rngHeaders} must be reset on each call...
    Set rngHeaders = Get_FirstRow(FromTable.CurrentRegion)
    '...using .CurrentRegion property
    
    Set rngMove = rngHeaders.Find( _
        What:=CStr(Named(i)), _
        LookIn:=xlValues, _
        LookAt:=IIf(PerfectMatch = True, xlWhole, xlPart), _
        MatchCase:=False, _
        SearchDirection:=xlNext _
    )
        
        'If column {Named(i)} exists...
        If TypeName(rngMove) <> "Nothing" Then
            
            '...set {rngBefore} based on {ToLocation}...
            Select Case ToLocation
                Case "Start": Set rngBefore = rngHeaders.Cells(1).Offset(0, jShifts)
                Case "End": Set rngBefore = rngHeaders.Cells(rngHeaders.Count).Offset(0, 1)
            End Select
            
            '...then cut and insert accordingly
            If rngMove <> rngBefore Then
                rngMove.EntireColumn.Cut
                rngBefore.EntireColumn.Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
            
            'To ensure columns are placed next to each other...
                jShifts = jShifts + 1
            '...track each column from {Named()} which did in fact exist
                    
        End If
Next i

rngHeaders.Columns.AutoFit
End Sub

Sub Filter_Dupes( _
    FromColsNamed As Variant, _
    rngTable As Range _
)

Dim TableCopy As Range, _
    rngHeaders As Range, _
    ColName As Variant, _
    rmFromColIndices() As Variant, _
    i As Long
    
'Make a copy of the table, as duplicates...
Set TableCopy = Fast_Copy(rngTable)
'...cannot be filtered from Excel tables
    
    Set rngHeaders = Get_FirstRow(TableCopy)
        
        'Loop through each header...
        For i = 1 To rngHeaders.Count
            
            '...and then compare it's name to the strings specified...
            For Each ColName In FromColsNamed
            
                '...when a header is matched, add it's index to an array
                If InStr(1, CStr(ColName), rngHeaders.Cells(i).Value) Then
                    Call ReDim_Add(rmFromColIndices, i)
                    Exit For
                End If
                
            Next ColName
            
        Next i
    
    'Remove duplicates from the columns specified above
    With ActiveSheet.Range(TableCopy.Address)
        .RemoveDuplicates _
            Columns:=(rmFromColIndices()), _
            Header:=xlYes
    End With

'Overwrite the current table with the newly filtered table
Call Overwrite_Table( _
    tblCurrent:=rngTable.CurrentRegion, _
    tblNew:=TableCopy _
)

End Sub

Function Fast_Copy( _
    rngToCopy As Range, _
    Optional rngOutput As Range _
)

If TypeName(rngOutput) = "Nothing" Then
    'Default {rngOutput} to the right of the original range + 1...
    Set rngOutput = rngToCopy.Offset(0, Get_FirstRow(rngToCopy).Count + 1)
    '...so that an Excel table does not automatically attach the range
End If

    With rngToCopy.AdvancedFilter( _
        Action:=xlFilterCopy, _
        CopyToRange:=rngOutput _
    ): End With
        
        Set Fast_Copy = rngOutput

End Function

Function Overwrite_Table( _
    tblCurrent As Range, _
    tblNew As Range _
) 'Note: Column dimensions must be the same (Excel table compatability)

'Ensure entire table regions are captured
Set tblCurrent = tblCurrent.CurrentRegion
Set tblNew = tblNew.CurrentRegion

Dim RowDif As Long: RowDif = tblCurrent.Cells(tblCurrent.Count).Row - tblNew.Cells(tblNew.Count).Row
Dim rngLastRow As Range: Set rngLastRow = Get_LastRow(tblCurrent)

If RowDif > 0 Then
    
    'Clear the rows which are no longer present in tblNew
    Range(rngLastRow.Offset(-RowDif + 1, 0), rngLastRow).Clear
        
        'If tblCurrent is a list object...
        If TypeName(tblCurrent.Cells(1).ListObject) <> "Nothing" Then
            
            '...also resize the Excel table object
            With ActiveSheet.ListObjects(tblCurrent.Cells(1).ListObject.Name)
                .Resize Range(tblCurrent.Cells(1), tblCurrent.Cells(tblCurrent.Count).Offset(-RowDif, 0))
            End With
        End If
End If

'Copy {tblNew} onto resized {tblCurrent} and clear {tblNew}
Set Overwrite_Table = Fast_Copy( _
    rngToCopy:=tblNew.CurrentRegion, _
    rngOutput:=tblCurrent.CurrentRegion _
)

tblNew.Clear
End Function

Function Get_FirstRow(rngFrom As Range)
    Set Get_FirstRow = Range(rngFrom.Cells(1), Cells(rngFrom.Cells(1).Row, rngFrom.Cells(rngFrom.Count).Column))
End Function

Function Get_LastRow(rngFrom As Range)
    Set Get_LastRow = Range(Cells(rngFrom.Cells(rngFrom.Count).Row, rngFrom.Cells(1).Column), rngFrom.Cells(rngFrom.Count))
End Function

Function Get_LastColumn(rngFrom As Range)
    Set Get_LastColumn = Range(Cells(rngFrom.Cells(1).Row, rngFrom.Cells(rngFrom.Count).Column), rngFrom.Cells(rngFrom.Count))
End Function

'===============================================================================================================================================================================================================================================================
'##  User Interface
'===============================================================================================================================================================================================================================================================

'NOTES:

'FaceIds: https://bettersolutions.com/vba/ribbon/face-ids-2003.htm
'FaceId = 1 is blank

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### MISC / FOR EXAMPLES
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

Sub MsgLines(MyText As String, Optional Repeat As Integer = 1)
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

Sub IdentifyMenus(Optional RemoveIndicators As Boolean = False)

Dim i As Long

For i = 1 To Application.CommandBars.Count
    On Error GoTo CannotAddToMenu
        If ToggleDirection = True Then
            With CommandBars(i).Controls.Add(before:=1)
               .Caption = "This is CommandBar(" & i & ")"
               .FaceId = 343
               .BeginGroup = True
            End With
        Else
            CommandBars(i).Reset
        End If
    GoTo MenuAdded
    
CannotAddToMenu:
On Error GoTo -1
Call Print_Named("Cannot add an item to " & Chr(34) & Application.CommandBars(i).Name & Chr(34) & " (menu " & i & ")")

MenuAdded:
Next i

End Sub

Function ResetCellMenu()
    CommandBars("Cell").Reset
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### Create Menu Command
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
Sub sssss()

Print_Named CommandBars("Cell").Index
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

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### Create Menu Command Section
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### Create Popup Menu
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### Create Add-Ins Ribbon Buttons
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### Create Button Shape
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Try_CreateButtonShape()
Range("A1").Select

MsgBox "Creating a [Blank Button] shape that does nothing:", vbInformation, "[Blank Button]"
    Call CreateButtonShape

MsgBox "Creating a button, calling it [Button One], and assigning some properties", vbInformation, "[Button One]"
    Call CreateButtonShape( _
        btnLabel:="Button One", _
        StrCommand:="MsgLines{'This is a message'}", _
        btnColor:=5242976, _
        Top:=40 _
    )

MsgBox "Creating a button, calling it [Button Two], assigning some properties, and using it in a sub...", vbInformation, "[Button Two]"
    Dim Button As Object
    Set Button = _
        CreateButtonShape( _
            btnLabel:="Button Two", _
            btnName:="btnTwo", _
            StrCommand:="MsgLines{MyText:='This is a message', Repeat:=4}", _
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
'## Extras
'===============================================================================================================================================================================================================================================================
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'### Legal Special Character Reference
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' ¶ € § Ø µ ª ° ¹ ² ³ · • ¿ ¡ ƒ × ¤ » « ‡ ¦ ± ÷ ¨ ¯ — ¬

'https://homepage.cs.uri.edu/faculty/wolfe/book/Readings/R02%20Ascii/completeASCII.htm


