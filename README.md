
#  QuickStartVBA ¬ github.com/ulchc (10-29-22)


## Overview


A collection of generic functions and subs for every action I had to implement in VBA more than once.

Prefix ƒ— denotes a function which has a notable load time or file interactions
outside ThisWorkbook. Since my intended use of the QuickStartVBA module/repo was to quickly
port in many potentially useful snippets of code, then build onto a secondary module for 
a specific use case, I've opted to use this uncommon character prefix instead of using 
Private Functions so that functions are available in any module.

#### Sections
  * [Functions](#functions)
  * [Subs](#subs)
  * [Data Transformation](#data-transformation)
  * [User Interface](#user-interface)

#### See Also
* [RscriptVBA](https://github.com/ulchc/RscriptVBA)

##  Important


If you intend to use the User Interface section, the following sub must be placed within ThisWorkbook:

``` VBA
Private Sub Workbook_BeforeClose(Cancel As Boolean)
   Call Remove_TempMenuCommands
   Call Remove_TempMenuCommandSections
End Sub
```
``` VBA
Public GlobalTempMenuCommands() As Variant
Public GlobalTempMenuSections() As Variant

'    Tracks menu commands or menu sections that have been added using
'    the CreateMenuCommand() or CreateMenuSection() commands with a
'    Temporary:=True property. Allows for the deletion of all user
'    created menus or menu items on the Workbook_BeforeClose() event.

```

##  Functions

``` VBA
  Get_Username()

'   Returns username by reading the environment variable.

```
``` VBA
  Get_DesktopPath()

'   Returns the desktop path regardless of platform with handling
'   for OneDrive hosted desktops.

```
``` VBA
  Get_DownloadsPath()

'   Returns the desktop path regardless of platform.

```
``` VBA
 Get_LatestFile( _
     FromFolder As String, _
     MatchingString As String, _
     FileType As String _
 )

'   Returns the latest file of the specified {FileType} with a name
'   that includes {MatchingString} from the directory {FromFolder}.

```
``` VBA
 ListFiles(FromFolder As String)

'   Returns an array of all file paths located in {FromFolder}

```
``` VBA
 ListFolders(FromFolder As String)

'   Returns an array of all file paths located in {FromFolder}

```
``` VBA
 CopySheets_FromFolder( _
     FromFolder As String, _
     Optional Copy_xlsx As Boolean, _
     Optional Copy_xlsm As Boolean, _
     Optional Copy_xls As Boolean, _
     Optional Copy_csv As Boolean _
 )
'   Opens all file types specified by the boolean parameters in the
'   directory {FromFolder}, copies all sheets to ThisWorkbook, then
'   returns an array of the new sheet names.

    Dim CopiedSheets(): CopiedSheets() = CopySheets_FromFolder(...)
    Sheets(CopiedSheets(1)).Activate

```
``` VBA
 PasteSheetVals_FromFolder( _
     FromFolder As String, _
     Optional Copy_xlsx As Boolean, _
     Optional Copy_xlsm As Boolean, _
     Optional Copy_xls As Boolean, _
     Optional Copy_csv As Boolean _
 )

'   Opens all file types specified by the boolean parameters in the
'   directory {FromFolder}, pastes cell values from each sheet to new
'   tabs in ThisWorkbook, then returns an array of the new sheet names.

    Dim PastedSheets(): PastedSheets() = PasteSheetVals_FromFolder(...)
    Sheets(PastedSheets(1)).Activate

```
``` VBA
 Clipboard_IsRange()

'   Returns True if a range is currently copied.

```
``` VBA
  Clipboard_Load(ByVal YourString As String)

'   Stores {YourString} in clipboard.

```
``` VBA
 ƒ—Clipboard_Read( _
     Optional IfRngConcatAllVals As Boolean = True, _
     Optional Sep As String = ", " _
 )

'   Returns text from the copied object (clipboard text or range).

```
``` VBA
  ƒ—Get_CopiedRangeVals()

'   If range copied, checks each Cell.Value in the range and
'   returns an array of each non-blank value.

```
``` VBA
 CopySheets_FromFile(FromFile As String)

'   Opens {FromFile}, copies all sheets within it to ThisWorkbook,
'   then returns an array of the new sheet names.

    Dim CopiedSheets(): CopiedSheets() = CopySheets_FromFile(...)
    Sheets(CopiedSheets(1)).Activate

```
``` VBA
 PasteSheetVals_FromFile(FromFile As String)

'   Opens {FromFile}, pastes cell values from all sheets within it
'   to ThisWorkbook, then returns an array of the new sheet names.

    Dim PastedSheets(): PastedSheets() = PasteSheetVals_FromFile(...)
    Sheets(PastedSheets(1)).Activate

```
``` VBA
 Get_FilesMatching( _
     FromFolder As String, _
     MatchingString As String, _
     FileType As String _
 )

'   Returns an array of file paths located in {FromFolder} which have
'   a file name containing {MatchingString} and a specific {FileType}.

```
``` VBA
 RenameSheet( _
     CurrentName As String, _
     NewName As String, _
     OverrideExisting As Boolean _
 )

'   Changes Sheets({CurrentName}).Name to {NewName} if {NewName}
'   is not already in use, otherwise, a bracketed number (n) is added
'   to {NewName}. The final name of the renamed sheet is returned.

'   If {OverrideExisting} = True and a sheet with the name {NewName}
'   exists, it will be deleted and Sheets({CurrentName}).Name will
'   always be set to {NewName}.

```
``` VBA
 Tabs_MatchingCodeName( _
     MatchCodeName As String, _
     ExcludePerfectMatch As Boolean _
 )

'   An array of tab names where {MatchCodeName} is within the CodeName
'   property (useful for detecting copies of a code-named template).

```
``` VBA
 WorksheetExists( _
     aName As String, _
     Optional wb As Workbook _
 )

'   True or False dependent on if tab name {aName} already exists.

```
``` VBA
 ƒ—Delete_FileAndFolder(ByVal aFilePath As String) as Boolean

'   Use with caution. Deletes the file supplied {aFilePath}, all
'   files in the same folder, and the directory itself.

'   Will exit the deletion procedure if {aFilePath} is a file
'   within the Desktop or Documents directory, or if the directory
'   is considered high level (it's within the user directory).

```
``` VBA
 PlatformFileSep()

'   Returns "\" or "/" depending on the operating system.

```
``` VBA
  MyOS()

'   Read the system environment OS variable and returns "Windows",
'   "Mac", or the unaltered Environ("OS") output if neither.

```
``` VBA
NOTE: Windows only (uses CreateObject("VBScript.RegExp"))

 Replace_SpecialChars( _
     YourString As String, _
     Replacement As String, _
     Optional ReplaceAll As Boolean, _
     Optional TrimWS As Boolean _
 )

'   Replaces `!@#$%^&“”*(")-=+{}\/?:;'.,<> from {YourString} with
'   {Replacement}.

```
``` VBA
NOTE: Windows only (uses CreateObject("VBScript.RegExp"))

 Function Replace_Any( _
     Of_Str As String, _
     With_Str As String, _
     Within_Str As String, _
     Optional TrimWS As Boolean _
 )

'   Replaces all characters {Of_Str} in the supplied {Within_Str}.
'   Distinct from VBA's Replace() in that all matched characters
'   are removed instead of perfect matches.

    Debug.Print Replace_Any(" '. ", "_", "Here's an example.")

```
``` VBA
 ExtractFirstInt_RightToLeft (aVariable)

'   Returns the first integer found in a string when searcing
'   from the right end of the string to the left.

    ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"

```
``` VBA
 ExtractFirstInt_LeftToRight (aVariable)

'   Returns the first integer found in a string when searcing
'   from the left end of the string to the right.

    ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"

```
``` VBA
 Truncate_Before_Int (YourString)

'   Removes characters before first integer in a sequence of characters.

    Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"

```
``` VBA
 Truncate_After_Int (YourString)

'   Removes characters after first integer in a sequence of characters.

    Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"

```
``` VBA
 IsInt_NoTrailingSymbols (aNumeric)

'   Checks if supplied value is both numeric, and contains no numeric
'   symbols (different from IsNumeric).

'   IsInt_NoTrailingSymbols(9999) = True
'   IsInt_NoTrailingSymbols(9999,) = False

```

## Subs

``` VBA
 ReDim_Add(ByRef aArr() As Variant, ByVal aVal)

'    Simplifies the addition of a value to a one dimensional array by
'    handling the initalization & resizing of an array in VBA

     Call ReDim_Add(aArr(), aVal) '-> last element of aArr() now aVal

```
``` VBA
 ReDim_Rem(ByRef aArr() As Variant)

'    Simplifies the sequential removal of the last element of a one
'    dimensional array by handing the resizing of the array as well
'    as the removal of the 0th value

     Call ReDim_Rem(aArr()) '-> last element of aArr() has been removed

```
``` VBA
 SaveToDownloads( _
    SaveTabNamed As String, _
    AsFileNamed As String, _
    OpenAfterSave As Boolean, _
    Optional SaveAsType As String = "xlsx" _
 )

'    {SaveTabNamed} is the ActiveSheet.Name property, {AsFileNamed}
'    is a plain string which is automatically combined with the local
'    download folder to create the full path to save to.

'    {SaveAsType} can be "xlsx", "xlsm", "xlsb", or "csv". A bracketed
'    (n) will automatically be added to the file name if it is
'    already taken.

```
``` VBA
 SaveToDownloads_Multiple( _
    SaveTabsNamed_Array As Variant, _
    AsFileNamed As String, _
    OpenAfterSave As Boolean, _
    Optional SaveAsType As String = "xlsx" _
 )

'    Operates the same as SaveToDownloads() but takes an array of
'    tab names.

```
``` VBA
 MergeAndCombine(MergeRange As Range, Optional SepValsByNewLine = True)

'    Concatenates each Cell.Value in a range & merges range as opposed
'    to Merge & Center which only keeps a single value

```
``` VBA
 AutoAdjustZoom(rngBegin As Range, rngEnd As Range)

'   Adjusts user view to the width of rngBegin to rngEnd

```
``` VBA
 LaunchLink (aLink)

'   Launches aLink in existing browser with error handling for
'   invalid Links

```
``` VBA
 ToggleDisplayMode()

'   Toggles display of ribbon, formula bar, status bar & headings

```
``` VBA
 CreateSlicer( _
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

'   Uses {tblKeyAddress} to determine the ListObject name,
'   creates a slicer for {ColumnName}, and then aligns it with the
'   cell specified by {HorizAlignAddress}.

'   Aligns the slicer with the top right corner of the cell when
'   {HorizAlignRight} = True and the top left corner when
'   {HorizAlignRight} = False.

```
``` VBA
 HorizAlignShape( _
     ShapeObject As Object, _
     AlignToRange As Range, _
     RightAlign As Boolean _
 )

'   Written for use in CreateSlicer(), but can be used to
'   skip the calculations involved to right align any shape.

```
``` VBA
 Get_AvenirSlicerStyle()

'   Creates the .TableStyle "AvenirSlicerStyle" for CreateSlicer().

```
``` VBA
 TableStyleExists(StyleNamed As String)

'   Returns True or False depending on if .TableStyle({StyleNamed})
'   exists.

```
``` VBA
 Create_Comment( _
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

'   Adds a comment to {rngComment} that has a cleaner look than
'   the base Excel comment, with each item of {arrTextLines} written
'   to the comment as a line of text seperated by a new line character,
'   and optional arguments to make changing the comment's properties
'   less convoluted.

'   More notably, automatically adjusts the dimensions of image
'   comments to match the aspect ratio of the image, and enables the
'   use of format strings to bolden specific sections of text in
'   {arrTextLines}.

'   If {arrTextLines} = Array("#Header#", "Point 1", "Point 2")

'   The comment would show as:

'   Header (bold)
'   Point 1
'   Point 2

```
``` VBA
 Get_AspectRatio(ImgPath As String)

'   Written for use in Create_Comment(), but will return the aspect
'   ratio (width / height) for any image.

```
``` VBA
 PrintEnvironVariables()

'   Print the environment variables to the Immediate window.

```
``` VBA
 Print_Named(Something, Optional Label)

'   Uses Debug.Print to add a space between each {Something} printed,
'   labels each {Something} if {Label} supplied.

```
``` VBA
 Print_Pad()

'   Uses Debug.Print to print a timestamped seperator of "======"

```

## Data Transformation

``` VBA
 Filter_By( _
     rngColumn As Range, _
     AdvFilterTerm As String, _
     Optional rngTable As Range _
 )

'   Removes filtered terms from either a range or ListObject by
'   copying the .CurrentRegion of {rngColumn} (unless {rngTable}
'   is specified), using .AdvancedFilter on the copy with the
'   {AdvFilterTerm} applied to {rngColumn}, then overwriting the
'   previous range or ListObject with the filtered result.

```
``` VBA
 Order_by( _
     rngColumn As Range, _
     Optional Descending As Boolean = True _
 )

'   Self-explanatory simplification of .Sort on a table.

```
``` VBA
 Pivot_Wider( _
     rngTable As Range, _
     NamesFrom As String, _
     ValuesFrom As String, _
     JoinFrom As String, _
     Optional PerfectMatchOnly As Boolean = True _
 )

'   Adds new columns to {rngTable} by seperating each category in
'   column {NamesFrom} into it's own column, with column values
'   obtained from the column named {ValuesFrom}.

'   If a column name should be approximately matched, {PerfectMatchOnly}
'   can be set equal to False.

'   After pivoting categories into columns, if it is found that there
'   are mutiple rows for a single value of the column named {JoinFrom},
'   the category values are consolidated into a single row and the
'   duplicate rows are removed.

'   Works similarily to dplyr pivot_wider() in R, with the original
'   columns removed after the pivot.

```
``` VBA
 ColumnSub( _
     rngColumn As Range, _
     strSubstitute As String, _
     strReplacement As String _
 )

'   Subsitutes each occurance of {strSubstitute} with {strReplacement}
'   for all values in {rngColumn}.

```
``` VBA
 Drop_Columns( _
     rngTable As Range, _
     strMatch As String, _
     Optional PerfectMatch As Boolean = False _
 )

'   Deletes any column with a header matching {strMatch} in {rngTable},
'   with optional parameter {PerfectMatch} to adjust match precision.

```
``` VBA
 Set_MinColWidth( _
     rngTable As Range, _
     MinWidth As Single, _
     Optional OverrideAll As Boolean = False _
 )

'   Adjusts columns widths of {rngTable} to at minimum be {MinWidth}
'   wide, with optional parameter {OverrideAll} to reset all widths
'   to {MinWidth}.

```
``` VBA
 Split_ColumnValues( _
     rngColumn As Range, _
     SplitTerm As String, _
     SplitKeepIndex As Long _
 )

'   Splits the values in {rngColumn} by {SplitTerm} and substitutes
'   column values with the split index specified: {SplitKeepIndex}.

```
``` VBA
 Reorder_Columns( _
     Named As Variant, _
     FromTable As Range, _
     Optional PerfectMatch As Boolean = False, _
     Optional ToLocation As String = "{Start} or {End}" _
 )

'   Rearranges a subset of columns specified in the array {Named}
'   by the order they were supplied, either to the "Start" or "End"
'   of the .CurrentRegion of the table, with optional parameter
'   {PerfectMatch} to adjust match precision of column names.

'   Note: {ToLocation} is not Optional. The default value is simply
'   a means to make the choice values visible when calling the sub.

```
``` VBA
 Filter_Dupes( _
     FromColsNamed As Variant, _
     rngTable As Range _
 )

'   Creates a non-ListObject copy of {rngTable}, filter duplicates
'   across all the columns specified in {FromColsNamed}, then
'   overwrites the original table with the filtered result.

```
``` VBA
 Overwrite_Table( _
     tblCurrent As Range, _
     tblNew As Range _
 )

'   Overwrites {tblCurrent} with {tblNew} and resizes the ListObject
'   that linked with {tblCurrent} if applicable.

'   Note: Column dimensions must be the same (Excel table compatability)

```
``` VBA
 Fast_Copy( _
     rngToCopy As Range, _
     Optional rngOutput As Range _
 )

'   Returns {rngOutput} after using .AdvancedFilter to copy
'   {rngToCopy} to the right of itself. If [rngOutput} is specified,
'   the default output location will be overridden.

```
``` VBA
 Get_FirstRow(rngFrom As Range)

'   Returns the first row of a given range.

```
``` VBA
 Get_LastRow(rngFrom As Range)

'   Returns the last row of a given range.

```
``` VBA
 Get_LastColumn(rngFrom As Range)

'   Returns the last column of a given range.

```

##  User Interface

``` VBA
 ConvertStrCommand( _
     CommandString As String, _
     Optional Verbose As Boolean = True _
 )

'   Automatically applied to all {StrCommand}'s passed to the menu
'   and button creation functions below (prior to linking the
'   macro to the object).

'   Changes existing apostrophes in {StrCommand} to quotation marks,
'   encases the full command in apostrophes, and substitutes curly
'   braces for brackets.

'   This is to make it easier to supply parameters to a sub or function
'   called by a menu or shape without having to include a long list of values
'   seperated with Chr(34) & "..." & Chr(34).

'   Original:   "MySub(Range{'NamedRange'}, 2)"
'   Converted: "'MySub(Range("NamedRange"), 2)'"

```
``` VBA
 ChangeMenuVisibility( _
     MenuItems_Array As Variant, _
     VisibleProperty As Boolean _
 )

'   Toggles the visibility of items on the menu shown by right
'   clicking a cell. For situations where the menu is becoming
'   overcrowded with custom commands.

'   Menu items can be refered to with the same string as their
'   display names *except* in the case of an underlined letter,
'   in which case, the true name includes an &. For example,
'   "Copy" is actually "&Copy".

'   All visibility modifications can be returned to default by
'   calling ResetCellMenu()

```
``` VBA
 ResetCellMenu()

'   Restores CommandBars("Cell") and ShortcutMenus(xlWorksheetCell)
'   to their default states.

```
``` VBA
 CreateMenuCommand( _
    MenuCommandName As String, _
    StrCommand As String, _
    Optional Temporary As Boolean = True, _
    Optional MenuFaceID As Long _
 )
PARAMETERS:
'    {MenuCommandName} = The name of the menu that will be created.
'    {StrCommand} = Command to run when clicked (see ConvertStrCommand).
'    {Temporary} = Whether the menu should be deleted on the WorkbookClose event.
'    {MenuFaceID} = The FaceId integer for the menu command.

EXPLANATION:
'    Adds an item to the top of the menu displayed when right clicking
'    a cell on a worksheet. Shows up with the caption {MenuCommandName}
'    and the icon specified with {MenuFaceID}.

'    Call RemoveMenuCommand(...) to remove

EXAMPLES: '(Ctrl+f to view & run)
     Sub Try_CreateMenuCommand()
```
``` VBA
 CreateMenuSection( _
    MenuSectionName As String, _
    Array_SectionMenuNames As Variant, _
    Array_StrCommands As Variant, _
    Optional Temporary As Boolean = True _
 )
PARAMETERS:
'    {MenuSectionName} = The name of the menu section that will be created.
'    {Array_SectionMenuNames} = Array of display names for each command.
'    {Array_StrCommands} = Array of commands for each menu item (see ConvertStrCommand).
'    {Temporary} = Whether the menu should be deleted on the WorkbookClose event.

EXPLANATION:
'    Adds a section to the top of the menu displayed when right clicking
'    a cell on a worksheet. Shows up with the caption {MenuSectionName}
'    and no icon.

'    When hovering over the menu section, a list of commands
'    specified by {Array_SectionMenuNames} will display, each running
'    it's corresponding macro specified in {Array_StrCommands}.

'    Call RemoveMenuSection(...) to remove

EXAMPLES: '(Ctrl+f to view & run)
     Sub Try_CreateMenuSection()
```
``` VBA
NOTE: Popup menus are Windows only

 CreatePopupMenu( _
    PopupMenuName As String, _
    Array_ItemNames As Variant, _
    Array_StrCommands As Variant, _
    Array_ItemFaceIDs As Variant, _
    Optional Temporary As Boolean = True _
 )
PARAMETERS:
'    {PopupMenuName} = The name of the menu that will be created.
'    {Array_ItemNames} = Array of display names for each command.
'    {Array_StrCommands} = Array of commands for each menu item (see ConvertStrCommand).
'    {Array_ItemFaceIDs} = Array of FaceId integers for each menu item.
'    {Temporary} = Whether the menu should be deleted on the WorkbookClose event.

EXPLANATION:
'    Creates a custom menu named {PopupMenuName} which can be displayed with
'    Application.CommandBars({PopupMenuName}).ShowPopup.

'    Each item from {Array_ItemNames} is included in the menu with
'    it's corresponding integer FaceId specified by {Array_ItemFaceIDs}.
'    When an item is clicked, it runs it's respective macro as
'    assigned by {Array_StrCommands}.

'    Call RemovePopupMenu(...) to remove

EXAMPLES: '(Ctrl+f to view & run)
     Sub Try_CreatePopupMenu()
     Sub Try_CreatePopupMenuColorful()
```
``` VBA
 CreateAddInButtons( _
    ButtonSectionName As String, _
    ButtonNames_Array As Variant, _
    ButtonTypes_Array As Variant, _
    ButtonStrCommands_Array As Variant, _
    Optional MenuFaceIDs_Array As Variant, _
    Optional Temporary As Boolean = True _
 )
PARAMETERS:
'    {ButtonSectionName} = Name of the row added to the Add-ins ribbon (visible on hover).
'    {ButtonNames_Array} = Array of names for each command (visible on hover).
'    {ButtonTypes_Array} = Array of types (1, 2 or 3) for the display of command buttons.
'    {ButtonStrCommands_Array} = Array of commands for each button (see ConvertStrCommand).
'    {MenuFaceIDs_Array} = Array of FaceId numbers (only applicable to ButtonTypes 1 and 3).
'    {Temporary} = Specifies whether the Add-ins section will automatically be removed when workbook closes.

EXPLANATION:
'    Creates a row of commands within the "Custom Toolbars" section
'    of the Add-ins ribbon and Debug.Prints the details.

'    Adds each command in {ButtonStrCommands_Array}
'    to the section with properties as specified in {ButtonTypes_Array},
'    {MenuFaceIDs_Array} and {ButtonNames_Array}. Each {..._Array}
'    parameter must be of equal length, but the item of {MenuFaceIDs_Array}
'    will be ignored if the corresponding element of {ButtonTypes_Array} is
'    2 given that it's a caption only display type.

     Call RemoveAddInSection(...) to remove

EXAMPLES: '(Ctrl+f to view & run)
     Sub Try_CreateAddInButtons_Type1()
     Sub Try_CreateAddInButtons_Type2()
     Sub Try_CreateAddInButtons_Type3()
```
``` VBA
 CreateButtonShape( _
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
PARAMETERS:
'    {StrCommand} = Commands to run when clicked (see ConvertStrCommand).
'    {btnLabel} = The display name of the shape.
'    {btnName} = The .Name property of the shape.
'    {ShapeType} = The look of the shape as specified by the integer type.
'    {btnColor} = Color code of the shape
'    {Lef} = .Left property of the shape
'    {Top} = .Top property of the shape
'    {Wid} = .Width property of the shape
'    {Hei} = .Hweight property of the shape

EXPLANATION:
'    Inserts a shape onto the sheet that has the appears of a button
'    and runs {StrCommand} when clicked.

EXAMPLES: '(Ctrl+f to view & run)
     Sub Try_CreateButtonShape()
```
``` VBA
 IdentifyMenus(Optional RemoveIndicators As Boolean = False)

'   Loops through each CommandBar in the workbook and adds a
'   new indicator command *This is CommandBar(i)* to the top
'   of the menu so that the index of the menu can be identified.

'   This is simply to enable the modification of CommandBars other
'   than the worksheet cell menus (ex. ListObject), which aren't
'   often named in an intuitive way.

```
