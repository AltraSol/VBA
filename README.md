
#  SimplifyVBA


A collection of functions to interface R with VBA, add functionality to Excel, or improve VBA debugging and readability.

Note: macOS compatability is still being tested


##  Functions


Prefix: ƒ— denotes a function which has a notable load time or file interactions outside ThisWorkbook. Only use within VBA.

``` VBA
Public GlobalUser As String

'   Prevents needless rerunning of the file search component of
'   Get_WindowsUsername() once the local user has been determined.

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
  ExtractFirstInt_RightToLeft (aVariable)

'   Returns the first integer found in a string when searcing
'   from the right end of the string to the left.

'   ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"

```
``` VBA
  ExtractFirstInt_LeftToRight (aVariable)

'   Returns the first integer found in a string when searcing
'   from the left end of the string to the right.

'   ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"

```
``` VBA
  Truncate_Before_Int (aString)

'   Removes characters before first integer in a sequence of characters.

'   Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"

```
``` VBA
  Truncate_After_Int (aString)

'   Removes characters after first integer in a sequence of characters.

'   Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"

```
``` VBA
  IsInt_NoTrailingSymbols (aNumeric)

'   Checks if supplied value is both numeric, and contains no numeric
'   symbols (different from IsNumeric).

'   IsInt_NoTrailingSymbols(9999) = True
'   IsInt_NoTrailingSymbols(9999,) = False

```
``` VBA
  MyOS()

'   "Windows",  "Mac", or "Neither Windows or Mac".

```
``` VBA
  Get_WindowsUsername()

'   Loops through folders to find paths matching C:\Users\...\AppData
'   then extracts the User from correct path. Superior to reading
'   .FullName of workbook which does not work for OneDrive.

```
``` VBA
  Get_MacUsername()

'   Reads ActiveWorkbook.FullName property to get Mac user.

```
``` VBA
  Get_Username()

'   Returns username regardless of Windows or Mac OS.

```
``` VBA
  Get_DesktopPath()

'   Returns Mac or Windows desktop directory (even if on OneDrive).

```
``` VBA
  Get_DownloadsPath()

'   Returns Mac or Windows downloads directory (even if on OneDrive).

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
  Clipboard_Load(ByVal aString As String)

'   Stores {aString} in clipboard.

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
 Clipboard_IsRange()

'   Returns True if a range is currently copied.

```
``` VBA
 PlatformFileSep()

'   Returns "\" or "/" depending on the operating system.

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
 ListFiles(FromFolder As String)

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
'   tabs in ThisWorkbook, then returnsan array of the new sheet names.

    Dim PastedSheets(): PastedSheets() = PasteSheetVals_FromFolder(...)
    Sheets(PastedSheets(1)).Activate

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
 Replace_SpecialChars( _
     YourString As String, _
     Replacement As String, _
     Optional ReplaceAll As Boolean, _
     Optional TrimWS As Boolean _
 )

'   Replaces `!@#$%^&“”*(")-=+{}\/?:;'.,<> from {YourString} with
'   {Replacement}.

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
 MergeAndCombine(MergeRange As Range, Optional SepValsByNewLine = True)

'    Concatenates each Cell.Value in a range & merges range as opposed
'    to Merge & Center which only keeps a single value

```
``` VBA
 AutoAdjustZoom(rngBegin, rngEnd)

'   Adjusts user view to the width of rngBegin to rngEnd

```
``` VBA
 LaunchLink (aLink)

'   Launches aLink in existing browser with error handling for
'   invalid Links

```
``` VBA
 InsertSlicer(NamedRange As String,
              NumCols As Integer,
              aHeight As Double,
              aWidth As Double)

'   Creates a slicer for the active sheet named range {NamedRange}
'   with {NumCols} buttons per slicer row, and with dimensions
'   {aHeight} by {aWidth}

```
``` VBA
  AlterSlicerColumns(SlicerName As String, NumCols)

'   Loops through workbook to find {SlicerName} and sets the number
'   of buttons per row to {NumCols}

```
``` VBA
  MoveSlicer(SlicerSelection,
            rngPaste As Range,
            leftOffset,
            IncTop)

'   Takes Selection as {SlicerSelection}, cuts & pastes it to a rough
'   location {rngPaste} to be incrementally adjusted from paste
'   location by {leftOffset} and {IncTop}

```
``` VBA
  ToggleDisplayMode()

'   Toggles display of ribbon, formula bar, status bar & headings

```
``` VBA
  Print_Pad()

'   Uses Debug.Print to print a timestamped seperator of "======"

```
``` VBA
  Print_Named(Something, Optional Label)

'   Uses Debug.Print to add a space between each {Something} printed,
'   labels each {Something} if {Label} supplied.

```

##  User Interface Additions


#### The following sub must be placed within ThisWorkbook:

     Private Sub Workbook_BeforeClose(Cancel As Boolean)
         Call Remove_TempMenuCommands
         Call Remove_TempMenuCommandsections
     End Sub

``` VBA
Public GlobalTempMenuCommands() As Variant
Public GlobalTempMenuSections() As Variant

'    Tracks menu commands or menu sections that have been added using
'    the CreateMenuCommand() or CreateMenuSection() commands with a
'    Temporary:=True property. Allows for the deletion of all user
'    created menus or menu items on the Workbook_BeforeClose() event.

```
``` VBA
Sub CreateAddInButtons( _
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
'
'    Adds each command in {ButtonStrCommands_Array}
'    to the section with properties as specified in {ButtonTypes_Array},
'    {MenuFaceIDs_Array} and {ButtonNames_Array}. Each {..._Array}
'    parameter must be of equal length, but the item of {MenuFaceIDs_Array}
'    will be ignored if the corresponding element of {ButtonTypes_Array} is
'    2 given that it's a caption only display type.

EXAMPLES: '(Ctrl+f to view & run)
     Sub Try_CustomToolbarsRow_Caption()
     Sub Try_CustomToolbarsRow_Icons()
     Sub Try_CustomToolbarsRow_CaptionIcon()
```

##  RSCRIPT


### TODO: Remove notification of deletion

``` VBA
  QuickRun_RScript(ByVal ScriptContents As String)

'   Writes a temporary .R script containing {ScriptContents}, runs
'   it, prompts for the deletion of the temporary script

```
``` VBA
  WriteTemp_RScript(ByVal ScriptContents As String)

'   Creates a random named temporary folder on desktop, creates an
'   .R file "Temp.R" containing {ScriptContents}, returns Temp.R path

```
``` VBA
  FindAndRun_RScript(ByVal ScriptLocation)

'   Takes a string or cell reference {RScriptPath} & runs it on the
'   latest version of R on the OS

```
``` VBA
  Run_RScript(ByVal RLocation As String, _
              ByVal ScriptLocation As String, _
              Optional ByVal Visibility As String)

'   Uses the RScript.exe pointed to by {RLocation} to run the script
'   found at {ScriptLocation}. Rscript.exe window displayed by default,
'   but {Visibility}:= "VeryHidden" or "Minimized" can be used

```
``` VBA
  Get_RExePath() As String

'   Returns the path to the latest version of Rscript.exe

```
``` VBA
  Get_LatestRVersion(ByVal RVersions As Variant)

'   Returns the latest version of R currently installed

```
``` VBA
  Get_RVersions(ByVal RFolderPath As String)

'   Returns an array of the R versions currently installed

```
``` VBA
  Get_RFolder() As String

'   Returns the parent R folder path which houses the installed
'   versions of R on the OS from which the sub is called

```
``` VBA
  Test_QuickRun_RScript()

'   Writes a computationally intensive script to Desktop and asks
'   if you want to run it (to visually verify all zRun_R f(x) worked)

```
