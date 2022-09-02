# VBA
Download VBA-All.xlsm for improved readability.

## Module: zPortable_Subs.bas
Portable module of subs which can be exported to any workbook and are only dependent on one-another (if at all)

```
 ReDim_Add(ByRef aArr() As Variant, ByVal aVal)

   Simplifies the addition of a value to a one dimensional array by
   handling the initalization & resizing of an array in VBA

   Call ReDim_Add(aArr(), aVal) -> last element of aArr() now aVal

```
```
 MergeAndCombine(MergeRange As Range, _
                 Optional SepValsByNewLine = True)

   Concatenates each Cell.Value in a range & merges range as opposed
   to Merge & Center which only keeps a single value

```
```
 MenuAdd_MergeAndCombine()

   Adds "Merge and Combine" menu option to cell right-click menu
   Note: Calls Sub "MergeAndCombine_Selection"

```
```
 MenuDelete_MergeAndCombine()

   Deletes "Merge and Combine" menu option

```
```
 AutoAdjustZoom(rngBegin, rngEnd)

   Adjusts user view to the width of rngBegin to rngEnd

```
```
 LaunchLink (aLink)

   Launches aLink in existing browser with error handling for
   invalid Links

```
```
 InsertSlicer(NamedRange As String,
              NumCols As Integer,
              aHeight As Double,
              aWidth As Double)

   Creates a slicer for the active sheet named range {NamedRange}
   with {NumCols} buttons per slicer row, and with dimensions
   {aHeight} by {aWidth}

```
```
 AlterSlicerColumns(SlicerName As String, NumCols)

   Loops through workbook to find {SlicerName} and sets the number
   of buttons per row to {NumCols}

```
```
 MoveSlicer(SlicerSelection,
            rngPaste As Range,
            leftOffset,
            IncTop)

   Takes Selection as {SlicerSelection}, cuts & pastes it to a rough
   location {rngPaste} to be incrementally adjusted from paste
   location by {leftOffset} and {IncTop}

```
```
 ToggleDisplayMode()

   Toggles display of ribbon, formula bar, status bar & headings

```

## Module: zPortable_Functions.bas
Portable module of functions which can be exported to any workbook and are only dependent on one-another (if at all)

```
 Tabs_MatchingCodeName(MatchCodeName As String,
                       ExcludePerfectMatch As Boolean)

   Returns array of tab names with MatchCodeName found in the CodeName
   property (useful for detecting copies of a codenamed template)

```
```
 WorksheetExists (aName)

   True or False dependent on if tab name {aName} already exists

```
```
 ExtractFirstInt_RightToLeft (aVariable)

   Returns the first integer found in a string when searcing
   from the right end of the string to the left

   ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"

```
```
 ExtractFirstInt_LeftToRight (aVariable)

   Returns the first integer found in a string when searcing
   from the left end of the string to the right

   ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"

```
```
 Truncate_Before_Int (aString)

   Removes characters before first integer in a sequence of characters

   Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"

```
```
 Truncate_After_Int (aString)

   Removes characters after first integer in a sequence of characters

   Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"

```
```
 IsInt_NoTrailingSymbols (aNumeric)

   Checks if supplied value is both numeric, and contains no numeric
   symbols (different from IsNumeric)

   IsInt_NoTrailingSymbols(9999) = True
   IsInt_NoTrailingSymbols(9999,) = False

```
```
 MyOS()

   Returns "Windows",  "Mac", or "Neither Windows or Mac"

```
```
 Get_WindowsUsername()

   Loops through folders to find paths matching C:\Users\...\AppData
   then extracts Username from correct path. Superior to reading
   .FullName of workbook which does not work for OneDrive files

```
```
 Get_MacUsername()

   Reads Activeworkbook.FullName property to get Mac user

```
```

 Get_Username()

   Returns username regardless of Windows or Mac OS

```
```
 Get_DesktopPath()

   Returns Mac or Windows desktop directory (even if on OneDrive)

```
```
 Delete_FileAndFolder(ByVal aFilePath As String)

   Read code directly prior to use

```
```
 Print_Pad()

   Uses Debug.Print to print a timestamped seperator of "======"

```
```
 Print_Named(Something, Optional Label)

   Uses Debug.Print to add a space between each {Something} printed,
   labels each {Something} if {Label} supplied

```
```
 Clipboard_Load(ByVal aString As String)

   Stores {aString} in clipboard

```
```
 Clipboard_Read(Optional IfRngConcatAllVals As Boolean = True,
                Optional Sep As String = ", ")

   Returns text from the copied object (clipboard text or range)

   >> NOT TO BE USED ON-SHEET << creates a sheet each refresh

```
```
 Get_CopiedRangeVals()

   If range copied, returns an array of each non-blank Cell.Value

   >> NOT TO BE USED ON-SHEET << creates a sheet each refresh

```
```
 Clipboard_IsRange()

   Returns True if a range is currently copied; only works in VBA

```

## Module: zRun_R.bas
Subs and Functions to interface with R in VBA; relies on zPortable_Subs and zPortable_Functions from github.com/AltraSol/VBA

```
 QuickRun_RScript(ByVal ScriptContents As String)

   Writes a temporary .R script containing {ScriptContents}, runs
   it, prompts for the deletion of the temporary script

```
```
 WriteTemp_RScript(ByVal ScriptContents As String)

   Creates a random named temporary folder on desktop, creates an
   .R file "Temp.R" containing {ScriptContents}, returns Temp.R path

```
```
 FindAndRun_RScript(ByVal ScriptLocation)

   Takes a string or cell reference {RScriptPath} & runs it on the
   latest version of R on the OS

```
```
 Run_RScript(ByVal RLocation As String, _
             ByVal ScriptLocation As String, _
             Optional ByVal Visibility As String)

   Uses the RScript.exe pointed to by {RLocation} to run the script
   found at {ScriptLocation}. Rscript.exe window displayed by default,
   but {Visibility}:= "VeryHidden" or "Minimized" can be used

```
```
 Get_RExePath() As String

   Returns the path to the latest version of Rscript.exe

```
```
 Get_LatestRVersion(ByVal RVersions As Variant)

   Returns the latest version of R currently installed

```
```
 Get_RVersions(ByVal RFolderPath As String)

   Returns an array of the R versions currently installed

```
```
 Get_RFolder() As String

   Returns the parent R folder path which houses the installed
   versions of R on the OS from which the sub is called

```
```
 Delete_FileAndFolder(ByVal aPath As String)

   Deletes {aPath} and it's container folder (including all other files)

```
