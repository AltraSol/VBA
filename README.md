# VBA

## zPortable_Functions
===================================================================
Portable module of functions which can be exported to any workbook
===================================================================
# VBA

## zPortable_Functions
===================================================================
Portable module of functions which can be exported to any workbook
===================================================================
```
 Tabs_MatchingCodeName(MatchCodeName As String,
                       ExcludePerfectMatch As Boolean)

   Returns array of tab names with MatchCodeName found in the CodeName
   property (useful for detecting copies of a code-named template)

```
 WorksheetExists (aName)

   True or False dependent on if tab name {aName} already exists

```
 ExtractFirstInt_RightToLeft (aVariable)

   Returns the first integer found in a string when searcing
   from the right end of the string to the left

   ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"

```
 ExtractFirstInt_LeftToRight (aVariable)

   Returns the first integer found in a string when searcing
   from the left end of the string to the right

   ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"

```
 Truncate_Before_Int (aString)

   Removes characters before first integer in a sequence of characters

   Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"

```
 Truncate_After_Int (aString)

   Removes characters after first integer in a sequence of characters

   Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"

```
 IsInt_NoTrailingSymbols (aNumeric)

   Checks if supplied value is both numeric, and contains no numeric
   symbols (different from IsNumeric)

   IsInt_NoTrailingSymbols(9999) = True
   IsInt_NoTrailingSymbols(9999,) = False

```
 Get_DesktopPath()

   Returns Windows desktop directory (even if hosted on OneDrive)

```

 Get_Username()

   Returns username regardless of Windows or Mac OS

```
 MyOS()

   'Returns "Windows",  "Mac", or "Neither Windows or Mac"

```
 Get_WindowsUsername()

   Loops through folders to find paths matching C:\Users\...\AppData
   then extracts Username from correct path. Superior to reading
   .FullName of workbook which does not work for OneDrive files

```
 Print_Pad()

   Uses Debug.Print to print a timestamped seperator of "======"

```
 Print_Named(Something, Optional Label)

   Uses Debug.Print to add a space between each {Something} printed,
   labels each {Something} if {Label} supplied

```
 Clipboard_Load(ByVal aString As String)

   Stores {aString} in clipboard

```
 Clipboard_Read(Optional IfRngConcatAllVals As Boolean = True,
                Optional Sep As String = ", ")

   Returns text from the copied object (clipboard text or range)

   '>> NOT TO BE USED ON-SHEET << creates a sheet each refresh

# VBA

## zPortable_Functions
===================================================================
Portable module of functions which can be exported to any workbook
===================================================================
```
 Tabs_MatchingCodeName(MatchCodeName As String,
                       ExcludePerfectMatch As Boolean)

   Returns array of tab names with MatchCodeName found in the CodeName
   property (useful for detecting copies of a code-named template)

```
 WorksheetExists (aName)

   True or False dependent on if tab name {aName} already exists

```
 ExtractFirstInt_RightToLeft (aVariable)

   Returns the first integer found in a string when searcing
   from the right end of the string to the left

   ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"

```
 ExtractFirstInt_LeftToRight (aVariable)

   Returns the first integer found in a string when searcing
   from the left end of the string to the right

   ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"

```
 Truncate_Before_Int (aString)

   Removes characters before first integer in a sequence of characters

   Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"

```
 Truncate_After_Int (aString)

   Removes characters after first integer in a sequence of characters

   Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"

```
 IsInt_NoTrailingSymbols (aNumeric)

   Checks if supplied value is both numeric, and contains no numeric
   symbols (different from IsNumeric)

   IsInt_NoTrailingSymbols(9999) = True
   IsInt_NoTrailingSymbols(9999,) = False

```
 Get_DesktopPath()

   Returns Windows desktop directory (even if hosted on OneDrive)

```

 Get_Username()

   Returns username regardless of Windows or Mac OS

```
 MyOS()

   'Returns "Windows",  "Mac", or "Neither Windows or Mac"

```
 Get_WindowsUsername()

   Loops through folders to find paths matching C:\Users\...\AppData
   then extracts Username from correct path. Superior to reading
   .FullName of workbook which does not work for OneDrive files

```
 Print_Pad()

   Uses Debug.Print to print a timestamped seperator of "======"

```
 Print_Named(Something, Optional Label)

   Uses Debug.Print to add a space between each {Something} printed,
   labels each {Something} if {Label} supplied

```
 Clipboard_Load(ByVal aString As String)

   Stores {aString} in clipboard

```
 Clipboard_Read(Optional IfRngConcatAllVals As Boolean = True,
                Optional Sep As String = ", ")

   Returns text from the copied object (clipboard text or range)

   '>> NOT TO BE USED ON-SHEET << creates a sheet each refresh

```
 Get_CopiedRangeVals()

   If range copied, returns an array of each non-blank Cell.Value

   >> NOT TO BE USED ON-SHEET << creates a sheet each refresh

```
 Clipboard_IsRange()

   Returns True if a range is currently copied; only works in VBA

``` Get_CopiedRangeVals()

   If range copied, returns an array of each non-blank Cell.Value

   >> NOT TO BE USED ON-SHEET << creates a sheet each refresh

```
 Clipboard_IsRange()

   Returns True if a range is currently copied; only works in VBA

```
