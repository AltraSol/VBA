Attribute VB_Name = "zPortable_Functions"
Option Explicit
'===================================================================
'Portable module of functions which can be exported to any workbook
'===================================================================
'-------------------------------------------------------------------
' Tabs_MatchingCodeName(MatchCodeName As String,
'                       ExcludePerfectMatch As Boolean)
'
'   Returns array of tab names with MatchCodeName found in the CodeName
'   property (useful for detecting copies of a code-named template)
'
'-------------------------------------------------------------------
' WorksheetExists(aName)
'
'   True or False dependent on if tab name {aName} already exists
'
'-------------------------------------------------------------------
' ExtractFirstInt_RightToLeft(aVariable)
'
'   Returns the first integer found in a string when searcing
'   from the right end of the string to the left
'
'   ExtractFirstInt_RightToLeft("Some12Embedded345Num") = "345"
'
'-------------------------------------------------------------------
' ExtractFirstInt_LeftToRight(aVariable)
'
'   Returns the first integer found in a string when searcing
'   from the left end of the string to the right
'
'   ExtractFirstInt_LeftToRight("Some12Embedded345Num") = "12"
'
'-------------------------------------------------------------------
' Truncate_Before_Int(aString)
'
'   Removes characters before first integer in a sequence of characters
'
'   Truncate_After_Int("Some12Embedded345Num") = "12Embedded345Num"
'
'-------------------------------------------------------------------
' Truncate_After_Int(aString)
'
'   Removes characters after first integer in a sequence of characters
'
'   Truncate_After_Int("Some12Embedded345Num") = "Some12Embedded345"
'
'-------------------------------------------------------------------
' IsInt_NoTrailingSymbols(aNumeric)
'
'   Checks if supplied value is both numeric, and contains no numeric
'   symbols (different from IsNumeric)
'
'   IsInt_NoTrailingSymbols(9999) = True
'   IsInt_NoTrailingSymbols(9999,) = False
'
'-------------------------------------------------------------------
' GetUserName() >> assumes Windows machine & file not cloud hosted <<
'
'   Extracts username from filepath (e.g. C:\Users\{YourName}\Folder...
'
'   GetUserName() = {YourName}
'
'-------------------------------------------------------------------
' Print_Pad()
'
'   Uses Debug.Print to print a timestamped seperator of "======"
'
'-------------------------------------------------------------------
' Print_Named(Something, Optional Label)
'
'   Uses Debug.Print to add a space between each {Something} printed,
'   labels each {Something} if {Label} supplied
'
'-------------------------------------------------------------------

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

Function GetUserName()
Dim SlashLocation, UserNameLength

    GetUserName = Replace(ActiveWorkbook.FullName, "C:\Users\", "")
        SlashLocation = InStr(1, GetUserName, "\")
            GetUserName = Left(GetUserName, SlashLocation - 1)

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





