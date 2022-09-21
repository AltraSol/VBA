Attribute VB_Name = "zPortable_Subs"
Option Explicit
'===================================================================
'## Module: zPortable_Subs.bas
'Portable module of subs which can be exported to any workbook and are
'only dependent on one-another (if at all)
'===================================================================
'------------------------------------------------------------------- VBA
' ReDim_Add(ByRef aArr() As Variant, ByVal aVal)
'
''    Simplifies the addition of a value to a one dimensional array by
''    handling the initalization & resizing of an array in VBA
'
''    Call ReDim_Add(aArr(), aVal) -> last element of aArr() now aVal
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' MergeAndCombine(MergeRange As Range, _
'                 Optional SepValsByNewLine = True)
'
''    Concatenates each Cell.Value in a range & merges range as opposed
''    to Merge & Center which only keeps a single value
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' MenuAdd_MergeAndCombine()
'
''    Adds "Merge and Combine" menu option to cell right-click menu
''    Note: Calls Sub "MergeAndCombine_Selection"
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' MenuDelete_MergeAndCombine()
'
''    Deletes "Merge and Combine" menu option
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' AutoAdjustZoom(rngBegin, rngEnd)
'
''   Adjusts user view to the width of rngBegin to rngEnd
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' LaunchLink (aLink)
'
''   Launches aLink in existing browser with error handling for
''   invalid Links
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
' InsertSlicer(NamedRange As String,
'              NumCols As Integer,
'              aHeight As Double,
'              aWidth As Double)
'
''   Creates a slicer for the active sheet named range {NamedRange}
''   with {NumCols} buttons per slicer row, and with dimensions
''   {aHeight} by {aWidth}
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  AlterSlicerColumns(SlicerName As String, NumCols)
'
''   Loops through workbook to find {SlicerName} and sets the number
''   of buttons per row to {NumCols}
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  MoveSlicer(SlicerSelection,
'            rngPaste As Range,
'            leftOffset,
'            IncTop)
'
''   Takes Selection as {SlicerSelection}, cuts & pastes it to a rough
''   location {rngPaste} to be incrementally adjusted from paste
''   location by {leftOffset} and {IncTop}
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  ToggleDisplayMode()
'
''   Toggles display of ribbon, formula bar, status bar & headings
'
'-------------------------------------------------------------------
'------------------------------------------------------------------- VBA
'  CopySheetsToWorkbook()
'
''   Copies each sheet within each workbook in a given folder path to 
'    the current workbook
'
'-------------------------------------------------------------------

Sub CopySheetsToWorkbook(ByVal FromFolder As String)
Application.ScreenUpdating = False

Dim aWorkbook As Workbook, _
    aSheet As Worksheet, _
    wbName As String: wbName = Dir(FromFolder & "*.xlsx")
        
        Do While wbName <> vbNullString
            Workbooks.Open Filename:=FromFolder & wbName, ReadOnly:=True
            Set aWorkbook = ActiveWorkbook
            
                For Each aSheet In aWorkbook.Sheets
                    aSheet.Copy After:=ThisWorkbook.Sheets(1)
                    Application.StatusBar = "Adding sheets from " & aWorkbook.Name & "..."
                Next aSheet
        
            aWorkbook.Close
            wbName = Dir()
        Loop

Application.ScreenUpdating = True
Application.StatusBar = False

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

Sub MergeAndCombine_Selection()
    'For MenuAdd_MergeAndCombine
    Call MergeAndCombine(Selection)
End Sub

Sub MenuAdd_MergeAndCombine()

On Error Resume Next
  Call MenuDelete_MergeAndCombine
On Error GoTo -1

Dim Menu_MergeAndCombine As Object
Set Menu_MergeAndCombine = CommandBars("Cell").Controls.Add(before:=1)

    With Menu_MergeAndCombine
       .Caption = "Merge and Combine"
       .OnAction = "MergeAndCombine_Selection"
       .FaceId = 402
       .BeginGroup = True
    End With

        Set Menu_MergeAndCombine = Nothing

End Sub

Sub MenuDelete_MergeAndCombine()

On Error Resume Next
   CommandBars("Cell").Controls("Merge and Combine").Delete
On Error GoTo -1

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

Sub InsertSlicer(NamedRange As String, NumCols As Integer, aHeight As Double, aWidth As Double)
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

Sub MoveSlicer(SlicerSelection, rngPaste As Range, leftOffset, IncTop)
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


