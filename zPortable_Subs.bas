Attribute VB_Name = "zPortable_Subs"
Option Explicit
'===================================================================
'Portable module of subs which can be exported to any workbook
'===================================================================
'-------------------------------------------------------------------
' ReDim_Add(ByRef aArr() As Variant, ByVal aVal)
'
'   Simplifies the addition of a value to a one dimensional array by
'   handling the initalization & resizing of an array
'
'   Call ReDim_Add(aArr(), aVal) -> last element of aArr() now aVal
'
'-------------------------------------------------------------------
' AutoAdjustZoom(rngBegin, rngEnd)
'
'   Adjusts user view to the width of rngBegin to rngEnd
'
'-------------------------------------------------------------------
' LaunchLink(aLink)
'
'   Launches aLink in existing browser with error handling for
'   invalid links
'
'-------------------------------------------------------------------
' InsertSlicer(NamedRange As String,
'              numCols As Integer,
'              aHeight As Double,
'              aWidth As Double)
'
'   Creates a slicer for the active sheet named range {NamedRange}
'   with {numCols} buttons per slicer row, and with dimensions
'   {aHeight} by {aWidth}
'
'-------------------------------------------------------------------
' AlterSlicerColumns(SlicerName As String, numCols)
'
'   Loops through workbook to find {SlicerName} and sets the number
'   of buttons per row to {numCols}
'
'-------------------------------------------------------------------
' MoveSlicer(slicerSelection,
'            rngPaste As Range,
'            leftOffset,
'            IncTop)
'
'   Takes Selection as {slicerSelection}, cuts & pastes it to a rough
'   location {rngPaste} to be incrementally adjusted from paste
'   location by {leftOffset} and {IncTop}
'
'-------------------------------------------------------------------

Sub MergeAndCombine()
    
    'Concatenates cells in a range

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
Call ReDim_Add(aArr(), aVal)

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

Sub InsertSlicer(NamedRange As String, numCols As Integer, aHeight As Double, aWidth As Double)
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
                
                    'Use f(x) AlterSlicerColumns to change {SlicerName}'s number of cols to {numCols}
                    Call AlterSlicerColumns(SlicerName, numCols)
End Sub

Sub AlterSlicerColumns(SlicerName As String, numCols)
On Error Resume Next

Dim i As Integer
    
    'Loop through each Slicer within workbook
    For i = 1 To ActiveWorkbook.SlicerCaches.Count
        'Neccesarily will error for all but one loop, when the correct Slicer
        'called {SlicerName} is found. using Slicers(1) or Slicers(j) does not
        'work consistently
         ActiveWorkbook.SlicerCaches(i).Slicers(SlicerName).NumberOfColumns = numCols
    Next i
    
    DoEvents

End Sub

Sub MoveSlicer(slicerSelection, rngPaste As Range, leftOffset, IncTop)
On Error Resume Next
    
    DoEvents
    slicerSelection.Cut 'Cut the slicer current selected, which is {slicerSelection}
    rngPaste.Select 'Select the range with which we're aligning {slicerSelection}'s top and left positions with
    ActiveSheet.Paste 'Paste the slicer onto cell {rngPaste}
    DoEvents
    
        'After pasting, {slicerSelection} is once again the selected object
        
        'Move {slicerSelection} to the RIGHT of {rngPaste} by {leftOffset}
        ActiveSheet.Shapes(Selection.Name).IncrementLeft leftOffset
        DoEvents
        
        'Move {slicerSelection} upwards by {IncTop}
        ActiveSheet.Shapes(Selection.Name).IncrementTop IncTop
        DoEvents
        
End Sub

