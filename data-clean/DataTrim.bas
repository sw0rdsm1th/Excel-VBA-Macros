Option Explicit

'Many ways to clean data in Excel and thankfully with macros

Sub CleamTrim()

Dim rg As Range
Dim SelectionArea As Range

'Check for formulas in selection
  If Selection.Cells.Count = 1 Then
    Set rg = Selection
  Else
    Set rg = Selection.SpecialCells(xlCellTypeConstants)
  End If

'Trim and Clean cell values
  For Each SelectionArea In rg.Areas
    SelectionArea.Value = Evaluate("IF(ROW(" & SelectionArea.Address & "),CLEAN(TRIM(" & SelectionArea.Address & ")))")
  Next SelectionArea
  
End Sub
