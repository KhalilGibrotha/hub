Sub FillBlanks()

Dim ws As Worksheet
Dim lRow As Long
Dim lCol As Long

Set ws = ActiveSheet

lRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
lCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

ws.Range(ws.Cells(1, 1), ws.Cells(lRow, lCol)).SpecialCells(xlCellTypeBlanks).Value = "NULL"

End Sub

