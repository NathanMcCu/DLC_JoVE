Attribute VB_Name = "Module2"
Sub Step2()
    AverageColumns
    DeleteNewColumns
    DeleteRows

End Sub
Sub DeleteNewColumns()
    Dim ws As Worksheet
    Dim columnsToDelete As Variant
    Dim i As Long
    Dim columnIndex As Long
    
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
   
    columnsToDelete = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O") ' Modify this array with the column letters you want to delete
    
    
    For i = LBound(columnsToDelete) To UBound(columnsToDelete) - 1
        For j = i + 1 To UBound(columnsToDelete)
            If ws.Range(columnsToDelete(i) & "1").Column < ws.Range(columnsToDelete(j) & "1").Column Then
                temp = columnsToDelete(j)
                columnsToDelete(j) = columnsToDelete(i)
                columnsToDelete(i) = temp
            End If
        Next j
    Next i
    
    
    For i = LBound(columnsToDelete) To UBound(columnsToDelete)
        columnIndex = ws.Range(columnsToDelete(i) & "1").Column
        ws.Columns(columnIndex).Delete
    Next i
End Sub
Sub DeleteRows()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ws.Rows("1:2").Delete
End Sub
Sub AverageColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim avgValue As Double
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ws.Columns("P:P").Insert Shift:=xlToRight
    
    ws.Cells(3, "P").Value = "EucD"
    
    For i = 4 To lastRow
        If Not (IsError(ws.Cells(i, "H").Value) Or IsError(ws.Cells(i, "O").Value)) Then
            If ws.Cells(i, "H").Value = "NaN" Or ws.Cells(i, "O").Value = "NaN" Then
                ws.Cells(i, "P").Value = "NaN"
            Else
                avgValue = Application.Evaluate("(H" & i & "+O" & i & ")/2")
                ws.Cells(i, "P").Value = avgValue
            End If
        Else
            ws.Cells(i, "P").Value = "NaN"
        End If
    Next i
End Sub

