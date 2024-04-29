Attribute VB_Name = "Module1"
Sub Step1()
    DeleteColumns
    InsertEucD_B
    InsertEucD_C
    ReplaceWithNaN
    

End Sub
Sub DeleteColumns()
    Dim ws As Worksheet
    Dim columnsToDelete As Variant
    Dim i As Long
    Dim columnIndex As Long
    
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
   
    columnsToDelete = Array("B", "C", "D", "E", "F", "G", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AI", "AJ", "AK", "AO", "AP", "AQ", "AU", "AV", "AW", "AX", "AY", "AZ") ' Modify this array with the column letters you want to delete
    
    
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
Sub ReplaceWithNaN()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outerCol As Variant
    Dim innerCol As Variant
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Specify the columns you want to check here
    Dim colsToCheck As Variant
    colsToCheck = Array("D", "G", "J", "M")
    
    ' Specify the value to check against (less than this value will be replaced)
    Dim threshold As Double
    threshold = 0.92
    
 For i = 4 To lastRow
        ' Check if any value in the specified columns is less than the threshold
        For Each outerCol In colsToCheck
            If ws.Cells(i, outerCol).Value < threshold Then
                ' Replace all values in the row with NaN, except for column A
                For Each innerCol In ws.Columns
                    If innerCol.Column <> 1 Then ' Skip Column A
                        If Not IsEmpty(ws.Cells(i, innerCol.Column).Value) Then
                            ws.Cells(i, innerCol.Column).Value = "NaN"
                        End If
                    End If
                Next innerCol
                ' Exit the outer loop since we've already found a value less than the threshold in this row
                Exit For
            End If
        Next outerCol
    Next i
End Sub
Sub InsertEucD_B()
    Dim ws As Worksheet
    Dim lastRow As Long
    
   
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    
    ws.Columns("H:H").Insert Shift:=xlToRight
    
    
    With ws
        
        .Range("H1").Value = "-"
        .Range("H2").Value = "-"
        
        .Range("H3").Value = "EucD"
        
        
        .Range("H4:H" & lastRow).Formula = "=SQRT((E4-B4)^2+(F4-C4)^2)"
        
        
        .Range("H4:H" & lastRow).Value = .Range("H4:H" & lastRow).Value
    End With
End Sub

Sub InsertEucD_C()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    
   
    ws.Columns("O:O").Insert Shift:=xlToRight

    With ws
        
        .Range("O1").Value = "-"
        .Range("O2").Value = "-"
        
        
        .Range("O3").Value = "EucD"
        
        
        .Range("O4:O" & lastRow).Formula = "=SQRT((L4-I4)^2+(M4-J4)^2)"
        
        
        .Range("O4:O" & lastRow).Value = .Range("O4:O" & lastRow).Value
    End With
End Sub



