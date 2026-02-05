Sub ПереносБлоками_J_в_NOP()

    Dim lastRow As Long
    Dim srcRow As Long
    Dim outRow As Long
    Dim cyclePos As Integer
    
    lastRow = Cells(Rows.Count, "J").End(xlUp).Row
    outRow = 1
    cyclePos = 1

    For srcRow = 1 To lastRow
        
        ' Проверка на любую границу между строками
        If srcRow > 1 Then
            If Cells(srcRow, "J").Borders(xlEdgeTop).LineStyle <> xlNone _
            Or Cells(srcRow - 1, "J").Borders(xlEdgeBottom).LineStyle <> xlNone Then
                
                outRow = outRow + 1
                cyclePos = 1
                
            End If
        End If
        
        ' Перенос значений в одну строку по циклу
        Select Case cyclePos
            Case 1
                Cells(outRow, "N").Value = Cells(srcRow, "J").Value
            Case 2
                Cells(outRow, "O").Value = Cells(srcRow, "J").Value
            Case 3
                Cells(outRow, "P").Value = Cells(srcRow, "J").Value
        End Select
        
        cyclePos = cyclePos + 1
        If cyclePos > 3 Then cyclePos = 1
        
    Next srcRow

End Sub
