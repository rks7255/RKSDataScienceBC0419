'I have used inclass activity Credit Card Checker VBA Script
'as a reference to get results for the homework.


Sub StockMarket()

Dim sm As Worksheet

Dim tckr As String
Dim vlm As Double

vlm = 0

For Each sm In ThisWorkbook.Worksheets '==> Run concurrently on all sheets
    sm.Cells(1, 9).Value = "<ticker>"
    sm.Cells(1, 10).Value = "<total stock volume>"
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2


   For i = 2 To sm.UsedRange.Rows.Count '==> Counts the range of used rows from row 2


       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
       tckr = sm.Cells(i, 1).Value
       vlm = vlm + sm.Cells(i, 7).Value
       
       sm.Cells(Summary_Table_Row, 9).Value = tckr
       sm.Cells(Summary_Table_Row, 10).Value = vlm

       Summary_Table_Row = Summary_Table_Row + 1

       vlm = 0

       Else
           vlm = vlm + sm.Cells(i, 7).Value

        End If
   Next i
Next

End Sub

