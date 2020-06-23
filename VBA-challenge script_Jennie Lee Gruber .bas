Attribute VB_Name = "Module2"
Sub Stockmarket()
Dim Ticker As String

Range("I1").Value = "Ticker"
Range("J1").Value = "Stock change"
Range("K1").Value = "% of change"
Range("L1").Value = "Stock volume"

Dim OpenPrice As Double
    OpenPrice = 0
Dim ClosePrice As Double
    ClosePrice = 0
Dim Volume As Double
    Volume = 0
Dim ChangeinPrice As Double
ChangeinPrice = 0
Dim Counter As Integer
    Counter = 0
Dim Summary_Row_Table As Integer
Summary_Row_Table = 2
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Range("I" & Summary_Row_Table).Value = Cells(i, 1).Value
        Range("L" & Summary_Row_Table).Value = Volume + Cells(i, 5).Value
        OpenPrice = Range("C" & (i - Counter)).Value
        ClosePrice = Range("F" & Summary_Row_Table).Value
        ChangeinPrice = ((ClosePrice - OpenPrice) / ClosePrice)
        Range("J" & Summary_Row_Table).Value = ClosePrice - OpenPrice
        Range("K" & Summary_Row_Table).Value = FormatPercent(ChangeinPrice)
            If Range("J" & Summary_Row_Table).Value < 0 Then
            Range("J" & Summary_Row_Table).Interior.ColorIndex = 3
            Else
            Range("J" & Summary_Row_Table).Interior.ColorIndex = 4
            End If
        Summary_Row_Table = Summary_Row_Table + 1
        Volume = 0
        Counter = 0
    Else
        Volume = Volume + Cells(i, 7).Value
        Counter = Counter + 1
        
       
End If
    
Next i
End Sub

