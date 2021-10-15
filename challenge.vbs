Sub WallStreet()

Dim Ticker As String
Dim TotalVol As Double
Dim YearOpen As Double
Dim YearClose As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim SummaryTableRow As Integer

For Each ws In Worksheets

LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
ws.Range("J:M").Delete

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Volume"

SummaryTableRow = 2
TotalVol = 0

YearOpen = ws.Cells(2, "C").Value

For i = 2 To LastRow
    TotalVol = TotalVol + ws.Cells(i + 1, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    YearClose = ws.Cells(i, "F").Value
    

    
'Formulas
    YearlyChange = YearClose - YearOpen
    If YearOpen = 0 Then
    PercentChange = 0
    Else
    PercentChange = YearlyChange / YearOpen * 100
    End If
    YearOpen = ws.Cells(i + 1, "C").Value

If PercentChange > 0 Then
    ws.Range("L" & SummaryTableRow).Interior.ColorIndex = 4
    ElseIf PercentChange < 0 Then
    ws.Range("L" & SummaryTableRow).Interior.ColorIndex = 3
    End If
    
    
'Where to put the data
ws.Range("J" & SummaryTableRow).Value = Ticker
ws.Range("K" & SummaryTableRow).Value = YearlyChange
ws.Range("L" & SummaryTableRow).Value = PercentChange
ws.Range("M" & SummaryTableRow).Value = TotalVol
SummaryTableRow = SummaryTableRow + 1
TotalVol = 0
        End If
    Next i
    
Next ws
End Sub
