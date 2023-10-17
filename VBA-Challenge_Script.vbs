VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub testing():

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Difference"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest%Increase"
    ws.Cells(3, 15).Value = "Greatest%Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    Dim ticker As String
    Dim lastrow As Long
    Dim summary_table_row As Long
    Dim openvalue As Double
    Dim closevalue As Double
    Dim TSV As LongLong
    Dim perchange As Double
    Dim maxincrease As Double
    Dim minincrease As Double
    Dim maxTSV As LongLong
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    summary_table_row = 2
    maxincrease = 0
    minincrease = 0
    maxTSV = 0
    
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            openvalue = ws.Cells(i, 3).Value
        End If
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            closevalue = ws.Cells(i, 6).Value
            TSV = TSV + ws.Cells(i, 7).Value
            ws.Range("I" & summary_table_row).Value = ws.Cells(i, 1).Value
            ws.Range("J" & summary_table_row).Value = closevalue - openvalue
                If ws.Range("j" & summary_table_row).Value > 0 Then
                    ws.Range("j" & summary_table_row).Interior.ColorIndex = 4
                Else
                    ws.Range("j" & summary_table_row).Interior.ColorIndex = 3
                End If
            ws.Range("L" & summary_table_row).Value = TSV
            perchange = (((closevalue / openvalue) - 1))
            ws.Range("K" & summary_table_row).Value = FormatPercent(perchange, 2)
            summary_table_row = summary_table_row + 1
            TSV = 0
        Else
            TSV = TSV + ws.Cells(i, 7).Value
        End If
    Next i
    For j = 2 To lastrow
        If ws.Cells(j, 11).Value > maxincrease Then
            maxincrease = ws.Cells(j, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(2, 17).Value = FormatPercent(maxincrease, 2)
        End If
        If ws.Cells(j, 11).Value < minincrease Then
            minincrease = ws.Cells(j, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(3, 17).Value = FormatPercent(minincrease, 2)
        End If
        If ws.Cells(j, 12).Value > maxTSV Then
            maxTSV = ws.Cells(j, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(4, 17).Value = maxTSV
        End If
    Next j
Next ws

End Sub
