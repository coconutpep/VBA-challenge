Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ws As Worksheet
For Each ws In Worksheets
    Count = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim open_val As Double
    Dim close_val As Double
    Dim total_val As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    For i = 2 To last_row
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
         ws.Cells(Count, 9).Value = ws.Cells(i, 1).Value
         open_val = 0
         open_val = open_val + ws.Cells(i, 3).Value
         total_val = 0
         total_val = total_val + ws.Cells(i, 7).Value
        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
         total_val = total_val + ws.Cells(i, 7).Value
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
         close_val = 0
         close_val = close_val + ws.Cells(i, 6).Value
         yearly_change = 0
         yearly_change = open_val - close_val
         ws.Cells(Count, 10).Value = yearly_change
         total_val = total_val + ws.Cells(i, 7).Value
         ws.Cells(Count, 12).Value = total_val
         If yearly_change > 0 Then
          ws.Cells(Count, 10).Interior.ColorIndex = 4
         ElseIf yearly_change < 0 Then
          ws.Cells(Count, 10).Interior.ColorIndex = 3
         End If
         If open_val <> 0 Then
          percent_change = 0
          percent_change = percent_change + (yearly_change / Abs(open_val))
          ws.Cells(Count, 11).Value = percent_change
          ws.Cells(Count, 11).NumberFormat = "0.00%"
         ElseIf open_val = 0 Then
          ws.Cells(Count, 11).Value = "N/A"
         End If
        Count = Count + 1
        End If
    Next i
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Columns.Item(11))
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Columns.Item(11))
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Columns.Item(12))
    ws.Range("P2:P3").NumberFormat = "0.00%"
    last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To last_row
        If ws.Cells(i, 11).Value = ws.Cells(2, 16).Value Then
         ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 16).Value Then
         ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 16).Value Then
         ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        End If
    Next i

    ws.Columns("I:P").AutoFit
Next ws
End Sub
