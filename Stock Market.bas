Attribute VB_Name = "Module1"
Sub Stock_market()

Dim ws As Worksheet
Dim Ticker As Variant
Dim Ticker_volume As Variant
Dim Lastrow As Variant
Dim I As Variant
Dim open_price As Variant
Dim close_price As Variant
Dim TickerRow As Variant
Dim NewRow As Variant
Dim closing_price As Variant
Dim closing_volume As String
Dim FoundCell As Variant
Dim MaxVal As Double
Dim MaxRow As Variant

For Each ws In ThisWorkbook.Worksheets

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Ticker = ws.Cells(2, 1).Value
open_price = ws.Cells(2, 6).Value
NewRow = 2
TickerRow = 1
closing_price = 0
closing_volume = 0

    For I = 2 To Lastrow
        If ws.Cells(I + 1, 1).Value <> Ticker Then
            close_price = ws.Cells(I, 6).Value
            TickerRow = TickerRow + 1
            If ws.Cells(TickerRow, 9).Value = Ticker Then
                ws.Cells(TickerRow, 10).Value = close_price - open_price
                ws.Cells(TickerRow, 11).Value = Format(((close_price - open_price) / (open_price)) * 100, "0.00")
                closing_volume = Format(Application.WorksheetFunction.Sum(ws.Range("G" & NewRow, "G" & I)), "#,##0")
                ws.Cells(TickerRow, 12).Value = closing_volume
                closing_volume = 0
            End If
            Ticker = ws.Cells(I + 1, 1).Value
            open_price = ws.Cells(I + 1, 6).Value
            NewRow = I + 1
        End If
    Next I
   
    If Lastrow + 1 = I Then
        ws.Cells(2, "Q").Value = WorksheetFunction.Max(ws.Range("K" & 2, "K" & Lastrow))
        Set FoundCell = ws.Range("K:K").Find(what:=ws.Cells(2, "Q").Value)
        ws.Cells(2, "P").Value = ws.Cells(FoundCell.Row, "I").Value
       
        ws.Cells(3, "Q").Value = WorksheetFunction.Min(ws.Range("K" & 2, "K" & Lastrow))
        Set FoundCell = ws.Columns("K:K").Find(what:=ws.Cells(3, "Q").Value, LookIn:=xlValues, lookat:=xlWhole)
        ws.Cells(3, "P").Value = ws.Cells(FoundCell.Row, "I").Value
       
       
       
        MaxVal = WorksheetFunction.Max(ws.Range("L" & 2, "L" & TickerRow))
        MaxRow = Application.Match(MaxVal, ws.Range("L" & 2, "L" & TickerRow), 0)
        ws.Cells(4, "Q").Value = MaxVal
        ws.Cells(4, "P").Value = ws.Cells(MaxRow + 1, "I").Value
        End If
   
Next ws


End Sub

