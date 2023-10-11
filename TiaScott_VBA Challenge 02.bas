Attribute VB_Name = "Module1"
Option Explicit
Sub Ticker()
  Const FIRST_DATA_ROW As Integer = 2
  Const IN_TICKER_COL As Integer = 1
  Const OPEN_COL As Integer = 3
  Const CLOSE_COL As Integer = 6
  Const IN_VOL_COL As Integer = 7
  Const OUT_TICKER_COL As Integer = 9
  Const YEARLY_CHG_COL As Integer = 10
  Const PERCENT_CHG_COL As Integer = 11
  Const OUT_TOTAL_VOL_COL As Integer = 12
  Const GREATEST_HIT_OUT As Integer = 15
  Const GREATEST_HIT_TICKER As Integer = 16
  Const GREATEST_HIT_VALUE As Integer = 17
  
  Dim ws As Worksheet
  Dim PrevTicker As String
  Dim CurrentTicker As String
  Dim NextTicker As String
  Dim Ticker_Total As Variant
  Dim YearlyChange As Double
  Dim YearlyChangeFrac As Double
  Dim NewYearOpenPrice As Double
  Dim YearEndClosePrice As Double
  Dim Output_Row As Long
  Dim Input_Row As Long
  Dim lastrow As Long
  Dim MaxPercentIncrease As Double
  Dim MaxPercentDecrease As Double
  Dim MaxTotalVolume As Double
  Dim MaxPercentIncreaseTicker As String
  Dim MaxPercentDecreaseTicker As String
  Dim MaxTotalVolumeTicker As String
  
  Ticker_Total = 0
  MaxPercentIncrease = 0
  MaxPercentDecrease = 0
  MaxTotalVolume = 0
  
  For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    Output_Row = FIRST_DATA_ROW
    lastrow = ws.Cells(Rows.Count, IN_TICKER_COL).End(xlUp).Row
    ws.Cells(1, OUT_TICKER_COL).Value = "Ticker"
    ws.Cells(1, YEARLY_CHG_COL).Value = "Yearly Change"
    ws.Cells(1, PERCENT_CHG_COL).Value = "Percent Change"
    ws.Cells(1, OUT_TOTAL_VOL_COL).Value = "Total Stock Volume"
    ws.Cells(2, GREATEST_HIT_OUT).Value = "Greatest % Increase"
    ws.Cells(3, GREATEST_HIT_OUT).Value = "Greatest % Decrease"
    ws.Cells(4, GREATEST_HIT_OUT).Value = "Greatest Total Volume"
    ws.Cells(1, GREATEST_HIT_TICKER).Value = "Ticker"
    ws.Cells(1, GREATEST_HIT_VALUE).Value = "Value"
  
  For Input_Row = FIRST_DATA_ROW To lastrow
    PrevTicker = ws.Cells((Input_Row - 1), IN_TICKER_COL).Value
    CurrentTicker = ws.Cells(Input_Row, IN_TICKER_COL).Value
    NextTicker = ws.Cells((Input_Row + 1), IN_TICKER_COL).Value
    Ticker_Total = Ticker_Total + ws.Cells(Input_Row, IN_VOL_COL).Value
      If PrevTicker <> CurrentTicker Then
        Ticker_Total = 0
        NewYearOpenPrice = ws.Cells(Input_Row, OPEN_COL).Value
      End If
      If NextTicker <> CurrentTicker Then
        YearEndClosePrice = ws.Cells(Input_Row, CLOSE_COL).Value
        YearlyChange = YearEndClosePrice - NewYearOpenPrice
        YearlyChangeFrac = YearlyChange / NewYearOpenPrice
        ws.Cells(Output_Row, OUT_TICKER_COL).Value = CurrentTicker
        ws.Cells(Output_Row, OUT_TOTAL_VOL_COL).Value = Ticker_Total
        ws.Cells(Output_Row, YEARLY_CHG_COL).Value = YearlyChange
            If YearlyChange > 0 Then
              ws.Cells(Output_Row, YEARLY_CHG_COL).Interior.Color = RGB(51, 153, 51)
            ElseIf YearlyChange < 0 Then
              ws.Cells(Output_Row, YEARLY_CHG_COL).Interior.Color = RGB(255, 0, 0)
            Else
              ws.Cells(Output_Row, YEARLY_CHG_COL).Interior.Color = RGB(255, 255, 255)
            End If
          ws.Cells(Output_Row, PERCENT_CHG_COL).Value = FormatPercent(YearlyChangeFrac)
            If ws.Cells(Output_Row, PERCENT_CHG_COL).Value > MaxPercentIncrease Then
              MaxPercentIncrease = ws.Cells(Output_Row, PERCENT_CHG_COL).Value
              MaxPercentIncreaseTicker = ws.Cells(Output_Row, OUT_TICKER_COL).Value
            ElseIf ws.Cells(Output_Row, PERCENT_CHG_COL).Value < MaxPercentDecrease Then
              MaxPercentDecrease = ws.Cells(Output_Row, PERCENT_CHG_COL).Value
              MaxPercentDecreaseTicker = ws.Cells(Output_Row, OUT_TICKER_COL).Value
            End If
            If ws.Cells(Output_Row, OUT_TOTAL_VOL_COL).Value > MaxTotalVolume Then
              MaxTotalVolume = ws.Cells(Output_Row, OUT_TOTAL_VOL_COL).Value
              MaxTotalVolumeTicker = ws.Cells(Output_Row, OUT_TICKER_COL).Value
            End If
          Output_Row = Output_Row + 1
      End If
  Next Input_Row
      ws.Cells(2, GREATEST_HIT_TICKER).Value = MaxPercentIncreaseTicker
      ws.Cells(3, GREATEST_HIT_TICKER).Value = MaxPercentDecreaseTicker
      ws.Cells(4, GREATEST_HIT_TICKER).Value = MaxTotalVolumeTicker
      ws.Cells(2, GREATEST_HIT_VALUE).Value = FormatPercent(MaxPercentIncrease)
      ws.Cells(3, GREATEST_HIT_VALUE).Value = FormatPercent(MaxPercentDecrease)
      ws.Cells(4, GREATEST_HIT_VALUE).Value = MaxTotalVolume
      
      ws.UsedRange.Columns.AutoFit
  Next ws
End Sub


