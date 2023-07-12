Sub stock_loop()

For Each ws In Worksheets

Dim ticker As String
Dim volume_total As Long
Dim column As Integer
Dim Percent_Change As Double
Dim Yearly_Open As Double
Dim Yearly_Close As Double

total_stock_volume = 0

column = 2

Open_counter = 2

LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRow

'Add headers for Ticker, Yearly Change, Percent Change and Total Stock Volume

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Value"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Value of Yearly_Open

Yearly_Open = ws.Cells(Open_counter, 3).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'obtain Yearly_Close Value

Yearly_Close = ws.Cells(i, 6).Value

ws.Range("J" & column).Value = Yearly_Close - Yearly_Open
ws.Range("K" & column).Value = (Yearly_Close - Yearly_Open) / Yearly_Open

' Conditional format with If statement, green for positive and red for negative values

If ws.Cells(column, 10).Value > 0 Then

ws.Cells(column, 10).Interior.ColorIndex = 4

Else

ws.Cells(column, 10).Interior.ColorIndex = 3

End If

If ws.Cells(column, 11).Value > 0 Then

ws.Cells(column, 11).Interior.ColorIndex = 4

Else

ws.Cells(column, 11).Interior.ColorIndex = 3

End If


'  Set the ticker value
ticker = ws.Cells(i, 1).Value

'Add to the total stock volume
total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

'Print the ticker value into the column
ws.Range("I" & column).Value = ticker

'Print the stock volume amount into column "L"

ws.Range("L" & column).Value = total_stock_volume

' change format to percentage

ws.Cells(column, 11).NumberFormat = "0.00%"

'Add one to the column to prevent overwriting
column = column + 1

'Reset the total stock volume

total_stock_volume = 0

' add one to open counter to move to the next ticker open price

Open_counter = i + 1


' If the ticker following a row is the same
    Else

      ' Add to the stock volume total
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value


End If


Next i

' Calculate Greatest % Increase, Greatest % Decrease and Greatest Total Volume with a new loop

Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double
Dim Greatest_Increase_Ticker As String
Dim Greatest_Decrease_Ticker As String
Dim Greatest_Volume_Ticker As String

Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Total_Volume = 0

Last_summary_row = ws.Cells(Rows.Count, "K").End(xlUp).Row

For j = 3 To Last_summary_row

If ws.Cells(j, 11).Value > Greatest_Increase Then

Greatest_Increase = ws.Cells(j, 11).Value

Greatest_Increase_Ticker = ws.Cells(j, 9).Value

ws.Range("Q2").Value = Greatest_Increase
ws.Range("P2").Value = Greatest_Increase_Ticker

End If

If ws.Cells(j, 11).Value < Greatest_Decrease Then

Greatest_Decrease = ws.Cells(j, 11).Value

Greatest_Decrease_Ticker = ws.Cells(j, 9).Value

ws.Range("Q3").Value = Greatest_Decrease
ws.Range("P3").Value = Greatest_Decrease_Ticker

End If

If ws.Cells(j, 12).Value > Greatest_Total_Volume Then

Greatest_Total_Volume = ws.Cells(j, 12).Value

Greatest_Volume_Ticker = ws.Cells(j, 9).Value

ws.Range("Q4").Value = Greatest_Total_Volume
ws.Range("P4").Value = Greatest_Volume_Ticker

End If

Next j



ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"



ws.Columns("A:Q").AutoFit



Next ws



End Sub
