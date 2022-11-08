Attribute VB_Name = "Module1"
Sub Stock_Analysis_NGirdharry()

'Define Variables

'Define Ticker
Dim Stock_Ticker As String

'Define Year Openning Price
Dim year_open As Double

'Define Year Closing Price
Dim year_close As Double

'Define the Yearly Change of the opening and closing price
Dim Yearly_Change As Double

'Define the total Stock Volume of the stock
Dim Total_Vol As Double

'Define the Percentage Change between the opening and closing price
Dim Percentage_Change As Double

'Define the Beginning of the dataset (Used to assign the beginning/ending numbers when running the code)
Dim data_start As Integer

'Define the worksheet
Dim ws As Worksheet

'Define the next ticker (Used to assign the beginning/ending numbers when running the code)
Dim next_ticker As Double

'Assign Column Names
For Each ws In Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Assign beginning and ending to numbers/code
data_start = 2
next_ticker = 1
Total_Vol = 0
EndRow = ws.Cells(Rows.Count, "a").End(xlUp).Row

'Assign loops
For i = 2 To EndRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Stock_Ticker = ws.Cells(i, 1).Value
next_ticker = next_ticker + 1
year_open = ws.Cells(next_ticker, 3).Value
year_close = ws.Cells(i, 6).Value

For j = next_ticker To i
Total_Vol = Total_Vol + ws.Cells(j, 7).Value
Next j

If year_open = 0 Then
Percentage_Change = year_close
Else
Yearly_Change = year_close - year_open
Percentage_Change = Yearly_Change / year_open
End If

'Assign values to cells in worksheet (NEW Table)
ws.Cells(data_start, 9).Value = Stock_Ticker
ws.Cells(data_start, 10).Value = Yearly_Change

'Percentage assignment for Percentage Change between opening price and new price in column J
ws.Cells(data_start, 11).Value = Percentage_Change
ws.Cells(data_start, 11).NumberFormat = "0.00%"
ws.Cells(data_start, 12).Value = Total_Vol

'Continue to next row and reset variables to 0
data_start = data_start + 1
Total_Vol = 0
Yearly_Change = 0
Percent_Change = 0
next_ticker = i
End If
Next i

'Colour formatting for percentage change in column J
jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
For j = 2 To jEndRow
If ws.Cells(j, 10) > 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4
Else
ws.Cells(j, 10).Interior.ColorIndex = 3
End If
Next j

'Continue onto next worksheet
Next ws
End Sub
