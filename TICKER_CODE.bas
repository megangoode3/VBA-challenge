Attribute VB_Name = "TICKER_CODE"
Sub Ticker_Summary():


For Each WS In Worksheets

'Set variables
Dim Worksheetname As String
Dim i As Integer
Dim j As Integer
Dim tickercount As Integer
Dim percentchange As Double

'Worksheet Name
Worksheetname = WS.Name

'Set the first column and look for the last row in that column
Dim lastrow As Integer
lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Set up summary titles
WS.Cells(1, 9).Value = "Ticker"
WS.Cells(1, 10).Value = "Yearly Change"
WS.Cells(1, 11).Value = "Percent Change"
WS.Cells(1, 12).Value = "Total Stock Volume"

'tickercount starting with the first row of data
tickercount = 2
j = 2

'Loop through rows in first column
For i = 2 To lastrow

'Check for different value/ticker. If not then
If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

'Ticker for column I (9)
WS.Cells(tickercount, 9).Value = WS.Cells(i, 1).Value

'Yearly Change for column J (10) (Close value in F(6) - open in C(3))
WS.Cells(tickercount, 10).Value = WS.Cells(i, 6).Value - WS.Cells(j, 3).Value

    'Format/Color for negative values, RED
    If WS.Cells(tickercount, 10).Value < 0 Then
    WS.Cells(tickercount, 10).Interior.ColorIndex = 3

    'Format if positive, green
    Else
    WS.Cells(tickercount, 10).Interior.ColorIndex = 4

    End If

    'Percent change for column K (11)
    If WS.Cells(j, 3).Value <> 0 Then
    percentchange = ((WS.Cells(i, 6).Value - WS.Cells(j, 3).Value) / WS.Cells(j, 3).Value)

    Else
    WS.Cells(tickercount, 11).Value = Format(0, "percent")

    End If

    'Total Stock volume for column L (12)
    WS.Cells(tickercount, 12).Value = WorksheetFunction.Sum(Range(WS.Cells(j, 7), WS.Cells(i, 7)))

    'increase tickercount to move to next
    tickercount = tickercount + 1

    'Start the new row for the ticker Summary
    j = i + 1

    End If

'Create loop through all If statements to build on Summary section
Next i


'Set variables
'GI as Greatest Increase, GD Greatest Decrease, GTV as greatest total volume
Dim GI As Double
Dim GD As Double
Dim GTV As Double

'Use percent change column K (11) as value for GI and GD
GI = WS.Cells(2, 11).Value
GD = WS.Cells(2, 11).Value
'Use stock volume column J (12) as value for GTV
GTV = WS.Cells(2, 12).Value

'Set second "summary" section
'Column Headers
WS.Cells(1, 16).Value = "Ticker"
WS.Cells(1, 17).Value = "Value"
'"Row headers"
WS.Cells(2, 15).Value = "Greatest % Increase"
WS.Cells(3, 15).Value = "Greatest % Decrease"
WS.Cells(4, 15).Value = "Greatest Total Volume"

'Find last row in NEW Ticker column I (9)
Newlastrow = WS.Cells(Rows.Count, 9).End(xlUp).Row

'Loop through new last row
For i = 2 To Newlastrow
            
    'GI, check for the greatest value.  If new row is greater, then use that, if not, stay with current GI
    If WS.Cells(i, 11).Value > GI Then
    GI = WS.Cells(i, 11).Value
    WS.Cells(2, 16).Value = WS.Cells(i, 9).Value

    Else

    GI = GI

    End If

    'GD, same format as GI
    If WS.Cells(i, 11).Value < GD Then
    GD = WS.Cells(i, 11).Value
    WS.Cells(3, 16).Value = WS.Cells(i, 9).Value

    Else

    GD = GD

    End If

    'GTV same as GI and GD
    If WS.Cells(i, 12).Value < GTV Then
    GTV = WS.Cells(i, 12).Value
    WS.Cells(4, 16).Value = WS.Cells(i, 9).Value

    Else

    GTV = GTV

    End If

'Print results in cells
WS.Cells(2, 17).Value = Format(GI, "Percent")
WS.Cells(3, 17).Value = Format(GD, "Percent")
WS.Cells(4, 17).Value = Format(GTV, "Scientific")

Next i
                
Next WS


End Sub
