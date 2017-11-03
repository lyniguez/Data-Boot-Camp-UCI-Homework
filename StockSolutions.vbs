Sub ModerateStockAnalysis()

Dim ticker As String
Dim TotalChange As Double
Dim Days As Double
Dim DailyChange As Double
Dim Change As Double
Dim AvgDailyChange As Double
Dim Summary As Double
Dim tstart As Double
Dim PercentChange As Single
Dim LastRow As Double
Dim volume As Double

Dim WS As Worksheet
Dim OWS As Worksheet

'create new worksheet for results
Sheets.Add.Name = "Results"
Set WS = Sheets("Results")


'set results column headers and greatest amounts
Sheets("Results").Range("A1").Value = "Ticker"
Sheets("Results").Range("B1").Value = "Yearly Change"
Sheets("Results").Range("C1").Value = "Percent Change"
Sheets("Results").Range("D1").Value = "Avg Daily Change"
Sheets("Results").Range("E1").Value = "Total Stock Volume"
Sheets("Results").Range("G2").Value = "Greatest Volume"
Sheets("Results").Range("G5").Value = "Greatest % Increase"
Sheets("Results").Range("G8").Value = "Greatest % Decrease"
Sheets("Results").Range("G11").Value = "Greatest Avg. Change"

'set active worksheet
Set OWS = Sheets("Stock_data_2016")
OWS.Activate

'set initial values
tstart = 2
Summary = 2
DailyChange = 0
volume = 0
TotalChange = 0


'get last row of data
LastRow = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRow

' ticker value change, start inputing results
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'total volume
volume = volume + Cells(i, 7).Value

'total change calculation from final close to opening ticker
TotalChange = Cells(i, 6).Value - Cells(tstart, 3).Value

'percent change calculation
PercentChange = (TotalChange / Round(Cells(tstart, 3), 6))

'day count
Days = (i - tstart) + 1

'calculate average daily change
AvgDailyChange = (DailyChange + Cells(i, 4).Value - Cells(i, 5).Value) / Days

'stock ticker
ticker = Cells(i, 1).Value

'reset start date for next ticker
tstart = i + 1

Sheets("Results").Range("A" & Summary).Value = ticker

Sheets("Results").Range("B" & Summary).Value = TotalChange

Sheets("Results").Range("C" & Summary).Value = PercentChange

Sheets("Results").Range("D" & Summary).Value = AvgDailyChange

Sheets("Results").Range("E" & Summary).Value = volume

'color assignments for stock tickers
Select Case TotalChange

'green
Case Is > 0
Sheets("Results").Range("B" & Summary).Interior.ColorIndex = 4

'red
Case Is < 0
Sheets("Results").Range("B" & Summary).Interior.ColorIndex = 3

'no color
Case Else
Sheets("Results").Range("B" & Summary).Interior.ColorIndex = 0

End Select

'reset variables
ticker = 0
Days = 0
DailyChange = 0
AvgDailyChange = 0
volume = 0
Summary = Summary + 1

' ticker is the same, add results up for each variable
Else

volume = volume + Cells(i, 7).Value

TotalChange = TotalChange + (Cells(i, 6).Value - Cells(i, 3).Value)

Change = Cells(i, 4).Value - Cells(i, 5).Value

DailyChange = DailyChange + Change


End If

Next i


'Greatest Amounts
WS.Activate

'results last row
Dim resultsrow As Double
Dim gvolume As Double
Dim gpincrease As Double
Dim gpdecrease As Double
Dim gavchange As Double
Dim gticker As String
Dim gphticker As String
Dim gplticker As String
Dim gavticker As String

gvolume = 0
gpincrease = 0
gavchange = 0

resultsrow = Cells(Rows.Count, "A").End(xlUp).Row

'for loop to find greatest volume
For i = 2 To resultsrow

If Cells(i, 5) > gvolume Then

gvolume = Cells(i, 5)
gticker = Cells(i, 1)

End If

Next i

'for loop to find greatest % increase
For i = 2 To resultsrow

If Cells(i, 3) > gpincrease Then

gpincrease = Cells(i, 3)
gphticker = Cells(i, 1)

End If

Next i

'for loop to find greatest % decrease
gpdecrease = Range("C2").Value

For i = 2 To resultsrow

If Cells(i, 3) <= gpdecrease Then

gpdecrease = Cells(i, 3)
gplticker = Cells(i, 1)

End If

Next i

'for loop to find greatest average change
For i = 2 To resultsrow

If Cells(i, 4) > gavchange Then

gavchange = Cells(i, 4)
gavticker = Cells(i, 1)

End If

Next i

'print values
Sheets("Results").Range("H2").Value = gvolume
Sheets("Results").Range("I2").Value = gticker

Sheets("Results").Range("H5").Value = gpincrease
Sheets("Results").Range("I5").Value = gphticker

Sheets("Results").Range("H8").Value = gpdecrease
Sheets("Results").Range("I8").Value = gplticker

Sheets("Results").Range("H11").Value = gavchange
Sheets("Results").Range("I11").Value = gavticker

'formatting cells
Sheets("Results").Range("A:H").EntireColumn.AutoFit
Sheets("Results").Range("C:C").NumberFormat = "0.00%"
Sheets("Results").Range("H5").NumberFormat = "0.00%"
Sheets("Results").Range("H8").NumberFormat = "0.00%"

End Sub
