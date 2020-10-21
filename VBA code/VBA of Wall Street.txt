Sub Run_All_Worksheets_At_Once()
'loop to run through all available worksheets
For x = 1 To Worksheets.Count
    ActiveWorkbook.Sheets(x).Activate
    'after activating each sheet, runs both subroutines on each sheet
    Format_Worksheet
    Run_Calculations
Next x
'after all sheets are populated, jumps back to first sheet
ActiveWorkbook.Sheets(1).Activate

End Sub

Sub Format_Worksheet()
'establishes the necessary headers and initial formatting via VBA code
Range("I1") = "Ticker"
Range("I1").ColumnWidth = 7
Range("J1") = "Yearly Change"
Range("J1").ColumnWidth = 12
Range("K1") = "Percent Change"
Range("K1").ColumnWidth = 13.5
Columns(11).NumberFormat = "0.00%"
Range("L1") = "Total Stock Volume"
Range("L1").ColumnWidth = 20
'Challenge Formatting
Range("O1") = "Ticker"
Range("O1").ColumnWidth = 7
Range("P1") = "Value"
Range("P1").ColumnWidth = 20
Range("N2") = "Greatest % Increase"
Range("P2").NumberFormat = "0.00%"
Range("N3") = "Greatest % Decrease"
Range("P3").NumberFormat = "0.00%"
Range("N4") = "Greatest Total Volume"
Range("N1").ColumnWidth = 19.5

End Sub

Sub Run_Calculations()

Dim TotalStockVolume As Double
Dim MinDateRow As Long
Dim TickerCounter As Long
Dim Green As Integer
Dim Red As Integer

'Challenge Variables
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestStockVolume As Double

Dim GreatestIncreaseRow As Long
Dim GreatestDecreaseRow As Long
Dim GreatestStockVolumeRow As Long

'avoids issues with headers
TickerCounter = 2
MinDateRow = 2
'notes colorindex for cleaner code
Green = 4
Red = 3

'all row calculations go under here
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    'sets code for getting to end of Individual Ticker - does all calculations at final entry in ticker block
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'does calculations inline and sets correspoinding cell values
        Cells(TickerCounter, 9).Value = Cells(i, 1).Value
        Cells(TickerCounter, 10).Value = Cells(i, 6).Value - Cells(MinDateRow, 3).Value
        On Error Resume Next
        Cells(TickerCounter, 11).Value = (Cells(i, 6).Value - Cells(MinDateRow, 3).Value) / Cells(MinDateRow, 3).Value
        Cells(TickerCounter, 12).Value = TotalStockVolume + Cells(i, 7).Value
       'determines if positive or negative value and assigns conditional formatting
        If Cells(TickerCounter, 10).Value >= 0 Then
            Cells(TickerCounter, 10).Interior.ColorIndex = Green
                Else
                    Cells(TickerCounter, 10).Interior.ColorIndex = Red
                End If
        'Challenge Code
        'looks for largest increase, saves value and row index to variables
        If Cells(TickerCounter, 11).Value > GreatestIncrease Then
            GreatestIncrease = Cells(TickerCounter, 11)
            GreatestIncreaseRow = TickerCounter
        End If
        'looks for largest decrease, saves value and row index to variables
        If Cells(TickerCounter, 11).Value < GreatestDecrease Then
            GreatestDecrease = Cells(TickerCounter, 11)
            GreatestDecreaseRow = TickerCounter
        End If
        'looks for largest stock volume, saves value and row index to variables
        If Cells(TickerCounter, 12).Value > GreatestStockVolume Then
            GreatestStockVolume = Cells(TickerCounter, 12)
            GreatestStockVolumeRow = TickerCounter
        End If
        'after documenting all relative data, increments "start" variables for next block
        MinDateRow = i + 1
        TickerCounter = TickerCounter + 1
        'after docmenting all relative data, resets Stock volume for next block
        TotalStockVolume = 0
        
    'if cells are the same, add the stock volume to the sum
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        'standard debug else - throws message box
                Else
                    MsgBox ("Error Code:1 - Error in IF Statement")
        
                End If
    
Next i
'populates challenge fields with values after all calculations are complete
Range("P2").Value = GreatestIncrease
Range("P3").Value = GreatestDecrease
Range("P4").Value = GreatestStockVolume
'pulls Ticker from list based on previously indexed rows
Range("O2").Value = Cells(GreatestIncreaseRow, 9).Value
Range("O3").Value = Cells(GreatestDecreaseRow, 9).Value
Range("O4").Value = Cells(GreatestStockVolumeRow, 9).Value

End Sub

