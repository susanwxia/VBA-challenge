'* Create a script that will loop through all the stocks for one year and output the following information.
'* The ticker symbol.
'* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'* The total stock volume of the stock.


Sub StockTrend()
Dim t, fst As Long
Dim OpenPri, ClosePri, PerChg As Variant
OpenPri = ActiveSheet.Range("C" & 2).Value
fst = 2
t = 2
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
        If ActiveSheet.Range("A" & i).Value <> ActiveSheet.Range("A" & i + 1).Value Then
        Cells(t, 9).Value = ActiveSheet.Range("A" & i).Value
        ClosePri = ActiveSheet.Range("F" & i).Value
        Cells(t, 10).Value = ClosePri - OpenPri
            If OpenPri = 0 Then
            Cells(t, 11).Value = Format(0, "0%")
            Else
            PerChg = Round((ClosePri - OpenPri) / OpenPri, 4)
            Cells(t, 11).Value = Format(PerChg, "0.00%")
            End If
        Cells(t, 12).Value = Application.WorksheetFunction.Sum(Range(Cells(fst, 7), Cells(i, 7)))
        OpenPri = ActiveSheet.Range("C" & i + 1).Value
        t = t + 1
        fst = i + 1
        End If
    Next i
ActiveSheet.Columns("I:L").AutoFit
End Sub

'---------------------------------------------------------------------------------------------------------------------------

'You should also have conditional formatting that will highlight positive change in green and negative change in red
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

Sub Formatter_GVal()
Dim GVol, GInc, GDec As Variant
Dim Tici, Ticd, Ticv As String
Dim LRowYr, GIncRow As Long
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
GInc = ActiveSheet.Range("K" & 2).Value
GDec = ActiveSheet.Range("K" & 2).Value
GVol = ActiveSheet.Range("L" & 2).Value
LRowYr = ActiveSheet.Cells(ActiveSheet.Rows.Count, 10).End(xlUp).Row
    For j = 2 To LRowYr
        If ActiveSheet.Range("J" & j).Value < 0 Then
        Range("J" & j).Interior.ColorIndex = 3
        ElseIf ActiveSheet.Range("J" & j).Value > 0 Then
        Range("J" & j).Interior.ColorIndex = 4
        Else
        Range("J" & j).Interior.ColorIndex = 2
        End If
   
        If GInc < ActiveSheet.Range("K" & j + 1).Value Then
        GInc = ActiveSheet.Range("K" & j + 1).Value
        Tici = ActiveSheet.Range("I" & j + 1).Value
        End If
        
        If GDec > ActiveSheet.Range("K" & j + 1).Value Then
        GDec = ActiveSheet.Range("K" & j + 1).Value
        Ticd = ActiveSheet.Range("I" & j + 1).Value
        End If
        
        If GVol < ActiveSheet.Range("L" & j + 1).Value Then
        GVol = ActiveSheet.Range("L" & j + 1).Value
        Ticv = ActiveSheet.Range("I" & j + 1).Value
        End If
               
    Next j
ActiveSheet.Range("Q2").Value = Format(GInc, "0.00%")
ActiveSheet.Range("Q3").Value = Format(GDec, "0.00%")
ActiveSheet.Range("Q4").Value = GVol
ActiveSheet.Range("P2").Value = Tici
ActiveSheet.Range("P3").Value = Ticd
ActiveSheet.Range("P4").Value = Ticv
ActiveSheet.Columns("O:Q").AutoFit
End Sub
