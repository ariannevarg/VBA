Attribute VB_Name = "Module1"
Sub StockMarketAnalysis1()
' Easy - kindaish, some code for other levels too

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
' Total Stock Volume
Total = 0
 
' Row summary for the table
Dim Row_Sum As Integer
Row_Sum = 2
 
Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
For i = 2 To LastRow
  Total = Total + Cells(i, 7)
 
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
 
  Cells(Row_Sum, 9).Value = Ticker
  Cells(Row_Sum, 12).Value = Total
 
  Row_Sum = Row_Sum + 1
 
  Total = 0
 
  End If
 
 Next i

End Sub
Sub StockMarketAnalysis2()
' Moderate

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
Dim Open_Column As Double
Dim Close_Column As Double
Dim Yearly_Change As Double
Dim Ticker_Name As String
Dim Percent_Change As Double

Dim Total_Volume As Long
Volume = 0

Dim Row As Double
Row = 2

Dim Column As Integer
Column = 1

Dim i As Long
        
' Open Price Column
    Open_Column = Cells(2, Column + 2).Value
        
For i = 2 To LastRow
        
        
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        Ticker_Name = Cells(i, Column).Value
        Range("I1").Value = "Ticker"

' Close Price Column
    Close_Column = Cells(i, Column + 5).Value

' Yearly Change Column
    Yearly_Change = Close_Column - Open_Column
    Range("J1").Value = "Yearly Change"

' Percent Change Column
    If (Open_Column = 0 And Close_Column = 0) Then
        Percent_Change = 0
    
    ElseIf (Open_Column = 0 And Close_Column <> 0) Then
        Percent_Change = 1
    
    Else
        Percent_Change = Yearly_Change / Open_Column
        Range("K1").Value = "Percent Change"
        Range("K1").NumberFormat = "0.00%"
                
    End If

' Total Volume Column
    Volume = Volume + Cells(i, Column + 6).Value
    Range("L1").Value = "Total Stock Volume"


    Open_Column = Cells(i + 1, Column + 2)
    Volume = 0
    
' Add a row to the summary table
    Row = Row + 1

    Else
    Volume = Volume + Cells(i, Column + 6).Value
    
        
        End If
    Next i
        
' Determine the Last Row of Yearly Change per work sheet
    YearlyChange_LastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row

' Set the Cell Colors for the Yearly Change Column
    For j = 2 To YearlyChange_LastRow
    
    If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
    Cells(j, Column + 9).Interior.ColorIndex = 4
    
    ElseIf Cells(j, Column + 9).Value < 0 Then
    
    Cells(j, Column + 9).Interior.ColorIndex = 3
        
        End If
        
    Next j
       
Next ws


End Sub
