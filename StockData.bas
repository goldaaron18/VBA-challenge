Attribute VB_Name = "Module1"
Sub Stockdata()


Dim EndPrice As Double
Dim StartPrice As Double
Dim ws As Worksheet
Dim Ticker As String
Dim YearlyChange As Double

 
For Each ws In Worksheets
  
   GreatestVolume = 0
   GreatestIncrease = 0
   GreatestDecrease = 0
   TickerRow = 2
   Volume = 0
   StartPrice = ws.Range("C2").Value
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   

   ws.Range("I1").Value = "Ticker"
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percent Change"
   ws.Range("L1").Value = "Total Stock Volume"
   ws.Range("O1").Value = "Ticker"
   ws.Range("P1").Value = "Value"
   ws.Range("N2").Value = "Greatest % Increase"
   ws.Range("N3").Value = "Greatest % Decrease"
   ws.Range("N4").Value = "Greatest Total Volume"
   
 For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker = ws.Cells(i, 1).Value
      Volume = Volume + ws.Cells(i, 7).Value
      EndPrice = ws.Cells(i, 6).Value
      
      YearlyChange = EndPrice - StartPrice
      If EndPrice <> 0 Then
      PercentChange = YearlyChange / EndPrice
      Else
      PercentChange = 0
      End If
      
      ws.Range("I" & TickerRow).Value = Ticker
      ws.Range("J" & TickerRow).Value = YearlyChange
      ws.Range("K" & TickerRow).Value = PercentChange
      ws.Range("L" & TickerRow).Value = Volume
      If YearlyChange < 0 Then
      ws.Range("J" & TickerRow).Interior.Color = 255
      Else
      ws.Range("J" & TickerRow).Interior.Color = 5296274
      End If
      
      If PercentChange > GreatestIncrease Then
      ws.Range("O2").Value = Ticker
      ws.Range("P2").Value = PercentChange
      GreatestIncrease = PercentChange
      End If
      If PercentChange < GreatestDecrease Then
      ws.Range("O3").Value = Ticker
      ws.Range("P3").Value = PercentChange
      GreatestDecrease = PercentChange
      End If
      If Volume > GreatestVolume Then
      ws.Range("O4").Value = Ticker
      ws.Range("P4").Value = Volume
      GreatestVolume = Volume
      End If
      
      TickerRow = TickerRow + 1
      Volume = 0
      StartPrice = ws.Cells(i + 1, 3).Value
    
    Else
      Volume = Volume + ws.Cells(i, 7).Value
    End If
Next i

   With ws.Columns("K:K")
    .Style = "Percent"
    .NumberFormat = "0.00%"
   End With
   ws.Range("P2").Style = "percent"
   ws.Range("P3").Style = "percent"
   
Next ws

End Sub
