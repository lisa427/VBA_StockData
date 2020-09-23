Attribute VB_Name = "Module1"
Sub SheetLoop():

Dim Sheet As Worksheet

Application.ScreenUpdating = False

For Each Sheet In Worksheets
    Sheet.Select
    Call TickerCalc
Next

Application.ScreenUpdating = True

End Sub

Sub TickerCalc():

'Used Variant type due to issue with my version of Excel for Mac

Dim Ticker As Variant
Dim OpenPrice As Variant
Dim OpenDate As Variant
Dim ClosePrice As Variant
Dim CloseDate As Variant
Dim Volume As Variant
Dim Counter As Variant
Dim LastRow As Variant
Dim YearlyChg As Variant
Dim PctChg As Variant
Dim LrgIncrease As Variant
Dim LrgIncreaseTicker As Variant
Dim LrgDecrease As Variant
Dim LrgDecreaseTicker As Variant
Dim LrgVolume As Variant
Dim LrgVolumeTicker As Variant

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range("A1:G" & LastRow).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

Ticker = Cells(2, 1).Value
OpenDate = Cells(2, 2).Value
OpenPrice = Cells(2, 3).Value
CloseDate = Cells(2, 2).Value
ClosePrice = Cells(2, 6).Value
Volume = Cells(2, 7).Value
Counter = 2

For i = 3 To (LastRow + 1)

If Cells(i, 1).Value = Ticker Then

    If Cells(i, 2).Value > CloseDate Then
    
    CloseDate = Cells(i, 2).Value
    ClosePrice = Cells(i, 6).Value
    
    ElseIf Cells(i, 2).Value < OpenDate Then
    
    OpenDate = Cells(i, 2).Value
    OpenPrice = Cells(i, 3).Value
    
    End If
    
Volume = Volume + Cells(i, 7).Value

Else
    YearlyChg = ClosePrice - OpenPrice
    
    If OpenPrice <> 0 Then
        PctChg = (ClosePrice - OpenPrice) / OpenPrice
    Else
        PctChg = 0
    End If
    
    Cells(Counter, 9).Value = Ticker
    Cells(Counter, 10).Value = YearlyChg
    Cells(Counter, 11).Value = PctChg
    Cells(Counter, 11).NumberFormat = "0.00%"
    Cells(Counter, 12).Value = Volume
    
    If Cells(Counter, 10).Value > 0 Then
        Cells(Counter, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(Counter, 10).Value < 0 Then
        Cells(Counter, 10).Interior.ColorIndex = 3
    
    End If
    
    If PctChg > LrgIncrease Then
        LrgIncrease = PctChg
        LrgIncreaseTicker = Ticker
        
    ElseIf PctChg < LrgDecrease Then
        LrgDecrease = PctChg
        LrgDecreaseTicker = Ticker
        
    End If
    
    If Volume > LrgVolume Then
        LrgVolume = Volume
        LrgVolumeTicker = Ticker
    End If
      
    Ticker = Cells(i, 1).Value
    OpenDate = Cells(i, 2).Value
    OpenPrice = Cells(i, 3).Value
    CloseDate = Cells(i, 2).Value
    ClosePrice = Cells(i, 6).Value
    Volume = Cells(i, 7).Value
    Counter = Counter + 1
    
End If

Next i

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

Cells(2, 16).Value = LrgIncrease
Cells(2, 16).NumberFormat = "0.00%"
Cells(2, 15).Value = LrgIncreaseTicker
Cells(3, 16).Value = LrgDecrease
Cells(3, 16).NumberFormat = "0.00%"
Cells(3, 15).Value = LrgDecreaseTicker
Cells(4, 16).Value = LrgVolume
Cells(4, 15).Value = LrgVolumeTicker

End Sub



