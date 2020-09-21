'Code works for all sheets, must call sub before running data

Sub AstanehVBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call MultipleSheets ' this will duplicate efforts on multiple worksheets
    Next ws
End Sub

Sub MultipleSheets()
    Debug.Print ActiveSheet.Name
    Call StockFormat ' enter both formatting and coding subs on this line
    Call stock
End Sub

Sub StockFormat()

' Begin by labeling columns

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' Format text

Range("A1:L1").ColumnWidth = 17
Range("A1:P1").Font.Bold = True
Range("A1:P1").Interior.ColorIndex = 56
Range("A1:P1").Font.ColorIndex = 2
Range("A1:P705714").HorizontalAlignment = xlCenter
Range("A1:P705714").VerticalAlignment = xlCenter
Range("A1:P705714").Font.Size = 10
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Range("N2:N4").ColumnWidth = 20
Range("N2:N4").HorizontalAlignment = xlCenter
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

End Sub

Sub stock()
Dim firstTicker As String
Dim homeworkTicker As String
Dim rowCounter As Double
Dim total As Double
Dim summary_row As Double
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double
Dim percentChange As Double

rowCounter = Cells(Rows.Count, "A").End(xlUp).Row
summary_row = 2
openPrice = Cells(2, 3).Value

For currentRow = 2 To rowCounter
    firstTicker = Cells(currentRow, 1).Value
    homeworkTicker = Cells(currentRow + 1, 1).Value
    total = total + Cells(currentRow, 7).Value
    
    If firstTicker <> homeworkTicker Then
        Cells(summary_row, "I").Value = firstTicker 'ticker
        Cells(summary_row, "L").Value = total 'total volume
        closePrice = Cells(currentRow, 6).Value 'yearly change
        
        yearlyChange = closePrice - openPrice
        Cells(summary_row, 10).Value = yearlyChange
        Cells(summary_row, 10).NumberFormat = "0.00"
        
        
    If yearlyChange > 0 Then ' format
        Cells(summary_row, 10).Interior.ColorIndex = 4
    
    ElseIf yearlyChange < 0 Then
        Cells(summary_row, 10).Interior.ColorIndex = 3
    
    Else
        Cells(summary_row, 10).Interior.ColorIndex = 0

End If
    
    If (openPrice = 0 And closePrice = 0) Then
        percentChange = 0
    
    ElseIf (openPrice = 0 And closePrice <> 0) Then
        percentChange = 1
    
    Else
        percentChange = (yearlyChange) / (openPrice)
        Cells(summary_row, 11).Value = percentChange
        Cells(summary_row, 11).NumberFormat = "0.0%"
        
End If
    
        total = 0
        openPrice = Cells(currentRow + 1, 3).Value
        summary_row = summary_row + 1
        
End If

Next currentRow

' Challenges
 
greatestIncrease = Application.WorksheetFunction.Max(Range("K2:K800000"))
    Cells(2, 16).Value = greatestIncrease
    Cells(2, 16).NumberFormat = "0.0%"

greatestDecrease = Application.WorksheetFunction.Min(Range("K2:K800000"))
    Cells(3, 16).Value = greatestDecrease
    Cells(3, 16).NumberFormat = "0.0%"

greatestVolume = Application.WorksheetFunction.Max(Range("L2:L800000"))
    Cells(4, 16).Value = greatestVolume

End Sub
