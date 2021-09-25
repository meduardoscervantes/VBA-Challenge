Attribute VB_Name = "Module1"
Sub TickerCalculations()
'This macro will take information from tickers and categorize and calculate yearly differences
'A = 1 <ticker>
'C = 3 <open>
'F = 6 <close>
'G = 7 <vol>
'I = 9 Ticker
'J = 10 Yearly Change
'K = 11 Percentage Change
'L = 12 Total Stock Volume
'O = 15 Greatest
'P = 16 (Bonus) Ticker
'Q = 17 Value
    
    'Declare local variables
    Dim i As Double 'For loop counter
    Dim iPosCounter As Double 'I column pos to identify working ticker
    Dim tempTotalStock As Double 'Holder for total stock volume
    Dim startPos As Double 'This value tells us when the new ticker started counting
    Dim tempMax As Double 'temp to find bonus question max
    Dim tempMaxPos As Double 'temp max pos for bonus questions
    Dim tempMin As Double  'temp to find bonus min
    Dim tempMinPos As Double 'temp min pos for bonus question
    Dim tempMaxVol As Double
    Dim tempMaxVolPos As Double
    
    Dim j As Integer
    
    For j = 1 To Worksheets.Count
        Worksheets(j).Activate
        'Declare column titles
        If IsEmpty(Cells(1, 9).Value) Then
            Cells(1, 9).Value = "Ticker"
        End If
        If IsEmpty(Cells(1, 10).Value) Then
            Cells(1, 10).Value = "Yearly Change"
        End If
        If IsEmpty(Cells(1, 11).Value) Then
            Cells(1, 11).Value = "Percentage Changed"
        End If
        If IsEmpty(Cells(1, 12).Value) Then
            Cells(1, 12).Value = "Total Stock Volume"
        End If
        
        'Assign first working Ticker and place into cell
        iPosCounter = 2
        Cells(iPosCounter, 9).Value = Cells(2, 1).Value
        tempTotalStock = 0
        startPos = 2
        
        'Populate the data for each ticker
        For i = 2 To Rows.Count
            'Check ticker at i and working ticker
            If Cells(i, 1).Value = Cells(iPosCounter, 9).Value Then
                tempTotalStock = tempTotalStock + Cells(i, 7).Value
            Else
                'Display the total stock for working ticker
                Cells(iPosCounter, 12).Value = tempTotalStock
                'Reset totalStock
                tempTotalStock = 0
                'Calculate yearly change
                Cells(iPosCounter, 10).Value = Cells((i - 1), 6).Value - Cells(startPos, 3).Value
                'Paint in the background
                If Cells(i - 1, 10).Value < 0 Then
                    Cells(i - 1, 10).Interior.Color = RGB(255, 0, 0) 'Change color to red
                ElseIf Cells(i - 1, 10).Value > 0 Then
                    Cells(i - 1, 10).Interior.Color = RGB(0, 255, 0) 'Change color to green
                End If
                'Calculate percentage change
                If Cells(startPos, 3).Value > 0 Then
                    Cells(iPosCounter, 11).Value = Cells(iPosCounter, 10).Value / Cells(startPos, 3).Value * 100 & "%"
                Else
                    Cells(iPosCounter, 11).Value = "0%"
                End If
                'updadte iPosCounter
                iPosCounter = iPosCounter + 1
                'update working ticker
                Cells(iPosCounter, 9).Value = Cells(i, 1).Value
                'update starting position
                startPos = i
            End If
        Next i
        
        'Declare Empty cells titles for bonus questions
        If IsEmpty(Cells(1, 16).Value) Then
            Cells(1, 16).Value = "Ticker"
        End If
        If IsEmpty(Cells(1, 17).Value) Then
            Cells(1, 17).Value = "Value"
        End If
        If IsEmpty(Cells(2, 15).Value) Then
            Cells(2, 15).Value = "Greatest % Increase"
        End If
        If IsEmpty(Cells(3, 15).Value) Then
            Cells(3, 15).Value = "Greatest % Decrease"
        End If
        If IsEmpty(Cells(4, 15).Value) Then
            Cells(4, 15).Value = "Greatest Total Volume"
        End If
        
        'Find greatest % increase, greatest % decrease, and Max Total Volume
        tempMax = 0
        tempMin = 0
        tempMaxVol = 0
        For i = 2 To Rows.Count
            If Not IsEmpty(Cells(i, 11).Value) And Cells(i, 11).Value > tempMax Then
                tempMax = Cells(i, 11).Value
                tempMaxPos = i
            End If
            If Not IsEmpty(Cells(i, 11).Value) And Cells(i, 11).Value < tempMin Then
                tempMin = Cells(i, 11).Value
                tempMinPos = i
            End If
            If Not IsEmpty(Cells(i, 12).Value) And Cells(i, 12).Value > tempMaxVol Then
                tempMaxVol = Cells(i, 12).Value
                tempMaxVolPos = i
            End If
        Next i
        Cells(2, 16).Value = Cells(tempMaxPos, 9).Value
        Cells(2, 17).Value = tempMax & "%"
        Cells(3, 16).Value = Cells(tempMinPos, 9).Value
        Cells(3, 17).Value = tempMin & "%"
        Cells(4, 16).Value = Cells(tempMaxVolPos, 9).Value
        Cells(4, 17).Value = tempMaxVol
    Next j
End Sub

