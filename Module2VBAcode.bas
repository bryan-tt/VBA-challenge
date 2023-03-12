Attribute VB_Name = "Module1"
Sub stockAnalysis():
    
    For Each ws In Worksheets
        Dim j As Integer
        Dim x As Integer
        Dim openPrice As Double
        Dim closePrice As Double
        Dim totalVolume As LongLong
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        totalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row
        j = 1
        x = 1
        openPrice = ws.Cells(2, 3).Value
        
        For i = 2 To totalRecords
            totalVolume = ws.Cells(i, 7).Value + totalVolume
            
            If (ws.Cells(i, 1).Value <> ws.Cells(x, 9).Value) Then
                x = x + 1
                ws.Cells(x, 9).Value = ws.Cells(i, 1).Value
            End If
            
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                closePrice = ws.Cells(i, 6).Value
                j = j + 1
                ws.Cells(j, 10).Value = closePrice - openPrice
                If (closePrice - openPrice >= 0) Then
                    ws.Cells(j, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(j, 10).Interior.Color = vbRed
                End If
                
                ws.Cells(j, 11).Value = (closePrice - openPrice) / openPrice
                ws.Cells(j, 11).NumberFormat = "0.00%"
                
                ws.Cells(j, 12).Value = totalVolume
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
            End If
        Next i
        
        Dim biggestInc As Double
        Dim biggestIncTik As String
        Dim biggestDec As Double
        Dim biggestDecTik As String
        Dim biggestVol As LongLong
        Dim biggestVolTik As String
        
        biggestInc = 0#
        biggestDec = 0#
        biggestVol = 0
        
        For i = 2 To 50000
            If (IsEmpty(ws.Cells(i, 9).Value)) Then
                Exit For
            End If
            
            If (ws.Cells(i, 11).Value > biggestInc) Then
                biggestInc = ws.Cells(i, 11).Value
                biggestIncTik = ws.Cells(i, 9).Value
            End If
            
            If (ws.Cells(i, 11).Value < biggestDec) Then
                biggestDec = ws.Cells(i, 11).Value
                biggestDecTik = ws.Cells(i, 9).Value
            End If
            
            If (ws.Cells(i, 12).Value > biggestVol) Then
                biggestVol = ws.Cells(i, 12).Value
                biggestVolTik = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        ws.Cells(2, 16).Value = biggestIncTik
        ws.Cells(2, 17).Value = biggestInc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = biggestDecTik
        ws.Cells(3, 17).Value = biggestDec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = biggestVolTik
        ws.Cells(4, 17).Value = biggestVol
        
        ws.Range("I1:Q4").Columns.AutoFit
        
        
    Next
End Sub
