Attribute VB_Name = "Module1"
Sub WallStreet()

  Dim wallSheet As Worksheet
  
  For Each wallSheet In ThisWorkbook.Worksheets
    wallSheet.Activate
  
    Dim lRow As Long
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim disRow As Integer
    disRow = 1
    
    Cells(disRow, 9).Value = ("Ticker")
    Cells(disRow, 10).Value = ("Yearly Change")
    Cells(disRow, 11).Value = ("Percent Change")
    Cells(disRow, 12).Value = ("Total Stock Volume")

    Dim curTick As String
    Dim prevTick As String
    Dim disTick As String
    Dim curOpen As Double
    Dim stOpen As Double
    Dim stClose As Double
    
    Dim curRow As Long
    curRow = 3
    
    stOpen = Cells(curRow - 1, 3).Value
    totVol = Cells(curRow - 1, 7).Value
        
    For curRow = 3 To (lRow + 1)
        curTick = Cells(curRow, 1).Value
        curOpen = Cells(curRow, 3).Value
        curClose = Cells(curRow, 6).Value
        curVol = Cells(curRow, 7).Value
        prevTick = Cells(curRow - 1, 1).Value
        
        If curTick <> prevTick Then
            disRow = disRow + 1
            Cells(disRow, 9).Value = disTick
            Cells(disRow, 10).Value = (stClose - stOpen)
                If ((stClose - stOpen) < 0) Then
                    Cells(disRow, 10).Interior.ColorIndex = 3
                ElseIf ((stClose - stOpen) > 0) Then
                    Cells(disRow, 10).Interior.ColorIndex = 4
                End If
            
                If stOpen <> 0 Then
                    Cells(disRow, 11).Value = (stClose - stOpen) / stOpen
                Else
                    Cells(disRow, 11).Value = 0
                End If
            Cells(disRow, 11).NumberFormat = "0.00%"
            Cells(disRow, 12).Value = totVol

            stOpen = curOpen
            totVol = Cells(curRow, 7).Value
        Else
            disTick = curTick
            stClose = curClose
            totVol = totVol + curVol
        End If

    Next curRow
    
    
    'Bonus
    Dim lsumRow As Integer
    lsumRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    Dim bull As Double
    Dim bullRow As Integer
    bull = Cells(2, 11).Value
    
    For sumRow = 2 To lsumRow
        If Cells(sumRow, 11).Value >= bull Then
            bull = Cells(sumRow, 11)
            bullRow = sumRow
        End If
    Next sumRow
    
    Dim bear As Double
    Dim bearRow As Integer
    bear = Cells(2, 11).Value
    
    For sumSow = 2 To lsumRow
        If Cells(sumSow, 11).Value <= bear Then
            bear = Cells(sumSow, 11)
            bearRow = sumSow
        End If
    Next sumSow
    
    Dim big As LongLong
    Dim bigRow As Integer
    big = Cells(2, 12).Value
    
    For sumTow = 2 To lsumRow
        If CLngLng(Cells(sumTow, 12).Value) >= big Then
            big = CLngLng(Cells(sumTow, 12).Value)
            bigRow = sumTow
        End If
    Next sumTow
    
    Range("O1").Value = "Ticker"
    Range("O2").Value = Cells(bullRow, 9).Value
    Range("O3").Value = Cells(bearRow, 9).Value
    Range("O4").Value = Cells(bigRow, 9).Value
    Range("P1").Value = "Value"
    Range("P2").NumberFormat = "0.00%"
    Range("P2").Value = bull
    Range("P3").NumberFormat = "0.00%"
    Range("P3").Value = bear
    Range("P4").Value = big
    Columns("J:N").AutoFit

Next
  
End Sub
