Sub TickAndChangeFinal():
' ----DECLARE VARIABLES----
    ' tickers
    Dim CurrTick As String
    Dim NextTick As String
    
    ' open, close, diff, volume
    Dim CurrOpen As Double
    Dim CurrClose As Double
    Dim Diff As Double
    Dim VolSum As Double
    
    ' current output table row
    Dim RowNum As Integer
    
    ' challenge variables
    Dim BiggestIncrease As Double
    Dim BiggestDecrease As Double
    Dim BiggestVolume As Double
    
    Dim BIticker As String
    Dim BDticker As String
    Dim BVticker As String
    
    Dim CurrPerc As Double
    Dim CurrVol As Double
    Dim ChallengeTicker As String
    
' ----LOOP THROUGH SHEETS----
    For Each ws In Worksheets

        
        
        ' set open, row num, volume back to base value
        CurrOpen = ws.Cells(2, 3).Value
        RowNum = 2
        VolSum = 0
        NextTick = ws.Cells(2, 1).Value
        
        ' set challenge values to 0
        BiggestIncrease = 0
        BiggestDecrease = 0
        BiggestVolume = 0

        ' header
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Year Change"
        ws.Cells(1, 11) = "% Change"
        ws.Cells(1, 12) = "Total Volume"
        
    ' ----LOOP THROUGH ROWS IN SHEET----
        For i = 2 To ws.Range("A:A").End(xlDown).Row
            
        ' redefine current & next tick
            CurrTick = NextTick
            NextTick = ws.Cells(i + 1, 1).Value
            
        ' add to volume
            VolSum = VolSum + ws.Cells(i, 7).Value
            
            
        ' LAST TICK
            If CurrTick <> NextTick Then
            ' redefine close
                CurrClose = ws.Cells(i, 6)
            
            ' ---INFO INTO TABLE---
                ' ticker
                ws.Cells(RowNum, 9) = CurrTick
                
                ' year change w/formatting
                Diff = CurrClose - CurrOpen
                ws.Cells(RowNum, 10) = Diff
                If Diff > 0 Then
                    ws.Cells(RowNum, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(RowNum, 10).Interior.Color = vbRed
                End If
                
                ' % change
                If CurrOpen <> 0 Then
                    ws.Cells(RowNum, 11) = Diff / CurrOpen
                End If
                ws.Cells(RowNum, 11).Style = "Percent"
                
                ' volume
                ws.Cells(RowNum, 12) = VolSum
            
            ' redefine open
                CurrOpen = ws.Cells(i + 1, 3)
                
            ' increment table num
                RowNum = RowNum + 1
                
            ' reset volume
                VolSum = 0
            
            End If
        Next i
        
        ' make % change column into right format
        'ws.Range("K:K").Style.NumberFormat = "0.00%"
        
        
        
' ----CHALLENGE TABLE----
        
        ' loop through ticker summary
        For j = 2 To ws.Range("I:I").End(xlDown).Row
            CurrPerc = ws.Cells(j, 11).Value
            CurrVol = ws.Cells(j, 12).Value
            ChallengeTicker = ws.Cells(j, 9).Value
            
            ' Biggest % Increase
            If BiggestIncrease < CurrPerc Then
                BiggestIncrease = CurrPerc
                BIticker = ChallengeTicker
            ' Biggest % Decrease
            ElseIf BiggestDecrease > CurrPerc Then
                BiggestDecrease = CurrPerc
                BDticker = ChallengeTicker
            End If
            
            ' Biggest Volume
            If BiggestVolume < CurrVol Then
                BiggestVolume = CurrVol
                BVticker = ChallengeTicker
            End If
        Next j
        
        ' Make table
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(2, 15) = BIticker
        ws.Cells(2, 16) = BiggestIncrease
        ws.Cells(2, 16).Style = "Percent"
        
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(3, 15) = BDticker
        ws.Cells(3, 16) = BiggestDecrease
        ws.Cells(3, 16).Style = "Percent"
        
        ' This effectively sets the format of all cells in the "Percent"
        ' style to .00% rather than the standard rounding to nearest integer.
        ws.Cells(3, 16).Style.NumberFormat = "0.00%"
        
        ws.Cells(4, 14) = "Greatest Total Volume"
        ws.Cells(4, 15) = BVticker
        ws.Cells(4, 16) = BiggestVolume
        

        'ws.Range("B:G").Style.NumberFormat = "General"
    Next ws
End Sub

