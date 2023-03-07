Sub StockAnalysis()
    
        Dim ws As Worksheet
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TickerRange As Range
        Dim TotalStockVolume As Double
        
        Dim MaxPercentChange As Double
        Dim MinPercentChange As Double
        Dim MaxVolume As Double
        
        Dim MaxTicker As String
        Dim MinTicker As String
        Dim MaxVolumeTicker As String
        
        Dim i As Long
        Dim j As Long
        Dim LastRow As Long
        
        
        For Each ws In ThisWorkbook.Worksheets
                
                OpenPrice = 0
                ClosePrice = 0
                YearlyChange = 0
                PercentChange = 0
                TotalStockVolume = 0
                
                MaxPercentChange = 0
                MinPercentChange = 0
                MaxVolume = 0
                
                MaxTicker = " "
                MinTicker = " "
                MaxVolumeTicker = " "
                
                TotalStockVolume = 0
                j = 2
                Start = 2
                LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
                
                    ws.Range("I1").Value = "Ticker"
                    ws.Range("J1").Value = "Yearly Change"
                    ws.Range("K1").Value = "Percent Change"
                    ws.Range("L1").Value = "Total Stock Volume"
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("Q1").Value = "Value"
                    ws.Range("O2").Value = "Greatest % Increase"
                    ws.Range("O3").Value = "Greatest % Decrease"
                    ws.Range("O4").Value = "Greatest Total Volume"
                        
                       For i = 2 To LastRow
                            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            
                            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)
                            
                            ClosePrice = ws.Cells(i, 6).Value
                            OpenPrice = ws.Cells(Start, 3).Value
                            YearlyChange = ClosePrice - OpenPrice
                            PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                            
                            ws.Range("I" & j).Value = ws.Cells(i, 1).Value
                            ws.Range("L" & j).Value = TotalStockVolume
                            ws.Range("J" & j).Value = YearlyChange
                            ws.Range("K" & j).Value = PercentChange
                            ws.Range("K" & j).NumberFormat = "0.00%"
                            Select Case YearlyChange
                                Case Is < 0
                                    ws.Range("J" & j).Interior.Color = vbRed
                                Case Is > 0
                                    ws.Range("J" & j).Interior.Color = vbGreen
                                   Case Else
                                    ws.Range("J" & j).Interior.ColorIndex = xlNone
                            End Select
                            
                            Start = i + 1
                            
                            j = j + 1
                            TotalStockVolume = 0
                            Else
                            TotalStockVolume = TotalStockVolume + Cells(i, 7)
    
                            End If
                        Next i
                        MaxPercentChange = WorksheetFunction.Max(Range("K2:K" & LastRow))
                        MinPercentChange = WorksheetFunction.Min(Range("K2:K" & LastRow))
                        MaxVolume = WorksheetFunction.Max(Range("L2:L" & LastRow))
                        
                        MaxTickerRow = WorksheetFunction.Match(MaxPercentChange, Range("K2:K" & LastRow), 0) + 1
                        MinTickerRow = WorksheetFunction.Match(MinPercentChange, Range("K2:K" & LastRow), 0) + 1
                        MaxVolumeRow = WorksheetFunction.Match(MaxVolume, Range("L2:L" & LastRow), 0) + 1
                        
                        MaxTicker = Cells(MaxTickerRow, 9).Value
                        MinTicker = Cells(MinTickerRow, 9).Value
                        MaxVolumeTicker = Cells(MaxVolumeRow, 9).Value
                        
                        ws.Range("Q2").Value = MaxPercentChange
                        ws.Range("Q2").NumberFormat = "0.00%"
                        ws.Range("Q3").Value = MinPercentChange
                        ws.Range("Q3").NumberFormat = "0.00%"
                        ws.Range("Q4").Value = MaxVolume
                        ws.Range("P2").Value = MaxVolumeTicker
                        ws.Range("P3").Value = MaxTicker
                        ws.Range("P4").Value = MinTicker
                        ws.Cells.EntireColumn.AutoFit
            Next ws
End Sub
