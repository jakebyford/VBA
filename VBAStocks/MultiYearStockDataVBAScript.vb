Sub MultiYearStockData()
Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
    
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    Dim lastrow As Long
    Dim writerow As Integer
    Dim tickercol As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim volume_col As Integer
    Dim max As Double
    Dim maxticker As String
    Dim min As Double
    Dim minticker As String
    Dim maxvolume As Double
    Dim maxvolticker As String
    Dim numerator As Double
    Dim denominator As Double
    
    
    
    tickercol = 1
    writerow = 2
    volumecol = 7
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   
    
    'Label Header/Title
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    'Begin loop
    For i = 1 To lastrow
            
         
        'If next ticker symbol is not the same
        If Cells(i + 1, tickercol).Value <> Cells(i, tickercol).Value Then
            
            'Grab next ticker symbol
            Cells(writerow, 9).Value = Cells(i + 1, tickercol).Value
           
                If (i <> 1) Then
                    
                    'Intialize Volume
                    Cells(writerow - 1, 12).Value = volume
                    volume = 0
                    
                    'Grab closing price of current price
                    close_price = Cells(i, 6).Value
            
                    'record the diff. between the tickers open & close prices
                    yearlychange = close_price - open_price
                    
                    'Record the percent difference
                    numerator = close_price - open_price
                    denominator = open_price
                    
    
                    
                    If numerator = 0 Then
                            numerator = 1
                        End If
                        
                    If denominator = 0 Then
                            denominator = 1
                        End If
                        
                    percentchange = numerator / denominator
                    
                        
                    'Format Color
                       If percentchange >= 0 Then
                        Cells(writerow - 1, 11).Interior.Color = vbGreen
                       Else: Cells(writerow - 1, 11).Interior.Color = vbRed
                       End If
                        
                
                    'Format Cells
                    Cells(writerow - 1, 10).Value = yearlychange
                    Cells(writerow - 1, 10).NumberFormat = "$0.00"
                    Cells(writerow - 1, 11).Value = percentchange
                    Cells(writerow - 1, 11).NumberFormat = "0.00%"
            
                End If
            
            'Grab new open price
            open_price = Cells(i + 1, 3).Value
            
            
            'Grabs next row
            writerow = writerow + 1
            
            
        End If
        'Take Volume
        volume = volume + Cells(i + 1, volumecol)
        
Next i
      
    max = Cells(2, 11).Value
    maxticker = Cells(2, 9).Value
    min = Cells(2, 11).Value
    minticker = Cells(2, 9).Value
    maxvolume = Cells(2, 12).Value
    maxvolticker = Cells(2, 9).Value
    
    
    
    
For j = 1 To writerow
    
    If Cells(j + 2, 11) > max Then
       
            max = Cells(j + 2, 11).Value
            maxticker = Cells(j + 2, 9).Value
            
    Else
            max = max
            maxticker = maxticker
    
        End If
         
    If Cells(j + 2, 11) < min Then
       
            min = Cells(j + 2, 11).Value
            minticker = Cells(j + 2, 9).Value
            
         Else
            min = min
            minticker = minticker
    
         End If
         
    If Cells(j + 2, 12) > maxvolume Then
       
            maxvolume = Cells(j + 2, 12).Value
            maxvolticker = Cells(j + 2, 9).Value
            
         Else
            maxvolume = maxvolume
            maxvolticker = maxvolticker
    
         End If
       
       
       
    Next j
    
    
    Cells(2, 16).Value = max
    Cells(2, 16).NumberFormat = ("0.00%")
    Cells(2, 15).Value = maxticker
    
    Cells(3, 16).Value = min
    Cells(3, 16).NumberFormat = ("0.00%")
    Cells(3, 15).Value = minticker
    
    Cells(4, 16).Value = maxvolume
    Cells(4, 15).Value = maxvolticker
    
    ws.Cells(1, 1) = "ticker" 'this sets cell A1 of each sheet to "ticker"
Next
starting_ws.Activate 'activate the worksheet that was originally active
End Sub
