Attribute VB_Name = "Module1"
Sub StockMarkert()

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Homework:

' 1-Create a script that will loop through all the stocks for one
'    year and output the following information.
    
' 2-The ticker symbol.
    
' 3-Yearly change from opening price at the beginning of a given
' year to the closing price at the end of that year.
    
'4-The percent change from opening price at the beginning of a given year
' to the closing price at the end of that year.
    
'5 - The total stock volume of the stock.
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------



' Define the variables
Dim WS_Count As Integer
Dim ICol As Long
Dim IRow As Long
Dim volume As Double
Dim opening_price As Long
Dim close_price As Long
Dim variation As Long
Dim num_tickers As Integer
Dim great_increase As Double
Dim great_decrease As Double
Dim great_tot_volume As Double



' Calculate the total worksheets
WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop throughout all the worksheets
For JJ = 1 To WS_Count
    
    'Activate the sheet
    Sheets(JJ).Activate
    
    ' Find the last column and last row number
    ICol = Cells(1, Columns.Count).End(xlToLeft).Column
    IRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Insert names on the columns
    Cells(1, ICol + 2).Value = "Ticker"
    Cells(1, ICol + 3).Value = "Yearly Change"
    Cells(1, ICol + 4).Value = "Percentage Change"
    Cells(1, ICol + 5).Value = "Total Stock Volume"
                
    ' Create some auxiliary variables
    volume = 0
    opening_price = Cells(2, 3).Value
    j = 2
    
    ' Loop through all rows to retrieve ticker label
    For I = 2 To IRow
    
        'Check if the ticker label changed and save the old value
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            Cells(j, ICol + 2).Value = Cells(I, 1).Value
                            
            'Calculate the values
            close_price = Cells(I, 6).Value
            variation = close_price - opening_price
            Cells(j, ICol + 3).Value = variation
                
            ' Add conditional format
            If variation < 0 Then
                Cells(j, ICol + 3).Interior.Color = vbRed
            Else
                Cells(j, ICol + 3).Interior.Color = vbGreen
            End If
                     
            'Calculate variation
            If opening_price = 0 Then
                'Avoid error in case the ticker has 0 opening price
                Cells(j, ICol + 4).Value = 0
            Else
                Cells(j, ICol + 4).Value = (variation / opening_price)
            End If
            
            ' Assign value to column
            Cells(j, ICol + 5).Value = volume + Cells(I, 7).Value
            
            'Reset the counters
            opening_price = Cells(I + 1, 3).Value
            volume = 0
            j = j + 1
                        
        Else
            ' Retrieve the volume for each row
            volume = volume + Cells(I, 7).Value
                        
        End If
    
    Next I
            
                
    'Autofit and cell formatting
    Cells(1, ICol + 2).Columns.AutoFit
    Cells(1, ICol + 3).Columns.AutoFit
    Cells(1, ICol + 4).Columns.AutoFit
    Cells(1, ICol + 5).Columns.AutoFit
    Range("K:K").NumberFormat = "0.00%"
                
                
    '----------------------------------------------------------------------
    ' Challenge
    '----------------------------------------------------------------------
    Cells(2, ICol + 8).Value = "Greatest % increase"
    Cells(3, ICol + 8).Value = "Greatest % decrease"
    Cells(4, ICol + 8).Value = "Greatest total volume"
    Cells(1, ICol + 9).Value = "Ticker"
    Cells(1, ICol + 10).Value = "Value"
                
               
    ' Calculate the number of tickers
    num_tickers = Cells(Rows.Count, ICol + 2).End(xlUp).Row
    great_increase = 0
    great_decrease = 0
    great_tot_volume = 0

    
    ' Loop through the different tickers
    For kk = 2 To num_tickers
        ' Calculate the great increase
        If Cells(kk, ICol + 4).Value > great_increase Then
            great_increase = Cells(kk, ICol + 4).Value
            ticker_great_increase = Cells(kk, ICol + 2).Value
        End If
        
        ' Calculate the great decrease
        If Cells(kk, ICol + 4).Value < great_decrease Then
            great_decrease = Cells(kk, ICol + 4).Value
            ticker_great_decrease = Cells(kk, ICol + 2).Value
        End If
                    
        ' Calculate the great volume
        If Cells(kk, ICol + 5).Value > great_tot_volume Then
            great_tot_volume = Cells(kk, ICol + 5).Value
            ticker_great_volume = Cells(kk, ICol + 2).Value
        End If
                    
    Next kk

    ' Assign the calculated for the challenge
    Cells(2, ICol + 9).Value = ticker_great_increase
    Cells(2, ICol + 10).Value = great_increase
    Cells(2, ICol + 10).NumberFormat = "0.00%"
    Cells(3, ICol + 9).Value = ticker_great_decrease
    Cells(3, ICol + 10).Value = great_decrease
    Cells(3, ICol + 10).NumberFormat = "0.00%"
    Cells(4, ICol + 9).Value = ticker_great_volume
    Cells(4, ICol + 10).Value = great_tot_volume

    ' Autofit the columns
    Cells(4, ICol + 8).Columns.AutoFit
    Cells(1, ICol + 9).Columns.AutoFit
    Cells(4, ICol + 10).Columns.AutoFit
    Range("A1").Select
    
    Next JJ
    
End Sub

