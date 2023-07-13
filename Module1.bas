
Sub multiple_year_stock()

'Declare worksheets
Dim ws As Worksheet

'Loop through all worksheets
For Each ws In Worksheets

'Creat colume headings
ws.Range("I1,P1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"


ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"

Dim Lastrow As Long
'Define last row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'The ticker symbol
    '........................................................
    Dim Ticker As String
    Dim Summary_table_row As Integer
    Summary_table_row = 2
    
    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
    '.........................................................
    Dim yearlychange As Double
    Dim openyear As Double
    Dim closeyear As Double
    'Define the opening price at the beginning of a year of each worksheet, before looping
    openyear = ws.Range("C2").Value

    'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    '...........................................................
    Dim percentchange As Double
    percentchange = 0

    'The total stock volumn is the sum of daily volumn
    Dim totalvolumn As Double
    totalvolumn = 0

    'BONUS:Find tickers that have greatest increase, greatest decrease,and greatest total volumn
    '.................................................................
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolumn As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volumn_ticker As String
    greatestincrease = 0
    greatestdecrease = 0
    greatestvolumn = 0
    greatest_increase_ticker = " "
    greatest_decrease_ticker = " "
    greatest_volumn_ticker = " "

        'Create Loop
        '..............................................
        
        'While in each row, loop through the ticket symbol, yearly change, percent change and total volumn
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Set the ticket symbol
            Ticker = ws.Cells(i, 1).Value

            'Print the ticket symbol in the summary table
            ws.Range("I" & Summary_table_row).Value = Ticker
            
            'Set the closing price of the ticker when there are no match between 2 tickers
            closeyear = ws.Cells(i, 6).Value
    
            'Set the yearly change equals the calculation and print it in the summary table
            yearlychange = closeyear - openyear
            ws.Range("J" & Summary_table_row).Value = yearlychange
    
            'Set the percent change equals the calculation and print it in the summary table
            percentchange = yearlychange / openyear
            ws.Range("K" & Summary_table_row).Value = percentchange
    
            
        'Looping through percent change to find greatest increase and greatest decrease
            'greatest year increase
            If percentchange > greatestincrease Then
            greatestincrease = percentchange
            greatest_increase_ticker = ws.Cells(i, 1).Value
            End If
            'greatest year decrease
            If percentchange < greatestdecrease Then
            greatestdecrease = percentchange
            greatest_decrease_ticker = ws.Cells(i, 1).Value
            End If
    
            'Looping through total stock volumn to find greatest total
            'greatest total volumn
            If totalvolumn > greatestvolumn Then
            greatestvolumn = totalvolumn
            greatest_volumn_ticker = ws.Cells(i, 1).Value
            End If
    
        

    'Apply conditional formatting for positive and negative values
    '............................................................
    'Formatting for Yearly Change
    Dim rg As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    
        'specify range to apply
        Set rg = ws.Range("J2", ws.Range("J2").End(xlDown))
        
        'clear existing conditional formatting
        rg.FormatConditions.Delete
        
        'apply conditional formatting
        Set condition1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")

        Set condition2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")
        
        'define format of each condition
        With condition1
        .Interior.Color = vbGreen
        End With
        
        With condition2
        .Interior.Color = vbRed
        End With
        
    'Formatting for Percent Change
    Dim rg1 As Range
    Dim condition3 As FormatCondition
    Dim condition4 As FormatCondition
    
        'specify range to apply
        Set rg1 = ws.Range("K2", ws.Range("K2").End(xlDown))
        
        'clear existing conditional formatting
        rg1.FormatConditions.Delete
        
        'apply conditional formatting:
        'when it's possitive
        Set condition3 = rg1.FormatConditions.Add(xlCellValue, xlGreater, "=0%")
        'when it's negative
        Set condition4 = rg1.FormatConditions.Add(xlCellValue, xlLess, "=0%")
        
        'set conditional format of each condition
        With condition3
        .Interior.Color = vbGreen
        End With
        
        With condition4
        .Interior.Color = vbRed
        End With

    

    
    
            'reset the openning value of each new ticker
            openyear = ws.Cells(i + 1, 3).Value
    
            'volumn
            totalvolumn = totalvolumn + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_table_row).Value = totalvolumn
    
            'Add next row to the table
            Summary_table_row = Summary_table_row + 1
            
            'reset total volumn
            totalvolumn = 0
    
    
            'When
            Else

            totalvolumn = totalvolumn + ws.Cells(i, 7).Value
            
    
            
   
    'End if loop
        End If

    'End For loop
    Next i
    
        
'Print the values and tickers on the Greatest table
ws.Range("P2").Value = greatest_increase_ticker
ws.Range("Q2").Value = greatestincrease
ws.Range("P3").Value = greatest_decrease_ticker
ws.Range("Q3").Value = greatestdecrease
ws.Range("P4").Value = greatest_volumn_ticker
ws.Range("Q4").Value = greatestvolumn
    
    
'apply percentage format
ws.Range("K2:K" & Lastrow).NumberFormat = "0.00%"
ws.Range("Q2", "Q3").NumberFormat = "0.00%"
        
        
    
Next ws

End Sub



          
