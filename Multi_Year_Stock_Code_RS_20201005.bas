Attribute VB_Name = "Module1"
'Ryan Simpson
'20201005
'VBA Challenge - Stock Ticker

'Create subroutine to manipulate and present stock ticker data
Sub StockTicker()

'Loop through each worksheet
For Each ws In Worksheets

'Set Column Headers on each sheet for ticker specific summary columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Set Column Headers and Row Headers for "Greatest" Summary Table
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Create Variable to track Current ticker and next row ticker during loop
Dim Current_Ticker As String
Dim Next_Ticker As String

'Create Variable to track open and close prices while looping through each unique ticker symbol
Dim Close_Price As Double
Dim Open_Price As Double

'Create Variable to track total volumne for each unique ticker symbol, initialize at 0
Dim Total_Volume As Double
Total_Volume = 0

'Create variable for loop iteration
Dim i As LongLong

'Create variable to track current row location in summary columns for recording data on each ticker; initialize at 2, as column data will start on row 2
Dim TableRow As LongLong
TableRow = 2

'Create variable to find the last row on the worksheet
Dim LastRow As LongLong

'Define last row of worksheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Record open price for first ticker symbol of data set
    Open_Price = ws.Cells(2, 3).Value
    
    'Loop through data from rows 2 to the last row on the worksheet to post data for summary columns
    For i = 2 To LastRow
        
        'Define cells in loop for current row ticker symbol and next row ticker symbol values
        Current_Ticker = ws.Cells(i, 1).Value
        Next_Ticker = ws.Cells(i + 1, 1).Value
        
        'Create conditional for comparing if the current row and next row ticker symbol values match; If they do not match then:
        If Current_Ticker <> Next_Ticker Then
            
            'Post current ticker symbol to correct row in summary columns per iteration
            ws.Range("I" & TableRow).Value = Current_Ticker
            
            'Record closing price for current ticker symbol
            Close_Price = ws.Cells(i, 6).Value
            
            'Post Yearly Change for current ticker symbol using closing and opening prices to summary column
            ws.Range("J" & TableRow).Value = Close_Price - Open_Price
            
                'Create conditional for missing data from data set (null values); will prevent breaks in yearly percentage calculation
                'If the open price is 0, then:
                If Open_Price = 0 Then
                    'Do not calculate Percentage Change; post 0 to summary column
                    ws.Range("K" & TableRow).Value = 0
                
                'If open price is greater than 0 then"
                Else
                'Post "Percentage Change" to summary column (calculated using Yearly Change calculation and Open Price)
                ws.Range("K" & TableRow).Value = (Close_Price - Open_Price) / Open_Price
                'Format range to percentage(%) with 2 decimals
                ws.Range("K" & TableRow).NumberFormat = "0.00%"
                
                    'Create Conditional to format color of "Percentage Change" cells based on value
                    'If the percentage change is a positive value, then:
                    If ws.Range("K" & TableRow).Value > 0 Then
                        'Fill cell with green
                        ws.Range("K" & TableRow).Interior.ColorIndex = 4
                    'If percentage change is a negative value, then:
                    ElseIf ws.Range("K" & TableRow).Value < 0 Then
                        'Fill cell with red
                        ws.Range("K" & TableRow).Interior.ColorIndex = 3
                    'End Conditional for formatting
                    End If
                    
                'End Conditional for percentage change
                End If
                
            'Add current row's volume value to the Total Volume counter
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            'Post the "Total Stock Volume" for the current ticker to the summary columns
            ws.Range("L" & TableRow).Value = Total_Volume
            
            'Reset the value for Total Volumne to 0 for the next ticker symbol/loop
            Total_Volume = 0
                        
            'Record the opening price for the next ticker symbol to be used in next loop
            Open_Price = ws.Cells(i + 1, 3).Value
            
            'Add 1 to the table row to store next loop's data in the next row of the summary columns
            TableRow = TableRow + 1
            
        'If the current ticker and next ticker are the same, then:
        Else
            'Add the volume for that row to the Total Volume counter
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
        'End conditional for comparing current and next row ticker values
        End If
        
    'Start next iteration of loop to populate summary columns
    Next i
            
    'Create loop to review summary columns and retrieve data for "Greatest" summary table
    'Loop through start of data row 2 to last row of data
    For i = 2 To LastRow
        
        'Create conditional to record "Greatest % Increase"
        'Compare the value for column K "Percent Change" from each iteration to the value currently noted for "Greatest % Increase"
        'If column K Percent Change value is greater than the current "Greastest % Increase" value, then:
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            'Post the value from "Percent Change" column K for the current iteration to the summary table for "Greatest % Increase"
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            'Post the ticker from "Ticker" column for the current iteration to the summary table
            ws.Range("P2").Value = ws.Range("I" & i).Value
       'End Conditional for "Greatest % Increase"
        End If
        
         'Create conditional to record "Greatest % Decrease"
        'Compare the value for column K Percent Change from each iteration to the value currently noted for "Greatest % Decrease"
        'If column K "Percent Change" value is less than the current "Greastest % Decrease" value, then:
        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            'Post the value from "Percent Change" column K for the current iteration to the summary table for "Greatest % Decrease"
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            'Post the ticker from "Ticker" column for the current iteration to the summary table
            ws.Range("P3").Value = ws.Range("I" & i).Value
        'End Conditional for "Greatest % Decrease"
        End If
        
         'Create conditional to record "Greatest Total Volume"
        'Compare the value for column L "Total Stock Volume" from each iteration to the value currently noted for "Greatest Total Volume"
        'If column L "Total Stock Volume" value is greater than the current "Greastest Total Volume" value, then:
        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
             'Post the value from "Total Stock Volume" for the current iteration to the summary table for "Greatest Total Volume"
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            'Post the ticker from "Ticker" coulmn for the current iteration to the summary table
            ws.Range("P4").Value = ws.Range("I" & i).Value
        'End Conditional for "Greatest Total Volume"
        End If
        
    'Start next iteration of loop to populate "Greatest" summary table
    Next i
    
    'Format the values for "Greatest % Increase" and "Greatest % Decrease" to percentages(%) with 2 decimals on each worksheet
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'Format column widths to autofit to data on each worksheet
    ws.Columns("A:Q").AutoFit
    
'Start loop iteration for the next worksheet
Next ws

'End subroutine
End Sub



