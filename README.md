# Stock Analysis with VBA

The objective of this project is to create a script that loops through all the stocks for one year and outputs the following information:

* The ticker symbol
* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
* The percentage change from the opening price at the beginning of any given year to the closing price at the end of that year
* The total stock volume of the stock.

# The Yearly Stock Performance Results for 2018
![2018_stock_data](https://github.com/faceadversity/VBA-challenge/assets/137361966/adade070-2749-4c90-a2a6-601ef6fd80c6)

# The Yearly Stock Performance Results for 2019
(already attached in the repo)

# The Yearly Stock Performance Results for 2020
(already attached in the repo)

Functionality was added to the VBA code that returned stock values with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". Each VBA script will have to be run individually to output the needed results in their corresponding yearly stock returns spreadsheet. The VBA files needed to perform this function is attached to the repository for reference purposes. A sample VBA code for the 2018 Stock Performance is included below:

Sub StockData()

    ' Define the variables
    Dim wrksht As Worksheet
    Dim last_row As Long, i As Long, start_row As Long
    Dim ticker_sym As String, open_price As Double, close_price As Double
    Dim yearly_change As Double, percent_change As Double, total_volume As Double
    Dim output_row As Long
    ' Variables for tracking greatest increase, decrease and total volume
    Dim greatest_inc As Double, greatest_dec As Double, greatest_Vol As Double
    Dim inc_ticker As String, dec_ticker As String, vol_ticker As String
    
    ' Set worksheet
    Set wrksht = Workbooks("Multiple_year_stock_data.xlsm").Sheets("2018")
    
    ' Set initial output row
    output_row = 2 ' Change to the row where you want the output to start
    
    ' Initialize tracking variables
    greatest_inc = 0
    greatest_dec = 0
    greatest_Vol = 0
    
    ' Find the last row of data
    last_row = wrksht.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Initialize start row
    start_row = 2
    
    ' Initialize ticker symbol
    ticker_sym = wrksht.Cells(start_row, 1).Value
    
    ' Initialize opening price
    open_price = wrksht.Cells(start_row, 3).Value
    
    ' Set column titles
    wrksht.Cells(1, 9).Value = "Ticker"
    wrksht.Cells(1, 10).Value = "Yearly Change"
    wrksht.Cells(1, 11).Value = "Percent Change"
    wrksht.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Looping through each row of data
    For i = start_row To last_row
        ' Check if we are sAll within the same Acker symbol
        If wrksht.Cells(i + 1, 1).Value <> ticker_sym Then
            ' Set the closing price
            close_price = wrksht.Cells(i, 6).Value
            
            ' Calculate the yearly change
            yearly_change = close_price - open_price
 
            ' Calculate the percentage change
            If open_price <> 0 Then
                percent_change = yearly_change / open_price
            Else
                percent_change = 0
            End If
            
            ' Calculate the total volume
            total_volume = Application.WorksheetFunction.Sum(wrksht.Range(wrksht.Cells(start_row, 7), wrksht.Cells(i, 7)))
            
            ' Output the results to the worksheet
            With wrksht.Cells(output_row, 10)
                .Value = yearly_change
                ' Clear any existing format conditions
                .FormatConditions.Delete
                ' Add conditional formatting
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address & "<0"
                .FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red is negative
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address & ">0"
                .FormatConditions(2).Interior.Color = RGB(0, 255, 0) ' Green is positive
            End With
 
            With wrksht.Cells(output_row, 11)
                .Value = percent_change * 100
                ' Clear any existing format conditions
                .FormatConditions.Delete
                ' Add conditional formatting
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address & "<0"
                .FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red is negative
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address & ">0"
                .FormatConditions(2).Interior.Color = RGB(0, 255, 0) ' Green is positive
            End With
 
            wrksht.Cells(output_row, 9).Value = ticker_sym
            wrksht.Cells(output_row, 12).Value = total_volume
            
            ' Check current ticker has greatest increase, decrease or total volume
            If percent_change > greatest_inc Then
                greatest_inc = percent_change
                inc_ticker = ticker_sym
            End If
            
            If percent_change < greatest_dec Then
                greatest_dec = percent_change
                dec_ticker = ticker_sym
            End If
 
            If total_volume > greatest_Vol Then
                greatest_Vol = total_volume
                vol_ticker = ticker_sym
            End If
 
            ' Move to the next output row
            output_row = output_row + 1
 
            ' Move to the next ticker
            start_row = i + 1
            ticker_sym = wrksht.Cells(start_row, 1).Value
            open_price = wrksht.Cells(start_row, 3).Value
        End If
    Next i
 
    ' Output the greatest increase, decrease, and total volume
    wrksht.Cells(1, 15).Value = ""
    wrksht.Cells(1, 16).Value = "Ticker"
    wrksht.Cells(1, 17).Value = "Value"
 
    wrksht.Cells(2, 15).Value = "Greatest % Increase"
    wrksht.Cells(2, 16).Value = inc_ticker
    wrksht.Cells(2, 17).Value = greatest_inc * 100
    
    wrksht.Cells(3, 15).Value = "Greatest % Decrease"
    wrksht.Cells(3, 16).Value = dec_ticker
    wrksht.Cells(3, 17).Value = greatest_dec * 100
    
    wrksht.Cells(4, 15).Value = "Greatest Total Volume"
    wrksht.Cells(4, 16).Value = vol_ticker
    wrksht.Cells(4, 17).Value = greatest_Vol
End Sub

code srcs: Close peer (Chidimma Oruche) with extensive Data Analytics experience. We both corrobarated ideas on this project and I was able to come up with the final outlines presented here. I also derived insights from Reddit and Stack Overflow but didn't use any specific coding formats pertaining to those sources. They were a bit too complicated especially when trying to merge all worksheets together to display needed results by running the macro one time. Doing this on an individual basis was within my abilities to comprehend and run successfully for all yearly stock outputs.
