Sub AnalyzeStocks()
' -------------------------------------------------------------------------
' VBA HOMEWORK - THE VBA OF WALL STREET
' GT DATA SCIENCE BOOTCAMP
' PROGRAMMED BY: ATUL NULKAR
' DATE: DECEMBER 2019

' This Module contains sub-routines to:
' 1) Analyze and display a yearly summary of real stock market data and
' 2) Conditionally format the data

' It performs the following functions:
' 1) Loops through all the stocks for a given year in each worksheet
' 2) Computes and displays the following summary information for each stock, within each worksheet:
'    * Ticker Symbol
'    * Yearly Price Change
'    * Yearly Percent Change
'    * Total Stock Volume
' -------------------------------------------------------------------------

    ' Define a current worksheet object
    ' We will use this to iterate through all worksheets in the active workbook
    Dim curr_worksheet As Worksheet

    ' Define a row variable to loop through all the stock rows on each worksheet
    Dim row As Long
    
    ' Define a lasRow variable to capture the last row number of the stock table in every worksheet
    Dim lastRow As Long
    
    ' Define row indices for ƒinding a stock's opening price and populating the summary information for stocks
    Dim openingPriceRowIndex, summaryRowIndex As Long

    ' Define the variables for the stock information that needs to be computed or displayed in the summary table
    Dim openingPrice, closingPrice, yearlyChange, percentChange As Variant
    Dim currentStockVolume, totalStockVolume As Variant

    ' Define current and next cells when checking and comparing the ticker symbols for each stock row
    Dim currentTicker, nextTicker As String

    ' Begin the loop to iterate through each worksheet in the active workbook
    For Each curr_worksheet In Worksheets
    
        ' The With statement here allows us to perform a series of statements on the curr_worksheet object..
        ' Without requalifying the name of the object
        With curr_worksheet
        
            ' Determine the last row of the stocks table in the current worksheet
            lastRow = .Cells(Rows.Count, 1).End(xlUp).row

            ' Set the initial row index to 2 since the opening price of the first stock is present in column C, row 2
            ' The row index will be changed in the for loop below and is used to grab the opening price of the stock when computing the yearly price change
            openingPriceRowIndex = 2

            ' Initialize the Total Stock Volume to 0
            totalStockVolume = 0

            ' Populate the summary table headers in columns I, J, K, L on each worksheet
            ' Set the initial row index for the summary table to 2 since we are displaying stock information from rows 2 and below
            .Range("I1") = "Ticker"
            .Range("J1") = "Yearly Change"
            .Range("K1") = "Percent Change"
            .Range("L1") = "Total Stock Volume"
            summaryRowIndex = 2

            ' Iterate through all of the stock rows in the current worksheet
            For row = 2 To lastRow

                ' Set current and next ticker cells to their respective variables
                currentTicker = .Cells(row, 1).Value
                nextTicker = .Cells(row + 1, 1).Value

                ' Set the volume for the current stock
                currentStockVolume = .Cells(row, 7).Value
                
                ' Add up the Total Volume for each stock (this maintains the cumulative total for each stock/ticker)
                totalStockVolume = totalStockVolume + currentStockVolume
                
                ' Check if the current stock/ticker is the same as the next one
                If currentTicker <> nextTicker Then
        
                    ' Grab the opening and closing prices of the stock
                    openingPrice = .Cells(openingPriceRowIndex, 3).Value
                    closingPrice = .Cells(row, 6).Value
                    
                    ' Compute the yearly change in price
                    yearlyChange = closingPrice - openingPrice
                    
                    ' Compute the percent change in price only if the opening price and yearly change is not 0
                    ' This covers for any divide by zero errors in the percent calculation below
                    If (openingPrice <> 0) And (yearlyChange <> 0) Then
                        
                        percentChange = (closingPrice - openingPrice) / openingPrice
                    
                    Else
                        
                        percentChange = 0
                        
                    End If

                    ' Once the next ticker symbol is found, display and format the stock information in the summary rows
                    .Range("I" & summaryRowIndex) = currentTicker
                    
                    .Range("J" & summaryRowIndex) = yearlyChange
                    .Range("J" & summaryRowIndex).NumberFormat = "0.00"
                    
                    .Range("K" & summaryRowIndex) = percentChange
                    .Range("K" & summaryRowIndex).NumberFormat = "0.00%"
                    
                    .Range("L" & summaryRowIndex) = totalStockVolume
                    .Range("L" & summaryRowIndex).NumberFormat = "#,##0"

                    ' Reset Total Stock Volume to 0 so it can be computed again for the next stock
                    totalStockVolume = 0

                    ' Increment the summaryRowIndex by 1 so information for the next stock is displayed in the next summary row
                    summaryRowIndex = summaryRowIndex + 1

                    ' Update the openinPriceRowIndex so the opening price can be grabbed for the next stock
                    openingPriceRowIndex = row + 1

                End If

            Next row
        
            MsgBox ("Completed processing worksheet: " + .Name)
            
            ' Call ApplyConditionalFormatting
        End With

    Next curr_worksheet

End Sub

Sub ApplyConditionalFormatting()

' This Sub applies conditional formatting to the Yearly Change column in all worksheets

    Columns("J:J").Select

End Sub