' ==============================================================================
' VBA HOMEWORK - THE VBA OF WALL STREET
' GT DATA SCIENCE BOOTCAMP
' PROGRAMMED BY: ATUL NULKAR
' Date: DECEMBER 2019
' ==============================================================================
' This Module contains sub-routines to:
' 1) Analyze and display a yearly summary of real stock market data
' 2) Conditionally format the data
' 3) Search and display stocks with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume" for a given year
' ==============================================================================

Sub AnalyzeStocks()

' -----------------------------------------------------------------------------------------------------------
' This sub-routine performs the following functions:
' 1) Loops through all the stocks for a given year in each worksheet
' 2) Computes and displays the following summary information for each stock, within each worksheet:
'    * Ticker Symbol
'    * Yearly Price Change
'    * Yearly Percent Change
'    * Total Stock Volume
' -----------------------------------------------------------------------------------------------------------
    
    'PURPOSE: Determine how many minutes it takes for the stock analysis code to finish running
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    
    Dim StartTime As Double
    Dim MinutesElapsed As String
    
    'Remember time when macro starts
    StartTime = Timer

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
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
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
                    ' This validation covers for any divide by zero errors in the percent change calculation
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
         
            ' Calls the sub-routine that applies conditional formatting on the Yearly Change column
            Call ApplyConditionalFormatting(curr_worksheet)
        
            ' Call the sub-routine to displays stocks with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume" for a given year
            Call DisplayMaxMinStockData(curr_worksheet)
           
            ' Message box that prints number of stocks processed on each worksheet for debugging purposes
            ' MsgBox ("Completed processing" + Str(lastRow - 1) + " stocks on worksheet: " + .Name)

        End With

    Next curr_worksheet
    
    'Determine how many seconds code took to run
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

    'Notify user in seconds
    MsgBox "The stock analysis was completed successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub

Sub ApplyConditionalFormatting(ws As Worksheet)

' -----------------------------------------------------------------------------------------------------------
' This sub-routine applies conditional formatting to the Yearly Change column in all worksheets
' -----------------------------------------------------------------------------------------------------------
    
    ' Define the Range object for the yearl change column
    Dim yearlyChangeRange As Range
    
    ' Define the conditional format objects for positive and negative yearly changes in price
    Dim positiveChangeFormat, negativeChangeFormat As FormatCondition
    
    ' Set the range from cell J2 and down the column
    Set yearlyChangeRange = ws.Range("J2", ws.Range("J2").End(xlDown))
    
    With yearlyChangeRange
        
        ' Clear any existing conditional formatting
        .FormatConditions.Delete

        ' Define the rule for each conditional format
        Set positiveChangeFormat = .FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set negativeChangeFormat = .FormatConditions.Add(xlCellValue, xlLess, "=0")
        
        ' Define the format applied for each conditional format
        ' Highlight positive change in green and negative change in red
        positiveChangeFormat.Interior.Color = vbGreen
        negativeChangeFormat.Interior.Color = vbRed
       
    End With

End Sub

Sub DisplayMaxMinStockData(ws As Worksheet)

' --------------------------------------------------------------------------------------------------------------
' This sub-routine displays stocks with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume" for a given year
' --------------------------------------------------------------------------------------------------------------

    With ws
         
         ' Populate summary table headers and row labels on each worksheet
        .Range("P1") = "Ticker"
        .Range("Q1") = "Value"
        .Range("O2") = "Greatest % Increase"
        .Range("O3") = "Greatest % Decrease"
        .Range("O4") = "Greatest Total Volume"
        
        ' Define variables to capture the Greatest % increase, Greatest % Decrease and Greatest total stock volume
        Dim maxPercentIncrease, maxPercentDecrease, maxTotalVolume
        
        ' Define variables to find the ticker symbol associated with the Greatest % increase, Greatest % Decrease and Greatest total stock volume
        Dim whatToFind, tickerValue
        
        ' Define a Range object that captures the value of the cell (ticker symbol) returned by the Range.Find function with searching for the ticker
        Dim findCell As Range
    
        ' Calculate and store the Greatest % increase
        maxPercentIncrease = Application.WorksheetFunction.Max(.Range("K2:K" & Rows.Count))
        .Range("Q2").NumberFormat = "0.00%"
        .Range("Q2") = maxPercentIncrease
        
        ' Calculate and store the Greatest % decrease
        maxPercentDecrease = Application.WorksheetFunction.Min(.Range("K2:K" & Rows.Count))
        .Range("Q3").NumberFormat = "0.00%"
        .Range("Q3") = maxPercentDecrease
        
        ' Calculate and store the Greatest total stock volume
        maxTotalVolume = Application.WorksheetFunction.Max(.Range("L2:L" & Rows.Count))
        .Range("Q4").NumberFormat = "#,##0"
        .Range("Q4") = maxTotalVolume
        
        ' Format the Greatest % increase value and store it in a variable so we can find the exact match
        whatToFind = FormatPercent(maxPercentIncrease, 2, vbUseDefault, vbUseDefault, vbFalse)
        
        ' Use the Range.Find function to find the Greatest % increase and pull the corresponding ticker symbol value from the returned Range/Cell
        Set findCell = .Range("K2:K" & Rows.Count).Find(What:=whatToFind, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not findCell Is Nothing Then
            
            ' Display the corresponding ticker symbol if a matching cell value is found
            tickerValue = .Cells(findCell.row, 9)
            .Range("P2").Value = tickerValue
        
        Else
        
            ' Display an error in case no ticker symbol is found
            .Range("P2").Value = "Not Found"
        
        End If
        
        ' Format the Greatest % decrease value and store it in a variable so we can find the exact match
        whatToFind = FormatPercent(maxPercentDecrease, 2)
        
        ' Use the Range.Find function to find the Greatest % decrease and pull the corresponding ticker symbol value from the returned Range/Cell
        Set findCell = .Range("K2:K" & Rows.Count).Find(What:=whatToFind, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not findCell Is Nothing Then
            
            ' Display the corresponding ticker symbol if a matching cell value is found
            tickerValue = .Cells(findCell.row, 9)
            .Range("P3").Value = tickerValue
        
        Else
        
            ' Display an error in case no ticker symbol is found
            .Range("P3").Value = "Not Found"
        
        End If
        
        ' Format the Greatest total stock volume value and store it in a variable so we can find the exact match
        whatToFind = Format(maxTotalVolume, "#,##0")
         
        ' Use the Range.Find function to find the Greatest total stock volume and pull the corresponding ticker symbol value from the returned Range/Cell
        Set findCell = .Range("L2:L" & Rows.Count).Find(What:=whatToFind, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not findCell Is Nothing Then
        
            ' Display the corresponding ticker symbol if a matching cell value is found
            tickerValue = .Cells(findCell.row, 9)
            .Range("P4").Value = tickerValue
            
        Else
        
            ' Display an error in case no ticker symbol is found
            .Range("P4").Value = "Not Found"
        
        End If

    End With

End Sub