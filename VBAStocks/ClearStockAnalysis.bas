Attribute VB_Name = "ClearStockAnalysis"
Sub ClearStockAnalysisInWorksheets()

' This sub-routine clears the stock analysis summary data on all the worksheets

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I:Q").Clear
    
Next ws

End Sub
