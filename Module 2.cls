VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim row As Long
    Dim summaryRow As Long
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Initialize variables for each sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        summaryRow = 2 ' Start summary output from row 2
        
        ' Set headers for summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Loop through each row in the sheet
        For row = 2 To lastRow
            ' Check if the ticker symbol has changed
            If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
                ' If not the first row, calculate and output summary for the previous ticker
                If row > 2 Then
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 11).Value = percentChange
                    ws.Cells(summaryRow, 12).Value = totalVolume
                    
                    ' Apply conditional formatting
                    If yearlyChange >= 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                    Else
                        ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                    End If
                    
                    ' Move to the next row in the summary
                    summaryRow = summaryRow + 1
                End If
                
                ' Reset variables for the new ticker
                ticker = ws.Cells(row, 1).Value
                openingPrice = ws.Cells(row, 3).Value
                totalVolume = 0
            End If
            
            ' Accumulate total volume
            totalVolume = totalVolume + ws.Cells(row, 7).Value
            
            ' If it's the last row for the ticker, calculate and output the summary
            'If row = lastRow Then
            If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
                closingPrice = ws.Cells(row, 6).Value
                yearlyChange = closingPrice - openingPrice
                
                ' Avoid division by zero error
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice) * 100
                Else
                    percentChange = yearlyChange * 100
                End If
                
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Apply conditional formatting
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If
            End If
        Next row
        
        ' Find and output the greatest percent increase, greatest percent decrease, and greatest total volume
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim maxIncreaseTicker As String
        Dim maxDecreaseTicker As String
        Dim maxVolumeTicker As String
        
        maxIncrease = ws.Cells(2, 11).Value
        maxDecrease = ws.Cells(2, 11).Value
        maxVolume = ws.Cells(2, 12).Value
        maxIncreaseTicker = ws.Cells(2, 9).Value
        maxDecreaseTicker = ws.Cells(2, 9).Value
        maxVolumeTicker = ws.Cells(2, 9).Value
        
        For row = 3 To summaryRow
            ' Find the greatest percent increase
            If ws.Cells(row, 11).Value > maxIncrease Then
                maxIncrease = ws.Cells(row, 11).Value
                maxIncreaseTicker = ws.Cells(row, 9).Value
            End If
            
            ' Find the greatest percent decrease
            If ws.Cells(row, 11).Value < maxDecrease Then
                maxDecrease = ws.Cells(row, 11).Value
                maxDecreaseTicker = ws.Cells(row, 9).Value
            End If
            
            ' Find the greatest total volume
            If ws.Cells(row, 12).Value > maxVolume Then
                maxVolume = ws.Cells(row, 12).Value
                maxVolumeTicker = ws.Cells(row, 9).Value
            End If
        Next row
        
        ' Output the results for each sheet
        ws.Cells(1, 15).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = maxIncrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = maxDecreaseTicker
        ws.Cells(4, 16).Value = maxDecrease & "%"
        
        ws.Cells(5, 15).Value = "Greatest Total Volume"
        ws.Cells(6, 15).Value = maxVolumeTicker
        ws.Cells(6, 16).Value = maxVolume
    Next ws
End Sub

