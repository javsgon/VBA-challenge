Sub StockAnalysis()

    ' Instructions for the second part:
    
    ' Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

    ' Loop through all the worksheets
    For Each ws In Worksheets
    
        ' Define all variables to hold the ticker name, the open price, the close price, the Yearly Change, the Percent Change
        ' and the Total Stock Volume
        Dim ticker As String
        Dim o_price As Double
        Dim c_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_vol As Long
        Dim last_row As Long
        Dim i, j, k As Integer
   
        Dim summary_table As Long
    
        ' Check for Last Row Dynamically in the data
        lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
        ' Set initial summary table row
        summary_table = 2
    
        ' Set initial value of open price that is in cell C2
        o_price = ws.Range("C2").Value
    
        'Create the column headings needed for the summary tables according to the instructions
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        ' Loop through all Tickers
        For i = 2 To lastRow
    
        ' Check if we are still within the same Ticker name, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                ' Get ticker name
                ticker = ws.Cells(i, 1).Value
        
                ' Calculate the total stock volume
                totalvol = totalvol + ws.Cells(i, 7).Value
        
                ' Print the Ticker name in the Summary Table
                ws.Range("I" & summary_table).Value = ticker

                ' Print the Ticker Amount to the Summary Table
                ws.Range("L" & summary_table).Value = totalvol
            
                ' Get the closing price, the yearly change and the percent change
                c_price = ws.Cells(i, 6).Value
                yearly_change = (c_price - o_price)
                percent_change = yearly_change / o_price
            
                ' Print the Yearly Change and percent change in the Summary Table
                ws.Range("J" & summary_table).Value = yearly_change
            
                ws.Range("K" & summary_table).Value = percent_change
                ws.Range("K" & summary_table).NumberFormat = "0.00%"
            
                ' Add one to the summary table row
                summary_table = summary_table + 1
      
                ' Reset the Totals
                totalvol = 0
                o_price = ws.Cells(i + 1, 3)
            
                ' If the cell immediately following a row has the same Ticker name...
                Else

                ' Add to the Ticker Total
                totalvol = totalvol + ws.Cells(i, 7).Value
            
            End If

        Next i
    
        ' Color the values depending on the value: website reviewed for color guides: http://dmcritchie.mvps.org/excel/colors.htm
        ' Check for Last Row Dynamically in the summary table
        lastrow_summarytable = ws.Range("I" & Rows.Count).End(xlUp).Row
    
        For j = 2 To lastrow_summarytable
                ' If the cell value is >0 color it green
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                    ' Otherwise color it red
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
        Next j
    
        For k = 2 To lastrow_summarytable
                ' Find the maximum percent change https://officetuts.net/excel/vba/find-the-maximum-and-minimum-value-in-the-range-in-vba/
                If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summarytable)) Then
                    ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                    ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                    ws.Cells(2, 17).NumberFormat = "0.00%"

                '   Find the minimum percent change https://officetuts.net/excel/vba/find-the-maximum-and-minimum-value-in-the-range-in-vba/
                ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summarytable)) Then
                    ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                    ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                    ws.Cells(3, 17).NumberFormat = "0.00%"
            
                ' Find the maximum volume of trade
                ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summarytable)) Then
                    ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                    ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            
                End If
        
            Next k
        ' Autofit summary table and data https://excelchamps.com/vba/autofit/
        ws.Range("I:Q").EntireColumn.AutoFit
        
        Next ws

End Sub