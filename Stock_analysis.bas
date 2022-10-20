Attribute VB_Name = "Module1"
Sub stock_Analysis()

    Dim ws As Worksheet

    'With for each loop code will be run for each and every worksheet in this workbook
    For Each ws In Worksheets

    ws.Activate

        'Determine variables for first summary table
        Dim Ticker As String
        Ticker = ""

        Dim initial_price As Double

        Dim last_price As Double

        Dim Yearly_change As Double

        Dim year_percentage_change As Double

        Dim total_volume As Double
        total_volume = 0

        Dim summary_table_row As Integer
        summary_table_row = 2

        'Determine variables for the second summary table
        Dim max_value As Double
        Dim min_value As Double
        Dim Total_value As Double


        'Print the headers of the tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "initial price"
        ws.Range("K1").Value = "last_price"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percent Change"
        ws.Range("N1").Value = "Total Stock Volume"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"

        initial_price = Cells(2, 3).Value
        ws.Range("j" & summary_table_row).Value = initial_price
            
        Endrow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

            For i = 2 To Endrow

                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                Ticker = Cells(i, 1).Value
                ws.Range("I" & summary_table_row).Value = Ticker
           
                total_volume = total_volume + ws.Cells(i, 7).Value
                ws.Range("N" & summary_table_row).Value = total_volume
            
                last_price = ws.Cells(i, 6)
                ws.Range("K" & summary_table_row).Value = last_price
            
                Yearly_change = last_price - initial_price
                ws.Range("L" & summary_table_row).Value = Yearly_change
                    
                    If ws.Range("L" & summary_table_row).Value >= 0 Then
                        ws.Range("L" & summary_table_row).Interior.ColorIndex = 4
                    ElseIf ws.Range("L" & summary_table_row).Value < 0 Then
                        ws.Range("L" & summary_table_row).Interior.ColorIndex = 3
                    End If
                        
                ws.Range("M" & summary_table_row).NumberFormat = "0.00%"
                Percent_change = (Yearly_change / initial_price)
                ws.Range("M" & summary_table_row).Value = Percent_change
          
                summary_table_row = summary_table_row + 1
            
                initial_price = ws.Cells(i + 1, 3)
                ws.Range("j" & summary_table_row).Value = initial_price
               
                total_volume = 0
        
                Else
             
                total_volume = total_volume + ws.Cells(i, 7).Value
                 
                End If
            
            Next i
            
    
        'Determine and print to the second summary table the Greatest Percentage increase
        Endroww = ws.Range("M" & Rows.Count).End(xlUp).Row
        
        max_value = 0

        For x = 2 To Endroww
        
            If Range("M" & x).Value > max_value Then
        
                max_value = Range("M" & x).Value
                ws.Range("R2").NumberFormat = "0.00%"
                Range("R2").Value = max_value
                Range("Q2").Value = Range("I" & x).Value
            End If
        
        Next x


        'Determine and print to the second summary table the Greatest Percentage decrease
        Endroww = ws.Range("M" & Rows.Count).End(xlUp).Row
        
        min_value = 0
        
        For y = 2 To Endroww
        
            If Range("M" & y).Value < min_value Then
        
                min_value = Range("M" & y).Value
                ws.Range("R3").NumberFormat = "0.00%"
                Range("R3").Value = min_value
                Range("Q3").Value = Range("I" & y).Value
            End If
        
        Next y
        
        
        'Determine and print to the second summary table the Greatest Total Volume
        Endrowww = ws.Range("N" & Rows.Count).End(xlUp).Row
        
        Total_value = 0
        
        For Z = 2 To Endrowww
        
            If Range("N" & Z).Value > Total_value Then
        
                Total_value = Range("N" & Z).Value
                Range("R4").Value = Total_value
                Range("Q4").Value = Range("I" & Z).Value
            End If
        
        Next Z

    Next ws

End Sub

