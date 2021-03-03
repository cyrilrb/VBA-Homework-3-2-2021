Sub Stocks()

'Part 1 (P1): Ticker and Total Stock Volume
'Part 2 (P2):Yearly change from opening stock price to closing stock price at end of year
'Part 3 (P3):Conditional fomatting hilighting the positive change in green and negitive change in red
        'Followed instructions from: https://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/
        'Note: This coding must be applied before the cells are populated or the range defined in the coding will stop at the first cell that has a value.
'Part 4 (P4): Designate and search for greatest % increase, decrease and greatest total volume Stock. Identify the Ticker designation and value for each.
       
        
'P1 To gather and summarize Ticker Information
  ' Define the variables
  Dim Ticker As String
  Dim TotalStockVol As Double
    TotalStockVol = 0
  Dim StockSummary As Integer
    StockSummary = 2 ' This designates the 2nd row as the place where entries start to get filled in as the loop moves down
  Dim ws As Worksheet
  'Dim Starting_ws As Worksheet
    'Set Starting_ws = ActiveSheet
  Dim LastRow As Double
  Dim CPrice As Double
  Dim Oprice1 As Double
    Oprice1 = 0
  Dim Oprice2 As Double
    Oprice2 = 0
  Dim Oprice3 As Double
  Dim Oprice4 As Double
  Dim Oprice5 As Double
    'P3 Define the variables
  Dim change As Range
  'set the condtion variables
  Dim cond1 As FormatCondition, cond2 As FormatCondition
  
  
'Loop to apply all actions to each worksheet in workbook

 
 'For Each ws In Worksheets
    

    'P1 Determine the number of rows on each sheet for i value boundaries
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'P1 Label the StockSummary Table headers
        Range("J1") = "Ticker"
        Range("K1") = "Yearly Change (Price)"
        Range("L1") = "% Change"
        Range("M1") = "Total Stock Volume"
        'Range("R1") = "Ticker"
        'Range("S1") = "Open Price Difference"
        'Range("T1") = "Last Open Price"
        'Range("U1") = "Open Price"
        'Range("V1") = "Close Price"
      
    'P3 Set values for variables
        Set change = Range("L2", Range("L2").End(xlDown))
        Set change = Range("K2", Range("K2").End(xlDown))
    
    'P3 Set the rules for the conditions
        Set cond1 = change.FormatConditions.Add(xlCellValue, xlGreater, "0")
        Set cond2 = change.FormatConditions.Add(xlCellValue, xlLess, "0")
        
    'P3 Set the colors for each condition if they are true
        With cond1
        .Interior.Color = vbGreen
        .Font.Color = vbBlack
        End With
                        
        With cond2
        .Interior.Color = vbRed
        .Font.Color = vbBlack
        End With
       
    'P1 Define the boundaries of the loop
        For i = 2 To LastRow
        
            'P1 Raster throught the sheet with if/then to capture all ticket types
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'P1 Designate where the ticker data is
                Ticker = Cells(i, 1).Value
                'Designate where the close price data is
                CPrice = Cells(i, 6).Value
                'Designate where the last open price is
                Oprice4 = Cells(i, 3).Value
                
                'P1 Sum up the Stock Volumes per ticker on your summary table
                TotalStockVol = TotalStockVol + Cells(i, 7).Value
                
                'P1 Output each ticker type from the table
                Cells(StockSummary, 10).Value = Ticker

                'P1 Output the Total Stock volume to column M
                Cells(StockSummary, 13).Value = TotalStockVol
            
                'Output each ticker type from the table
                'Cells(StockSummary, 18).Value = Ticker
                'Output Last Open Price
                'Cells(StockSummary, 20).Value = Oprice4
                'Output Close Price
                'Cells(StockSummary, 22).Value = CPrice
                

                            'Designate where the openprice data is. Designated as the difference between the first open price and last open price.
                            Oprice3 = Oprice1 - Oprice2
                            'Output Open Price
                            'Cells(StockSummary, 19).Value = Oprice3
                            
                            'Reset open pricing
                                Oprice1 = 0
                                Oprice2 = 0
                                
                                                                            
                          Dim YearChange As Double
                            
                                Oprice5 = Oprice3 + Oprice4
                                'Cells(StockSummary, 21).Value = Oprice5
                            
                                YearChange = CPrice - Oprice5
                                    Cells(StockSummary, 11).Value = Format(YearChange, "currency")
                                                                    
                                    Dim pChange As Double
            
                                    If Oprice5 <> 0 Then
                                        pChange = YearChange / Oprice5
                                        Cells(StockSummary, 12).Value = Format(pChange, "Percent")
                                    Else
                                        Cells(StockSummary, 12).Value = Format(0, "Percent")
                                    End If
                                    
                'P1 Move down one
                StockSummary = StockSummary + 1
      
                'P1 Reset the Stock Volume
                TotalStockVol = 0
                         
            'P1 If the Ticker is the same as the next entry
            Else

                'P1 Keep adding to the Total Stock Volume
                TotalStockVol = TotalStockVol + Cells(i, 7).Value

                            'Sum up the openprice data is
                            Oprice1 = Oprice1 + Cells(i, 3).Value
                            'Sum up the openprice data is
                            Oprice2 = Oprice2 + Cells(i + 1, 3).Value

            End If

        Next i
       
'Part 4 (P4): Designate and search for greatest % increase, decrease and greatest total volume Stock. Identify the Ticker designation and value for each.

'P4 Set the variable
    Dim Great As Double
    Dim Great_Ticker As String
    Dim k As Integer
    Dim LastRow_2 As Double
    
    
    'P4 Label the headers for the Greatest Table
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
      
    LastRow_2 = Cells(Rows.Count, 10).End(xlUp).Row
   
    For k = 2 To LastRow_2
         
        'P4 Set the values for the variables
        Great = Cells(k, 12)
        Great_Ticker = Cells(k, 10)
        Great_volume = Cells(k, 13)
    

        'P4 Max and min values for the % change and total stock volume
        'Coding help from: https://www.ozgrid.com/forum/index.php?thread/87224-min-max-values-from-row-column-using-vba/
        'Coding help for $,%,etc. from: https://www.techonthenet.com/excel/formulas/format_number.php
        Cells(2, 17).Value = Format(WorksheetFunction.Max(Columns("L")), "percent")
        Cells(3, 17).Value = Format(WorksheetFunction.Min(Columns("L")), "percent")
        Cells(4, 17).Value = WorksheetFunction.Max(Columns("M"))
        
        'P4 To grab the ticker information for each greatest value
        If Great = Range("Q2") Then
            Range("P2") = Great_Ticker
            
            ElseIf Great = Range("Q3") Then
            Range("P3") = Great_Ticker
            
            ElseIf Great_volume = Range("Q4") Then
            Range("P4") = Great_Ticker
        End If
    
    Next k


'Code to ensure action cycles through each worksheet
'MsgBox ("Sheet " & ws.Name & " has " & LastRow & " rows")

'Next ws


End Sub

