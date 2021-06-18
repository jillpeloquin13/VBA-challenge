Attribute VB_Name = "Module2"
Sub StockoutputWS()

For Each ws In Worksheets


'Find the last row
Dim lastRow As LongLong
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1


'Define the summary-table row
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Create column headers
ws.Range("I1").Value = "ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
   
'Set the stock volume and opening price as variables and give them an inital value
Dim Stock_Volume As LongLong
Stock_Volume = 0
     
Dim Opening_price_counter As Double
Opening_price_counter = 0

For i = 2 To lastRow
        'Set the Ticker name
        Dim ticker As String
        ticker = ws.Cells(i, 1).Value
     
        'Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
          
        'start counting the stock volume
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        Opening_price_counter = Opening_price_counter + 1
          
        Else
        
            'count the last cell for stock volume, but do not increment the ticker
             Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

                ' Print the Credit Card Brand in the Summary Table
                Dim Closing_price As Double
                Dim Opening_price As Double
                Dim Percent_change As Double
                Closing_price = ws.Cells(i, 6).Value
                Opening_price = ws.Cells(i - Opening_price_counter, 3).Value
                 
                          'handle errors when dividing by 0
                          If Opening_price > 0 Then
                               Percentage_change = ((Closing_price - Opening_price) / Opening_price) * 100
                          
                          Else
                               Percentage_change = 0
                          
                          End If
                
                ws.Range("I" & Summary_Table_Row).Value = ticker
                ws.Range("J" & Summary_Table_Row).Value = Closing_price - Opening_price
                ws.Range("K" & Summary_Table_Row).Value = Percentage_change
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                
                'reset the counters for the next ticker
                Stock_Volume = 0
                Opening_price_counter = 0
                
 
        End If
Next i


    'add the condtional formatting
      Dim MyRange As Range
    
    'Create range object
    Set MyRange = ws.Range("J" & Summary_Table_Row)
    'Add first rule
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        MyRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        'Add second rule
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        MyRange.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
    

'Start the Bonus section


'Find the last row

Dim lastRowSummary As Long
lastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1

'Create column and row headers
ws.Range("P1").Value = "ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim Max_perc_Increase As Double
Dim Max_perc_Decrease As Double
Dim Max_Volume As LongLong


    For a = 2 To lastRowSummary
    If ws.Cells(a, 11) > Max_perc_Increase Then
        Max_perc_Increase = ws.Cells(a, 11).Value
        ws.Range("P2").Value = ws.Cells(a, 9).Value
        ws.Range("Q2").Value = Max_perc_Increase
    End If
    Next a

    For b = 2 To lastRowSummary
    If Cells(b, 11) < Max_perc_Decrease Then
        Max_perc_Decrease = ws.Cells(b, 11).Value
        ws.Range("P3").Value = ws.Cells(b, 9).Value
        ws.Range("Q3").Value = Max_perc_Decrease
    End If
    Next b

    For c = 2 To lastRowSummary
    If ws.Cells(c, 12) > Max_Volume Then
        Max_Volume = ws.Cells(c, 12).Value
        ws.Range("P4").Value = ws.Cells(c, 9).Value
        ws.Range("Q4").Value = Max_Volume
    End If
    Next c


 Next ws

End Sub



