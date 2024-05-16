Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()


For Each ws In Worksheets

  ' Set an initial variable for holding the Ticker Name
  Dim Ticker_Name As String

  ' Declaring variables for holding the summary table columns, Stock total and price calculations
  Dim Quarterly_Change As Double
  Dim Percent_Change As Double
  Dim Stock_Total As Double
  Stock_Total = 0
  Quarterly_Change = 0
  Percent_Change = 0
  Dim LastRow As Long
  Dim FirstRow As Long
  Dim Opening_price As Double
  Dim Closing_price As Double
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  
  
  'Label first summary table columns
  
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Quarterly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   
   
  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Opening value of stock at beginning of quarter at row 2 to get opening price of first Ticker on each sheet
   Opening_price = ws.Cells(2, 3)

  ' Loop through all Tickers
  For i = 2 To LastRow
  
      
      
    ' Check if we are still within the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
      ' Set the Ticker Name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
      Closing_price = ws.Cells(i, 6).Value
      
      


      'Calculate quarterly change
      
      Quarterly_Change = Closing_price - Opening_price
      
      'Calculate Percent Change
      
      Percent_Change = (Quarterly_Change / Opening_price)
      
      ' Print the Ticker Name in the Summary Table
      
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      'Print the Quarterly Change in the Summary table
      ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
      
      'Print the Percent Change in the Summary table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      'Convert to percentage format
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      

      ' Print the Stock Volume Total to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total,closing price, qaurterly change and percent change
      Stock_Total = 0
      Closing_price = 0
      Quarterly_Change = 0
      Percent_Change = 0
      
      'Setting the opening price to first row of quarter opening price corresponding to the next Ticker
      Opening_price = ws.Cells(i + 1, 3).Value


    ' If the cell immediately following a row is the same Ticker
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
    
    
    End If
    
    
       'Declaring variables for cell formatting
       
       
           Next i
          
           'Declaring variables for cell formatting
          Dim QC_Last_Row As Long
       QC_Last_Row = ws.Cells(Rows.Count, 10).End(xlUp).Row
       
       'Formatting for Color Coding
       
       For i = 2 To QC_Last_Row
       
       If ws.Cells(i, 10).Value > 0 Then
           ws.Cells(i, 10).Interior.ColorIndex = 4
           
        ElseIf ws.Cells(i, 10).Value < 0 Then
          ws.Cells(i, 10).Interior.ColorIndex = 3
        
        End If
        
        Next i
        
        'Second Summary Table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest%Increase"
        ws.Cells(3, 15).Value = "Greatest%Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
          
         'Declaring variables to find min and max of percentage change
  
          Dim PC_Last_Row As Long
          PC_Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
          Dim PC_max As Double
          PC_max = 0
          Dim PC_min As Double
          PC_min = 0
  
    For i = 2 To PC_Last_Row

         If PC_max < ws.Cells(i, 11).Value Then
           PC_max = ws.Cells(i, 11).Value
           ws.Cells(2, 17).Value = PC_max
           ws.Cells(2, 17).NumberFormat = "0.00%"
           ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
           
    ElseIf PC_min > ws.Cells(i, 11).Value Then
    
         PC_min = ws.Cells(i, 11).Value
         ws.Cells(3, 17).Value = PC_min
         ws.Cells(3, 17).NumberFormat = "0.00%"
         ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
         
    End If
     
      Next i

      'Declaring variable for total of the total stock volume

       Dim Last_Row_Ttl_Vol As Long
       Last_Row_Ttl_Vol = ws.Cells(Rows.Count, 12).End(xlUp).Row
       Dim Ttl_Vol_Max As Double
       Ttl_Vol_Max = 0


 
      For i = 2 To Last_Row_Ttl_Vol

     'Calculating max Total Volume from all Ticker totals

      If Ttl_Vol_Max < ws.Cells(i, 12).Value Then
       Ttl_Vol_Max = ws.Cells(i, 12).Value
       ws.Cells(4, 17).Value = Ttl_Vol_Max
       ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
       
        End If
      
      
      Next i
    
    Next ws
          
End Sub
