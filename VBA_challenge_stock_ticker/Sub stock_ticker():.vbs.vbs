Sub stock_ticker():

'Loop through all sheets
For Each ws In Worksheets
 
 'Set Headers for column
 ws.Range("K1").Value = "Ticker"
 ws.Range("L1").Value = "Yearly_Change"
 ws.Range("M1").Value = "Percentage Change"
 ws.Range("N1").Value = "Total Ticker Volume"
 ws.Range("R1").Value = "Ticker"
 ws.Range("S1").Value = "Value"

 
  ' Set an initial variable for ticker
  Dim ticker_name As String
  Dim rowCount As Long
  Dim open_price As Double
  Dim closing_price As Double
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim increase As Double
  Dim increase_ticker As String


Dim WS_Count As Integer

  ' Set an initial variable for holding the total per ticker""
    Dim Ticker_Total As Double
    Ticker_Total = 0

  ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    open_price = Cells(2, 3).Value
    increase = 0
  
  ' Loop through all ticker
    Dim lastrow As Double
    lastrow = Cells(Rows.Count, "a").End(xlUp).Row
    For i = 2 To lastrow
  
  
    ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    

      ' Set the ticker name
        ticker_name = Cells(i, 1).Value

      ' Add to the ticker Total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        close_price = Cells(i, 6).Value
        yearly_change = close_price - open_price
        percent_change = Round(yearly_change / open_price, 2)
        If yearly_change > 0 Then
        Cells(Summary_Table_Row, 12).Interior.ColorIndex = 4
       
       Else
         
        Cells(Summary_Table_Row, 12).Interior.ColorIndex = 3
        
        End If
        
        If percent_change > increase Then
        increase = percent_change
        increase_ticker = ticker_name
        
        End If
         
        
      ' Print the ticker in the Summary Table
        Range("k" & Summary_Table_Row).Value = ticker_name
        Range("L" & Summary_Table_Row).Value = yearly_change
        Range("m" & Summary_Table_Row).Value = percent_change
        
      ' Print the ticker Amount to the Summary Table
        Range("N" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker Total
        Ticker_Total = 0
        open_price = Cells(i + 1, 3).Value
      

    ' If the cell immediately following a row is the same ticker...
    
    Else

      ' Add to the Ticker_Total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value

    End If

  Next i
  
   'print the ticker in the summary table'
   Range("Q" & Summary_Table_Row).Value = Greatest_percentage_increase
   Range("R" & Summary_Table_Row).Value = Greatest_percentage_decrease
   Range("S" & Summary_Table_Row).Value = Greatest_percentage_Total_Volume
   
   Next ws
   


End Sub

