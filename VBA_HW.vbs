Sub RunAll()
    Call Ticker
    Call PriceChange
    Call PrecentageChange
    Call Total_Stock_Volume
End Sub


Sub Ticker():

 'Create the Ticker Symbol variable
 Dim TickerS As String
 
 'create a summary table variable
 Dim Summary_Table_Row As Integer
 
  ' Make summary table equal to 2 so that the table starts at row two.
  Summary_Table_Row = 2
  
 'Create column header name
  Range("i1") = "Ticker Symbol"
  
  'make the data autofit
  Columns("A:P").AutoFit
  
  'Count number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    'create loops for rows
    For i = 2 To lastrow
    
        'Set condition so that it checks if cells match
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
          'Set TickerS varible to show the value in the cell above the one that is different
          TickerS = Cells(i, 1).Value
          
          'Set the column I so that the variable TickerS show in that column
          Range("I" & Summary_Table_Row).Value = TickerS
          
          ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
          
        'End the conditional
        End If
        
        'End the loop
        Next i
    
  
End Sub


Sub PriceChange():
    
 'Create column header name
 Range("j1") = "Yearly Change"
     
 'make the data autofit
  Columns("A:P").AutoFit

  ' Set an initial variable for holding the total per open price
  Dim OpenPriceSummary As Double
  OpenPriceSummary = 0
  
  'Set an initial variable for closed price summary
  Dim ClosedPriceSummary As Double
  ClosedPriceSummary = 0

  ' Keep track of the location for each price summary
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Count number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the prices
  For i = 2 To lastrow

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Add to the Open Price total
      OpenPriceSummary = OpenPriceSummary + Cells(i, 3).Value
      
      'Add to Closed Price Total
      ClosedPriceSummary = ClosedPriceSummary + Cells(i, 6).Value

      ' Print the the difference in open and closed prices
      Range("J" & Summary_Table_Row).Value = ClosedPriceSummary - OpenPriceSummary
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the open price
      OpenPriceSummary = 0
      
      'Reset the closed price
      ClosedPriceSummary = 0

  'Set conditional for color
      If Cells(Summary_Table_Row, 10).Value > 0 Then
            
           'Set color to be green if value is greater than 0
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
        Else
        
           'set color to be red for anything else
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3



    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the open price total
      OpenPriceSummary = OpenPriceSummary + Cells(i, 3).Value
      
      ' Add to the closed price total
      ClosedPriceSummary = ClosedPriceSummary + Cells(i, 6).Value
     
      
      
      End If
      
  Next i
     
     
     
End Sub

Sub PrecentageChange():
    
 'Create column header name
 Range("k1") = "Percentage Change"
     
 'make the data autofit
  Columns("A:P").AutoFit

  ' Set an initial variable for holding the total per open price
  Dim OpenPriceSummary As Double
  OpenPriceSummary = 0
  
  'Set an initial variable for closed price summary
  Dim ClosedPriceSummary As Double
  ClosedPriceSummary = 0

  ' Keep track of the location for each price summary
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Count number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the prices
  For i = 2 To lastrow
      
    'Format column "K" to be a percentage
     Cells(i, 11).NumberFormat = "0.00%"
     
    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Add to the Open Price total
      OpenPriceSummary = OpenPriceSummary + Cells(i + 1, 3).Value
      
      'Add to Closed Price Total
      ClosedPriceSummary = ClosedPriceSummary + Cells(i, 6).Value
      
      ' Print the the difference in open and closed prices
      Range("K" & Summary_Table_Row).Value = (ClosedPriceSummary - OpenPriceSummary) / OpenPriceSummary
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the open price
      OpenPriceSummary = 0
      
      'Reset the closed price
      ClosedPriceSummary = 0

    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the open price total
      OpenPriceSummary = OpenPriceSummary + Cells(i, 3).Value
      
      ' Add to the closed price total
      ClosedPriceSummary = ClosedPriceSummary + Cells(i, 6).Value
     
      
      End If
      
  Next i
     
     
     
End Sub

Sub Total_Stock_Volume()

'Create column header name
 Range("L1") = "Total Stock Volume"
     
 'make the data autofit
  Columns("A:P").AutoFit

  ' Set an initial variable for holding the total Stock volume
  Dim TotStockVol As Double
  TotStockVol = 0

  ' Keep track of the location for each stock volume summary
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Count number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the prices
  For i = 2 To lastrow

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Add to the Open Stock Volume Total
      TotStockVol = TotStockVol + Cells(i, 7).Value
      
      ' Print the the total stock volume in the correct cells
      Range("L" & Summary_Table_Row).Value = TotStockVol
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total stock volume
      TotStockVol = 0

    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the stock volume total
      TotStockVol = TotStockVol + Cells(i, 7).Value
     
     
    End If
      
      
  Next i
     
     
     
End Sub





Sub PrecentageChange():
    
 'Create column header name
 Range("k1") = "Percentage Change"
     
 'make the data autofit
  Columns("A:P").AutoFit

  ' Set an initial variable for holding the total per open price
  Dim OpenPriceSummary As Double
  OpenPriceSummary = 0
  
  'Set an initial variable for closed price summary
  Dim ClosedPriceSummary As Double
  ClosedPriceSummary = 0

  ' Keep track of the location for each price summary
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Count number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the prices
  For i = 2 To lastrow
      
    'Format column "K" to be a percentage
     Cells(i, 11).NumberFormat = "0.00%"
     
    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Add to the Open Price total
      OpenPriceSummary = OpenPriceSummary + Cells(i + 1, 3).Value
      
      'Add to Closed Price Total
      ClosedPriceSummary = ClosedPriceSummary + Cells(i, 6).Value
      
      ' Print the the difference in open and closed prices
      Range("K" & Summary_Table_Row).Value = (ClosedPriceSummary - OpenPriceSummary) / OpenPriceSummary
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the open price
      OpenPriceSummary = 0
      
      'Reset the closed price
      ClosedPriceSummary = 0

    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the open price total
      OpenPriceSummary = OpenPriceSummary + Cells(i, 3).Value
      
      ' Add to the closed price total
      ClosedPriceSummary = ClosedPriceSummary + Cells(i, 6).Value
     
      
      End If
      
  Next i
     
     
     
End Sub

Sub Total_Stock_Volume()

'Create column header name
 Range("L1") = "Total Stock Volume"
     
 'make the data autofit
  Columns("A:P").AutoFit

  ' Set an initial variable for holding the total Stock volume
  Dim TotStockVol As Double
  TotStockVol = 0

  ' Keep track of the location for each stock volume summary
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Count number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the prices
  For i = 2 To lastrow

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Add to the Open Stock Volume Total
      TotStockVol = TotStockVol + Cells(i, 7).Value
      
      ' Print the the total stock volume in the correct cells
      Range("L" & Summary_Table_Row).Value = TotStockVol
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total stock volume
      TotStockVol = 0

    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the stock volume total
      TotStockVol = TotStockVol + Cells(i, 7).Value
     
     
    End If
      
      
  Next i
     
     
     
End Sub








