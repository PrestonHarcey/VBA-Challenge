Attribute VB_Name = "Module1"
Sub Stock_Check()
    'Create Variables to hold requested information
'Variable for holding Stock Ticker
    Dim Ticker As String
'Variable for holding yearly Change
    Dim Yearly_Change As Double
'Variable for holding Percentage Change
    Dim Percent_Change As Double
'Variable for holding The Total Volume
    Dim Total_Volume As Long
    'Keep track of the location for each stock ticker in the stock table
    Dim Stock_Table_Row As Integer
    Summary_Table_Row = 2
'Loop through the dataset
    Dim I As Long
    For I = 1 To Rows.Count


    
    'set ticker
        Ticker = Cells(I, 1).Value
        'check if cells are still using the same ticker, if not,
        If Cells(I + 1).Value <> Cells(I, 1).Value Then
        
        
        
        'create a range for the data
            Dim rng As Range, cell As Range
            Set rng = Range("C2:F2")
 
'Set the yearly change
    
        
            Yearly_Change = Application.WorksheetFunction.StDev(rng)
        
        
        

'Set the yearly percentage and format to percentage
            Precent_Change = FormatPercent(Application.WorksheetFunction.StDev(rng), 2, vbUseDefault, vbUseDefault, vbUseDefault)
        


'add total volume for that stock
       

'Print the Ticker Name in the sheet
            Range("I" & Summary_Table_Row).Value = Ticker

'Print the Yearly change
            Range("J" & Summary_Table_Row).Value = Yearly_Change
'Print the Percent Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
'Print the Total
            Range("L" & Summary_Table_Row).Value = Total_Volume

'add 1 to the rows to continue populating.
            Summary_Table_Row = Summary_Table_Row + 1
'reset the ticker total
            Total_Volume = 0
'if the cells are still of the same ticker
      
'continue adding the total
           
        End If

   Next I

End Sub
