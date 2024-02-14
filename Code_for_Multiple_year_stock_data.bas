Attribute VB_Name = "code"
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock. The result should match the following image:

Sub Stock_Analysis()

'THIS IS THE CORRECT CODE FOR GRADING

'Set Initial Variables for Ticker
Dim Ticker As String
'Set Initial Variables for Ticker Total
Dim Ticker_Volume_Total As LongLong
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim LR As Long
'For understanding LR = Last Row
'Dim LastRow As Long

Dim WS As Worksheet
'used to go through each worksheet

For Each WS In ThisWorkbook.Worksheets
'Going through each worksheet

'Create Summary Table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Summary Table Headers
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
         'Aka Ticker_Volume_Total

'Set Values for each Summary worksheet
Ticker_Volume_Total = 0
Yearly_Change = 0
Close_Price = 0
Open_Price = WS.Cells(2, 3).Value
'Comparing to the row below

'get the last row number with data
LR = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
MsgBox LR

'Find data from 2 to last row
For i = 2 To LR
    
'-------------------------------

'Yearly Change Calculation: Yearly change = Close Price - Open Price

'Check to see if we are still in Ticker Name and if not...
If WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value Then

    'Set Ticker Name
    Ticker = WS.Cells(i, 1).Value
    
  'Total Volume
    Ticker_Volume_Total = Ticker_Volume_Total + WS.Cells(i, 7).Value
    'Volume = volume  + Cells(i, 7).Value
    
    'Closing Price
    Close_Price = WS.Cells(i, 6).Value
    
    'Calculate Yearly Changes
    Yearly_Change = (Close_Price - Open_Price)
    
    'Calculate Percent Change
    ' Percent_Change = Yearly_Change / Open_Price
    
    If Open_Price = 0 Then
        Percent_Change = 0
        
        Else
        Percent_Change = Yearly_Change / Open_Price
        
        End If
    
    'Add yearly change in summary table
    'Cells(1, 10).Value = "Yearly Change"
    
    'Add to Summary Table
    WS.Cells(Summary_Table_Row, 9).Value = Ticker
    
    WS.Cells(Summary_Table_Row, 10).Value = Yearly_Change
    
    WS.Cells(Summary_Table_Row, 11).Value = Percent_Change
    WS.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
    
    WS.Cells(Summary_Table_Row, 12).Value = Ticker_Volume_Total
    
    Summary_Table_Row = Summary_Table_Row + 1
    
     'open price compared to row below
     Open_Price = WS.Cells(i + 1, 3)
    
    'reset Volume
    Ticker_Volume_Total = 0
    
    Else
    
    're-estabilish volume if same
    Ticker_Volume_Total = WS.Cells(i, 7).Value + Ticker_Volume_Total
   
    End If
    
Next i

MsgBox "Test 1"
'------------------------------------------

'Color Coding Summary Table based on Positive or Negative for Yearly Change

'Color code to LR
For i = 2 To LR

If WS.Cells(i, 10).Value > 0 Then
    WS.Cells(i, 10).Interior.Color = vbGreen

Else
WS.Cells(i, 10).Interior.Color = vbRed

End If

Next i

MsgBox "test 2"

'===============================================
'Calculate Min Max
'All three of the following values are calculated correctly and displayed in the output:
'Greatest % Increase (5 points)
'Greatest % Decrease (5 points)
'Greatest Total Volume (5 points)
'Pulling % increase and Decrease from Row K aka Percent Change
'Pulling Total Volum from Row L aka Total Stsock Value

'Set Variables
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Total_Volume As Long
Dim Greatest_Increase_Ticker As String
Dim Greatest_Decrease_Ticker As String
Dim Greatest_Volume_Ticker As String

'Create Min Max Summary Table
Dim Min_Max_Summary_Table As Integer
Min_Max_Summary_Table = 2
'not sure if this is needed

'Min Max Summary Table Lables
WS.Cells(1, 15).Value = "Ticker"
WS.Cells(1, 16).Value = "Value"
WS.Cells(2, 14).Value = "Greatest Percent Increase"
WS.Cells(3, 14).Value = "Greatest percent Decrease"
WS.Cells(4, 14).Value = "Greatest Total Volume"
'Range("Q" And Summary_Ticker_Row).NumberFormat = "Percent"

MsgBox "Test 3"
'==============================================
'Set values
'Percent_Change = 0
'Ticker_Volume_Total = 0
'Greatest_Percent_Increase = 0
'Greatest_Percent_Decrease = 0
'Greatest_Total_Volume = 0
'Ticker = ""
'Greatest_Increase_Ticker = ""
'Greatest_Decrease_Ticker = ""
'Greatest_Volume_Ticker = ""
'
'For i = 2 To Summary_Table_Row - 1
'
''Finding Greatest % Increase
'If WS.Cells(i, 11).Value > Greatest_Percent_Increase Then
'    Greatest_Percent_Increase = WS.Cells(i, 11).Value
'    Greatest_Increase_Ticker = WS.Cells(i, 9).Value
'
'End If
'
''Find Greatest % Decrease
'
'If WS.Cells(i, 11).Value < Greatest_Percent_Decrease Then
'    Greatest_Percent_Decrease = WS.Cells(i, 11).Value
'    Greatest_Decrease_Ticker = WS.Cells(i, 9).Value
'
'End If
'
''Finding Greatest Total Volume
'If WS.Cells(i, 12).Value > Greatest_Total_Volume Then
'    Greatest_Total_Volume = WS.Cells(i, 12).Value
'    Greatest_Volume_Ticker = WS.Cells(i, 9).Value
'
'End If
'
''Add results to Min_Max_Summary_Table
'
''Add results for Greatest Percent Increase
'    WS.Cells(2, 15).Value = Greatest_Increase_Ticker
'    WS.Cells(2, 16).Value = Greatest_Percent_Increase
'    WS.Cells(2, 16).NumberFormat = "0.00%"
'
'
''Add results for Greatest Percent Increase decrease
'    WS.Cells(3, 15).Value = Greatest_Decrease_Ticker
'    WS.Cells(3, 16).Value = Greatest_Percent_Decrease
'    WS.Cells(3, 16).NumberFormat = "0.00%"
'
'
''Add results for Greatest_Total_Volume
'    WS.Cells(4, 15).Value = Greatest_Volume_Ticker
'    WS.Cells(4, 16).Value = Greatest_Total_Volume

'Next i

Next WS

End Sub

