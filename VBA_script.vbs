Attribute VB_Name = "Module1"
Sub Test_Data()

'Create a script that will loop through all the stocks for one year and output the following information.
'
'The ticker symbol. (column A)

Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets


Dim LastRow As Double
Dim ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim ClosePrice As Double
Dim OpenPrice As Double

TotalStockVolume = 0

Dim SummaryTableRow As Integer
SummaryTableRow = 2

LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
OpenPrice = WS.Cells(2, 3).Value


For i = 2 To LastRow

TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
    'check if ticker is different
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    
       
    ClosePrice = WS.Cells(i, 6).Value
    YearlyChange = ClosePrice - OpenPrice
    
    If OpenPrice = 0 Then
    PercentChange = 0
    Else
    PercentChange = YearlyChange / OpenPrice
    
    End If
    
    OpenPrice = WS.Cells(i + 1, 3).Value
        
    
    'set each ticker name
        ticker = WS.Cells(i, 1).Value
               
        WS.Cells(SummaryTableRow, 9).Value = ticker
        WS.Cells(SummaryTableRow, 10).Value = YearlyChange
        WS.Cells(SummaryTableRow, 11).Value = PercentChange
        WS.Cells(SummaryTableRow, 12).Value = TotalStockVolume
        
                                         
        TotalStockVolume = 0
           
      If WS.Cells(SummaryTableRow, 10).Value >= 0 Then
      WS.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
      Else
      WS.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            
      End If
      
      SummaryTableRow = SummaryTableRow + 1
              
    End If
        
Next i

Next WS


'YearlyChange = ClosePrice... WS.Cells(i +1, 6) & (20161230) - OpenPrice (20160101)


'PercentChange = YearlyChange / OpenPrice
'YearlyChange = WS.Cells(i +1,10) &i,10).value
'OpenPrice = WS.Cells(i +1,3)

'TotalStockVolume = TotalStockVolume +WS.Cells(i,7).value

'
'Yearly change from opening price (column C) at the beginning of a given year to the closing price (column F) at the end of that year.
'
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'The total stock volume of the stock.
'



'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'have conditional formatting that will highling positive change in green and negative change in red.

End Sub

