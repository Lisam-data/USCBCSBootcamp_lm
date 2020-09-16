'VBA Code for homework
Attribute VB_Name = "Module2"
Sub MasterTickerCal():

'Declare Worksheet as ws
Dim ws As Worksheet

'Loop through every worksheet in a workbook
For Each ws In Worksheets
  ws.Activate
'Declare an initial variable for holding the ticker name
Dim Ticker_Name As String

'Declare an initial variable for holding the total per ticker
Dim Ticker_Total As Double
Ticker_Total = 0

'Keep track of the location for each ticker name and ticker total in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
    
'Declare an initial variable for holding the open price
Dim openprice As Double
openprice = 0

'Declare an initial variable for holding the close price
Dim closeprice As Double

 'Declare an initial variable for holding the change price and the calculation of changeprice is the difference between endprice and open price
Dim changeprice As Double
changeprice = endprice - openprice

'Declare an initial variable for percent change
Dim percentchange As Double
percentchange = 0
 
'print the headers under these columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Open Price"
Range("K1").Value = "End Price"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

'set the maxChange as a variable, this one is for the challenges part
Dim maxChange As Double
maxChange = 0

'declare a variable for greatest change ticker, greatest decrease, and greatest volume
Dim maxChangeticker As String
Dim minChangeticker As String
Dim maxVolumeticker As String
'set the minChange as a variable
Dim minChange As Double
minChange = 0

'set the maxtotalvolume as a variable
Dim maxtotalvolume As Double
maxtotalvolume = 0


'print these title
Range("Q2") = "Greatest % increase"
Range("Q3") = "Greatest % Decrease"
Range("Q4") = "Greatest Total Volume"

    ' Loop through all ticker names under column A
    For Row = 2 To ws.Cells(1, 1).End(xlDown).Row
        
        ' if the next cell value (tickername) is not the same as the current cell (tickername) value, Then
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
           'set the ticker name
            Ticker_Name = ws.Cells(Row, 1).Value
            'add the volume to the ticker total and divide by 1,000
            Ticker_Total = Ticker_Total + (ws.Cells(Row, 7).Value / 1000)
            'set close price which is under column F
            closeprice = ws.Cells(Row, 6).Value
            'set the change price
            changeprice = (closeprice - openprice)
            
            'print the ticker name, ticker total, close price and change price in these columns in the summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            ws.Range("N" & Summary_Table_Row).Value = Ticker_Total
            ws.Range("K" & Summary_Table_Row).Value = closeprice
            ws.Range("L" & Summary_Table_Row).Value = changeprice
           
           'set the cell color to green if the change price is positive
            If changeprice > 0 Then
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                 
            'Else, set the cell color to red if the change price is a negative number
            Else
                 ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                 
            End If
            
            'If the open price is 0 such as ticker PLNT, then force the percentchange to 0
            If openprice = 0 Then
                        
                 percentchange = 0
                 
             'else the percentchange is equal to changeprice / open price
            Else
                percentchange = changeprice / openprice
                
            End If
          
             'set the percent change in column M
            ws.Range("M" & Summary_Table_Row).Value = percentchange
            'change the column M to Percentage format
            ws.Range("M" & Summary_Table_Row).NumberFormat = "#.##%"
            'set the total volume number format
            ws.Range("N" & Summary_Table_Row).NumberFormat = "#,###,##0"


            'if the percentchange is greater than the maxchange variable then
            If percentchange > maxChange Then
            
            'set the percentchange to be maxchange
                maxChange = percentchange
             
            End If
            
            'if percenchange < minchange then
           If percentchange < minChange Then
           
           'set the percentchange = to minchange
              minChange = percentchange
              
           End If
            
            If Ticker_Total > maxtotalvolume Then
            
                maxtotalvolume = Ticker_Total
                
            End If
            
            
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
             ' Reset the ticker Total, so it will calculate the new ticker total
            Ticker_Total = 0
            'Reset the open price to zero
            openprice = 0   'reset to zero
        
        
        Else
            If openprice = 0 Then
                'set open price
                openprice = ws.Cells(Row, 3).Value
                'print open price to column J
                ws.Range("J" & Summary_Table_Row).Value = openprice
            End If
              
            Ticker_Total = Ticker_Total + (ws.Cells(Row, 7).Value / 1000)
            ws.Range("N" & Summary_Table_Row).NumberFormat = "#,###,###"
        
        End If
        
       
    Next Row
       
       'print the result of maxchange, minChange, and maxtotalvolume and the number format
              
        ws.Range("S2") = maxChange
        ws.Range("S2").NumberFormat = "#,###.##%"
        
        'for loop through column M
        For i = 2 To ws.Cells(1, 13).End(xlDown).Row
        
        'when the value of column = maxChange
        If ws.Cells(i, 13).Value = ws.Range("S2").Value Then
        
        'set the ticker name from column I to be maxChangeticker
            maxChangeticker = ws.Cells(i, 9)
            
        End If
        
        'print max change ticker in Cell R2
        ws.Range("R2") = maxChangeticker
        
        'print the lowest change in Cell S3 and set the number format
        ws.Range("S3") = minChange
        ws.Range("S3").NumberFormat = "#,###.##%"
        
        'when the value under Column M = minChange
         If ws.Cells(i, 13).Value = ws.Range("S3").Value Then
        
        'Set the ticker name from column I to be minchangeticker
            minChangeticker = ws.Cells(i, 9)
            
        End If
        'print minchangeticker in Cell R3
        ws.Range("R3") = minChangeticker
        Next i
        
        'Print maxtotalvolume and set the number format
        ws.Range("S4") = maxtotalvolume
        ws.Range("S4").NumberFormat = "#,###,###"
        
        'loop for column N Total Volume
        For j = 2 To ws.Cells(1, 14).End(xlDown).Row
        
        'when cell value in column N = maxtotalvolume
        If ws.Cells(j, 14).Value = ws.Range("S4").Value Then
        
        'set the ticker name from column I to be the maxVolumeticker
          maxVolumeticker = ws.Cells(j, 9)
        
        End If
        
        ws.Range("R4") = maxVolumeticker
        
        Next j
        
 Next ws


End Sub

