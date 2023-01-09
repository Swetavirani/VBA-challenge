Sub VBAchallenge()

' Loop Through all sheets

For Each ws In Worksheets

Dim Ticker As String
Dim Yearchange As Double
Dim Percentchange As Double
Dim yearopendate As String
Dim yearclosedate As String
Dim Tickeryearopenprice As Double
Dim Tickeryearendprice As Double
Dim LastRow As Long
Dim i As Long
Dim Summary_Table_Row As Long
Dim Volume_total As Double
Dim Greatest_percentage_Increase As Double
Dim Greatest_percentage_Decrease As Double
Dim Greatest_total_volume As Double
Dim j As Long


' Creating headers for the both the summary tables

ws.Range("L1").Value = "Ticker"
ws.Range("M1").Value = "Yearly Change"
ws.Range("N1").Value = "Percent Change"
ws.Range("O1").Value = "Total Stock Volume"
ws.Range("S1").Value = "Ticker"
ws.Range("T1").Value = "Value"
ws.Range("R2").Value = "Greatest % Increase"
ws.Range("R3").Value = "Greatest % Decrease"
ws.Range("R4").Value = "Greatest Total Volume"

' Retrieve the year from Worksheet name , store Year start and end dates

yearopendate = ws.Name & "0102"
yearclosedate = ws.Name & "1231"

' Counts the number of rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox (LastRow & "is last row")

' Keeping track of the location for each ticker in the summary table starting from Row 2
Summary_Table_Row = 2

' Loop through all the tickers

For i = 2 To LastRow

    ' Checks if this is ticker's beginning of the year price details, retrieves and stores the open price
        
    If ws.Cells(i, 2).Value = yearopendate Then
       Tickeryearopenprice = ws.Cells(i, 3).Value

    End If

' Checks if this is ticker's end of the year price details, retrieves and stores the close price

    If ws.Cells(i, 2).Value = yearclosedate Then
       Tickeryearendprice = ws.Cells(i, 6).Value
   
    End If

    ' Checks if we are still within the same ticker, if it is not...
            
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

       Ticker = ws.Cells(i, 1).Value
       Yearchange = Tickeryearendprice - Tickeryearopenprice
       Percentchange = Yearchange / Tickeryearopenprice
       Volume_total = Volume_total + ws.Cells(i, 7).Value
       ws.Range("L" & Summary_Table_Row).Value = Ticker
       ws.Range("M" & Summary_Table_Row).Value = Yearchange
       ws.Range("N" & Summary_Table_Row).Value = Percentchange
       ws.Range("O" & Summary_Table_Row).Value = Volume_total
       
       ' Moving the counter to the next row for next ticker and resetting the volume total counter for next ticker
         Summary_Table_Row = Summary_Table_Row + 1
         Volume_total = 0
       
   ' Conditional formatting to highlight postive change in green and negative in red
         
    If Yearchange < 0 Then
       ws.Range("M" & Summary_Table_Row - 1).Interior.ColorIndex = 3
   
    ElseIf Yearchange > 0 Then
           ws.Range("M" & Summary_Table_Row - 1).Interior.ColorIndex = 4
           
    End If
           
    Else
        ' Add to the Volume Total
          Volume_total = Volume_total + ws.Cells(i, 7).Value
        
    End If
        
Next i


' Looping through all the tickers in first summary table for greatest changes

' Calculating Greatest % increase, retrieving the respective ticker and % increase

For j = 2 To Summary_Table_Row

    Greatest_percentage_Increase = WorksheetFunction.Max(ws.Range("N2" & ":" & "N" & Summary_Table_Row))
       
           
    If ws.Cells(j, 14).Value = Greatest_percentage_Increase Then
       ws.Cells(2, 19).Value = ws.Cells(j, 12)
       ws.Cells(2, 20).Value = ws.Cells(j, 14)
       
    End If

' Calculating Greatest % decrease, retrieving the respective ticker and % decrease
    
    
    Greatest_percentage_Decrease = WorksheetFunction.Min(ws.Range("N2" & ":" & "N" & Summary_Table_Row))
    
    If ws.Cells(j, 14).Value = Greatest_percentage_Decrease Then
       ws.Cells(3, 19).Value = ws.Cells(j, 12)
       ws.Cells(3, 20).Value = ws.Cells(j, 14)
              
    End If
       
 ' Calculating Greatest total volume , retrieving the respective ticker and total voume
  
  Greatest_total_volume = WorksheetFunction.Max(ws.Range("O2" & ":" & "O" & Summary_Table_Row))
    
    If ws.Cells(j, 15).Value = Greatest_total_volume Then
        
       ws.Cells(4, 19).Value = ws.Cells(j, 12)
       ws.Cells(4, 20).Value = ws.Cells(j, 15)
       
    End If
    
Next j
                   
' Formatting

ws.Columns("M").NumberFormat = "0.00"
ws.Columns("N").NumberFormat = "0.00%"
ws.Range("T2:T3").NumberFormat = "0.00%"
ws.Columns("L:T").AutoFit


'MsgBox ws.Name

Next ws

MsgBox ("All Done")


End Sub



