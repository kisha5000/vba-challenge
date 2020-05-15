Option Explicit

Sub LoopStockXchange()
'Define loop function to go through pages
    Dim Page As Byte
    Dim i As Long
'go through all worksheet count - declare and define. allsheets means "all worksheet"
    Dim allsheets As Byte
    allsheets = ThisWorkbook.Worksheets.Count

'loop through worksheet
    For Page = 1 To allsheets
'Activate current workbook = choosing the sheet's tab.
    ThisWorkbook.Worksheets(Page).Activate
'Declare everything in original worksheet set
    Dim ticker As String, yr_open As Double, yr_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim totalvol As Variant
    totalvol = 0
    
'Define the summary logic
    Dim Rollup_cell_all As Integer
    Rollup_cell_all = 2

    Dim lastrow As Variant
    
    Dim tickercounter As Long
    tickercounter = 0
    
'counts the number of rows
    lastrow = Cells(1, 1).End(xlDown).Row



    
'Populate Column Headers on every Worksheet
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"

        
'loop through cells on each worksheet
        For i = 2 To lastrow

'create if statement for combining like tickers. "<>" means not equal
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            

 ' print the tickers
                ticker = Cells(i, 1).Value
                yr_open = Cells((i) - tickercounter, 3).Value
                yr_close = Cells(i, 6).Value
              
                
' Add to the volume Total
                totalvol = totalvol + Cells(i, 7).Value
             
                
                            
'calculations for yearly change and percent change
                    yearly_change = yr_close - yr_open


                If yr_open <> 0 Then
                    percent_change = (yearly_change# / yr_open#)
                Else
                    percent_change = 0
                End If
                       
                     

'populate rollup values
                Cells(Rollup_cell_all, 9).Value = ticker
                Cells(Rollup_cell_all, 10).Value = yearly_change
                Cells(Rollup_cell_all, 11).Value = percent_change
                Cells(Rollup_cell_all, 12).Value = totalvol
  'populate interior cell color
 'Add one to the roll up row
                Rollup_cell_all = Rollup_cell_all + 1
                
                
 'Reset the Total Volume, Yearly change volume and percent change volume
            totalvol = 0
            tickercounter = 0
                
             Else
             
  ' Add to the volume Total
                totalvol = totalvol + Cells(i, 7).Value
                tickercounter = tickercounter + 1
               
                 
                
                 
            End If
            
'format color
            
             If Cells(Rollup_cell_all, 10).Value >= 0 Then
                        Cells(Rollup_cell_all, 10).Interior.ColorIndex = 4
                         Else
                    Cells(Rollup_cell_all, 10).Interior.ColorIndex = 3
                    End If
                 
'move to next row
        Next i
        
 'format percentage
 
          Columns("K").NumberFormat = "0.00%"
                   
'move to next page
    Next Page
    
End Sub
