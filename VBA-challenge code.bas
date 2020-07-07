Attribute VB_Name = "Module1"
Sub VBA_Challenge()

For Each ws In Worksheets

' Label Columns

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
' Adjust all columns

    Cells.Columns.AutoFit

' Variables List - Volume/Ticker

    Dim Total_Stock_Volume As Single
    Dim Summary_Table_Row As Double
    Dim i As Long
    
 ' Values List - Volume/Ticker
 
      Total_Stock_Volume = 0
      Summary_Table_Row = 2
      
' Volume Variables/Parameters
    
    Dim Ticker_Range As Range
    Dim Volume_Range As Range
    Set Ticker_Range = ws.Range("A:A")
    Set Volume_Range = ws.Range("G:G")

      
' Variables List - Yearly Change

    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Yearly_Open As Double
    Dim Yearly_Close As Double
    Dim Counter As Double
    
' Values List - Yearly Change

    Yearly_Change = 0
    Percent_Change = 0
    Yearly_Open = 0
    Yearly_Close = 0
    Counter = 0

' Yearly Change, Percent Change, and Volume - SumIf taken from: https://docs.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.sumif

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ws.Cells(Summary_Table_Row, 12).Value = WorksheetFunction.SumIfs(Volume_Range, Ticker_Range, ws.Cells(Summary_Table_Row, 9).Value)
        
             ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i + 1, 1).Value
             
             Yearly_Open = ws.Cells(i + 1, 3).Value
             
             Counter = WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(Summary_Table_Row, 9).Value)
             
             Yearly_Close = ws.Cells(i + Counter, 6).Value
             
             ws.Cells(Summary_Table_Row, 10).Value = Yearly_Close - Yearly_Open
             
                If Yearly_Open = 0 Then
                    ws.Cells(Summary_Table_Row, 11).Value = 0
                
                Else
                    ws.Cells(Summary_Table_Row, 11).Value = (Yearly_Close - Yearly_Open) / Yearly_Open
                
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
            
            End If
            
        Next i
        
        
' I for the life of me can't figure out why the code is skipping the "A" ticker....
        


' Formatting
    
    ws.Range("J:J").NumberFormat = "0.00"
    ws.Range("K:K").NumberFormat = "0.00%"
    
' Color Formatting

    YCR = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To YCR
    
        If ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i

    
    Next ws

End Sub

