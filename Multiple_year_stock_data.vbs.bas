Attribute VB_Name = "Module1"
Sub Stockmarket()
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
  
    Dim Ticker As String

    Dim Vol_Total As Double
    Vol_Total = 0

    Dim Close_Price As Double
  
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    For i = 2 To lastrow

   
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
      Close_Price = ws.Cells(i, 6).Value
      
      Ticker = ws.Cells(i, 1).Value

      Vol_Total = Vol_Total + ws.Cells(i, 7).Value
      
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ws.Range("L" & Summary_Table_Row).Value = Vol_Total
      
      ws.Range("N" & Summary_Table_Row).Value = Close_Price
     
     
      Summary_Table_Row = Summary_Table_Row + 1
      Vol_Total = 0
      
      Else
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value
     
      End If
     
      If ws.Cells(i, 2).Value = 20200102 Or ws.Cells(i, 2).Value = 20190102 Or ws.Cells(i, 2).Value = 20180102 Then
      ws.Range("M" & Summary_Table_Row).Value = ws.Cells(i, 3).Value
      
      End If

    Next i
    
    lastrow_Summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
      
      For i = 2 To lastrow_Summary
    
      ws.Cells(i, 10).Value = ws.Cells(i, 14).Value - ws.Cells(i, 13).Value
    
      If ws.Cells(i, 10).Value < 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 3
      Else
      ws.Cells(i, 10).Interior.ColorIndex = 4
      End If
    
      ws.Cells(i, 11).Value = ws.Cells(i, 10).Value / ws.Cells(i, 13).Value
    
      ws.Range("K:K").NumberFormat = "0.00%"
      ws.Range("L:L").NumberFormat = "0"
      ws.Range("S3") = WorksheetFunction.Max(ws.Range("K:K"))
      ws.Range("S4") = WorksheetFunction.Min(ws.Range("K:K"))
      ws.Range("S3:S4").NumberFormat = "0.00%"
      ws.Range("S5") = WorksheetFunction.Max(ws.Range("L:L"))
      ws.Range("S5").NumberFormat = "0"
      
      If ws.Cells(i, 11).Value = ws.Range("S3") Then
      ws.Range("R3").Value = ws.Cells(i, 9).Value
      End If
      
      If ws.Cells(i, 11).Value = ws.Range("S4") Then
      ws.Range("R4").Value = ws.Cells(i, 9).Value
      End If
      
      If ws.Cells(i, 12).Value = ws.Range("S5") Then
      ws.Range("R5").Value = ws.Cells(i, 9).Value
      End If
      
      Next i
    
    ws.Columns("M:N").Hidden = True
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(3, 17).Value = "Greatest % increase"
    ws.Cells(4, 17).Value = "Greatest % decrease"
    ws.Cells(5, 17).Value = "Greatest total volume"
    ws.Cells(2, 18).Value = "Ticker"
    ws.Cells(2, 19).Value = "Value"
    


Next ws

End Sub


