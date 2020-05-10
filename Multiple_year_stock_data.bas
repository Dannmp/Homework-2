Sub stock()

For Each ws In Worksheets

    'Declarar la variable del ticket
    Dim ticker As String

    'Declarar la variable de la varación de costo de año
    Dim yearvar As Double

    'Declarar la variable del porcentaje
    Dim percet As Double

    'Declarar total de las acciones
    Dim totalstock As Double
    totalstock = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim i As Long

    Dim start As Long

    start = 2

    'Inicia el loop

    For i = 2 To lastrow
 
        openprice = ws.Cells(start, 3).Value
 
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
        closeprice = ws.Cells(i, 6).Value
        
    'Calculo yearly change

        yearch = closeprice - openprice
    
    ws.Range("J" & Summary_Table_Row).Value = yearch
    
        start = i + 1

    'Calculo del porcentaje

    If (openprice = 0 And closeprice = 0) Then

        Percent = 0
    
    Else: Percent = (((closeprice / openprice) - 1))
    ws.Range("K" & Summary_Table_Row).Value = Percent

    End If
    

    'Calculo del total de volumen
 
        ticker = ws.Cells(i, 1).Value
           
        totalstock = totalstock + ws.Cells(i, 7).Value
    
        ws.Range("I" & Summary_Table_Row).Value = ticker
    
        ws.Range("L" & Summary_Table_Row).Value = totalstock
    
        Summary_Table_Row = Summary_Table_Row + 1
        
        totalstock = 0
    
    Else

        totalstock = totalstock + ws.Cells(i, 7).Value
    End If

    'Conditional Fomartting
    If ws.Cells(i, 10).Value > 0 Then

    ws.Cells(i, 10).Interior.Color = vbGreen
    
    ElseIf ws.Cells(i, 10).Value < 0 Then

    ws.Cells(i, 10).Interior.Color = vbRed
  
    End If
       
    Next i
    
    'Calculo de max / min percent y max vol
    ylastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
  
  
    For j = 2 To ylastrow
    If ws.Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & ylastrow)) Then
        ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(j, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
    ElseIf ws.Cells(j, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & ylastrow)) Then
        ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(j, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
    ElseIf ws.Cells(j, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & ylastrow)) Then
        ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
    End If

    Next j


Next ws

End Sub