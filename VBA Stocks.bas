Attribute VB_Name = "Módulo1"
Sub Stock()
Dim WS As Worksheet
Dim lRow1 As Long
Dim Max As Double
Dim Row1 As Double
Dim Ticker As String
Dim TotalStocks As Double
Dim Row As Integer
Dim lRow As Long
Dim YearlyChg As Double
Dim OpenValue As Double
Dim CloseValue As Double

For Each WS In ActiveWorkbook.Worksheets        'Recorre todas las hojas del libro
WS.Activate



lRow = Cells(Rows.Count, 1).End(xlUp).Row           'Busca la ultima celda con datos
Total = 0
Row = 2
OpenValue = Cells(2, 3).Value                       'Inicia el valor de OpenValue en celda (2,3)
Range("I1").Value = "Ticker"                        'Inserta el nombre de las celdas I1, J1, K1, L1
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"



For i = 2 To lRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'revisa que las celdas de la columna A sean diferentes empezando en renglon 2
Ticker = Cells(i, 1).Value                         'Establece el valor de Ticker de celda (2,1)
TotalStocks = TotalStocks + Cells(i, 7).Value      ' Inicia la suma de TotalStocks con celda (1,7)
Cells(Row, 9).Value = Ticker                       'Imprime el valor de Ticker en celda (2,9)
Cells(Row, 12).Value = TotalStocks                  ' Imprime el valor de TotalStocks en celda (2,12)
CloseValue = Cells(i, 6)                       ' Establece valor de CloseValue de la ultima celda
TotalStocks = 0                                     ' Reinicia valor de TotalStocks

If OpenValue = 0 Or CloseValue = 0 Then         'Revisa que los valores no sean igual a 0 para evitar division entre 0
Cells(Row, 11).Value = 0
Cells(Row, 10).Value = 0
Else
Cells(Row, 10).Value = CloseValue - OpenValue       'Imprime la resta de CloseValue - OpenValue
If Cells(Row, 10).Value > 0 Then                    'Da formato si el valor de la resta es mayor a 0 o negativo
Cells(Row, 10).Interior.ColorIndex = 4
Else
Cells(Row, 10).Interior.ColorIndex = 3
End If
Cells(Row, 11).Value = Cells(Row, 10).Value / OpenValue 'Imprime el porcentaje de Percent Change
Cells(Row, 11).NumberFormat = "0.00%"               'Da formato de porcentaje a la celda Percent Change
End If
Row = Row + 1                                       'Aumenta el valor de Row para imprimir tabla resumen
OpenValue = Cells(i + 1, 3).Value                   'Establece el valor de Opnevalue al del siguiente Ticker

Else
TotalStocks = Cells(i, 7).Value + TotalStocks       'Realiza la suma de TotalStocks siempre que las celdas de la columna A sean iguales
End If
Next i



Range("N2").Value = "Greatest % Increase"       'Pone el nombre de las celdas requerido
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

Max = 0

lRow1 = Cells(Rows.Count, 12).End(xlUp).Row
For i = 2 To lRow1
If Max < Cells(i, 12) Then
Ticker1 = Cells(i, 9)
Max = Cells(i, 12)
Row1 = i
End If

Next i
Range("P4").Value = Max
Range("O4").Value = Ticker1

Max = 0

lRow1 = Cells(Rows.Count, 11).End(xlUp).Row
For i = 2 To lRow1
If Max < Cells(i, 11) Then
Ticker1 = Cells(i, 9)
Max = Cells(i, 11)
Row1 = i
End If
Next i
Range("P2").Value = Max
Range("P2").NumberFormat = "0.00%"
Range("O2").Value = Ticker1

Max = 0
lRow1 = Cells(Rows.Count, 11).End(xlUp).Row
For i = 2 To lRow1
If Max > Cells(i, 11) Then
Ticker1 = Cells(i, 9)
Max = Cells(i, 11)
Row1 = i
End If
Next i
Range("P3").Value = Max
Range("P3").NumberFormat = "0.00%"
Range("O3").Value = Ticker1

Next WS







End Sub


