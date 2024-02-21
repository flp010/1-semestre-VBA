Attribute VB_Name = "Módulo1"
Option Explicit

Sub ex_5()
'Declaração de variaveis
Dim vraio As Double, comp As Double, area As Double, vol As Double
Dim pi As Double


'Entrada de dados
vraio = ActiveCell.Offset(0, -1).Value
comp = ActiveCell.Offset(0, 0).Value
area = ActiveCell.Offset(0, 1).Value
vol = ActiveCell.Offset(0, 2).Value
pi = 3.14

'Calculos e operações
comp = 2 * pi * vraio
area = 4 * pi * vraio ^ 2
vol = (4 * pi * vraio ^ 3) / 3

'Saida de dados
ActiveCell.Offset(0, 0) = comp
ActiveCell.Offset(0, 1) = area
ActiveCell.Offset(0, 2) = vol
ActiveCell.Offset(1, 0).Select

End Sub

Sub ex_05_01()

'Declaração de variaveis
Dim vraio As Double, comp As Double, area As Double, vol As Double
Dim pi As Double, L As Integer
L = Range("b50").End(xlUp).Row + 1

'Entrada de dados
vraio = Cells(L, 1).Value
comp = Cells(L, 2).Value
area = Cells(L, 3).Value
vol = Cells(L, 4).Value
pi = 3.14

'Calculos e operações
comp = 2 * pi * vraio
area = 4 * pi * vraio ^ 2
vol = (4 * pi * vraio ^ 3) / 3

'Saida de dados
Cells(L, 2).Value = comp
Cells(L, 3).Value = area
Cells(L, 4).Value = vol

End Sub
