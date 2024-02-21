Attribute VB_Name = "Módulo1"
Option Explicit
Sub Ex_01()
'DECLARAÇÃO DE VARIAVEIS
Dim n1 As Double, n2 As Double, n3 As Double, n4 As Double
Dim soma, med01, med02
'LEITURA DE DADOS
n1 = ActiveCell.Offset(0, -4).Value
n2 = ActiveCell.Offset(0, -3).Value
n3 = ActiveCell.Offset(0, -2).Value
n4 = ActiveCell.Offset(0, -1).Value
'CALCULOS E OPERAÇÕES
soma = n1 + n2 + n3 + n4
med01 = soma / 4
med02 = n1 * 0.2 + n2 * 0.3 + n3 * 0.3 + n4 * 0.2
'SAIDA DE DADOS
ActiveCell.Offset.Value = soma
ActiveCell.Offset(0, 1).Value = med01
ActiveCell.Offset(0, 2).Value = med02
ActiveCell.Offset(1, 0).Select



End Sub
Sub Ex_01_2()
'DECLARAÇÃO DE VARIAVEIS
Dim n1 As Double, n2 As Double, n3 As Double, n4 As Double
Dim soma, med01, med02
Dim L As Integer
L = Range("e200").End(xlUp).Row + 1
'LEITURA DE DADOS
n1 = Cells(L, 1).Value
n2 = Cells(L, 2).Value
n3 = Cells(L, 3).Value
n4 = Cells(L, 4).Value
'CALCULOS E OPERAÇÕES
soma = n1 + n2 + n3 + n4
med01 = soma / 4
med02 = n1 * 0.2 + n2 * 0.3 + n3 * 0.3 + n4 * 0.2
'SAIDA DE DADOS
Cells(L, 5) = soma
Cells(L, 6) = med01
Cells(L, 7) = med02


End Sub
