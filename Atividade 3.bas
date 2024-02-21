Attribute VB_Name = "Módulo1"
Option Explicit
Sub ex_02()

'DECLARÃO DE VARIAVEIS
Dim SALARIO As Double, KW As Double
Dim VKW As Double, VALOR As Double
Dim DESCONTO As Double, VFINAL As Double

'LEITURA DE DADOS
SALARIO = ActiveCell.Offset(0, -2).Value
KW = ActiveCell.Offset(0, -1).Value
VKW = ActiveCell.Value
VALOR = ActiveCell.Offset(0, 1).Value
DESCONTO = ActiveCell.Offset(0, 2).Value
VFINAL = ActiveCell.Offset(0, 3).Value

'CALCULOS E OPERAÇÕES
VKW = (SALARIO * 1) / 5
VALOR = KW * VKW
DESCONTO = (VALOR * 15) / 100
VFINAL = VALOR - DESCONTO

'SAIDA DE DADOS
ActiveCell.Value = VKW
ActiveCell.Offset(0, 1).Value = VALOR
ActiveCell.Offset(0, 2).Value = DESCONTO
ActiveCell.Offset(0, 3).Value = VFINAL
ActiveCell.Offset(1, 0).Select

End Sub

Sub ex_02_01()

'DECLARÃO DE VARIAVEIS
Dim SALARIO As Double, KW As Double
Dim VKW As Double, VALOR As Double
Dim DESCONTO As Double, VFINAL As Double
Dim L As Integer
L = Range("c50").End(xlUp).Row + 1


'LEITURA DE DADOS
SALARIO = Cells(L, 1).Value
KW = Cells(L, 2).Value
VKW = Cells(L, 3).Value
VALOR = Cells(L, 4).Value
DESCONTO = Cells(L, 5).Value
VFINAL = Cells(L, 6).Value

'CALCULOS E OPERAÇÕES
VKW = (SALARIO * 1) / 5
VALOR = KW * VKW
DESCONTO = (VALOR * 15) / 100
VFINAL = VALOR - DESCONTO

'SAIDA DE DADOS
Cells(L, 3).Value = VKW
Cells(L, 4).Value = VALOR
Cells(L, 5).Value = DESCONTO
Cells(L, 6).Value = VFINAL

End Sub




