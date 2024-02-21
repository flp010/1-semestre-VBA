Attribute VB_Name = "Módulo1"
Option Explicit
Sub ex_02()
'DECLARAÇÃO DE VARIAVEIS
Dim horas As Double, salario As Double
Dim vhora As Double, sbruto As Double
Dim imposto As Double, sliquido As Double

'LEITURA DE DADOS
horas = ActiveCell.Offset(0, -2).Value
salario = ActiveCell.Offset(0, -1).Value
vhora = ActiveCell.Offset.Value
sbruto = ActiveCell.Offset(0, 1).Value
imposto = ActiveCell.Offset(0, 2).Value
sliquido = ActiveCell.Offset(0, 3).Value

'CALCULOS E OPERAÇÕES
vhora = (salario / horas) / 2
sbruto = horas * vhora
imposto = (sbruto * 3) / 100
sliquido = sbruto - imposto

'SAIDA DE DADOS
ActiveCell.Value = vhora
ActiveCell.Offset(0, 1).Value = sbruto
ActiveCell.Offset(0, 2).Value = imposto
ActiveCell.Offset(0, 3).Value = sliquido
ActiveCell.Offset(1, 0).Select

End Sub

Sub ex_02_02()

'DECLARAÇÃO DE VARIAVEIS
Dim horas As Double, salario As Double, vhora As Double, sbruto As Double, imposto As Double, sliquido As Double
Dim L As Integer
L = Range("C50").End(xlUp).Row + 1

'LEITURA DE DADOS
horas = Cells(L, 1).Value
salario = Cells(L, 2).Value
vhora = Cells(L, 3).Value
sbruto = Cells(L, 4).Value
imposto = Cells(L, 5).Value
sliquido = Cells(L, 6).Value

'CALCULOS E OPERAÇÕES
vhora = (salario / horas) / 2
sbruto = horas * vhora
imposto = (sbruto * 3) / 100
sliquido = sbruto - imposto

'SAIDA DE DADOS
Cells(L, 3).Value = vhora
Cells(L, 4).Value = sbruto
Cells(L, 5).Value = imposto
Cells(L, 6).Value = sliquido

End Sub

