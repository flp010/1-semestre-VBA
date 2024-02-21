Attribute VB_Name = "Módulo2"
Option Explicit

Sub Ex_03()
'Declaração de variaveis
Dim anonasc As Double, idanos As Double, idmeses As Double, iddias As Double, idsemanas As Double
Dim soma As Double, soma1 As Double, soma2 As Double, soma3 As Double, anoatual As Double

'Leitura de dados
anonasc = ActiveCell.Offset(0, -1).Value
idanos = ActiveCell.Offset(0, 0).Value
idmeses = ActiveCell.Offset(0, 1).Value
idsemanas = ActiveCell.Offset(0, 2).Value
iddias = ActiveCell.Offset(0, 3).Value
anoatual = ActiveCell.Offset(0, 4).Value

'Calculos e operações
soma = anoatual - anonasc
soma1 = soma * 12
soma2 = soma * 48
soma3 = soma * 365

'Saida de dados
ActiveCell.Offset(0, 0).Value = soma
ActiveCell.Offset(0, 1).Value = soma1
ActiveCell.Offset(0, 2).Value = soma2
ActiveCell.Offset(0, 3).Value = soma3
ActiveCell.Offset(1, 0).Select


End Sub

Sub ex_03_1()
'declaração de variaveis
Dim anonasc As Double, idanos As Double, idmeses As Double, iddias As Double, idsemanas As Double
Dim soma As Double, soma1 As Double, soma2 As Double, soma3 As Double, anoatual As Double
Dim L As Integer
L = Range("b100").End(xlUp).Row + 1

'Leitura de dados
anonasc = Cells(L, 1).Value
idanos = Cells(L, 2).Value
idmeses = Cells(L, 3).Value
idsemanas = Cells(L, 4).Value
iddias = Cells(L, 5).Value
anoatual = Cells(L, 6).Value

'calculos e operações
soma = anoatual - anonasc
soma1 = soma * 12
soma2 = soma * 48
soma3 = soma * 365

'Saida de dados
Cells(L, 2).Value = soma
Cells(L, 3).Value = soma1
Cells(L, 4).Value = soma2
Cells(L, 5).Value = soma3






End Sub




