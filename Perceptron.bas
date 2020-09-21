VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Perceptron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public biasWeight As Double
Public inputA As Variant
Public inputWeightA As Variant

Const BIAS = 1
Const INPUT_COUNT = 3
Const LEARNING_RATE = 0.05
Const TEST_ARRAY_LENGHT = 8

Public Sub Init()
    Dim i As Integer
    Randomize (Rnd)
    biasWeight = Rnd
    ReDim inputA(1 To INPUT_COUNT)
    ReDim inputWeightA(1 To INPUT_COUNT)
    For i = 1 To INPUT_COUNT
        Randomize (Rnd)
        inputWeightA(i) = Rnd
    Next i
    Call Train
End Sub

Public Sub Train()
    Dim data As Variant
    Dim inputA(1 To INPUT_COUNT) As Double
    Dim answer As Double
    Dim calcValue As Double
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    data = GetLearnData()
    For k = 1 To 1000
        For i = 1 To UBound(data, 1)
            For j = 1 To INPUT_COUNT
                inputA(j) = CDbl(data(i, j))
            Next j
            answer = CDbl(data(i, INPUT_COUNT + 1))
            calcValue = Calc(inputA)
            If answer <> calcValue Then
                For j = 1 To INPUT_COUNT
                    inputWeightA(j) = DeltaWeight(inputWeightA(j), answer - calcValue, inputA(j))
                Next j
                biasWeight = DeltaWeight(biasWeight, answer - calcValue, BIAS)
            End If
        Next i
    Next k
End Sub

Public Function Calc(inputA As Variant) As Double
    Dim i As Integer
    Dim result As Double
    For i = 1 To INPUT_COUNT
        result = result + (inputA(i) * inputWeightA(i))
    Next i
    Calc = Logistic(result + (BIAS * biasWeight))
End Function

Public Function Test()
    Dim data As Variant
    Dim inputA(1 To INPUT_COUNT) As Double
    Dim answer As Double
    Dim i As Integer
    Dim j As Integer
    Dim calcValues(1 To TEST_ARRAY_LENGHT, 1 To 1) As Double
    Dim errValues(1 To TEST_ARRAY_LENGHT, 1 To 1) As Double
    data = GetTestData()
    For i = 1 To UBound(data, 1)
        For j = 1 To INPUT_COUNT
            inputA(j) = data(i, j)
        Next j
        answer = data(i, INPUT_COUNT + 1)
        calcValues(i, 1) = Calc(inputA)
        errValues(i, 1) = answer - calcValues(i, 1)
    Next i
    Call TestToTable(calcValues, errValues)
End Function

Private Function GetLearnData()
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet
    GetLearnData = ws.Range(ws.Cells(5, 1), ws.Cells(12, 4))
End Function

Private Function GetTestData()
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet
    GetTestData = ws.Range(ws.Cells(5, 6), ws.Cells(12, 9))
End Function

Private Function Logistic(net As Double) As Double
    Logistic = 1 / (1 + Exp(-1 * net))
End Function

Private Function DeltaWeight(ByVal weight As Double, delta As Double, x As Double) As Double
    DeltaWeight = weight + (LEARNING_RATE * delta * x)
End Function

Private Sub TestToTable(calcValues As Variant, errValues As Variant)
    Dim ws As Worksheet
    Dim r As Range
    
    Set ws = Application.ActiveSheet
    Set r = ws.Range(ws.Cells(5, 10), ws.Cells(5 + TEST_ARRAY_LENGHT - 1, 10))
    r = calcValues
    Set r = ws.Range(ws.Cells(5, 11), ws.Cells(5 + TEST_ARRAY_LENGHT - 1, 11))
    r = errValues
End Sub
