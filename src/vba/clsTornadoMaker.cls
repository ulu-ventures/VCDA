'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTornadoMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Public Function makeTornadoUsing(outputs() As clsOutput, Inputs() As clsInput)
        Dim outputRange() As clsOutputRange
        ReDim outputRange(1 To UBound(outputs), 1 To UBound(Inputs))
        Dim inputIndex, outputIndex As Integer
        For inputIndex = 1 To UBound(Inputs)
            Dim lowBaseHighAllOutputs As Variant
            lowBaseHighAllOutputs = Inputs(inputIndex).LowBaseHighFor(outputs)
            For outputIndex = 1 To UBound(outputs)
                Set outputRange(outputIndex, inputIndex) = lowBaseHighAllOutputs(outputIndex)
            Next outputIndex
        Next inputIndex
        Call sortBySwing(outputRange)
        makeTornadoUsing = outputRange
End Function

Private Sub sortBySwing(outputRange As Variant)
    Dim outputIndex As Integer
    Dim eachOutput As Variant
    Dim multiDimUtility As clsMultiDimUtility
    Set multiDimUtility = New clsMultiDimUtility
    Dim sorter As clsSorter
    Set sorter = New clsSorter
    For outputIndex = 1 To UBound(outputRange)
        eachOutput = multiDimUtility.getOneLine(outputRange, outputIndex)
        sorter.Quicksort eachOutput, 1, UBound(eachOutput)
        multiDimUtility.setOneLine outputRange, outputIndex, eachOutput
    Next outputIndex
End Sub



