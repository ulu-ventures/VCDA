'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private i_description As String
Private i_units As String
Private i_indexCellRef As String
Private i_sheetName As String
Private i_low As Double
Private i_base As Double
Private i_high As Double

Public Property Get Description() As String
    Description = i_description
End Property
Public Property Let Description(ByVal strDescription As String)
    i_description = strDescription
End Property
Public Property Get Units() As String
    Units = i_units
End Property
Public Property Let Units(ByVal strUnits As String)
    i_units = strUnits
End Property

Public Property Let sheetName(ByVal strSheetName As String)
    i_sheetName = strSheetName
End Property

Public Property Let High(ByVal dblHigh As Double)
    i_high = dblHigh
End Property

Public Property Get High() As Double
    High = i_high
End Property

Public Property Let Base(ByVal dblBase As Double)
    i_base = dblBase
End Property

Public Property Get Base() As Double
    Base = i_base
End Property

Public Property Let Low(ByVal dblLow As Double)
    i_low = dblLow
End Property

Public Property Get Low() As Double
    Low = i_low
End Property


Public Property Get IndexCellRef() As String
    IndexCellRef = i_indexCellRef
End Property
Public Property Let IndexCellRef(ByVal strIndexCellRef As String)
    i_indexCellRef = strIndexCellRef
End Property

Public Function LowBaseHighFor(outputs() As clsOutput) As clsOutputRange()
    Dim outputRanges() As clsOutputRange
    ReDim outputRanges(1 To UBound(outputs))
    Dim outputIndex, inputIndex As Integer
    Dim val As Double
    For outputIndex = 1 To UBound(outputs)
        Set outputRanges(outputIndex) = New clsOutputRange
        outputRanges(outputIndex).OutputDefn = outputs(outputIndex)
        outputRanges(outputIndex).inputDefn = Me
    Next outputIndex
    For inputIndex = 1 To 3
        Sheets(i_sheetName).Range(i_indexCellRef) = inputIndex
        For outputIndex = 1 To UBound(outputs)
            val = Sheets(outputs(outputIndex).sheetName).Range(outputs(outputIndex).CellRef)
            Call outputRanges(outputIndex).setOutput(inputIndex, val)
        Next outputIndex
    Next inputIndex
    Sheets(i_sheetName).Range(i_indexCellRef) = 2
    LowBaseHighFor = outputRanges
End Function

Public Function OutputForInputSetTo(lowBaseOrHigh As Integer, output As clsOutput) As Double
    Dim val As Double
    Sheets(i_sheetName).Range(i_indexCellRef) = lowBaseOrHigh
    val = Sheets(output.sheetName).Range(output.CellRef)
    OutputForInputSetTo = val
End Function
Public Sub SetInputTo(lowBaseOrHigh As Integer)
    Sheets(i_sheetName).Range(i_indexCellRef) = lowBaseOrHigh
End Sub

