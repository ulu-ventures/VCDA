'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutputRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private or_outputDefn As clsOutput
Private or_outputSwing(1 To 3) As Double
Private or_inputDefn As clsInput

Public Property Get inputDefn() As clsInput
    Set inputDefn = or_inputDefn
End Property
Public Property Let inputDefn(ByVal clsInputDefn As clsInput)
    Set or_inputDefn = clsInputDefn
End Property
Public Property Get OutputDefn() As clsOutput
    Set OutputDefn = or_outputDefn
End Property
Public Property Let OutputDefn(ByVal clsOutputDefn As clsOutput)
    Set or_outputDefn = clsOutputDefn
End Property

Public Property Let Low(val As Double)
    or_outputSwing(1) = val
End Property
Public Property Get Low() As Double
    Low = or_outputSwing(1)
End Property
Public Property Let Base(val As Double)
    or_outputSwing(2) = val
End Property
Public Property Get Base() As Double
    Base = or_outputSwing(2)
End Property
Public Property Let High(val As Double)
    or_outputSwing(3) = val
End Property
Public Property Get High() As Double
    High = or_outputSwing(3)
End Property
Public Sub setOutput(index As Integer, val As Double)
    or_outputSwing(index) = val
End Sub
Public Property Get Swing() As Double
    Swing = or_outputSwing(3) - or_outputSwing(1)
End Property

Public Property Get SwingSquare() As Double
    SwingSquare = Swing * Swing
End Property

Public Property Get KeyValue() As Double
    KeyValue = SwingSquare
End Property
