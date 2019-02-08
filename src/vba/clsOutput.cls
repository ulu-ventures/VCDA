'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private o_description As String
Private o_units As String
Private o_cellRef As String
Private o_sheetName As String

Public Property Get Description() As String
    Description = o_description
End Property
Public Property Let Description(ByVal strDescription As String)
    o_description = strDescription
End Property

Public Property Get Units() As String
    Units = o_units
End Property
Public Property Let Units(ByVal strUnits As String)
    o_units = strUnits
End Property

Public Property Get CellRef() As String
    CellRef = o_cellRef
End Property
Public Property Let CellRef(ByVal strCellRef As String)
    o_cellRef = strCellRef
End Property

Public Property Let sheetName(ByVal strSheetName As String)
    o_sheetName = strSheetName
End Property
Public Property Get sheetName() As String
    sheetName = o_sheetName
End Property

Public Property Get CurrentValue() As Double
    CurrentValue = Sheets(o_sheetName).Range(o_cellRef)
End Property
