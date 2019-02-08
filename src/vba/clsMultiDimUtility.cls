'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultiDimUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'''https://stackoverflow.com/questions/14173622/return-an-entire-row-of-a-multidimensional-array-in-vba-to-a-one-dimensional-arr
Public Function getOneLine(array2D As Variant, lineIndex As Integer) As Variant
' returning one column or row of a 2D array
' array2D: 2 dimensional Array
' lineIndex: the index of column or row, starting at 0
    Dim i, n As Integer
    Dim oneLine As Variant
    n = UBound(array2D, 2)
    ReDim oneLine(1 To n)

        For i = 1 To n
            Set oneLine(i) = array2D(lineIndex, i)
        Next

    getOneLine = oneLine

End Function

Public Sub setOneLine(ByRef array2D As Variant, lineIndex As Integer, newLine As Variant)
    Dim i As Integer
    For i = 1 To UBound(newLine)
        Set array2D(lineIndex, i) = newLine(i)
    Next i
End Sub
