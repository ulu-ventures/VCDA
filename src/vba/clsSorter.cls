'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''Adapted from https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
''' Assumptions
''' The caller must provide an array of objects where each object has the following property defined
''' KeyValue returns the value of the key variable on which the sort is being performed

Sub Quicksort(ByRef vArray As Variant, arrLbound As Long, arrUbound As Long)
    'Sorts a one-dimensional VBA array from smallest to largest
    'using a very fast quicksort algorithm variant.
    Dim pivotVal As Variant
    Dim vSwap    As Variant
    Dim tmpLow   As Long
    Dim tmpHi    As Long
     
    tmpLow = arrLbound
    tmpHi = arrUbound
    '''pivotVal = vArray((arrLbound + arrUbound) \ 2).SwingSquare
    pivotVal = vArray((arrLbound + arrUbound) \ 2).KeyValue
     
    While (tmpLow <= tmpHi) 'divide
       While (vArray(tmpLow).KeyValue < pivotVal And tmpLow < arrUbound)
          tmpLow = tmpLow + 1
       Wend
      
       While (pivotVal < vArray(tmpHi).KeyValue And tmpHi > arrLbound)
          tmpHi = tmpHi - 1
       Wend
     
       If (tmpLow <= tmpHi) Then
          Set vSwap = vArray(tmpLow)
          Set vArray(tmpLow) = vArray(tmpHi)
          Set vArray(tmpHi) = vSwap
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If
    Wend
 
  If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi 'conquer
  If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound 'conquer
End Sub
