'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSummaryMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const TWENTY_FIVE As Double = 0.25
Private Const Fifty As Double = 0.5
Public Function makeSummaryUsing(tornadoResults As Variant)
    Dim outputIndex As Integer
    Dim multiDimUtility As clsMultiDimUtility
    Set multiDimUtility = New clsMultiDimUtility
    Dim summaries() As clsJointDistribution
    ReDim summaries(1 To UBound(tornadoResults))
    Dim eachOutput As Variant
    For outputIndex = 1 To UBound(tornadoResults)
        eachOutput = multiDimUtility.getOneLine(tornadoResults, outputIndex)
        Dim jointDistribution As clsJointDistribution
        Set jointDistribution = New clsJointDistribution
        summarize eachOutput, 5, 5, 1#, jointDistribution
        jointDistribution.sortCumulateAndExtract
        Set summaries(outputIndex) = jointDistribution
        cleanUp eachOutput, 5
    Next outputIndex
    makeSummaryUsing = summaries
End Function
Private Sub cleanUp(ByRef eachOutput As Variant, r_depth As Integer)
    Dim i As Integer
    For i = 1 To r_depth
        inputIndex = UBound(eachOutput) - (r_depth - i)
        Set outputRange = eachOutput(inputIndex)
        Set inputDefn = outputRange.inputDefn
        inputDefn.SetInputTo 2
    Next i
End Sub
Private Sub summarize(ByRef eachOutput As Variant, r_depth As Integer, currentDepth As Integer, joint As Double, ByRef jointDistribution As clsJointDistribution)
    Dim inputIndex As Integer
    Dim outputRange As clsOutputRange
    Dim inputDefn As clsInput
    Dim val As Double
    If currentDepth = 0 Then
        inputIndex = UBound(eachOutput) - r_depth + 1
        Set outputRange = eachOutput(inputIndex)
        jointDistribution.addJoint joint, outputRange.OutputDefn.CurrentValue
    Else
        inputIndex = UBound(eachOutput) - (r_depth - currentDepth)
        Set outputRange = eachOutput(inputIndex)
        Set inputDefn = outputRange.inputDefn
            
        inputDefn.SetInputTo 1
        Dim newJoint As Double
        newJoint = joint * TWENTY_FIVE
        Dim newDepth As Integer
        newDepth = currentDepth - 1
        summarize eachOutput, r_depth, newDepth, newJoint, jointDistribution
    
        inputDefn.SetInputTo 2
        newJoint = joint * Fifty
        summarize eachOutput, r_depth, newDepth, newJoint, jointDistribution
        inputDefn.SetInputTo 3
        newJoint = joint * TWENTY_FIVE
        summarize eachOutput, r_depth, newDepth, newJoint, jointDistribution
    End If
End Sub
