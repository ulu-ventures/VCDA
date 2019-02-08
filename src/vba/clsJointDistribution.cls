'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJointDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private jd_joints() As clsJoint
Private jd_mean As Double
Private jd_ten As Double
Private jd_fifty As Double
Private jd_ninety As Double
Public Property Get joints() As clsJoint()
    joints = jd_joints
End Property
Public Property Get Mean() As Double
    Mean = jd_mean
End Property

Public Property Get Ten() As Double
    Ten = jd_ten
End Property
Public Property Get Fifty() As Double
    Fifty = jd_fifty
End Property
Public Property Get Ninety() As Double
    Ninety = jd_ninety
End Property

Public Sub addJoint(joint As Double, val As Double)
    On Error GoTo handler
    ReDim Preserve jd_joints(1 To UBound(jd_joints) + 1)
    GoTo add
handler:
    ReDim jd_joints(1 To 1)
    jd_mean = 0#
add:
    Set jd_joints(UBound(jd_joints)) = New clsJoint
    jd_joints(UBound(jd_joints)).Probability = joint
    jd_joints(UBound(jd_joints)).Value = val
    jd_mean = jd_mean + joint * val
End Sub

Public Sub sortCumulateAndExtract()
    Dim sorter As New clsSorter
    sorter.Quicksort jd_joints, 1, UBound(jd_joints)
    Dim jointIndex As Integer
    Dim Cume As Double
    Dim prevCume As Double
    Dim joint As clsJoint
    prevCume = 0#
    For jointIndex = 1 To UBound(jd_joints)
        prevCume = Cume
        Set joint = jd_joints(jointIndex)
        Cume = Cume + joint.Probability
        joint.Cume = Cume
        If prevCume < 0.1 And Cume >= 0.1 Then
            jd_ten = joint.Value
        ElseIf prevCume < 0.5 And Cume >= 0.5 Then
            jd_fifty = joint.Value
        ElseIf prevCume < 0.9 And Cume >= 0.9 Then
            jd_ninety = joint.Value
        End If
    Next jointIndex
End Sub

