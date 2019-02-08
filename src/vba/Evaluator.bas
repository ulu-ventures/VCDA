'''THE MIT LICENSE
'''Copyright 2018 Somik Raha, Clint Korver, Ulu Ventures
'''Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'''The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Attribute VB_Name = "Evaluator"
''' Tornado Page wiring
Sub UpdateTornado()
Application.ScreenUpdating = False
    Dim outputTableName As String
    outputTableName = "Tornado_Output" & Sheets(ActiveSheet.Name).Range("Tornado_TableIndex")
    Set outputTable = Sheets(ActiveSheet.Name).Range(outputTableName)
    
    Dim numRows As Integer
    numRows = outputTable.Cells(1000, 1).End(xlUp).Row - outputTable.Cells(1, 1).Row + 1
    Sheets(ActiveSheet.Name).ChartObjects("TornadoChart").Activate
    ActiveChart.Axes(xlValue).CrossesAt = outputTable.Cells(5, 7)
    ActiveChart.ChartTitle.Text = "Tornado for " & Sheets(ActiveSheet.Name).Range("Tornado_SelectedOutputName")
    ActiveChart.PlotArea.Select
    lowRange = Range(outputTable.Cells(4, 6), outputTable.Cells(numRows, 6))
    highRange = Range(outputTable.Cells(4, 8), outputTable.Cells(numRows, 8))
    xRange = Range(outputTable.Cells(4, 1), outputTable.Cells(numRows, 1))
    ActiveChart.SeriesCollection(1).Values = lowRange
    ActiveChart.SeriesCollection(2).Values = highRange
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).XValues = xRange
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = outputTable.Cells(4, 2)
    For i = 1 To numRows - 3
        ActiveChart.SeriesCollection(1).DataLabels.Select
        ActiveChart.SeriesCollection(1).Points(i).DataLabel.Select
        ActiveChart.SeriesCollection(1).Points(i).DataLabel.Text = outputTable.Cells(3 + i, 3)
        ActiveChart.SeriesCollection(1).Points(i).DataLabel.HorizontalAlignment = xlLeft
        ActiveChart.SeriesCollection(2).DataLabels.Select
        ActiveChart.SeriesCollection(2).Points(i).DataLabel.Select
        ActiveChart.SeriesCollection(2).Points(i).DataLabel.Text = outputTable.Cells(3 + i, 5)
        ActiveChart.SeriesCollection(1).Points(i).DataLabel.HorizontalAlignment = xlRight
    Next i
    Sheets(ActiveSheet.Name).ChartObjects("CDF").Activate
    ActiveChart.ChartTitle.Text = "CDF for " & Sheets(ActiveSheet.Name).Range("Tornado_SelectedOutputName")
    ActiveChart.Axes(xlCategory).AxisTitle.Text = outputTable.Cells(4, 2)
    numRows = outputTable.Cells(1000, 11).End(xlUp).Row - outputTable.Cells(1, 1).Row + 1
    ActiveChart.PlotArea.Select
    xRange = Range(outputTable.Cells(4, 13), outputTable.Cells(numRows, 13))
    cdfRange = Range(outputTable.Cells(4, 12), outputTable.Cells(numRows, 12))
    ActiveChart.SeriesCollection(1).Values = cdfRange
    ActiveChart.SeriesCollection(1).XValues = xRange
    Application.ScreenUpdating = True
End Sub

''' Main Macro to be wired onto Evaluate Button
Sub EvaluateModel()
    Application.ScreenUpdating = False
    Set tornadoMaker = New clsTornadoMaker
    Dim outputs() As clsOutput
    outputs = outputsOfInterest(ActiveSheet.Name)
    tornadoResults = tornadoMaker.makeTornadoUsing(outputs, inputsOfInterest(ActiveSheet.Name))
    Set summaryMaker = New clsSummaryMaker
    
    Dim summaryResults() As clsJointDistribution
    summaryResults = summaryMaker.makeSummaryUsing(tornadoResults)
    showSummaryResults ActiveSheet.Name, summaryResults
    updateTornadoPage "Tornado", outputs, summaryResults, tornadoResults
    Application.ScreenUpdating = True
End Sub
Sub updateTornadoPage(sheetName As String, outputs() As clsOutput, summaryResults() As clsJointDistribution, tornadoResults As Variant)
    Set outputsList = Sheets(sheetName).Range("Tornado_OutputsList")
    Dim outputIndex As Integer
    Dim dataTableRangeName As String
    Dim output As clsOutput
    Dim summaryResult As clsJointDistribution
    Dim jointIndex As Integer
    For outputIndex = 1 To UBound(outputs)
        Set output = outputs(outputIndex)
        outputsList.Cells(1 + outputIndex, 1) = output.Description
        dataTableRangeName = sheetName & "_Output" & Trim(Str(outputIndex))
        Set dataTableRange = Sheets(sheetName).Range(dataTableRangeName)
        Dim lastRow As Long
        lastRow = dataTableRange.Cells(4, 1).End(xlDown).Row
        Set clearingRange = Range(dataTableRange.Cells(4, 1), dataTableRange.Cells(lastRow, 10))
        clearingRange.ClearContents
        dataTableRange.Cells(1, 1) = "Output: " & output.Description
        dataTableRange.Cells(4, 1) = "Combined Unc"
        dataTableRange.Cells(4, 2) = output.Units
        Set summaryResult = summaryResults(outputIndex)
        dataTableRange.Cells(4, 6) = summaryResult.Ten
        dataTableRange.Cells(4, 7) = summaryResult.Fifty
        dataTableRange.Cells(4, 8) = summaryResult.Ninety
        For jointIndex = 1 To UBound(summaryResult.joints)
            Dim joints() As clsJoint
            joints = summaryResult.joints
            Dim joint As clsJoint
            Set joint = joints(jointIndex)
            dataTableRange.Cells(3 + jointIndex, 11) = joint.Probability
            dataTableRange.Cells(3 + jointIndex, 12) = joint.Cume
            dataTableRange.Cells(3 + jointIndex, 13) = joint.Value
        Next jointIndex
        Dim tornadoResult As Variant
        Dim multiDimUtility As clsMultiDimUtility
        Set multiDimUtility = New clsMultiDimUtility
        tornadoResult = multiDimUtility.getOneLine(tornadoResults, outputIndex)
        For inputIndex = 1 To UBound(tornadoResult)
            Dim inputDefn As clsInput
            Set inputDefn = tornadoResult(inputIndex).inputDefn
            dataTableRange.Cells(4 + inputIndex, 1) = inputDefn.Description
            dataTableRange.Cells(4 + inputIndex, 2) = inputDefn.Units
            dataTableRange.Cells(4 + inputIndex, 3) = inputDefn.Low
            dataTableRange.Cells(4 + inputIndex, 4) = inputDefn.Base
            dataTableRange.Cells(4 + inputIndex, 5) = inputDefn.High
            dataTableRange.Cells(4 + inputIndex, 6) = tornadoResult(inputIndex).Low
            dataTableRange.Cells(4 + inputIndex, 7) = tornadoResult(inputIndex).Base
            dataTableRange.Cells(4 + inputIndex, 8) = tornadoResult(inputIndex).High
            dataTableRange.Cells(4 + inputIndex, 9) = tornadoResult(inputIndex).Swing
            dataTableRange.Cells(4 + inputIndex, 10) = tornadoResult(inputIndex).SwingSquare
        Next inputIndex
    Next outputIndex
End Sub

Sub showSummaryResults(sheetName As String, summaryResults() As clsJointDistribution)
    Set outputsTable = Sheets(sheetName).Range("outputDefnTable")
    Dim outputIndex As Integer
    For outputIndex = 1 To UBound(summaryResults)
        outputsTable.Cells(outputIndex, 6) = summaryResults(outputIndex).Mean
        outputsTable.Cells(outputIndex, 7) = summaryResults(outputIndex).Ten
        outputsTable.Cells(outputIndex, 8) = summaryResults(outputIndex).Fifty
        outputsTable.Cells(outputIndex, 9) = summaryResults(outputIndex).Ninety
    Next outputIndex
End Sub
Function outputsOfInterest(sheetName) As clsOutput()
    Set outputsTable = Sheets(sheetName).Range("outputDefnTable")
    Dim currentRow, numRows As Integer
    Dim numOutputs As Integer
    numRows = outputsTable.Rows.Count
    numOutputs = outputsTable.Cells(1, 1).End(xlDown).Row - outputsTable.Cells(1, 1).Row + 1
    Dim outputs() As clsOutput
    ReDim outputs(1 To numOutputs)
    For i = 1 To numOutputs
        Set outputs(i) = New clsOutput
        outputs(i).Description = outputsTable.Cells(i, 1)
        outputs(i).Units = outputsTable.Cells(i, 3)
        outputs(i).CellRef = outputsTable.Cells(i, 5).Address
        outputs(i).sheetName = sheetName
    Next i
    outputsOfInterest = outputs
End Function

Function inputsOfInterest(sheetName) As clsInput()
    Set inputsTable = Sheets(sheetName).Range("inputTable")
    Dim numRows, numInputs As Integer
    numRows = inputsTable.Rows.Count
    currentRow = 1
    numInputs = 0
    Dim Inputs() As clsInput
    Dim firstTime As Boolean
    firstTime = True
    While (currentRow < numRows)
        If inputsTable.Cells(currentRow, 5) <> "" Then
            numInputs = numInputs + 1
            If firstTime Then
                ReDim Inputs(1 To 1)
                firstTime = False
            Else
                ReDim Preserve Inputs(1 To numInputs)
            End If
            Set Inputs(numInputs) = New clsInput
            Inputs(numInputs).Description = inputsTable.Cells(currentRow, 2)
            Inputs(numInputs).Units = inputsTable.Cells(currentRow, 3)
            Inputs(numInputs).IndexCellRef = inputsTable.Cells(currentRow, 4).Address
            Inputs(numInputs).Low = inputsTable.Cells(currentRow, 5)
            Inputs(numInputs).Base = inputsTable.Cells(currentRow, 6)
            Inputs(numInputs).High = inputsTable.Cells(currentRow, 7)
            Inputs(numInputs).sheetName = sheetName
        End If
        currentRow = currentRow + 1
    Wend
    inputsOfInterest = Inputs
End Function


