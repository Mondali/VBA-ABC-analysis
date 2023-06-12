Sub RunApplicationPW()

    Dim relativePathPW As String
    relativePathPW = ThisWorkbook.path
    
    Dim startPW As Worksheet
    Dim reportPW As Worksheet
    Set startPW = ThisWorkbook.Sheets("Start")
    Set reportPW = ThisWorkbook.Sheets("Raport")

    Dim zestawienieFilePW As String
    Dim rejestrFilePW As String
    zestawienieFilePW = relativePathPW & "\" & startPW.Range("B4").value
    rejestrFilePW = relativePathPW & "\" & startPW.Range("B3").value

    Dim faultCostPW As Object
    Set faultCostPW = LoadFaultCost(zestawienieFilePW)
    
    Dim faultOccurancesPW As Object
    Set faultOccurancesPW = CountOccurences(rejestrFilePW, 1)
    
    Dim faultsTotalCostPW As Object
    Set faultsTotalCostPW = CountFaultTotalCost(faultCostPW, faultOccurancesPW)
    
    Dim sortedMapPW As Object
    Set sortedMapPW = bubblesortMap(faultsTotalCostPW)
    
    GenerateRaport reportPW, sortedMapPW
    
    GenerateChart reportPW
    
    MsgBox "Raport wygenerowany prawidlowo"
    
End Sub

Function CountOccurences(filePath As String, columnNumber As Integer) As Object
    
    Dim mapPW As Object
    Set mapPW = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    Open filePath For Input As #1
    
    Dim line As String
    Dim data() As String
    
    Line Input #1, line
    
    While Not EOF(1)
        Line Input #1, line
        data = Split(line, ";")
        
        If mapPW.Exists(data(columnNumber - 1)) Then
            mapPW(data(columnNumber - 1)) = mapPW(data(columnNumber - 1)) + 1
        Else
            mapPW.Add data(columnNumber - 1), 1
        End If
    Wend
    
    Close #1
    
    Set CountOccurences = mapPW
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error. Check file or file path " & Err.Description
    Set CountOccurences = Nothing
End Function
Function LoadFaultCost(filePath As String) As Object

    Dim mapPW As Object
    Set mapPW = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    Open filePath For Input As #1
    
    Dim line As String
    Dim data() As String
    
    Line Input #1, line
    
    While Not EOF(1)
        Line Input #1, line
        data = Split(line, ";")
        
        If Not mapPW.Exists(data(0)) Then
            mapPW.Add data(0), data(1)
        End If
        
    Wend
    
    Close #1
    
    Set LoadFaultCost = mapPW
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error. Check file or file path " & Err.Description
    Set LoadFaultCost = Nothing
End Function

Function CountFaultTotalCost(faultCosts As Object, faultOccurences As Object) As Object

    Dim totalCostPW As Object
    Set totalCostPW = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    
    For Each key In faultOccurences.keys
        If faultCosts.Exists(key) Then
            totalCostPW.Add key, faultCosts(key) * faultOccurences(key)
        End If
    Next key
    
    Set CountFaultTotalCost = totalCostPW

End Function

Function bubblesortMap(map As Object) As Object

    Dim keys() As Variant
    Dim values() As Variant
    
    keys = map.keys
    values = map.Items
    
    Dim i As Long, j As Long
    Dim temporary As Variant
    
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(i) < values(j) Then
 
                temporary = values(i)
                values(i) = values(j)
                values(j) = temporary

                temporary = keys(i)
                keys(i) = keys(j)
                keys(j) = temporary
            End If
        Next j
    Next i
    
    Dim sortedMapPW As Object
    Set sortedMapPW = CreateObject("Scripting.Dictionary")
    
    For i = LBound(keys) To UBound(keys)
        sortedMapPW.Add keys(i), values(i)
    Next i

    Set bubblesortMap = sortedMapPW

End Function

Sub GenerateRaport(reportPW As Worksheet, map As Object)

    Dim key As Variant
    Dim lastRowPW As Long
    lastRowPW = reportPW.Cells(Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    i = 1

    If lastRowPW > 2 Then
        reportPW.Range("A3:D" & lastRowPW).Clear
    End If
    
    Dim cumulatedCostPW As Double
    cumulatedCostPW = 0

    For Each key In map.keys
        reportPW.Cells(i + 2, "A").value = key
        reportPW.Cells(i + 2, "B").value = map(key)
        cumulatedCostPW = cumulatedCostPW + map(key)
        reportPW.Cells(i + 2, "C").value = cumulatedCostPW
        i = i + 1
    Next key
    
    Dim borderA As Double
    Dim borderB As Double
    borderA = 0.8 * cumulatedCostPW
    borderB = 0.95 * cumulatedCostPW

    For i = 3 To reportPW.Cells(Rows.Count, "C").End(xlUp).Row
        If reportPW.Cells(i, "C").value <= borderA Then
            reportPW.Cells(i, "D").value = "A"
        ElseIf reportPW.Cells(i, "C").value <= borderB Then
            reportPW.Cells(i, "D").value = "B"
        Else
            reportPW.Cells(i, "D").value = "C"
        End If
    Next i

End Sub
Sub GenerateChart(reportPWSheet As Worksheet)

    Dim chartPW As ChartObject

    If reportPWSheet.ChartObjects.Count > 0 Then
        For Each chartPW In reportPWSheet.ChartObjects
            chartPW.Delete
        Next chartPW
    End If

    Dim lastRowPW As Long
    lastRowPW = reportPWSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim chartPWObject As ChartObject
    Set chartPWObject = reportPWSheet.ChartObjects.Add(Left:=330, Width:=700, Top:=10, Height:=400)
    
    With chartPWObject.chart
        .HasTitle = True
        .ChartTitle.Text = "Pareto-Lorenz chart"
        .HasLegend = False
    End With
    
    Dim cumulativeCostSeriesPW As Series
    Set cumulativeCostSeriesPW = chartPWObject.chart.SeriesCollection.NewSeries
    With cumulativeCostSeriesPW
        .values = reportPWSheet.Range("C3:C" & lastRowPW)
        .XValues = reportPWSheet.Range("A3:A" & lastRowPW)
        .ChartType = xlLineMarkers
        .AxisGroup = 1
    End With

    Dim totalCostSeriesPW As Series
    Set totalCostSeriesPW = chartPWObject.chart.SeriesCollection.NewSeries
    With totalCostSeriesPW
        .values = reportPWSheet.Range("B3:B" & lastRowPW)
        .XValues = reportPWSheet.Range("A3:A" & lastRowPW)
        .ChartType = xlColumnClustered
        .AxisGroup = 2
    End With

    chartPWObject.chart.Axes(xlValue, xlSecondary).Delete
    
End Sub
