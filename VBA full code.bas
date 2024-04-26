Attribute VB_Name = "Module1"
Sub CalculateStockMetrics()
    Dim yearSheet As Worksheet
    Set yearSheet = ThisWorkbook.Sheets("2018")
    
    Dim analysisSheet As Worksheet
    On Error Resume Next
    Set analysisSheet = ThisWorkbook.Sheets("Analysis Results")
    If analysisSheet Is Nothing Then
        Set analysisSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        analysisSheet.Name = "Analysis Results"
    Else
        analysisSheet.Cells.Clear
    End If
    On Error GoTo 0
    
    analysisSheet.Cells(1, 1) = "Stock"
    analysisSheet.Cells(1, 2) = "Annual Change"
    analysisSheet.Cells(1, 3) = "Change in %"
    analysisSheet.Cells(1, 4) = "Total Traded Volume"
    
    Dim totalEntries As Long
    totalEntries = yearSheet.Cells(yearSheet.Rows.Count, 1).End(xlUp).Row
    
    Dim stock As String
    Dim annualChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim resultsRow As Long
    resultsRow = 2
    
    totalVolume = 0
    openPrice = yearSheet.Cells(2, 3).Value
    
    Dim rowIndex As Long
    For rowIndex = 2 To totalEntries
        If yearSheet.Cells(rowIndex + 1, 1).Value <> yearSheet.Cells(rowIndex, 1).Value Then
            stock = yearSheet.Cells(rowIndex, 1).Value
            totalVolume = totalVolume + yearSheet.Cells(rowIndex, 6).Value
            closePrice = yearSheet.Cells(rowIndex, 4).Value
            
            annualChange = closePrice - openPrice
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = (annualChange / openPrice) * 100
            End If
            
            analysisSheet.Cells(resultsRow, 1) = stock
            analysisSheet.Cells(resultsRow, 2) = annualChange
            analysisSheet.Cells(resultsRow, 3) = Round(percentChange, 2) & "%"
            analysisSheet.Cells(resultsRow, 4) = totalVolume
            
            resultsRow = resultsRow + 1
            
            If rowIndex + 1 <= totalEntries Then
                openPrice = yearSheet.Cells(rowIndex + 1, 3).Value
                totalVolume = 0
            End If
        Else
            totalVolume = totalVolume + yearSheet.Cells(rowIndex, 6).Value
        End If
    Next rowIndex

    Dim changeRange As Range
    Set changeRange = analysisSheet.Range("B2:B" & resultsRow - 1)
    With changeRange
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = vbGreen
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(1).Interior.Color = vbRed
    End With

    Dim percentRange As Range
    Set percentRange = analysisSheet.Range("C2:C" & resultsRow - 1)
    With percentRange
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = vbGreen
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(1).Interior.Color = vbRed
    End With

    analysisSheet.Columns("A:D").AutoFit
End Sub

