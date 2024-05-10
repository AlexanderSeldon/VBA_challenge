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
    
    ' Headers
    analysisSheet.Cells(1, 1) = "Stock"
    analysisSheet.Cells(1, 2) = "Annual Change"
    analysisSheet.Cells(1, 3) = "Change in %"
    analysisSheet.Cells(1, 4) = "Total Traded Volume"
    analysisSheet.Cells(1, 5) = "Greatest % Increase"
    analysisSheet.Cells(1, 6) = "Greatest % Decrease"
    analysisSheet.Cells(1, 7) = "Greatest Total Volume"
    
    Dim totalEntries As Long
    totalEntries = yearSheet.Cells(yearSheet.Rows.Count, 1).End(xlUp).Row
    
    Dim stock As String
    Dim annualChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim stockMaxInc As String
    Dim stockMaxDec As String
    Dim stockMaxVol As String
    
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
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
            
            ' Update maximums
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                stockMaxInc = stock
            End If
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                stockMaxDec = stock
            End If
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                stockMaxVol = stock
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

    ' Update the greatest values on the top of the sheet
    analysisSheet.Cells(2, 5) = stockMaxInc & " (" & Round(maxIncrease, 2) & "%)"
    analysisSheet.Cells(2, 6) = stockMaxDec & " (" & Round(maxDecrease, 2) & "%)"
    analysisSheet.Cells(2, 7) = stockMaxVol & " (" & maxVolume & ")"

    ' Apply conditional formatting
    Call ApplyConditionalFormatting(analysisSheet.Range("B2:B" & resultsRow - 1))
    Call ApplyConditionalFormatting(analysisSheet.Range("C2:C" & resultsRow - 1))
    
    analysisSheet.Columns("A:G").AutoFit
End Sub

Sub ApplyConditionalFormatting(rng As Range)
    With rng
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = vbGreen
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = vbRed
    End With
End Sub
