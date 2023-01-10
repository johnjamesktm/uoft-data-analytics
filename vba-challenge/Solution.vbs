Attribute VB_Name = "Module1"

Function ColumnNumberToLetter(ws As Worksheet, columnNumber As Integer) As String
    ColumnNumberToLetter = Trim(Split(ws.Cells(1, columnNumber).Address, "$")(1))
End Function

Sub ClearAnalysisOnWorksheet(ws As Worksheet)
Dim lastColumn
lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
If lastColumn > 7 Then
    For i = 8 To lastColumn
        ws.Columns(8).EntireColumn.Delete
    Next i
End If
End Sub

Sub ClearAnalysis()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Call ClearAnalysisOnWorksheet(ws)
Next ws
End Sub

Sub ApplyAnalysisToWorksheet(ws As Worksheet)
Dim lastRow
Dim lastColumn
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

' Insert ten new columns in the end
For i = 1 To 10
    ws.Columns(lastColumn + i).EntireColumn.Insert
Next i

ws.Cells(1, lastColumn + 2).Value = "Ticker"
ws.Cells(1, lastColumn + 3).Value = "Yearly Change"
ws.Cells(1, lastColumn + 4).Value = "Percent Change"
ws.Cells(1, lastColumn + 5).Value = "Total Stock Volume"

ws.Cells(2, lastColumn + 8).Value = "Greatest % Increase"
ws.Cells(3, lastColumn + 8).Value = "Greatest % Decrease"
ws.Cells(4, lastColumn + 8).Value = "Greatest Total Volume"

ws.Cells(1, lastColumn + 9).Value = "Ticker"
ws.Cells(1, lastColumn + 10).Value = "Value"

Dim tickerRowIndex As Integer
tickerRowIndex = 2

Dim firstOpenValue
firstOpenValue = ws.Cells(2, 3).Value

Dim totalStockVolume
totalStockVolume = 0

Dim greatestIncreaseValue
Dim greatestDecreaseValue
Dim greatestTotalVolumeValue

greatestIncreaseValue = 0
greatestDecreaseValue = 0
greatestTotalVolumeValue = 0

Dim greatestIncreaseTicker
Dim greatestDecreaseTicker
Dim greatestTotalVolumeTicker

For i = 2 To lastRow
    totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
    If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
        Dim currentTicker
        currentTicker = ws.Cells(i, 1).Value
        
        Dim yearlyChange
        yearlyChange = ws.Cells(i, 6).Value - firstOpenValue
        
        Dim percentChange
        percentChange = yearlyChange / firstOpenValue
        
        ws.Cells(tickerRowIndex, lastColumn + 2).Value = currentTicker
        ws.Cells(tickerRowIndex, lastColumn + 3).Value = yearlyChange
        ws.Cells(tickerRowIndex, lastColumn + 4).Value = percentChange
        ws.Cells(tickerRowIndex, lastColumn + 5).Value = totalStockVolume
        
        If percentChange > greatestIncreaseValue Then
            greatestIncreaseValue = percentChange
            greatestIncreaseTicker = currentTicker
        End If
        
        If percentChange < greatestDecreaseValue Then
            greatestDecreaseValue = percentChange
            greatestDecreaseTicker = currentTicker
        End If
        
        If totalStockVolume > greatestTotalVolumeValue Then
            greatestTotalVolumeValue = totalStockVolume
            greatestTotalVolumeTicker = currentTicker
        End If
        
        'Apply conditional formatting
        Dim conditionalColor As Integer
        If yearlyChange > 0 Then
            conditionalColor = 4
        Else
            conditionalColor = 3
        End If
        ws.Cells(tickerRowIndex, lastColumn + 3).Interior.ColorIndex = conditionalColor
        ws.Cells(tickerRowIndex, lastColumn + 4).Interior.ColorIndex = conditionalColor
        
        'Prepare for the next ticker
        tickerRowIndex = tickerRowIndex + 1
        firstOpenValue = ws.Cells(i + 1, 3).Value
        totalStockVolume = 0
    End If
Next i

Dim columnLetter
columnLetter = ColumnNumberToLetter(ws, 11)
ws.Range(columnLetter & "2:" & columnLetter & Trim(Str(tickerRowIndex - 1))).NumberFormat = "0.00%"

ws.Cells(2, lastColumn + 9).Value = greatestIncreaseTicker
ws.Cells(3, lastColumn + 9).Value = greatestDecreaseTicker
ws.Cells(4, lastColumn + 9).Value = greatestTotalVolumeTicker

ws.Cells(2, lastColumn + 10).Value = greatestIncreaseValue
ws.Cells(3, lastColumn + 10).Value = greatestDecreaseValue
ws.Cells(4, lastColumn + 10).Value = greatestTotalVolumeValue

columnLetter = ColumnNumberToLetter(ws, lastColumn + 10)
ws.Range(columnLetter & "2:" & columnLetter & "3").NumberFormat = "0.00%"
ws.Cells(4, lastColumn + 10).NumberFormat = "0.00E+0"

'Auto-fit all new columns
For i = 1 To 10
    ws.Columns(lastColumn + i).EntireColumn.AutoFit
Next i
End Sub

Sub ApplyAnalysisToWorkbook()
Dim ws As Worksheet
Call ClearAnalysis
For Each ws In ThisWorkbook.Worksheets
    Call ApplyAnalysisToWorksheet(ws)
Next ws
End Sub
