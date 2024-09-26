
Sub AnalyzeAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim i As Long
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim increaseTicker As String, decreaseTicker As String, volumeTicker As String

    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize the start row (assuming headers are in row 1)
        startRow = 2
        
        ' Create new columns for output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        
        ' Loop through each row of stock data
        For i = startRow To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' New ticker symbol detected
                If i > startRow Then
                    ' Perform calculations for the previous ticker
                    closePrice = ws.Cells(i - 1, 6).Value ' Column F: <close>
                    totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(i - 1, 7))) ' Column G: <vol>
                    quarterlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = (quarterlyChange / openPrice) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Output results for the previous ticker
                    ws.Cells(startRow, 9).Value = ticker
                    ws.Cells(startRow, 10).Value = totalVolume
                    ws.Cells(startRow, 11).Value = quarterlyChange
                    ws.Cells(startRow, 12).Value = percentChange
                    
                    ' Apply conditional formatting
                    Call ApplyConditionalFormatting(ws, startRow, 11, 12)
                    
                    ' Check for greatest values
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        increaseTicker = ticker
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        decreaseTicker = ticker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        volumeTicker = ticker
                    End If
                End If
                
                ' Reset for the new ticker
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value ' Column C: <open>
                startRow = i
            End If
        Next i
        
        ' Output for the last ticker in the sheet
        closePrice = ws.Cells(lastRow, 6).Value
        totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(lastRow, 7)))
        quarterlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentChange = (quarterlyChange / openPrice) * 100
        Else
            percentChange = 0
        End If
        ws.Cells(startRow, 9).Value = ticker
        ws.Cells(startRow, 10).Value = totalVolume
        ws.Cells(startRow, 11).Value = quarterlyChange
        ws.Cells(startRow, 12).Value = percentChange
        Call ApplyConditionalFormatting(ws, startRow, 11, 12)
    Next ws
    
    ' Output the greatest values
    MsgBox "Greatest % Increase: " & increaseTicker & " (" & greatestIncrease & "%)" & vbCrLf & _
           "Greatest % Decrease: " & decreaseTicker & " (" & greatestDecrease & "%)" & vbCrLf & _
           "Greatest Total Volume: " & volumeTicker & " (" & greatestVolume & ")"
End Sub

Sub ApplyConditionalFormatting(ws As Worksheet, rowNum As Long, changeCol As Long, percentCol As Long)
    Dim changeRng As Range, percentRng As Range
    Set changeRng = ws.Cells(rowNum, changeCol)
    Set percentRng = ws.Cells(rowNum, percentCol)
    
    ' Clear existing formats
    changeRng.FormatConditions.Delete
    percentRng.FormatConditions.Delete
    
    ' Apply conditional formatting to the quarterly change (column changeCol)
    With changeRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Green for positive change
    End With
    With changeRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Red for negative change
    End With
    
    ' Apply conditional formatting to the percent change (column percentCol)
    With percentRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0)
    End With
    With percentRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0)
    End With
End Sub
