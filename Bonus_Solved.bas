Attribute VB_Name = "Module2"
Sub bonus()

Dim ws As Worksheet
Dim i As Long
Dim firstRow As Integer
Dim columnNumber As Integer
Dim max As Double
Dim min As Double
Dim vol As Double
Dim maxticker As String
Dim minticker As String
Dim volticker As String

firstRow = 2
columnNumber = 11

For Each ws In ThisWorkbook.Worksheets
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        
If ws.UsedRange.Rows.Count <= 1 Then max = -1 Else max = ws.Cells(2, 11)
If ws.UsedRange.Rows.Count >= 1 Then min = 0 Else min = ws.Cells(2, 11)
If ws.UsedRange.Rows.Count <= 1 Then vol = 0 Else vol = ws.Cells(2, 12)
        
For i = firstRow To ws.UsedRange.Rows.Count

    If ws.Cells(i, 11) > max Then
    
    max = ws.Cells(i, 11)
    maxticker = ws.Cells(i, 9)
    
    End If
    
    If ws.Cells(i, 11) < min Then
    
    min = ws.Cells(i, 11)
    minticker = ws.Cells(i, 9)
    
    End If
    
    If ws.Cells(i, 12) > vol Then
    
    vol = ws.Cells(i, 12)
    volticker = ws.Cells(i, 9)
    
    End If
          
Next i

ws.Cells(2, 17) = max
ws.Cells(2, 16) = maxticker
ws.Cells(3, 17) = min
ws.Cells(3, 16) = minticker
ws.Cells(4, 17) = vol
ws.Cells(4, 16) = volticker

ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub
