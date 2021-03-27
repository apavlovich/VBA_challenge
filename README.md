# VBA_challenge

VBA Code:

    Sub tickertotal()

    Dim ws As Worksheet
    Dim ticker As String
    Dim volume As Long
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim summaryTableRow As Integer

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"
    
    summaryTableRow = 2
    
    For i = 2 To ws.UsedRange.Rows.Count
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        yearOpen = ws.Cells(i, 3).Value
        yearClose = ws.Cells(i, 6).Value
        
        yearlyChange = yearClose - yearOpen
        percentChange = (yearClose - yearOpen) / yearClose
        
        ws.Cells(summaryTableRow, 9).Value = ticker
        ws.Cells(summaryTableRow, 10).Value = yearlyChange
        ws.Cells(summaryTableRow, 11).Value = percentChange
        ws.Cells(summaryTableRow, 12).Value = volume
        summaryTableRow = summaryTableRow + 1
        
        vol = 0
                
    End If
    Next i

    ws.Columns("K").NumberFormat = "0.00%"

    Dim rng As Range
    Dim g As Long
    Dim c As Long
    Dim colors As Range

    Set rng = ws.Range("J2", Range("J2").End(xlDown))
    c = rng.Cells.Count

    For g = 1 To c
      Set colors = rng(g)
    Select Case colors
        Case Is >= 0
            With colors
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With colors
                .Interior.Color = vbRed
            End With
    End Select
    Next g

    Next ws

    End Sub
    
  Bonus Question:
    
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


  
