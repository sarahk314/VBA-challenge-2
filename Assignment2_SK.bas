Attribute VB_Name = "Module1"
Sub mult_stocks()

Dim ws As Worksheet

For Each ws In Worksheets

Dim WorksheetName As String
Dim ticker As String
Dim volume As Double
Dim openyr As Double
Dim closeyr As Double
Dim yearchange As Double
Dim perchange As Double
Dim stocksum As Integer
Dim i As Long
Dim j As Long

WorksheetName = ws.Name

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

Dim lastA As Long
lastA = ws.Cells(Rows.Count, 1).End(xlUp).Row
MsgBox (lastA)

stocksum = 2
j = 2

For i = 2 To lastA

ticker = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value

openyr = ws.Cells(j, 3).Value
closeyr = ws.Cells(i, 6).Value
yearchange = closeyr - openyr
perchange = (closeyr - openyr) / openyr

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
ws.Cells(stocksum, 9).Value = ticker
ws.Cells(stocksum, 10).Value = yearchange
    If ws.Cells(stocksum, 10).Value < 0 Then
    ws.Cells(stocksum, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(stocksum, 10).Interior.ColorIndex = 4
    End If
ws.Cells(stocksum, 11).Value = perchange
    If ws.Cells(j, 3).Value <> 0 Then
    ws.Cells(stocksum, 11).Value = Format(perchange, "Percent")
    Else
    ws.Cells(stocksum, 11).Value = Format(0, "Percent")
    End If

ws.Cells(stocksum, 12).Value = volume

stocksum = stocksum + 1

j = i + 1

volume = 0

End If

Next i

Dim lastI As Long
lastI = ws.Cells(Rows.Count, 9).End(xlUp).Row

Dim greatinc As Double
Dim greatdecr As Double
Dim greatvol As Double

greatincr = ws.Cells(2, 11).Value
greatdecr = ws.Cells(2, 11).Value
greatvol = ws.Cells(2, 12).Value

For i = 2 To lastI

If ws.Cells(i, 12).Value > greatvol Then
greatvol = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
Else
greatvol = greatvol
End If

If ws.Cells(i, 11).Value > greatincr Then
greatincr = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
Else
greatincr = greatincr
End If

If ws.Cells(i, 11).Value < greatdecr Then
greatdecr = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
Else
greatdecr = greatdecr
End If

ws.Cells(2, 17).Value = Format(greatincr, "Percent")
ws.Cells(3, 17).Value = Format(greatdecr, "Percent")
ws.Cells(4, 17).Value = Format(greatvol, "Scientific")

Next i

Worksheets(WorksheetName).Columns("A:Z").AutoFit

Next ws

End Sub
