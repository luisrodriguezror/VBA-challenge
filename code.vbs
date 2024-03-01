Sub hwk()

For Each ws In Worksheets


Dim n As Long
Dim i As Long
Dim num As String
Dim inicial As String
Dim final As String
Dim openy As String
Dim closey As String
Dim increase As String
Dim columnapor As Range
Dim max As Double
Dim columnamin As Range
Dim min As Double
Dim columamax As Range
Dim maxpor As Double


n = 1

num = 0
inicial = 0
final = 0
openy = ws.Cells(2, 3).Value
closey = 0
increase = 0
Set columnapor = ws.Range("L:L")
Set columnamin = ws.Range("K:K")
Set columnamax = ws.Range("k:k")



For i = 2 To 753001


If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
num = num + ws.Cells(i, 7).Value

End If

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
n = n + 1
ws.Cells(n, 9).Value = ws.Cells(i, 1).Value
ws.Cells(n, 12).Value = num + ws.Cells(i, 7).Value
closey = ws.Cells(i, 6).Value
ws.Cells(n, 10).Value = closey - openy
ws.Cells(n, 11).Value = ((closey - openy) / openy)
ws.Cells(n, 11).NumberFormat = "0.00%"
openy = ws.Cells(i + 1, 3).Value
num = 0

If ws.Cells(n, 10).Value > 0 Then
ws.Cells(n, 10).Interior.Color = RGB(0, 255, 0)

Else
ws.Cells(n, 10).Interior.Color = RGB(255, 0, 0)
End If



End If





If ws.Cells(n, 11).Value > increase Then
increase = ws.Cells(n, 11).Value
ws.Range("p2").Value = ws.Cells(n, 9).Value
ws.Range("q2").Value = increase
ws.Range("q2").NumberFormat = "0.00%"
End If

maxpor = WorksheetFunction.max(columnamax)
ws.Range("q2").Value = maxpor

If ws.Range("q2").Value = ws.Cells(n, 11).Value Then
ws.Range("p2").Value = ws.Cells(n, 9).Value
End If


max = WorksheetFunction.max(columnapor)
ws.Range("q4").Value = max

If ws.Range("q4").Value = ws.Cells(n, 12).Value Then
ws.Range("p4").Value = ws.Cells(n, 9).Value
End If

min = WorksheetFunction.min(columnamin)
ws.Range("q3").Value = min

If ws.Range("q3").Value = ws.Cells(n, 11).Value Then
ws.Range("p3").Value = ws.Cells(n, 9).Value
End If


Next i



ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
ws.Range("q3").NumberFormat = "0.00%"


 
Next ws
End Sub