Attribute VB_Name = "Module1"
Sub functionalscript()
Dim w As Worksheet
For Each w In Worksheets
Dim i, j As Integer
Dim cp, op, maxa, mina As Double
Dim total, lastr, tt, y As Long
Dim str, stri, strd, strt As String
maxa = 0
mina = 0
tt = 0
total = 0
j = 2
lastr = w.Cells(Rows.Count, 1).End(xlUp).Row
w.Cells(1, 9).Value = "Ticker"
w.Cells(1, 10).Value = "Quaterly Change($)"
w.Cells(1, 11).Value = "Percentage change"
w.Cells(1, 12).Value = "Total stock volume"
w.Cells(1, 12).Value = "Total stock volume"
w.Cells(2, 15).Value = "Greatest % Increase"
w.Cells(3, 15).Value = "Greatest % decrease"
w.Cells(4, 15).Value = "Greatest Total Volume"
w.Range("p1") = "Ticker"
w.Range("q1") = "Value"
For i = 2 To lastr
    str = w.Cells(i, 1).Value
    w.Cells(j, 9).Value = str
    op = w.Cells(i, 3).Value
        Do While w.Cells(i, 1).Value = str
            total = total + w.Cells(i, 7).Value
            cp = w.Cells(i, 6).Value
            i = i + 1
        Loop
        w.Cells(j, 10).Value = (cp - op)

        If cp > op Then
            w.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf cp < op Then
            w.Cells(j, 10).Interior.ColorIndex = 3
            Else
            w.Cells(j, 10).Interior.ColorIndex = 2
        End If
        w.Cells(j, 11).Value = FormatPercent((cp - op) / op)
            If maxa < ((cp - op) / op) Then
                maxa = ((cp - op) / op)
                stri = str
            ElseIf mina > ((cp - op) / op) Then
                mina = ((cp - op) / op)
                strd = str
            End If
        w.Cells(j, 12).Value = total
            If tt < total Then
                tt = total
                strt = str
            End If

            
        total = 0
        j = j + 1
        i = i - 1
Next i
    w.Range("q4") = tt
    w.Range("q2") = FormatPercent(maxa)
    w.Range("q3") = FormatPercent(mina)
    w.Range("p2") = stri
    w.Range("p3") = strd
    w.Range("p4") = strt
Next w
End Sub




