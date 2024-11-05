Attribute VB_Name = "Module1"
Sub functionalscript()
Dim i, j As Integer
Dim cp, op, maxa, mina As Double
Dim total, lastr, tt, y As Long
Dim str, stri, strd, strt As String
maxa = 0
mina = 0
tt = 0
total = 0
j = 2
lastr = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quaterly Change($)"
Cells(1, 11).Value = "Percentage change"
Cells(1, 12).Value = "Total stock volume"
Cells(1, 12).Value = "Total stock volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Range("p1") = "Ticker"
Range("q1") = "Value"
For i = 2 To lastr
    str = Cells(i, 1).Value
    Cells(j, 9).Value = str
    op = Cells(i, 3).Value
        Do While Cells(i, 1).Value = str
            total = total + Cells(i, 7).Value
            cp = Cells(i, 6).Value
            i = i + 1
        Loop
        Cells(j, 10).Value = (cp - op)

        If cp > op Then
            Cells(j, 10).Interior.ColorIndex = 4
            ElseIf cp < op Then
            Cells(j, 10).Interior.ColorIndex = 3
            Else
            Cells(j, 10).Interior.ColorIndex = 2
        End If
        Cells(j, 11).Value = FormatPercent((cp - op) / op)
            If maxa < ((cp - op) / op) Then
                maxa = ((cp - op) / op)
                stri = str
            ElseIf mina > ((cp - op) / op) Then
                mina = ((cp - op) / op)
                strd = str
            End If
        Cells(j, 12).Value = total
            If tt < total Then
                tt = total
                strt = str
            End If

            
        total = 0
        j = j + 1
        i = i - 1
Next i
    Range("q4") = tt
    Range("q2") = FormatPercent(maxa)
    Range("q3") = FormatPercent(mina)
    Range("p2") = stri
    Range("p3") = strd
    Range("p4") = strt
    
End Sub



