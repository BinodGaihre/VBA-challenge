Attribute VB_Name = "Module1"
Sub quateronescript()
Dim i, j As Integer
Dim cp, op As Double
Dim total, lastr As Long
Dim str As String
total = 0
j = 2
lastr = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quaterly Change($)"
Cells(1, 11).Value = "Percentage change"
Cells(1, 12).Value = "Total stock volume"
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
            Cells(j, 12).Value = total
            
            total = 0
            j = j + 1
            i = i - 1
Next i
End Sub



