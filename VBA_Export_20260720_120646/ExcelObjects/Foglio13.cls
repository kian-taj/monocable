VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub grafico_tens_posa_Click()

        'trova la tensione di posa nel range di temperature da -20 a + 30
        
        riga = 100
        For tamb = -20 To 30
                            Sheets("F20").Cells(9, 7) = tamb
                            Call seek_step
                            Sheets("F20").Cells(riga, 1) = tamb
                            Sheets("F20").Cells(riga, 2) = Sheets("F20").Cells(14, 6)
                            Sheets("F20").Cells(riga, 3) = Sheets("F20").Cells(14 + ns - 2, 6)
                            riga = riga + 1
        Next tamb
        Sheets("F20").Cells(100, 1).Select
        
End Sub

Public Sub tensione_posa_Click()
         
         'trova la tensione di posa alla temperatura ambiente impostata nella cella
         
         Call seek_step

End Sub



Public Sub seek_step()

Sheets("F20").Select
Sheets("F20").Cells(1, 1).Select
Sheets("F20").Range("A14:D94,M14:P94,F14").Select
Range("F14").Activate
Selection.ClearContents
Sheets("F20").Cells(1, 1).Select


Call assegna_generali
Call leggi_linea

tempamb = Sheets("F20").Cells(9, 7)

Sheets("F20").Cells(9, 7) = 0

'completa tabella salita
rr = 14
dd = 0
For j = 1 To ns - 1
                   Sheets("F20").Cells(rr, 1) = n$(1, j)
                   Sheets("F20").Cells(rr, 2) = n$(1, j + 1)
                   Sheets("F20").Cells(rr, 3) = l(1, j)
                   Sheets("F20").Cells(rr, 4) = d(1, j)
                   rr = rr + 1
                   dd = dd + d(1, j)
Next j

'assegna tensione di posa alla temperatura di riferimento di 0°C
If tt$ = "VALLE" Then
                     tirorif = contr / 2
Else
        tirorif = contr / 2 - mf * 9.81 / 10 * dd
End If

Sheets("F20").Cells(4, 7) = tirorif
Sheets("F20").Cells(14, 6) = tirorif

'completa tabella discesa
rr = 14
For j = 1 To ns - 1
                   Sheets("F20").Cells(rr, 13) = n$(2, j)
                   Sheets("F20").Cells(rr, 14) = n$(2, j + 1)
                   Sheets("F20").Cells(rr, 15) = l(2, j)
                   Sheets("F20").Cells(rr, 16) = d(2, j)
                   rr = rr + 1
Next j

'fune a terra alla temperatura di riferimento = zero °C
sv1 = Sheets("F20").Cells(11, 9)
sv2 = Sheets("F20").Cells(11, 21)
ae1 = Sheets("F20").Cells(11, 10)
at1 = Sheets("F20").Cells(11, 11)
ae2 = Sheets("F20").Cells(11, 22)
at2 = Sheets("F20").Cells(11, 23)
Sheets("F20").Cells(6, 7) = sv1 + sv2 - ae1 - at1 - ae2 - at2

Sheets("F20").Cells(9, 7) = tempamb

'fune a terra alla temp.ambiente
sv1 = Sheets("F20").Cells(11, 9)
sv2 = Sheets("F20").Cells(11, 21)
ae1 = Sheets("F20").Cells(11, 10)
at1 = Sheets("F20").Cells(11, 11)
ae2 = Sheets("F20").Cells(11, 22)
at2 = Sheets("F20").Cells(11, 23)


Sheets("F20").Range("G8").GoalSeek Goal:=0, ChangingCell:=Range("F14")

End Sub
