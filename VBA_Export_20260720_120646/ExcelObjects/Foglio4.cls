VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
Dim numc, numr, colore, indirizzo As Integer

numr = Target.Row
numc = Target.Column

Sheets("F09").Select
Sheets("F09").Range("B27:H127").Select
Selection.ClearContents
Sheets("F09").Cells(1, 1).Select

If Cells(numr, numc) = "OOO" Or Cells(numr, numc) = "XXX" Then
        kk% = numr - 5
        xx% = numc - 6
        ipoipo = (kk% - 1) * 6 + xx%
        Call crea_disegno
End If

Cells(1, 1).Select

End Sub


Public Sub crea_disegno()
           Call assegna_generali
           fgl = "F09"
           path$ = "C:\TMP_SIF\"
           nome$ = "valori"
           rr% = 27
           marcia$ = "AVANTI"
           rr13 = 4
           cc13 = 1 + ipoipo
           ttvv = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
           GoSub leggi_pompa
           rr% = 27
           marcia$ = "INDIETRO"
           rr13 = 4
           cc13 = 93 + ipoipo
           ttvv = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
           GoSub leggi_pompa
           
           Exit Sub
           
leggi_pompa:
           testdati = 0
           While testdati = 0
                          truotas = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                          truotad = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                          corsa = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                          spz = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                          Fmotrice = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                          If truotas > 0 Then
                            If marcia$ = "AVANTI" Then
                               Sheets("F09").Cells(rr%, 2) = CDbl(spz)
                               Sheets("F09").Cells(rr%, 3) = CDbl(truotas)
                               Sheets("F09").Cells(rr%, 4) = CDbl(truotad)
                               Sheets("F09").Cells(rr%, 5) = CDbl(Fmotrice)
                            Else
                               Sheets("F09").Cells(rr%, 6) = CDbl(truotas)
                               Sheets("F09").Cells(rr%, 7) = CDbl(truotad)
                               Sheets("F09").Cells(rr%, 8) = CDbl(Fmotrice)
                            End If
                            rr% = rr% + 1
                          Else
                               testdati = 1
                          End If
           Wend
           Close
           Return
           
End Sub
