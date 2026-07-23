VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Sub stampa_foglio()
           Dim area, areastampa, areaprint As String
           Call leggi_linea
           'leggi da archivio i valori calcolati e stampa su foglio
           fgl = "F08"
           path$ = "C:\TMP_SIF\"
           nome$ = "valori"
           Sheets(fgl).Range("B30:P200").Select
           Selection.ClearContents
           Sheets(fgl).Range("B23:P26").Select
           Selection.ClearContents
           Sheets(fgl).Cells(1, 1).Select
           
           Call salti_pagina(fgl)
           Sheets(fgl).Cells(23, 2) = Sheets("F01").Cells(3, 4)
           Sheets(fgl).Cells(24, 2) = Sheets("F13").Cells(kk%, 1) + " : " + Sheets("F13").Cells(xx%, 5)
           Sheets(fgl).Cells(23, 9) = Sheets("F05").Cells(5, 11)
           Sheets(fgl).Cells(24, 9) = Sheets("F05").Cells(6, 11)
           Sheets(fgl).Cells(25, 9) = Sheets("F05").Cells(7, 11)
           Sheets(fgl).Cells(23, 13) = Sheets("F05").Cells(5, 15)
           Sheets(fgl).Cells(24, 13) = Sheets("F05").Cells(6, 15)
           Sheets(fgl).Cells(25, 13) = Sheets("F05").Cells(7, 15)
           Sheets(fgl).Cells(24, 16) = Date
           Sheets(fgl).Cells(25, 16) = Time
           
           rr% = 30
           Sheets(fgl).Cells(rr%, 2) = Sheets("F13").Cells(1, 8)
           rr% = rr% + 1
           If opt_00.value Then Ipocalc = 0
           If opt_01.value Then Ipocalc = 1
           If opt_02.value Then Ipocalc = 2
           If opt_03.value Then Ipocalc = 3
           Close
           Select Case Ipocalc
                  Case 0
                        rr6 = 5
                  Case 1
                        rr6 = 4005
                  Case 3
                        rr6 = 6005
                  Case 2
                        rr6 = 2005
           End Select
           cc6 = 1 + ipoipo
           
           flagmed% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           ns = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           nd = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           'lettura e stampa del ramo salita -------------------------------------------------
           ssdd% = 1
           cc6 = 1 + ipoipo
           For j% = 1 To ns - 1
                          GoSub leggi_dati_foglio
           Next j%
           
           'inserimento di pagina nuova per stampa ramo discesa-------------------------------
           Cells(rr%, 1).Select
           ActiveWindow.SelectedSheets.HPageBreaks.add Before:=ActiveCell
           
           Sheets(fgl).Cells(rr%, 2) = Sheets("F13").Cells(2, 8)
           rr% = rr% + 1
           
           'lettura e stampa del ramo discesa ------------------------------------------------
           ssdd% = 2
           Select Case Ipocalc
                  Case 0
                        rr6 = 5
                  Case 1
                        rr6 = 4005
                  Case 3
                        rr6 = 6005
                  Case 2
                        rr6 = 2005
           End Select
           rr6 = rr6 + 3
           cc6 = 92 + ipoipo
           For j% = 1 To nd - 1
                          GoSub leggi_dati_foglio
           Next j%
           Close
           
           'imposta l'area di stampa
           area = Format(rr%, "####")
           areastampa = "A23:P" + area
           areaprint = "$A$23:$P$" + area
           Range(areastampa).Select
           ActiveSheet.PageSetup.PrintArea = areaprint
           Exit Sub
           
leggi_dati_foglio:
           cval = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           cmon = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           tvv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           tmm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           ffcc = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           avv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           amm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           GoSub riempi_celle_1
           If flagmed% = 0 Then
                        tvv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        tmm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        ffcc = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        avv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        amm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        cval = "": cmon = ""
                        GoSub riempi_celle_1
           End If
           If (ssdd% = 1 And j% <> ns - 1) Or (ssdd% = 2 And j% <> nd - 1) Then
              nsost = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
              tpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
              devpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
              pspp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
              attrpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
              GoSub riempi_celle_2
              If flagmed% = 0 Then
                        tpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        devpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        pspp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        attrpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        nsost = ""
                        GoSub riempi_celle_2
              End If
           End If
           Return
           
riempi_celle_1:
           Sheets(fgl).Cells(rr%, 2) = cval
           Sheets(fgl).Cells(rr%, 3) = cmon
           Sheets(fgl).Cells(rr%, 4) = CDbl(tvv)
           Sheets(fgl).Cells(rr%, 5) = CDbl(tmm)
           Sheets(fgl).Cells(rr%, 6) = CDbl(ffcc)
           Sheets(fgl).Cells(rr%, 7) = CDbl(avv * 180 / PI#)
           Sheets(fgl).Cells(rr%, 8) = CDbl(amm * 180 / PI#)
           rr% = rr% + 1
           Return
riempi_celle_2:
           Sheets(fgl).Cells(rr%, 9) = nsost
           Sheets(fgl).Cells(rr%, 10) = CDbl(tpp)
           Sheets(fgl).Cells(rr%, 11) = CDbl(devpp * 180 / PI#)
           Sheets(fgl).Cells(rr%, 12) = CDbl(pspp)
           Sheets(fgl).Cells(rr%, 13) = CDbl(attrpp)
           rrulli = rulli(ssdd%, j% + 1)
           If nsost <> "" And rrulli <> 0 Then Sheets(fgl).Cells(rr%, 14) = rrulli
           If rrulli <> 0 Then
                             Sheets(fgl).Cells(rr%, 15) = CDbl(devpp / rrulli * 180 / PI#)
                             Sheets(fgl).Cells(rr%, 16) = CDbl(pspp / rrulli)
           End If
           rr% = rr% + 1
           Return

End Sub


Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
Dim numc, numr, colore, indirizzo As Integer

numr = Target.Row
numc = Target.Column

If Cells(numr, numc) = "OOO" Or Cells(numr, numc) = "XXX" Then
        kk% = numr - 5
        xx% = numc - 6
        ipoipo = (kk% - 1) * 6 + xx%
        Call stampa_foglio
End If

Cells(1, 1).Select

End Sub

