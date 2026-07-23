VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub calcola_ancorata_Click()

       Open "C:\TMP_SIF\MONOFUNI_TEST_PSWKEY" For Input As #1
       Input #1, testkey
       Close
       If testkey <> "activate" Then
                                    Beep
                                    MsgBox ("Codice di attivazione programma NON valido")
                                    End
       End If
       XIndex = 4
       Ipocalc = 0
       nomefoglio = "F05"
       'Call elimina_dati_generali(nomefoglio)
       'Call elimina_store13(nomefoglio)
       'Call elimina_store05(nomefoglio)
       Call prepara_calcolo
       Call esegui_verifica(XIndex)
       'genera tabella max_min
       Sheets("F10").Select
       Call stampa_maxmin(1)
       Sheets("F05").Select
       'genera tabella potenze
       Sheets("F11").Select
       nomeipotesi$ = Sheets("F11").ver_01.Caption
       Call stampa_potenze(1, nomeipotesi$)
       Sheets("F05").Select
       'genera foglio dati generali
       Sheets("F12").Select
       Call rel_gen
       Sheets("F05").Select
       
       'MsgBox ("FINE ELABORAZIONE")
End Sub

Private Sub calcola_idraul_Click()

       Open "C:\TMP_SIF\MONOFUNI_TEST_PSWKEY" For Input As #1
       Input #1, testkey
       Close
       If testkey <> "activate" Then
                                    Beep
                                    MsgBox ("Codice di attivazione programma NON valido")
                                    End
       End If

       XIndex = 1
       Ipocalc = 0
       nomefoglio = "F05"
       'Call elimina_dati_generali(nomefoglio)
       'Call elimina_store13(nomefoglio)
       'Call elimina_store05(nomefoglio)
       Call prepara_calcolo
       Call esegui_verifica(XIndex)
       idraulic_input.Show
End Sub

Private Sub calcola_normale_Click()

       Open "C:\TMP_SIF\MONOFUNI_TEST_PSWKEY" For Input As #1
       Input #1, testkey
       Close
       If testkey <> "activate" Then
                                    Beep
                                    MsgBox ("Codice di attivazione programma NON valido")
                                    End
       End If

       XIndex = 0
       Ipocalc = 0
       nomefoglio = "F05"
       'Call elimina_dati_generali(nomefoglio)
       'Call elimina_store13(nomefoglio)
       'Call elimina_store05(nomefoglio)
       Call prepara_calcolo
       'verifica con tensione nominale
       Call esegui_verifica(XIndex)
       'verifica +/- 10%
       Call calcola_varten_Click
       'genera tabella max_min
       Sheets("F10").Select
       Call stampa_maxmin(1)
       Sheets("F05").Select
       'genera tabella potenze
       Sheets("F11").Select
       nomeipotesi$ = Sheets("F11").ver_01.Caption
       Call stampa_potenze(1, nomeipotesi$)
       Sheets("F05").Select
       'genera foglio dati generali
       Sheets("F12").Select
       Call rel_gen
       Sheets("F05").Select
       'If Sheets("F05").Cells(1, 17) <> "AUTOCAD" Then
       '                                             MsgBox ("FINE ELABORAZIONE")
       'Else
       '      Sheets("F05").Cells(1, 17) = ""
       'End If
End Sub

Private Sub calcola_varten_Click()

       Open "C:\TMP_SIF\MONOFUNI_TEST_PSWKEY" For Input As #1
       Input #1, testkey
       Close
       If testkey <> "activate" Then
                                    Beep
                                    MsgBox ("Codice di attivazione programma NON valido")
                                    End
       End If

       XIndex = 5
       Ipocalc = 1
       nomefoglio = "F05"
       Call elimina_dati_generali(nomefoglio)
       Call prepara_calcolo
       Call esegui_verifica(XIndex)
End Sub



Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub nv_p_Click()
                    Call assegna_generali
                    lanello = Sheets("F05").Cells(3, 15)
                    If ZETA = "F" Then
                        nveic = Sheets("F05").Cells(4, 15)
                        port = Sheets("F05").Cells(7, 15)
                        If nveic > 0 Then
                                         eg = lanello / nveic
                        Else
                            eg = eg0
                            nveic = lanello / eg0
                        End If
                        dd = eg
                        vel = port * dd / 3600 / posti
                        porteff = port
                        Sheets("F05").Cells(5, 15) = eg
                        Sheets("F05").Cells(6, 15) = vel
                        Sheets("F01").Cells(41, 5) = port
                        Sheets("F01").Cells(42, 5) = vel
                    End If
End Sub

Private Sub nv_v_Click()
                    Call assegna_generali
                    lanello = Sheets("F05").Cells(3, 15)
                    If ZETA = "F" Then
                        nveic = Sheets("F05").Cells(4, 15)
                        vel = Sheets("F05").Cells(6, 15)
                        If nveic > 0 Then
                                         eg = lanello / nveic
                        Else
                            eg = eg0
                            nveic = lanello / eg0
                        End If
                        dd = eg
                        porteff = 3600 / dd * posti * vel
                        port = porteff
                        Sheets("F05").Cells(5, 15) = eg
                        Sheets("F05").Cells(7, 15) = porteff
                        Sheets("F01").Cells(41, 5) = port
                        Sheets("F01").Cells(42, 5) = vel
                    End If
End Sub

Private Sub offset_fine_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = offset_fine.Text
        Call virgola_to_punto(tstringa)
        offset_fine.Text = tstringa
End Sub

Private Sub offset_inizio_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
        Sheets("F05").Select
        Sheets("F05").Cells(1, 1).Select
        Sheets("F05").Range("T59:AF3000").Select
        Selection.ClearContents
        Sheets("F05").Cells(1, 1).Select
        
        Ins% = ns + 4 + 2   'massimo numero di sostegni per ramo salita
        nr% = 2              'massimo numero di rami impianto
        u% = 40             'massimo veicoli in campata
        T% = 400            'massimo numero di veicoli per ramo di impianto
        '----------------------------dim shared  ARRAY---------------------------------------
        ReDim Pr(nr%, T%), i(nr%)

       'vedi posizione carichi per un determinato spiazzamento
        MyVar = Sheets("F05").offset_inizio.Text
        MyCheck = IsNumeric(MyVar)
        If MyCheck = True Then spz = CDbl(Sheets("F05").offset_inizio.Text) Else spz = 0
        Call posizione_veicoli(spz)

End Sub

Private Sub offset_inizio_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = offset_inizio.Text
        Call virgola_to_punto(tstringa)
        offset_inizio.Text = tstringa
End Sub



Private Sub p_v_Click()
                    Call assegna_generali
                    lanello = Sheets("F05").Cells(3, 15)
                    port = Sheets("F05").Cells(7, 15)
                    vel = Sheets("F05").Cells(6, 15)
                    If ZETA = "F" Then
                        eg0 = 3600 / port * vel * posti
                        nveic = lanello / eg0
                        resto = Abs(Int(nveic) - nveic)
                        If resto > 0 Then
                             If resto > 0.5 Then
                                    nveic = Int(nveic) + 1
                             Else
                                  nveic = Int(nveic) - 1
                             End If
                        End If
                        eg = eg0
                        porteff = port
                        dd = eg
                        Sheets("F05").Cells(4, 15) = nveic
                        Sheets("F05").Cells(5, 15) = eg
                        Sheets("F01").Cells(41, 5) = port
                        Sheets("F01").Cells(42, 5) = vel
                    End If
End Sub



Private Sub reset_all_Click()
Dim numc, numr, colore As Integer
Cells(1, 1).Select
ActiveSheet.Unprotect ("vit210147vit")
For numr = 10 To 24
     For numc = 4 To 9
         Cells(numr, numc).Select
         Selection.Interior.ColorIndex = 2
         Cells(numr, numc) = ""
     Next numc
Next numr
Cells(1, 1).Select
ActiveSheet.Protect password:="vit210147vit"
End Sub

Private Sub salto_termico_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = salto_termico.Text
        Call virgola_to_punto(tstringa)
        salto_termico.Text = tstringa
End Sub


Private Sub v_passo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = v_passo.Text
        Call virgola_to_punto(tstringa)
        v_passo.Text = tstringa
End Sub

Private Sub var_tensione_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = var_tensione.Text
        Call virgola_to_punto(tstringa)
        var_tensione.Text = tstringa
End Sub

Private Sub Worksheet_Activate()

        passo = 0
        spziniz = 0
        spzfin = 0

        'lettura dati generali dell'impianto-----------------------------
        lanello = 0: nveic = 0
        Call assegna_generali
        
        'lettura linea --------------------------------------------------
        Call leggi_linea
        
        If ns < 2 Then
                      Beep
                      MsgBox ("ATTENZIONE : MANCA INPUT DATI")
                      Sheets("F05").Select
                      Exit Sub
        End If
        '--------------------------------------------------------
        'ns=sostegni salita (compreso asse ruote ed avanstazioni)
        'nd=sostegni discesa(      "         "           "      )
        'Prs(1,x) = progressiva sostegni salita
        'Prs(2,x) = progressiva sostegni discesa
        '--------------------------------------------------------

        'salita
         Lsvm1 = 0: Prs(1, 1) = 0: Lavm1 = 0
         For X = 1 To ns - 1
                         Lsvm1 = Lsvm1 + Sqr(d(1, X) ^ 2 + l(1, X) ^ 2)
                         Prs(1, X + 1) = Lsvm1
                         If X > 1 And X < ns - 1 Then
                                      Lavm1 = Lavm1 + Sqr(d(1, X) ^ 2 + l(1, X) ^ 2)
                         End If
         Next X
        'discesa
        Lsvm2 = 0: Prs(2, 1) = 0: Lavm2 = 0
        For X = 1 To nd - 1
                         Lsvm2 = Lsvm2 + Sqr(d(2, X) ^ 2 + l(2, X) ^ 2)
                         Prs(2, X + 1) = Lsvm2
                         If X > 1 And X < nd - 1 Then
                                      Lavm2 = Lavm2 + Sqr(d(2, X) ^ 2 + l(2, X) ^ 2)
                         End If
        Next X

        'lunghezza complessiva dell'anello di fune
        lanello = Lsvm1 + Lsvm2 + PI# * (Drt / 1000 + Drtrin / 1000) / 2

        If ZETA = "I" Then
                       eg = lanello / (2 * ng)   'equidistanza dei grappoli
                       porteff = port
                       nveic = 2 * ng * vg
        Else
           hg0 = 3600 / port * vel * posti    'equidistanza teorica dei veicoli
           eg0 = hg0
           If ZETA = "F" Then
                    'monofuni ad ammorsamento fisso
                    'in funzione della portata e velocità
                    If Sheets("F05").p_v.value Then
                        nveic = lanello / eg0
                        resto = Abs(Int(nveic) - nveic)
                        If resto > 0 Then
                             If resto > 0.5 Then
                                    nveic = Int(nveic) + 1
                             Else
                                  nveic = Int(nveic) - 1
                             End If
                        End If
                        eg = eg0
                        porteff = port
                        dd = eg
                    End If
                    If Sheets("F05").nv_v.value Then
                        nveic = Sheets("F05").Cells(4, 15)
                        If nveic > 0 Then
                                         eg = lanello / nveic
                        Else
                            eg = eg0
                            nveic = lanello / eg0
                        End If
                        dd = eg
                        porteff = 3600 / dd * posti * vel
                    End If
                    If Sheets("F05").nv_p.value Then
                        nveic = Sheets("F05").Cells(4, 15)
                        If nveic > 0 Then
                                         eg = lanello / nveic
                        Else
                            eg = eg0
                            nveic = lanello / eg0
                        End If
                        dd = eg
                        vel = port * dd / 3600 / posti
                        porteff = port
                    End If
           Else
               'monofuni ad ammorsamento automatico
               eg0 = hg0: eg = eg0: dd = eg: porteff = port
               nveic = (Lavm1 + Lavm2) / eg0
               resto = nveic - Int(nveic)
               If resto > 0.5 Then
                                nveic = Int(nveic) + 1
               Else
                   nveic = Int(nveic)
               End If
           End If
           Nvg = 1
        End If
        Sheets("F05").Cells(3, 15) = lanello
        Sheets("F05").Cells(4, 15) = nveic
        Sheets("F05").Cells(5, 15) = eg
        Sheets("F05").Cells(6, 15) = vel
        Sheets("F05").Cells(7, 15) = porteff

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
Dim numc, numr, colore, indirizzo As Integer

numr = Target.Row
numc = Target.Column
Cells(1, 1).Select
ActiveSheet.Unprotect ("vit210147vit")

If numc >= 4 And numc <= 9 Then
  If numr >= 10 And numr <= 24 Then
    Cells(numr, numc).Select
    colore = Selection.Interior.ColorIndex
    'MsgBox (Str(colore))
    If colore = 2 Then Selection.Interior.ColorIndex = 5: Cells(numr, numc) = "OOO"
    If colore = 5 Then Selection.Interior.ColorIndex = 3: Cells(numr, numc) = "XXX"
    If colore = 3 Then Selection.Interior.ColorIndex = 2: Cells(numr, numc) = ""
  End If
End If

Cells(1, 1).Select
ActiveSheet.Protect password:="vit210147vit"

If numc = 1 Then
  If numr >= 15 And numr <= 24 Then
       Sheets("F06").Select
       indirizzo = numr - 14
       indirizzo = 3 + (indirizzo - 1) * 2
       Sheets("F06").Cells(12, indirizzo).Select
       Exit Sub
  End If
End If

Cells(1, 1).Select

End Sub


Public Sub prepara_calcolo()
       Call assegna_generali
       Call leggi_linea
       
       Sheets("F09").Select
       Sheets("F09").Range("B27:H127").Select
       Selection.ClearContents
       Sheets("F09").Cells(1, 1).Select
       
       Sheets("F08").Select
       Sheets("F08").Range("B30:P200").Select
       Selection.ClearContents
       Sheets("F08").Range("B23:P26").Select
       Selection.ClearContents
       Sheets("F08").Cells(30, 2).Select
       
       Sheets("F10").Select
       Sheets("F10").Range("B30:P200").Select
       Selection.ClearContents
       Sheets("F10").Range("B23:P26").Select
       Selection.ClearContents
       Sheets("F10").Cells(30, 2).Select
       
       Sheets("F11").Select
       Sheets("F11").Range("B30:Q200").Select
       Selection.ClearContents
       Sheets("F11").Range("B23:Q26").Select
       Selection.ClearContents
       Sheets("F11").Cells(30, 2).Select
       
       Sheets("F05").Select
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
       rr = Target.Row
       cc = Target.Column
       If rr = 1 And cc = 50 Then
                                 'avvia esecuzione calcolo normale
                                 Sheets("F10").test_grafico.value = False
                                 Sheets("F05").Cells(1, 1).Select
                                 Application.WindowState = xlMaximized
                                 Call calcola_normale_Click
                                 Application.WindowState = xlMinimized
       End If
       If rr = 2 And cc = 50 Then
                                 'avvia esecuzione calcolo ancorata
                                 Sheets("F10").test_grafico.value = False
                                 Sheets("F05").Cells(1, 1).Select
                                 Application.WindowState = xlMaximized
                                 Call calcola_ancorata_Click
                                 Application.WindowState = xlMinimized
       End If
       If rr = 3 And cc = 50 Then
                                 'avvia esecuzione calcolo idraulica
                                 Sheets("F10").test_grafico.value = False
                                 Sheets("F05").Cells(1, 1).Select
                                 Application.WindowState = xlMaximized
                                 Call calcola_idraul_Click
                                 Application.WindowState = xlMinimized
       End If
       
End Sub

