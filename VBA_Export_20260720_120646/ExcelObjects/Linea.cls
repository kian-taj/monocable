VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




Private Sub crea_relativi_Click()
       Dim j, k, w, Z As Integer
       
       
       nomefoglio = "F03"
       Call elimina_F04(nomefoglio)

       'controlla profilo terreno ramo salita
       nbs = 1
       w = 15
       j = 4
       k = 5
       p_terr = Sheets("F02").Cells(w, j)
       q_terr = Sheets("F02").Cells(w, k)
       While (p_terr + q_terr) <> 0
                  nbs = nbs + 1
                  w = w + 1
                  p_terr = Sheets("F02").Cells(w, j)
                  q_terr = Sheets("F02").Cells(w, k)
       Wend
       If nbs < 2 Then nbs = 0
       'controlla profilo terreno ramo discesa
       nbd = 1
       w = 15
       j = 9
       k = 10
       p_terr = Sheets("F02").Cells(w, j)
       q_terr = Sheets("F02").Cells(w, k)
       While (p_terr + q_terr) <> 0
                  nbd = nbd + 1
                  w = w + 1
                  p_terr = Sheets("F02").Cells(w, j)
                  q_terr = Sheets("F02").Cells(w, k)
       Wend
       If nbd < 2 Then nbd = nbs
             
       'ramo salita
       k = 1
       nc(k) = 0
       If Sheets("F03").Cells(15, 4) = 0 And Sheets("F03").Cells(15, 7) = 0 Then
                           Beep
                           MsgBox ("Mancanza Dati : verificare!")
                           Exit Sub
       Else
             nomesost(k, 1) = Sheets("F03").Cells(15, 3)
             progsost(k, 1) = Sheets("F03").Cells(15, 4)
             quotasost(k, 1) = Sheets("F03").Cells(15, 7)
             Sheets("F04").Cells(6, 5) = progsost(k, 1)
             Sheets("F04").Cells(7, 5) = quotasost(k, 1)
       End If
       w = 15
       j = 1
       cek_pr = Val(Sheets("F03").Cells(w + 1, 4))
       cek_qp = Val(Sheets("F03").Cells(w + 1, 7))
       cek_tot = cek_pr + cek_qp
       While cek_tot <> 0
             w = w + 1
             j = j + 1
             nomesost(k, j) = Sheets("F03").Cells(w, 3)
             progsost(k, j) = Sheets("F03").Cells(w, 4)
             quotasost(k, j) = Sheets("F03").Cells(w, 7)
             nomecamp(k, j - 1) = nomesost(k, j - 1) + "-" + nomesost(k, j)
             Lorizz(k, j - 1) = progsost(k, j) - progsost(k, j - 1)
             Disliv(k, j - 1) = quotasost(k, j) - quotasost(k, j - 1)
             Corda(k, j - 1) = Sqr(Lorizz(k, j - 1) ^ 2 + Disliv(k, j - 1) ^ 2)
             Sheets("F04").Cells(w - 1, 3) = nomecamp(k, j - 1)
             Sheets("F04").Cells(w - 1, 4) = Lorizz(k, j - 1)
             Sheets("F04").Cells(w - 1, 5) = Disliv(k, j - 1)
             nc(k) = nc(k) + 1
             cek_pr = Val(Sheets("F03").Cells(w + 1, 4))
             cek_qp = Val(Sheets("F03").Cells(w + 1, 7))
             cek_tot = cek_pr + cek_qp
      Wend
      For j = w To 94
        For Z = 3 To 5
             Sheets("F04").Cells(j, Z) = ""
        Next Z
      Next j
      
       'ramo discesa
       k = 2
       nc(k) = 0
       If Sheets("F03").Cells(15, 11) = 0 And Sheets("F03").Cells(15, 14) = 0 Then
                           'ramo discesa = ramo salita salita
                           nc(2) = nc(1)
                           For j = 1 To nc(2)
                                 nomesost(2, j) = nomesost(1, j)
                                 progsost(2, j) = progsost(1, j)
                                 quotasost(2, j) = quotasost(1, j)
                                 If j > 1 Then
                                      nomecamp(2, j - 1) = nomecamp(1, j - 1)
                                      Lorizz(2, j - 1) = Lorizz(1, j - 1)
                                      Disliv(2, j - 1) = Disliv(1, j - 1)
                                      Corda(2, j - 1) = Corda(1, j - 1)
                                 End If
                           Next j
       Else
             nomesost(k, 1) = Sheets("F03").Cells(15, 10)
             progsost(k, 1) = Sheets("F03").Cells(15, 11)
             quotasost(k, 1) = Sheets("F03").Cells(15, 14)
             Sheets("F04").Cells(6, 5) = progsost(k, 1)
             Sheets("F04").Cells(7, 5) = quotasost(k, 1)
       End If
       w = 15
       j = 1
       While Val(Sheets("F03").Cells(w + 1, 11)) <> 0 And Val(Sheets("F03").Cells(w + 1, 14)) <> 0
             w = w + 1
             j = j + 1
             nomesost(k, j) = Sheets("F03").Cells(w, 10)
             progsost(k, j) = Sheets("F03").Cells(w, 11)
             quotasost(k, j) = Sheets("F03").Cells(w, 14)
             nomecamp(k, j - 1) = nomesost(k, j - 1) + "-" + nomesost(k, j)
             Lorizz(k, j - 1) = progsost(k, j) - progsost(k, j - 1)
             Disliv(k, j - 1) = quotasost(k, j) - quotasost(k, j - 1)
             Corda(k, j - 1) = Sqr(Lorizz(k, j - 1) ^ 2 + Disliv(k, j - 1) ^ 2)
             Sheets("F04").Cells(w - 1, 7) = nomecamp(k, j - 1)
             Sheets("F04").Cells(w - 1, 8) = Lorizz(k, j - 1)
             Sheets("F04").Cells(w - 1, 9) = Disliv(k, j - 1)
             nc(k) = nc(k) + 1
      Wend
      For j = w To 94
        For Z = 7 To 9
             Sheets("F04").Cells(j, Z) = ""
        Next Z
      Next j
      
      'calcolo delle quote terreno ed altezze verticali sostegni
      'salita
      If nbs > 0 Then
        For j = 1 To nc(1) + 1
                   w = 15
                   While Sheets("F02").Cells(w, 4) < progsost(1, j)
                           p_terr = Sheets("F02").Cells(w, 4)
                           q_terr = Sheets("F02").Cells(w, 5)
                           p_terrs = Sheets("F02").Cells(w + 1, 4)
                           q_terrs = Sheets("F02").Cells(w + 1, 5)
                           w = w + 1
                   Wend
                   If p_terr = progsost(1, j) Then
                                         Sheets("F03").Cells(14 + j, 6) = q_terr
                                         Sheets("F03").Cells(14 + j, 7) = quotasost(1, j) - q_terr
                   Else
                        p_delta = p_terrs - p_terr
                        q_delta = q_terrs - q_terr
                        q_terrpalo = q_terr + (progsost(1, j) - p_terr) * q_delta / p_delta
                        Sheets("F03").Cells(14 + j, 5) = q_terrpalo
                        Sheets("F03").Cells(14 + j, 6) = quotasost(1, j) - q_terrpalo
                   End If
        Next j
      End If
      'calcolo delle quote terreno ed altezze verticali sostegni
      'discesa
      If nbd > 0 Then
         If nc(1) <> nc(2) Then
            For j = 1 To nc(2) + 1
                   w = 15
                   While Sheets("F02").Cells(w, 4) < progsost(2, j)
                           p_terr = Sheets("F02").Cells(w, 9)
                           q_terr = Sheets("F02").Cells(w, 10)
                           p_terrs = Sheets("F02").Cells(w + 1, 9)
                           q_terrs = Sheets("F02").Cells(w + 1, 10)
                           w = w + 1
                   Wend
                   If p_terr = progsost(2, j) Then
                                         Sheets("F03").Cells(14 + j, 12) = q_terr
                                         Sheets("F03").Cells(14 + j, 13) = quotasost(1, j) - q_terr
                   Else
                        p_delta = p_terrs - p_terr
                        q_delta = q_terrs - q_terr
                        q_terrpalo = q_terr + (progsost(2, j) - p_terr) * q_delta / p_delta
                        Sheets("F03").Cells(14 + j, 12) = q_terrpalo
                        Sheets("F03").Cells(14 + j, 13) = quotasost(2, j) - q_terrpalo
                   End If
            Next j
         End If
      End If
      Sheets("F04").Select
      Sheets("F04").Range("C15").Select

End Sub

Private Sub elimina_Click()

      Dim j, k As Integer
If (numcolonna = 2 Or numcolonna = 9) And numriga > 14 Then
      If cek_ins_del Then
              For j = numriga To 93
                     Cells(j, numcolonna + 1) = Cells(j + 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j + 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j + 1, numcolonna + 3)
                     Cells(j, numcolonna + 4) = Cells(j + 1, numcolonna + 4)
                     Cells(j, numcolonna + 5) = Cells(j + 1, numcolonna + 5)
                     Cells(j, numcolonna + 6) = Cells(j + 1, numcolonna + 6)
              Next j
      Else
          Beep
          MsgBox ("DELETE ROW : TO BE A FAILURE")
      End If
Else
          Beep
          MsgBox ("DELETE ROW : TO BE A FAILURE")
End If

End Sub

Private Sub Image1_Click()
        Sheets("MENU").Select
        Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub inserisci_Click()
      
      Dim j, k As Integer
If (numcolonna = 2 Or numcolonna = 9) And numriga > 14 Then
      If cek_ins_del Then
              For j = 94 To numriga + 1 Step -1
                     Cells(j, numcolonna + 1) = Cells(j - 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j - 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j - 1, numcolonna + 3)
                     Cells(j, numcolonna + 4) = Cells(j - 1, numcolonna + 4)
                     Cells(j, numcolonna + 5) = Cells(j - 1, numcolonna + 5)
                     Cells(j, numcolonna + 6) = Cells(j - 1, numcolonna + 6)
              Next j
              Cells(numriga, numcolonna + 1) = " "
              Cells(numriga, numcolonna + 2) = " "
              Cells(numriga, numcolonna + 3) = " "
              Cells(numriga, numcolonna + 4) = " "
              Cells(numriga, numcolonna + 5) = " "
              Cells(numriga, numcolonna + 6) = " "
      Else
          Beep
          MsgBox ("INSERT ROW : TO BE A FAILURE")
      End If
Else
          Beep
          MsgBox ("INSERT ROW : TO BE A FAILURE")
End If

End Sub

Private Sub Worksheet_Activate()
       nome_foglio = "ASSOLUTO"
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
       
       On Local Error GoTo err_range
       
       numriga = Target.Row
       numcolonna = Target.Column
       ntower = Sheets("F03").Cells(numriga, numcolonna)
       If IsNumeric(ntower) Then
                                numsost = Format(ntower, "###")
       Else
             numsost = ntower
       End If
       
If Sheets("HOME").flag_italiano.value Then
       If (numriga > 14 And numriga < 94) And (numcolonna = 2 Or numcolonna = 9) Then
             testo1 = Sheets("F03").Cells(209, 2)
             testo2 = Sheets("F03").Cells(210, 2)
             cek_ins_del = True
             If numcolonna = 2 Then
                                   riga_corrente.Caption = testo1 + " " + numsost + " : " + testo2
                                   riga_corrente_d.Caption = ""
             Else
                  riga_corrente_d.Caption = testo1 + " " + numsost + " : " + testo2
                  riga_corrente.Caption = ""
             End If
       Else
             testo1 = Sheets("F03").Cells(209, 2)
             testo2 = Sheets("F03").Cells(211, 2)
             riga_corrente.Caption = testo1 + " " + numsost + " : " + testo2
             riga_corrente_d.Caption = testo1 + " " + numsost + " : " + testo2
             cek_ins_del = False
       End If
Else
       If (numriga > 14 And numriga < 94) And (numcolonna = 2 Or numcolonna = 9) Then
             testo1 = Sheets("F03").Cells(224, 2)
             testo2 = Sheets("F03").Cells(225, 2)
             cek_ins_del = True
             If numcolonna = 2 Then
                                   riga_corrente.Caption = testo1 + " " + numsost + " : " + testo2
                                   riga_corrente_d.Caption = ""
             Else
                  riga_corrente_d.Caption = testo1 + " " + numsost + " : " + testo2
                  riga_corrente.Caption = ""
             End If
       Else
             testo1 = Sheets("F03").Cells(224, 2)
             testo2 = Sheets("F03").Cells(226, 2)
             riga_corrente.Caption = testo1 + " " + numsost + " : " + testo2
             riga_corrente_d.Caption = testo1 + " " + numsost + " : " + testo2
             cek_ins_del = False
       End If
End If
              
   'controllo per calcolo quote terreno o quote funi
   '---------------------------------------------------------
       If numcolonna = 5 Then GoSub calcola_terreno_salita
       If numcolonna = 4 Then
                            numcolonna = 5
                            numriga = numriga - 1
                            If numriga < 15 Then numriga = 15
                            GoSub calcola_terreno_salita
       End If
       If numcolonna = 12 Then GoSub calcola_terreno_discesa
       If numcolonna = 11 Then
                            numcolonna = 12
                            numriga = numriga - 1
                            If numriga < 15 Then numriga = 15
                            GoSub calcola_terreno_discesa
       End If
       If numcolonna = 7 Then GoSub calcola_qfune
       If numcolonna = 6 Then
                            numcolonna = 7
                            numriga = numriga - 1
                            If numriga < 15 Then numriga = 15
                            GoSub calcola_qfune
       End If
       If numcolonna = 14 Then GoSub calcola_qfune
       If numcolonna = 13 Then
                            numcolonna = 14
                            numriga = numriga - 1
                            If numriga < 15 Then numriga = 15
                            GoSub calcola_qfune
       End If
       If numcolonna = 8 Then GoSub calcola_hsostegno_salita
       If numcolonna = 7 Then
                            numcolonna = 8
                            numriga = numriga - 1
                            If numriga < 15 Then numriga = 15
                            GoSub calcola_hsostegno_salita
       End If
       If numcolonna = 15 Then GoSub calcola_hsostegno_discesa
       If numcolonna = 14 Then
                            numcolonna = 15
                            numriga = numriga - 1
                            If numriga < 15 Then numriga = 15
                            GoSub calcola_hsostegno_discesa
       End If
       Exit Sub
       
calcola_terreno_salita:
      'calcolo delle quote terreno ed altezze verticali sostegni
      'salita
      Call leggi_terreno
      w = 1
      p_fune = Sheets("F03").Cells(numriga, numcolonna - 1)
      If p_fune > prog_terr(1, nbs) Then
                                      Beep
                                      MsgBox ("Attenzione : errore progressiva terreno")
                                      Exit Sub
      End If
      If p_fune = 0 Then
                        'correzione per prima battuta =0
                        p_fune = 0.001
                        Sheets("F03").Cells(numriga, numcolonna - 1) = p_fune
      End If
      'verifica untima progressiva terreno > progressiva fune
      If p_fune >= prog_terr(1, nbs) Then
                                        'aggiunta battura terreno
                                        Sheets("F02").Cells(15 + nbs, 4) = p_fune + 5
                                        Sheets("F02").Cells(15 + nbs, 5) = Sheets("F02").Cells(15 + nbs - 1, 5)
                                        Call leggi_terreno
      End If
      While prog_terr(1, w) < p_fune
                                w = w + 1
      Wend
      If prog_terr(1, w - 1) = p_fune Then
                                Sheets("F03").Cells(numriga, numcolonna) = quot_terr(1, w - 1)
      Else
          p_delta = prog_terr(1, w) - prog_terr(1, w - 1)
          q_delta = quot_terr(1, w) - quot_terr(1, w - 1)
          q_terrpalo = quot_terr(1, w - 1) + (p_fune - prog_terr(1, w - 1)) * q_delta / p_delta
          Sheets("F03").Cells(numriga, numcolonna) = q_terrpalo
      End If
      Return
      
calcola_terreno_discesa:
      'discesa
      Call leggi_terreno
      w = 1
      p_fune = Sheets("F03").Cells(numriga, numcolonna - 1)
      If p_fune = 0 Then Beep: Exit Sub
      While prog_terr(2, w) < p_fune
                                w = w + 1
      Wend
      If prog_terr(2, w - 1) = p_fune Then
                                Sheets("F03").Cells(numriga, numcolonna) = quot_terr(2, w - 1)
      Else
          p_delta = prog_terr(2, w) - prog_terr(2, w - 1)
          q_delta = quot_terr(2, w) - quot_terr(2, w - 1)
          q_terrpalo = quot_terr(2, w - 1) + (p_fune - prog_terr(2, w - 1)) * q_delta / p_delta
          Sheets("F03").Cells(numriga_numcolonna) = q_terrpalo
      End If
      Return
       
calcola_qfune:
      h_fusto = Sheets("F03").Cells(numriga, numcolonna - 1)
      q_terrpalo = Sheets("F03").Cells(numriga, numcolonna - 2)
      If h_fusto > 0 And q_terrpalo > 0 Then
                    Sheets("F03").Cells(numriga, numcolonna) = q_terrpalo + h_fusto
      End If
      Return
      
calcola_hsostegno_salita:
       q_fune = Sheets("F03").Cells(numriga, numcolonna - 1)
       q_terrpalo = Sheets("F03").Cells(numriga, numcolonna - 3)
       h_fusto = q_fune - q_terrpalo
       If Sheets("F03").Cells(numriga, 4) > 0 Then
                                Sheets("F03").Cells(numriga, numcolonna - 2) = h_fusto
       End If
       Return
       
calcola_hsostegno_discesa:
       q_fune = Sheets("F03").Cells(numriga, numcolonna - 1)
       q_terrpalo = Sheets("F03").Cells(numriga, numcolonna - 3)
       h_fusto = q_fune - q_terrpalo
       If Sheets("F03").Cells(numriga, 11) > 0 Then
                                Sheets("F03").Cells(numriga, numcolonna - 2) = h_fusto
       End If
       Return
       
      
       
err_range:
          Beep
          MsgBox ("INVALID SELECTION")
          Exit Sub
End Sub
