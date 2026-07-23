VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linea1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




Private Sub elimina_Click()

      Dim j, k As Integer
  If (numcolonna = 2 Or numcolonna = 6) And numriga > 14 Then
      If cek_ins_del Then
              For j = numriga To 93
                     Cells(j, numcolonna + 1) = Cells(j + 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j + 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j + 1, numcolonna + 3)
              Next j
      Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo = Sheets("F04").Cells(200, 11)
          Else
                testo = Sheets("F04").Cells(215, 11)
          End If
          MsgBox (testo)
      End If
  Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo = Sheets("F04").Cells(200, 11)
          Else
                testo = Sheets("F04").Cells(215, 11)
          End If
          MsgBox (testo)
  End If
End Sub

Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub inserisci_Click()
      
      Dim j, k As Integer
If (numcolonna = 2 Or numcolonna = 6) And numriga > 14 Then
      If cek_ins_del Then
              For j = 94 To numriga + 1 Step -1
                     Cells(j, numcolonna + 1) = Cells(j - 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j - 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j - 1, numcolonna + 3)
              Next j
              Cells(numriga, numcolonna + 1) = " "
              Cells(numriga, numcolonna + 2) = " "
              Cells(numriga, numcolonna + 3) = " "
      Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo = Sheets("F04").Cells(200, 11)
          Else
                testo = Sheets("F04").Cells(215, 11)
          End If
          MsgBox (testo)
      End If
Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo = Sheets("F04").Cells(200, 11)
          Else
                testo = Sheets("F04").Cells(215, 11)
          End If
          MsgBox (testo)
End If

End Sub

Private Sub calc_assoluti_Click()

       Dim j, k, w, p As Integer
       Dim ncamp As String
       
       'pulizia dei campi valori del foglio F03
       nomefoglio = "F04"
       Call elimina_F03(nomefoglio)
       
       'controlla profilo terreno ramo salita
       Call leggi_terreno
       
       'ramo salita
       k = 1
       nc(k) = 0
       If Sheets("F04").Cells(15, 4) = 0 Then
                           Beep
                           If Sheets("HOME").flag_italiano Then
                                              testo = Sheets("F04").Cells(201, 11)
                           Else
                               testo = Sheets("F04").Cells(216, 11)
                           End If
                           MsgBox (testo)
                           Exit Sub
       Else
             progsost(k, 1) = Sheets("F04").Cells(6, 5)
             quotasost(k, 1) = Sheets("F04").Cells(7, 5)
             If quotasost(k, 1) + progsost(k, 1) = 0 Then
                           Beep
                           If Sheets("HOME").flag_italiano Then
                                              testo = Sheets("F04").Cells(202, 11)
                           Else
                               testo = Sheets("F04").Cells(217, 11)
                           End If
                           MsgBox (testo)
                           Exit Sub
             End If
             Sheets("F03").Cells(15, 4) = progsost(k, 1)
             Sheets("F03").Cells(15, 7) = quotasost(k, 1)
       End If
       
       w = 15
       j = 1
       While Sheets("F04").Cells(w, 4) <> 0
             nomecamp(k, j) = Sheets("F04").Cells(w, 3)
             Lorizz(k, j) = Sheets("F04").Cells(w, 4)
             Disliv(k, j) = Sheets("F04").Cells(w, 5)
             Corda(k, j) = Sqr(Lorizz(k, j) ^ 2 + Disliv(k, j) ^ 2)
             nc(k) = nc(k) + 1
             progsost(k, j + 1) = progsost(k, j) + Lorizz(k, j)
             quotasost(k, j + 1) = quotasost(k, j) + Disliv(k, j)
             Sheets("F03").Cells(w + 1, 4) = progsost(k, j + 1)
             Sheets("F03").Cells(w + 1, 7) = quotasost(k, j + 1)
             p = InStr(nomecamp(k, j), "-")
             If p <> 0 Then
                         Sheets("F03").Cells(w, 3) = mid(nomecamp(k, j), 1, p - 1)
                         Sheets("F03").Cells(w + 1, 3) = mid(nomecamp(k, j), p + 1, Len(nomecamp(k, j)))
             End If
             w = w + 1
             j = j + 1
      Wend
      For j = w + 1 To 94
             Sheets("F03").Cells(j, 3) = ""
             Sheets("F03").Cells(j, 4) = ""
             Sheets("F03").Cells(j, 5) = ""
             Sheets("F03").Cells(j, 6) = ""
             Sheets("F03").Cells(j, 7) = ""
             Sheets("F03").Cells(j, 8) = ""
      Next j
      
       'completa la tabella delle quote terreno ed altezze sostegni
       
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
       End If
       w = 15
       j = 1
       While Val(Sheets("F03").Cells(w + 1, 4)) <> 0 And Val(Sheets("F03").Cells(w + 1, 7)) <> 0
             w = w + 1
             j = j + 1
             nomesost(k, j) = Sheets("F03").Cells(w, 3)
             progsost(k, j) = Sheets("F03").Cells(w, 4)
             quotasost(k, j) = Sheets("F03").Cells(w, 7)
             nomecamp(k, j - 1) = nomesost(k, j - 1) + "-" + nomesost(k, j)
             Lorizz(k, j - 1) = progsost(k, j) - progsost(k, j - 1)
             Disliv(k, j - 1) = quotasost(k, j) - quotasost(k, j - 1)
             Corda(k, j - 1) = Sqr(Lorizz(k, j - 1) ^ 2 + Disliv(k, j - 1) ^ 2)
             nc(k) = nc(k) + 1
      Wend
      
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
             nc(k) = nc(k) + 1
      Wend
      
      'calcolo delle quote terreno ed altezze verticali sostegni
      'salita
      If nbs > 0 Then
         For j = 1 To nc(1) + 1
                   w = 1
                   While prog_terr(1, w) < progsost(1, j)
                                                     w = w + 1
                   Wend
                   If prog_terr(1, w - 1) = progsost(1, j) Then
                                         Sheets("F03").Cells(14 + j, 5) = quot_terr(1, w - 1)
                                         Sheets("F03").Cells(14 + j, 6) = quotasost(1, j) - quot_terr(1, w - 1)
                   Else
                        p_delta = prog_terr(1, w) - prog_terr(1, w - 1)
                        q_delta = quot_terr(1, w) - quot_terr(1, w - 1)
                        q_terrpalo = quot_terr(1, w - 1) + (progsost(1, j) - prog_terr(1, w - 1)) * q_delta / p_delta
                        Sheets("F03").Cells(14 + j, 5) = q_terrpalo
                        Sheets("F03").Cells(14 + j, 6) = quotasost(1, j) - q_terrpalo
                   End If
         Next j
      End If
      'discesa
      If nbd > 0 Then
        If nc(1) <> nc(2) Then
           For j = 1 To nc(2) + 1
                   w = 1
                   While prog_terr(2, w) < progsost(2, j)
                                                    w = w + 1
                   Wend
                   If prog_terr(2, w - 1) = progsost(2, j) Then
                                         Sheets("F03").Cells(14 + j, 12) = quot_terr(2, w - 1)
                                         Sheets("F03").Cells(14 + j, 13) = quotasost(1, j) - quot_terr(2, w - 1)
                   Else
                        p_delta = prog_terr(2, w) - prog_terr(2, w - 1)
                        q_delta = quot_terr(2, w) - quot_terr(2, w - 1)
                        q_terrpalo = quot_terr(2, w - 1) + (progsost(2, j) - prog_terr(2, w - 1)) * q_delta / p_delta
                        Sheets("F03").Cells(14 + j, 12) = q_terrpalo
                        Sheets("F03").Cells(14 + j, 13) = quotasost(2, j) - q_terrpalo
                   End If
           Next j
        End If
      End If
      
      Sheets("F03").Select
      Sheets("F03").Range("C15").Select

End Sub


Private Sub Worksheet_Activate()
        nome_foglio = "RELATIVO"
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
       
       On Local Error GoTo err_range_F04
       
       ntower = "": numsost = ""
       numriga = Target.Row
       numcolonna = Target.Column
       ntower = Sheets("F04").Cells(numriga, numcolonna)
       If IsNumeric(ntower) Then
                                numsost = Format(ntower, "###")
       Else
             numsost = ntower
       End If
       
   If Sheets("HOME").flag_italiano.value Then
             If numcolonna < 6 Then
                                   rramo$ = Sheets("F04").Cells(203, 11)
             Else
                   rramo$ = Sheets("F04").Cells(204, 11)
             End If
             If (numriga > 14 And numriga < 94) And (numcolonna = 2 Or numcolonna = 6) Then
                     testo = Sheets("F04").Cells(208, 11) + " = " + numsost + "  " + rramo$ + " : " + Sheets("F04").Cells(205, 11)
                     riga_relativa.Caption = testo
                     cek_ins_del = True
             Else
                     testo = Sheets("F04").Cells(208, 11) + "  " + rramo$ + " : " + Sheets("F04").Cells(206, 11)
                     riga_relativa.Caption = testo
                     cek_ins_del = False
             End If
   Else
             If numcolonna < 6 Then
                                   rramo$ = Sheets("F04").Cells(218, 11)
             Else
                  rramo$ = Sheets("F04").Cells(219, 11)
             End If
             If (numriga > 14 And numriga < 94) And (numcolonna = 2 Or numcolonna = 6) Then
                     testo = Sheets("F04").Cells(223, 11) + " = " + numsost + "  " + rramo$ + " : " + Sheets("F04").Cells(220, 11)
                     riga_relativa.Caption = testo
                     cek_ins_del = True
             Else
                     testo = Sheets("F04").Cells(223, 11) + "  " + rramo$ + " : " + Sheets("F04").Cells(221, 11)
                     riga_relativa.Caption = testo
                     cek_ins_del = False
             End If
  End If
  
       Exit Sub
       
err_range_F04:
          Beep
          If Sheets("HOME").flag_italiano.value Then
                testo = Sheets("F04").Cells(207, 11)
                MsgBox (testo)
          Else
                testo = Sheets("F04").Cells(222, 11)
                MsgBox (testo)
          End If
          
          Exit Sub
End Sub
