Attribute VB_Name = "Modulo6"

Sub posizione_veicoli(spz)


'----------------------< inizio procedura geoveicoli >---------------------
'determinazione posizione vetture in linea RAMO SALITA
'spz     = spiazzamento iniziale del primo treno dall'avanstazione a valle
'Lanello =lunghezza sviluppata dalle corde delle campate + giri ruota
'Eg      = equidistanza dei grappoli/equidistanza veicoli
'Ng      = numero dei grappoli per ramo/numero di veicoli in linea
'Nvg     = numero dei veicoli per grappolo
'Ev      = equidistanza dei veicoli nel grappolo
'Drt     = diametro puleggia motrice
'Drtrin  = diametro ruota di rinvio
'Pr(k,i) = progressiva della vettura i sul ramo k (salita=1/discesa=2)
'I(k)    = contatore dei veicoli totali per ramo
'Vuoto   = tratto di fune libera tra due grappoli consecutivi
geoveicoli:
         ii% = 1
         Vuoto = eg - (Nvg - 1) * Ev
         If spz <= Vuoto Then Pr(1, ii%) = spz: Lneg = 0: GoTo ciclos
         Lis = spz - Vuoto: Lneg = ((Nvg - 1) * Ev - Lis)
         Ne = Int(Lis / Ev): If Ne > 0 Then Ne = Ne - 1
         Lio = Lis - Ne * Ev
         Pr(1, ii%) = Lio
         For X = 1 To Ne
         ii% = ii% + 1
         Pr(1, ii%) = Pr(1, ii% - 1) + Ev
         Next X
         ii% = ii% + 1
         Pr(1, ii%) = Pr(1, ii% - 1) + Vuoto
ciclos:
         For X = 1 To Nvg - 1
         ii% = ii% + 1
         Pr(1, ii%) = Pr(1, ii% - 1) + Ev
         If Pr(1, ii%) > Lsvm1 Then i(1) = ii% - 1: GoTo discesa
         Next X
         ii% = ii% + 1
         Pr(1, ii%) = Pr(1, ii% - 1) + Vuoto
         If Pr(1, ii%) > Lsvm1 Then
                                   i(1) = ii% - 1
                                   GoTo discesa
         End If
         GoTo ciclos

'determinazione posizione vetture in linea RAMO DISCESA
'Lgsv    = lunghezza del giro stazione a valle da avanstaz. ad avanstazione
'Lid     = distanza prima vettura-avanstazione sulla discesa
'Lidf    = distanza avanstazione-ultima vettura del grappolo sulla discesa

discesa:
         ii% = 1
         If Lneg = 0 Then
         Lid = eg - spz - PI# * Drt / 1000 / 2 - (Nvg - 1) * Ev
             If Lid >= 0 Then
                      Pr(2, ii%) = Lid
                      GoTo ciclod
             End If
         End If
         Lidf = eg - spz - PI# * Drt / 1000 / 2
         If Lidf < 0 Then
                      Pr(2, ii%) = Vuoto + Lidf
                      GoTo ciclod
         End If
         Ne = Int(Lidf / Ev): If Ne > 0 Then Ne = Ne - 1
         Liod = Lidf - Ne * Ev
         Pr(2, ii%) = Liod
         For X = 1 To Ne
         ii% = ii% + 1
         Pr(2, ii%) = Pr(2, ii% - 1) + Ev
         Next X
         ii% = ii% + 1
         Pr(2, ii%) = Pr(2, ii% - 1) + Vuoto
ciclod:
         For X = 1 To Nvg - 1
         ii% = ii% + 1
         Pr(2, ii%) = Pr(2, ii% - 1) + Ev
         If Pr(2, ii%) > Lsvm2 Then i(2) = ii% - 1: Return
         Next X
         ii% = ii% + 1
         Pr(2, ii%) = Pr(2, ii% - 1) + Vuoto
         If Pr(2, ii%) > Lsvm2 Then
                                   i(2) = ii% - 1
                                   GoSub disegna_veicoli
                                   Exit Sub
         End If
         GoTo ciclod
         
Exit Sub

disegna_veicoli:
         
        'disegna veicoli salita discesa
        If Sheets("F05").ceck_avanzamento.value Then
             'assegnazione progressiva alla tabella x grafico
             '-----------------------------------------------------------------------
             'sostegni in linea : salita
             '-----------------------------------------------------------------------
             k = 1: rr = 60: cc = 20
             For jj = 0 To ns
                      Sheets("F05").Cells(rr, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr + 1, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr + 2, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr + 3, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr, cc + 1) = 0
                      Sheets("F05").Cells(rr + 1, cc + 1) = -0.5
                      Sheets("F05").Cells(rr + 3, cc + 1) = 0:
                      rr = rr + 4
             Next jj
             Sheets("F05").Cells(rr, cc) = 0
             Sheets("F05").Cells(rr, cc + 1) = 0
             rr = rr + 1
             Sheets("F05").Cells(rr, cc) = 0
             Sheets("F05").Cells(rr, cc + 1) = -0.5
             rr = rr + 1
             '------------------------------------------------------------------------
             'carichi sul ramo salita
             '-----------------------------------------------------------------------
                For yy = 1 To i(1)
                      Sheets("F05").Cells(rr, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 1, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 2, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 3, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 4, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr, cc + 2) = -0.5
                      Sheets("F05").Cells(rr + 1, cc + 2) = -0.6
                      Sheets("F05").Cells(rr + 2, cc + 2) = -0.5
                      Sheets("F05").Cells(rr + 3, cc + 2) = -0.4
                      Sheets("F05").Cells(rr + 4, cc + 2) = -0.5
                      rr = rr + 5
                Next yy
                
             '-----------------------------------------------------------------------
             'sostegni in linea : discesa
             '-----------------------------------------------------------------------
             k = 2:  cc = 20
             For jj = 0 To nd
                      Sheets("F05").Cells(rr, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr + 1, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr + 2, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr + 3, cc) = Prs(k, jj)
                      Sheets("F05").Cells(rr, cc + 1) = 0
                      Sheets("F05").Cells(rr + 2, cc + 1) = 0.5
                      Sheets("F05").Cells(rr + 3, cc + 1) = 0:
                      rr = rr + 4
             Next jj
             Sheets("F05").Cells(rr, cc) = 0
             Sheets("F05").Cells(rr, cc + 1) = 0
             rr = rr + 1
             Sheets("F05").Cells(rr, cc) = 0
             Sheets("F05").Cells(rr, cc + 1) = 0.5
             rr = rr + 1
             '------------------------------------------------------------------------
             'carichi sul ramo discesa
             '-----------------------------------------------------------------------
                For yy = 1 To i(2)
                      Sheets("F05").Cells(rr, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 1, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 2, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 3, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr + 4, cc) = Pr(k, yy)
                      Sheets("F05").Cells(rr, cc + 2) = 0.5
                      Sheets("F05").Cells(rr + 1, cc + 2) = 0.6
                      Sheets("F05").Cells(rr + 2, cc + 2) = 0.5
                      Sheets("F05").Cells(rr + 3, cc + 2) = 0.4
                      Sheets("F05").Cells(rr + 4, cc + 2) = 0.5
                      rr = rr + 5
                Next yy
                'Application.Wait (Now + TimeValue("0:00:01"))
                Sheets("F05").ChartObjects(1).Select
        End If
        Return

End Sub

