Attribute VB_Name = "Modulo1"



Public mia_path As String
Public tipo_impianto, tstringa, fgl, lingua As String
Public test_ricalcolo As Boolean

Public Const PI# = 3.141
Public Const G# = 0.981
Public Const KTERMICO# = 0.000012

'variabili comuni al calcolo di linea (Calcolo.BI)
Public scelta%, kk%, xx%, ipoipo, descrizione$, spool$, Aze$, Nazione$
Public ramo%, nomeramo$, mask$, vval$, fPress
Public workarea$, versione$
Public XIndex As Integer
Public cella_num As Integer
Public rr13, cc13, rr12, cc12, rr6, cc6, rr5, cc5 As Integer

Public Av(), Am(), Tcv(), Tcm(), Tp()
Public Tm0(), Sv0()
Public Dp(), Pp(), attrito()
Public Bo(), Pr(), ncnc(), alfa(), PV()
Public i()
Public maxps(), minps()
Public Fm(), s()
Public FV(), FF(), Allung()
Public Hc(), Fvc(), Fmmax(), Fmmin(), Fvt()
Public maxav(), maxam(), minav(), minam(), maxdev()
Public mindev(), maxtv(), maxtm(), mintv(), mintm()
Public maxattr(), minattr(), maxtp(), mintp(), maxsq()
Public corsaz(14, 3)


Public impianto$, port, porteff, vel, posti, tipof$, dia, qv
Public m$, tt$, contr, ZETA, ng, scart, flagdiscesa, buco, flagvento
Public morse, areac, areav, acc, dec, fren, rap, frinv, vento2, vento1
Public argano, modulo, scansione, area, mf, fili, res, dias, ms, ms1, ds1
Public srott, areas, rullier$, mapp, mrit, dapp, drit, diaapp, diarit
Public cmaxapp, cmaxrit, devmaxapp, devmaxrit, flagfinecorsa
Public attr0, pstazione, add, dd, qc, tsegn, ymin, ymax, vg, eg, persona, attr1, eg0, hg0, nveic

Public nc(2)
Public din(6), attr(6), qs(6), qd(6)
Public ipotesi$(30), modo$(6), ipo$(2), nomeipotesi$
Public l(2, 80), d(2, 80), Prs(2, 80), tat$(2, 80), atreg(2, 80), atfre(2, 80)
Public r1(80), r2(80)
Public quota(2, 80), prog(2, 80), hv(2, 80), rulli(2, 80), camp$(2, 80)
Public n$(2, 80), SiNo%(15, 6)
Public truotas, truotad, corsa
Public path$, nome$, spziniz, spzfin, passo, elimina$, ns, nd
Public Nvg, Ev, fflagdiscesa$
Public rendarg
Public aa$, tipoattr$, tipocontr$, testdevo$, rrulli$, Drt, Drtrin
Public a, car, Qf, lanello, Lsvm1, Lsvm2
Public corsacil
Public autorulli
Public DEtiprul$, DEdiamrul, DEmassarul, DEcarmax
Public tipoveic$, sensomarcia$, Intstaz
Public descagg$, lungpass, tipotiro$
Public Ipocalc
Public Ftrave

Public nbs, nbd As Integer
Public prog_terr(2, 800), quot_terr(2, 800) As Single

Public numsost, ntower As String
Public numriga, numcolonna As Integer

Public cek_ins_del, cek_ins_del_d
Public nomesost(2, 80) As String
Public progsost(2, 80), quotasost(2, 80) As Single
Public nomecamp(2, 80), Lorizz(2, 80), Disliv(2, 80), Corda(2, 80) As Single

Public Sub salti_pagina(fgl)

On Local Error GoTo err_salta

For Each interrpag In Worksheets(fgl).HPageBreaks
    If interrpag.Extent = xlPageBreakFull Then
        cIntere = cIntere + 1
    Else
        cParziali = cParziali + 1
    End If
Next
'MsgBox cIntere & " interruzioni di pagina di schermo intero, " & cParziali & _
    " interruzioni di pagina di area di stampa"
For j% = 1 To cIntere
        Worksheets(fgl).HPageBreaks(j%).Delete
Next j%
For j% = 1 To cParziali
        Worksheets(fgl).HPageBreaks(j%).Delete
Next j%

Exit Sub

err_salta:
         Beep
End Sub

Public Sub leggi_terreno()
       'controlla profilo terreno ramo salita
       nbs = 1
       w = 15
       j = 4
       k = 5
       prog_terr(1, nbs) = Sheets("F02").Cells(w, j)
       quot_terr(1, nbs) = Sheets("F02").Cells(w, k)
       While prog_terr(1, nbs) + quot_terr(1, nbs) <> 0
                  nbs = nbs + 1
                  w = w + 1
                  prog_terr(1, nbs) = Sheets("F02").Cells(w, j)
                  quot_terr(1, nbs) = Sheets("F02").Cells(w, k)
       Wend
       nbs = nbs - 1
       If nbs < 2 Then
                      nbd = 0
                      Beep
                      MsgBox ("Mancanza profilo terreno")
                      Exit Sub
       End If
       'controlla profilo terreno ramo discesa
       nbd = 1
       w = 15
       j = 9
       k = 10
       prog_terr(2, nbd) = Sheets("F02").Cells(w, j)
       quot_terr(2, nbd) = Sheets("F02").Cells(w, k)
       While prog_terr(2, nbd) + quot_terr(2, nbd) <> 0
                  nbd = nbd + 1
                  w = w + 1
                  prog_terr(2, nbd) = Sheets("F02").Cells(w, j)
                  quot_terr(2, nbd) = Sheets("F02").Cells(w, k)
       Wend
       nbd = nbd - 1
End Sub


Public Sub leggi_linea()
                      
           'lettura ramo salita
           j% = 1
           riga% = 15
           quota(1, j%) = Sheets("F03").Cells(riga%, 7)
           prog(1, j%) = Sheets("F03").Cells(riga%, 4)
           n$(1, j%) = Sheets("F03").Cells(riga%, 3)
           hv(1, j%) = Sheets("F03").Cells(riga%, 6)
           rulli(1, j%) = CDbl(Sheets("F03").Cells(riga%, 8)): Sheets("F03").Cells(riga%, 8) = rulli(1, j%)
           rulli(2, j%) = CDbl(Sheets("F03").Cells(riga%, 15)): Sheets("F03").Cells(riga%, 15) = rulli(2, j%)
           r1(j%) = rulli(1, j%)
           camp$(1, j%) = " "
           riga% = riga% + 1
           While Val(Sheets("F03").Cells(riga%, 4)) <> 0 And Val(Sheets("F03").Cells(riga%, 7)) <> 0
                      j% = j% + 1
                      quota(1, j%) = Sheets("F03").Cells(riga%, 7)
                      prog(1, j%) = Sheets("F03").Cells(riga%, 4)
                      d(1, j% - 1) = quota(1, j%) - quota(1, j% - 1)
                      l(1, j% - 1) = prog(1, j%) - prog(1, j% - 1)
                      n$(1, j%) = Sheets("F03").Cells(riga%, 3)
                      hv(1, j%) = Sheets("F03").Cells(riga%, 6)
                      rulli(1, j%) = CDbl(Sheets("F03").Cells(riga%, 8)): Sheets("F03").Cells(riga%, 8) = rulli(1, j%)
                      r1(j%) = rulli(1, j%)
                      rulli(2, j%) = CDbl(Sheets("F03").Cells(riga%, 15)): Sheets("F03").Cells(riga%, 15) = rulli(2, j%)
                      camp$(1, j%) = n$(1, j% - 1) + " - " + n$(1, j%)
                      riga% = riga% + 1
           Wend
           ns = j%
           
           'lettura ramo discesa
           'controllo se celle ramo discesa sono riempite
           j% = 2: riga% = 15
           prog(2, j%) = Sheets("F03").Cells(16, 11)
           If Val(prog(2, j%)) = 0 Then
               'salita = discesa
               nd = ns
               For j% = 1 To nd
                      quota(2, j%) = quota(1, j%)
                      prog(2, j%) = prog(1, j%)
                      d(2, j%) = d(1, j%)
                      l(2, j%) = l(1, j%)
                      n$(2, j%) = n$(1, j%)
                      hv(2, j%) = hv(1, j%)
                      If rulli(2, j%) = 0 Then rulli(2, j%) = rulli(1, j%): Sheets("F03").Cells(riga%, 15) = rulli(2, j%)
                      r2(j%) = rulli(2, j%)
                      camp$(2, j%) = camp$(1, j%)
                      riga% = riga% + 1
               Next j%
           Else
               j% = 1
               riga% = 15
               quota(2, j%) = Sheets("F03").Cells(riga%, 14)
               prog(2, j%) = Sheets("F03").Cells(riga%, 11)
               d(2, j%) = 0
               l(2, j%) = 0
               n$(2, j%) = Sheets("F03").Cells(riga%, 10)
               hv(2, j%) = Sheets("F03").Cells(riga%, 13)
               rulli(2, j%) = CDbl(Sheets("F03").Cells(riga%, 15)): Sheets("F03").Cells(riga%, 15) = rulli(2, j%)
               r2(j%) = rulli(2, j%)
               camp$(2, j%) = " "
               riga% = riga% + 1
               While CDbl(Sheets("F03").Cells(riga%, 4)) <> 0 And CDbl(Sheets("F03").Cells(riga%, 7)) <> 0
                      j% = j% + 1
                      quota(2, j%) = Sheets("F03").Cells(riga%, 14)
                      prog(2, j%) = Sheets("F03").Cells(riga%, 11)
                      d(2, j% - 1) = quota(2, j%) - quota(2, j% - 1)
                      l(2, j% - 1) = prog(2, j%) - prog(2, j% - 1)
                      n$(2, j%) = Sheets("F03").Cells(riga%, 10)
                      hv(2, j%) = Sheets("F03").Cells(riga%, 13)
                      rulli(2, j%) = CDbl(Sheets("F03").Cells(riga%, 15)): Sheets("F03").Cells(riga%, 15) = rulli(2, j%)
                      r2(j%) = rulli(2, j%)
                      camp$(2, j%) = n$(2, j% - 1) + " - " + n$(2, j%)
                      riga% = riga% + 1
               Wend
               nd = j%
           End If
           
           'generazione tabella attriti di default
           riga% = 12
           For j% = 2 To ns - 1
                   If Val(Sheets("F07").Cells(riga%, 4)) = 0 Then
                                       Sheets("F07").Cells(riga%, 2) = n$(1, j%)
                                       Sheets("F07").Cells(riga%, 3) = "%"
                                       Sheets("F07").Cells(riga%, 4) = Sheets("F07").Cells(4, 9)
                                       Sheets("F07").Cells(riga%, 5) = Sheets("F07").Cells(5, 9)
                   End If
                   riga% = riga% + 1
           Next j%
           riga% = 12
           For j% = 2 To nd - 1
                   If Val(Sheets("F07").Cells(riga%, 9)) = 0 Then
                                       Sheets("F07").Cells(riga%, 7) = n$(2, j%)
                                       Sheets("F07").Cells(riga%, 8) = "%"
                                       Sheets("F07").Cells(riga%, 9) = Sheets("F07").Cells(4, 9)
                                       Sheets("F07").Cells(riga%, 10) = Sheets("F07").Cells(5, 9)
                   End If
                   riga% = riga% + 1
           Next j%
          
End Sub


Public Sub assegna_generali()
                
        Dim err_help As String
        Dim nrig, ncol, num_err As Integer
        
        'assegna da opzioni
        If Sheets("F01").ancorato_fisso.value Then tipotiro$ = "A"
        If Sheets("F01").argano_mon.value Then m$ = "MONTE"
        If Sheets("F01").argano_val.value Then m$ = "VALLE"
        If Sheets("F01").F.value Then ZETA = "F"
        If Sheets("F01").i.value Then ZETA = "I"
        If Sheets("F01").in_tensione.value Then tipotiro$ = "T"
        If Sheets("F01").rotaz_antior.value Then sensomarcia$ = Sheets("F01").rotaz_antior.Caption
        If Sheets("F01").rotaz_or.value Then sensomarcia$ = Sheets("F01").rotaz_or.Caption
        If Sheets("F01").s.value Then ZETA = "S"
        If Sheets("F01").T.value Then ZETA = "T"
        If Sheets("F01").tensione_mon.value Then tt$ = "MONTE"
        If Sheets("F01").tensione_val.value Then tt$ = "VALLE"
        If Sheets("F01").tipo_tend_grav.value Then tipotensione$ = "G"
        If Sheets("F01").tipo_tend_idra.value Then tipotensione$ = "I"
        tipocontr$ = tipotensione$
        If Sheets("F01").car_salita.value Then flagdiscesa = 0
        If Sheets("F01").car_salita_discesa.value Then flagdiscesa = 1
        If Sheets("F01").test_vento_fe_si.value Then flagvento = 1
        If Sheets("F01").test_vento_fe_no.value Then flagvento = 0
        
        attr0 = Sheets("F01").Cells(53, 5): If attr0 = 0 Then num_err = 2: GoSub test_err
        attr1 = Sheets("F01").Cells(54, 5): If attr0 = 0 Then num_err = 3: GoSub test_err
        If Sheets("F07").Cells(4, 9) = 0 Then Sheets("F07").Cells(4, 9) = attr0
        If Sheets("F07").Cells(5, 9) = 0 Then Sheets("F07").Cells(5, 9) = attr1
        
        nome$ = Sheets("F01").Cells(3, 4)
        impianto$ = Sheets("F01").Cells(4, 4)
        port = Sheets("F01").Cells(41, 5): If port = 0 Then num_err = 1: GoSub test_err
        vel = Sheets("F01").Cells(42, 5): If vel = 0 Then num_err = 4: GoSub test_err
        posti = Sheets("F01").Cells(43, 5): If posti = 0 Then num_err = 5: GoSub test_err
        persona = Sheets("F01").Cells(44, 5): If persona = 0 Then num_err = 6: GoSub test_err
        tipof$ = Sheets("F01").Cells(7, 4)
        dia = Sheets("F01").Cells(8, 5)
        qv = Sheets("F01").Cells(45, 5)
        qc = Sheets("F01").Cells(46, 5)
        contr = Sheets("F01").Cells(40, 5)
        ng = Sheets("F01").Cells(32, 10)
        vg = Sheets("F01").Cells(33, 10)
        eg = Sheets("F01").Cells(34, 10)
        descagg$ = Sheets("F01").Cells(5, 4)
        lungpass = 0
        
        scart = Sheets("F01").Cells(59, 5)
        buco = Sheets("F01").Cells(39, 5)
        morse = Sheets("F01").Cells(61, 5)
        areac = Sheets("F01").Cells(65, 5)
        areav = Sheets("F01").Cells(64, 5)
        acc = Sheets("F01").Cells(47, 5)
        dec = Sheets("F01").Cells(48, 5)
        fren = Sheets("F01").Cells(49, 5)
        rap = Sheets("F01").Cells(50, 5)
        frinv = 0
        vento2 = Sheets("F01").Cells(66, 5)
        vento1 = Sheets("F01").Cells(67, 5)
        argano = Sheets("F01").Cells(51, 5)
        rendarg = Sheets("F01").Cells(52, 5)
        modulo = Sheets("F01").Cells(13, 5)
        
        area = Sheets("F01").Cells(9, 5)
        mf = Sheets("F01").Cells(10, 5)
        fili = Sheets("F01").Cells(11, 5)
        res = Sheets("F01").Cells(12, 5)
        dias = Sheets("F01").Cells(24, 5)
        ds1 = Sheets("F01").Cells(27, 5)
        ms = Sheets("F01").Cells(26, 5)
        ms1 = 0  ' calcolare
        srott = Sheets("F01").Cells(28, 5)
        areas = Sheets("F01").Cells(25, 5)
        
        rullier$ = Sheets("F01").Cells(68, 5)
        mapp = Sheets("F01").Cells(69, 5)
        mrit = Sheets("F01").Cells(70, 5)
        dapp = Sheets("F01").Cells(71, 5)
        drit = Sheets("F01").Cells(72, 5)
        diaapp = dapp
        diarit = drit
        cmaxapp = Sheets("F01").Cells(75, 5)
        cmaxrit = Sheets("F01").Cells(76, 5)
        devmaxapp = Sheets("F01").Cells(73, 5)
        devmaxrit = Sheets("F01").Cells(74, 5)
        pstazione = Sheets("F01").Cells(55, 5)
        Ftrave = Sheets("F01").Cells(56, 5)
        DEtiporul$ = Sheets("F01").Cells(77, 5)
        DEdiamrul = Sheets("F01").Cells(78, 5)
        DEmassarul = Sheets("F01").Cells(79, 5)
        DEcarmax = Sheets("F01").Cells(80, 5)
        tipoveic$ = Sheets("F01").Cells(60, 5)
        Intstaz = Sheets("F01").Cells(57, 5)
        Drt = Sheets("F01").Cells(57, 5)
        Drtrin = Sheets("F01").Cells(58, 5)
        
        dd = 3600 / port * posti * vel
        
        ttel = tsegn
        Qf = mf * G#
        If ZETA = "I" Then
                        Nvg = vg
                        Ev = eg
        Else
            Nvg = 1
            Ev = dd
        End If
        
        Exit Sub
        
test_err:
            Beep
            Sheets("F01").Select
            Select Case num_err
                   Case 1
                        err_help = "portata = 0"
                        nrig = 41: ncol = 5
                   Case 2
                        err_help = "Attrito e regime = 0"
                        nrig = 53: ncol = 5
                   Case 3
                        err_help = "attrito in frenatura = 0"
                        nrig = 54: ncol = 5
                   Case 4
                        err_help = "velocità di esercizio = 0"
                        nrig = 42: ncol = 5
                   Case 5
                        err_help = "numero posti veicolo = 0"
                        nrig = 43: ncol = 5
                   Case 6
                        err_help = "peso persona = 0"
                        nrig = 44: ncol = 5
                        
                        
                        
            End Select
            MsgBox ("Warning : " & err_help)
            Sheets("F01").Cells(nrig, ncol).Select

End Sub

Public Sub virgola_to_punto(tstringa)

        Dim lstringa As Integer
        lstringa = Len(tstringa)
        If lstringa > 0 Then
           If mid(tstringa, lstringa, 1) = "," Then
                          Mid(tstringa, lstringa, 1) = "."
           End If
        End If

End Sub

Public Sub esegui_verifica(XIndex As Integer)

   
        Dim nsost, cval, cmon As String
        Dim tvv, tmm, ffcc, avv, amm, aa1, aa2 As Double
        Dim saltot As Integer
        Dim stato_prg As String
        
        Ins% = ns + 4 + 2   'massimo numero di sostegni per ramo salita
        nr% = 2              'massimo numero di rami impianto
        u% = 40             'massimo veicoli in campata
        T% = 400            'massimo numero di veicoli per ramo di impianto
        '----------------------------dim shared  ARRAY---------------------------------------
        ReDim Av(nr%, Ins%), Am(nr%, Ins%), Tcv(nr%, Ins%), Tcm(nr%, Ins%), Tp(nr%, Ins%)
        ReDim Tm0(nr%, Ins%), Sv0(nr%, Ins%)
        ReDim Dp(nr%, Ins%), Pp(nr%, Ins%), attrito(nr%, Ins%)
        ReDim Bo(nr%, Ins%, u%), Pr(nr%, T%), ncnc(nr%, Ins%), alfa(nr%, Ins%), PV(nr%, Ins%)
        ReDim i(nr%)
        ReDim maxps(nr%, Ins%), minps(nr%, Ins%)
        ReDim Fm(nr%, Ins%), s(nr%, Ins%)
        ReDim FV(u%), FF(u%), Allung(nr%, Ins%)
        ReDim Hc(nr%, Ins%), Fvc(u%), Fmmax(nr%, Ins%), Fmmin(nr%, Ins%), Fvt(u%)
        ReDim maxav(nr%, Ins%), maxam(nr%, Ins%), minav(nr%, Ins%), minam(nr%, Ins%), maxdev(nr%, Ins%)
        ReDim mindev(nr%, Ins%), maxtv(nr%, Ins%), maxtm(nr%, Ins%), mintv(nr%, Ins%), mintm(nr%, Ins%)
        ReDim maxattr(nr%, Ins%), minattr(nr%, Ins%), maxtp(nr%, Ins%), mintp(nr%, Ins%), maxsq(Ins%)

        path$ = "C:\TMP_SIF\"
        nome$ = "valori"
        'stato_prg = Sheets("F05").in_esecuzione.Caption
        Sheets("F05").Cells(26, 6) = ""
        Sheets("F05").in_esecuzione.Visible = True
        
        If Sheets("F05").Option1_maxmin.value Then
                                                    'vedi massimi e minimi
                                                    flagmed% = 0
        Else
             'vedi valori medi
             flagmed% = 9
        End If
        
        'leggi la matrice ipotesi di calcolo (SiNo%(15,6)
        For numr = 10 To 24
              For numc = 4 To 9
                   kk% = numr - 9
                   xx% = numc - 3
                   SiNo%(kk%, xx%) = 0
                   If Cells(numr, numc) = "OOO" Then
                                                    SiNo%(kk%, xx%) = 1
                   End If
                   If Cells(numr, numc) = "XXX" Then SiNo%(kk%, xx%) = 2
              Next numc
        Next numr
        
        If ZETA = "F" Then
                    eg = Sheets("F05").Cells(5, 15)
                    Ev = eg
        Else
            If ZETA <> "I" Then eg = dd: Ev = eg
        End If
                        
        'lettura passo di spostamento veicoli e spiazzamenti iniziale e finale
        'controllo valore del passo
        avanzamento = Sheets("F05").step_car.value
        spziniz = 0
        spzfin = 0
        
        If ZETA = "I" Then
                        passo = Ev
                        spzfin = eg
                        Sheets("F05").Cells(10, 15) = Format(passo, "###0.00")
        Else
              passo = eg / avanzamento: Sheets("F05").Cells(10, 15) = Format(passo, "###0.00")
              spzfin = eg
        End If
        
        
        'possibilità di calcolare una sola posizione dei veicoli
        '-------------------------------------------------------
        If passo = -1 Then
                          flagmed% = 9
                          passo = 1
                          spzfin = spziniz
        End If
        '-------------------------------------------------------
        
        lll% = Int((spzfin - spziniz) / passo) + 1
        If lll% > 800 Then
                        Beep
                        MsgBox ("Error : n°step > 800")
                        Exit Sub
        End If
        k% = 0
        precalc% = 0
'--------------------------< INIZIO DEL CICLO DI VERIFICA LINEA >------------
        If XIndex = 1 Then
                        'precalcolo tensione idraulica
                        precalc% = 1
                        GoSub precalcolo
                        precalc% = 0
                        XIndex = 1

                        For zzz% = 5 To 7
                            For uu% = 1 To 3
                                  corsaz(zzz%, uu%) = corsaz(zzz% - 4, uu%) - corsaz(zzz% - 3, uu%)
                            Next uu%
                        Next zzz%
                        Load idraulic_input
                        
                        idraulic_input.TextBox1.Text = 30: saltot = 30
                        corsaz(8, 1) = KTERMICO# * lanello / 2 * saltot
                        corsaz(9, 1) = 0.3
                        corsaz(10, 1) = 0.1
                        
                        If ZETA <> "F" Then
                                corsaz(11, 1) = corsaz(5, 1) + corsaz(6, 1) + corsaz(7, 1) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                                corsaz(12, 1) = corsaz(5, 2) + corsaz(6, 2) + corsaz(7, 2) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                                corsaz(13, 1) = corsaz(5, 3) + corsaz(6, 3) + corsaz(7, 3) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                        Else
                                corsaz(11, 1) = corsaz(6, 1) + corsaz(7, 1) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                                corsaz(12, 1) = corsaz(6, 2) + corsaz(7, 2) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                                corsaz(13, 1) = corsaz(6, 3) + corsaz(7, 3) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                        End If
                        maxcorsa = 0
                        If corsaz(11, 1) > maxcorsa Then maxcorsa = corsaz(11, 1)
                        If corsaz(12, 1) > maxcorsa Then maxcorsa = corsaz(12, 1)
                        If corsaz(13, 1) > maxcorsa Then maxcorsa = corsaz(13, 1)
                        corsaz(14, 1) = Int(corsaz(11, 1) * 10) / 10 + 0.1
                        
                        mask$ = "####0.00"
                        For zzz% = 1 To 7
                           For ii% = 1 To 3
                               uu% = (zzz% - 1) * 3 + ii%
                               tstringa = Format(corsaz(zzz%, ii%), mask$)
                               Call codifica(tstringa)
                               Select Case uu%
                                              Case 1
                                                     idraulic_input.Label13.Caption = tstringa
                                              Case 2
                                                     idraulic_input.Label14.Caption = tstringa
                                              Case 3
                                                     idraulic_input.Label15.Caption = tstringa
                                              Case 4
                                                     idraulic_input.Label16.Caption = tstringa
                                              Case 5
                                                     idraulic_input.Label17.Caption = tstringa
                                              Case 6
                                                     idraulic_input.Label18.Caption = tstringa
                                              Case 7
                                                     idraulic_input.Label19.Caption = tstringa
                                              Case 8
                                                     idraulic_input.Label20.Caption = tstringa
                                              Case 9
                                                     idraulic_input.Label21.Caption = tstringa
                                              Case 10
                                                     idraulic_input.Label22.Caption = tstringa
                                              Case 11
                                                     idraulic_input.Label23.Caption = tstringa
                                              Case 12
                                                     idraulic_input.Label24.Caption = tstringa
                                              Case 13
                                                     idraulic_input.Label25.Caption = tstringa
                                              Case 14
                                                     idraulic_input.Label26.Caption = tstringa
                                              Case 15
                                                     idraulic_input.Label27.Caption = tstringa
                                              Case 16
                                                     idraulic_input.Label28.Caption = tstringa
                                              Case 17
                                                     idraulic_input.Label29.Caption = tstringa
                                              Case 18
                                                     idraulic_input.Label30.Caption = tstringa
                                              Case 19
                                                     idraulic_input.Label31.Caption = tstringa
                                              Case 20
                                                     idraulic_input.Label32.Caption = tstringa
                                              Case 21
                                                     idraulic_input.Label33.Caption = tstringa
                               End Select
                           Next ii%
                        Next zzz%
                        tstringa = Format(corsaz(8, 1), mask$)
                        Call codifica(tstringa): idraulic_input.Label34.Caption = tstringa
                        tstringa = Format(corsaz(9, 1), mask$)
                        Call codifica(tstringa): idraulic_input.TextBox2.Text = tstringa
                        tstringa = Format(corsaz(10, 1), mask$)
                        Call codifica(tstringa): idraulic_input.TextBox3.Text = tstringa
                        tstringa = Format(corsaz(14, 1), mask$)
                        Call codifica(tstringa): idraulic_input.TextBox4.Text = tstringa
                        tstringa = Format(corsaz(11, 1), mask$)
                        Call codifica(tstringa): idraulic_input.Label35.Caption = tstringa
                        tstringa = Format(corsaz(12, 1), mask$)
                        Call codifica(tstringa): idraulic_input.Label36.Caption = tstringa
                        tstringa = Format(corsaz(13, 1), mask$)
                        Call codifica(tstringa): idraulic_input.Label37.Caption = tstringa

                        'picture1.Visible = 0
                        'If ZETA = "F" Then option2(1).Value = -1
                        Sheets("F05").Cells(26, 6) = ""
                        Sheets("F05").in_esecuzione.Visible = False
                        Exit Sub
        End If

        If XIndex = 2 Then
                        corsacil = Val(idraulic_input.TextBox4)
                        If idraulic_input.Opt1.value Then
                                         corsazero = Val(idraulic_input.Label13.Caption)
                        Else
                            corsazero = Val(idraulic_input.Label16.Caption)
                        End If
                        XIndex = 1
                        precalc% = 0
        End If

        If XIndex = 3 Then Close: Exit Sub

        If XIndex = 4 Then
                         'salto_termico
                         DeltaT = 0
                         KT = 0.000012
                         GoSub parametri
                         XIndex = 0
                         'fune nuda ferma : riferimento tensioni e sviluppi di posa
                         kk% = 5: xx% = 1
                         GoSub precalcsub
                         XIndex = 4
                         Svtot0 = 0
                         AEtot0 = 0
                         ATtot0 = 0
                         For j% = 1 To ns - 1
                                  'sviluppi catenarie salita
                                  Ci = Sqr(d(1, j%) ^ 2 + l(1, j%) ^ 2)
                                  Tm0(1, j%) = (Tcv(1, j%) + Tcm(1, j%)) / 2
                                  Sv0(1, j%) = Ci + s(1, j%)
                                  Svtot0 = Svtot0 + Sv0(1, j%)
                                  'allungamenti elastici salita
                                  AEtot0 = AEtot0 + Sv0(1, j%) * 10 * Tm0(1, j%) / modulo / area
                                  'allungamenti termici in salita
                                  ATtot0 = ATtot0 + KT * Sv0(1, j%) * DeltaT
                         Next j%
                         For j% = 1 To nd - 1
                                  'sviluppi catenarie discesa
                                  Ci = Sqr(d(2, j%) ^ 2 + l(2, j%) ^ 2)
                                  Tm0(2, j%) = (Tcv(2, j%) + Tcm(2, j%)) / 2
                                  Sv0(2, j%) = Ci + s(2, j%)
                                  Svtot0 = Svtot0 + Sv0(2, j%)
                                  'allungamenti elastici discesa
                                  AEtot0 = AEtot0 + Sv0(2, j%) * 10 * Tm0(2, j%) / modulo / area
                                  'allungamenti termici in salita
                                  ATtot0 = ATtot0 + KT * Sv0(2, j%) * DeltaT
                         Next j%
                         'calcolo della fune a terra
                         LRF_FUNE_A_TERRA = Svtot0 - AEtot0 - ATtot0
        End If

        If XIndex = 0 Or XIndex = 4 Or XIndex = 5 Then
            corsacil = 0
        End If

        test% = 0
        If XIndex = 2 Or XIndex = 1 Then test% = 3
        If XIndex <> 5 Then
                         GoSub eseguiver
        Else
            'tensione contrappeso +10%
            kcontrapp = 1 + Sheets("F05").Cells(17, 17) / 100
            test% = 1: GoSub eseguiver
            Ipocalc = 1: Call analisi
            'tensione contrappeso -10%
            kcontrapp = 1 - Sheets("F05").Cells(17, 17) / 100
            test% = 2: GoSub eseguiver
            Ipocalc = 3: Call analisi
        End If

        Close
        
        Call vedi_linea
        If XIndex <> 5 Then Call analisi
        
        Sheets("F05").Select
        Sheets("F05").Range("D10:I24").Select
        Selection.Copy
        Cells(1, 1).Select
        Sheets("F08").Select
        Sheets("F08").Range("G6:L20").Select
        ActiveSheet.Paste
        Cells(30, 2).Select
        Sheets("F09").Select
        Sheets("F09").Range("G6:L20").Select
        ActiveSheet.Paste
        Cells(27, 2).Select
        Sheets("F10").Select
        Sheets("F10").Range("G6:L20").Select
        ActiveSheet.Paste
        Cells(27, 2).Select
        Sheets("F11").Select
        Sheets("F11").Range("J6:O20").Select
        ActiveSheet.Paste
        Cells(27, 1).Select
        
        Sheets("F05").Cells(26, 6) = ""
        Sheets("F05").in_esecuzione.Visible = False
        'Beep
        'MsgBox ("FINE ELABORAZIONE")
        
        Exit Sub


eseguiver:
        GoSub parametri
        GoSub pulisci_geoveicoli
        
        aaa$ = "####"
        marcia$ = "AVANTI"
        If XIndex = 0 Or XIndex = 4 Then rr12 = 4
        If XIndex = 1 Then rr12 = 600
        If XIndex = 5 Then
           If test% = 1 Then
                            rr12 = 200
           Else
               rr12 = 400
           End If
        End If
        For kk% = 1 To 15
           For xx% = 1 To 6
               If SiNo%(kk%, xx%) > 0 Then
                                         cc12 = 187
                                         GoSub ciclocalcolo
               End If
           Next xx%
        Next kk%
        Close
        marcia$ = "INDIETRO"
        GoSub pulisci_geoveicoli
        If XIndex = 0 Or XIndex = 4 Then rr12 = 100
        If XIndex = 1 Then rr12 = 700
        If XIndex = 5 Then
           If test% = 1 Then
                            rr12 = 300
           Else
               rr12 = 500
           End If
        End If
        For kk% = 1 To 15
           For xx% = 2 To 6
               If SiNo%(kk%, xx%) > 0 Then
                                         cc12 = 187
                                         GoSub ciclocalcolo
               End If
           Next xx%
        Next kk%
        Return

'----------------< FINE DEL CICLO DI CALCOLO E MEMORIZZAZIONE >--------------
precalcolo:
        GoSub parametri
        aaa$ = "####"
        For zzz% = 1 To 14: For ii% = 1 To 3: corsaz(zzz%, ii%) = 0: Next ii%: Next zzz%
        'fune nuda ferma
        kk% = 5: xx% = 1: zzz% = 1: GoSub precalcsub

        'fune scarica ferma
        kk% = 2: xx% = 1: zzz% = 2: GoSub precalcsub

        'fune sal.carica/dis.scarica
        kk% = 1: xx% = 2: zzz% = 3: GoSub precalcsub

        If flagdiscesa = 1 Then
                             'fune sal.carica/dis.carica
                             kk% = 4: xx% = 2: zzz% = 4: GoSub precalcsub
        Else
            corsaz(4, 1) = corsaz(3, 1)
            corsaz(4, 2) = corsaz(3, 2)
            corsaz(4, 3) = corsaz(3, 3)
        End If
        Return

precalcsub:
        Close
        marcia$ = "AVANTI"
        kcontrapp = 1: GoSub ciclocalcolo
        corsaz(zzz%, 1) = (sviluppomax + sviluppomin) / 4
        kcontrapp = 1.1: GoSub ciclocalcolo
        corsaz(zzz%, 2) = (sviluppomax + sviluppomin) / 4
        kcontrapp = 0.9: GoSub ciclocalcolo
        corsaz(zzz%, 3) = (sviluppomax + sviluppomin) / 4
        Close
        Return

ciclocalcolo:
             If marcia$ = "AVANTI" Then test_marcia = "FORWARD GEAR" Else test_marcia = "REVERSE GEAR"
             tentativo% = 0
             If XIndex = 5 Or precalc% = 1 Then
                             Tinizio = contr * kcontrapp
             Else
                  Tinizio = contr
             End If
             nomeipotesi$ = ipotesi$(kk%): mmodo$ = modo$(xx%)
             ipoipo = (kk% - 1) * 6 + xx%
             Sheets("F05").in_esecuzione.Caption = "RUNIG PROGRAM " + " : " + test_marcia + " >" + "    Ipo: " + Format(ipoipo, "###") + "  " + nomeipotesi$ + "   " + mmodo$ + "..."
             GoSub creacalcolo
             GoSub risultati
             Return

creacalcolo:
             Do While 0 = 0
                       Select Case kk%
                                   Case 5
                                          GoSub nuda
                                   Case Is > 5
                                          ok% = 0
                                          For j% = 1 To ns - 1
                                                      PV(1, j%) = 0
                                          Next j%
                                          For j% = 1 To nd - 1
                                                      PV(2, j%) = 0
                                          Next j%
                                          'leggi carico imposto in linea ------------------------------------------
                                          'kk% = caso di carico > 5 ( i primi 5 sono automatici)
                                          'xx% = condizione di moto dll'impianto ( da 1 a 6 )
                                          'inzio riga dati carico imposto     (F06) = 12
                                          'inizio colonna dati carico imposto (F06) = 3
                                          '------------------------------------------------------------------------
                                          riga% = 11
                                          colonna% = 3 + (kk% - 6) * 2
                                          nw = Chr(10) + Chr(13)
                                          For j% = 1 To ns - 1
                                                    'controllo compatibilità valori numerici nelle celle
                                                    miavar = Sheets("F06").Cells(riga% + j%, colonna%)
                                                    If IsNumeric(miavar) Then
                                                             PV(1, j%) = CDbl(Sheets("F06").Cells(riga% + j%, colonna%)) * G#
                                                    Else
                                                         Beep
                                                         miavar = "CALCOLO NON ESEGUITO CAUSA ERRORE FORMATO DATI NUMERICI" + nw + nw
                                                         miavar = miavar + "ERRORE NEI DATI DEL CASO DI CARICO NUM.  " + Str(kk% - 5) + nw + nw
                                                         miavar = miavar + "VERIFICARE DATI FOGLIO F06 -->  RIGA " + Str(riga% + j%) + "   COLONNA" + Str(colonna%) + nw
                                                         MsgBox (miavar)
                                                         Exit Sub
                                                    End If
                                                    miavar = Sheets("F06").Cells(riga% + j%, colonna% + 1)
                                                    If IsNumeric(miavar) Then
                                                             PV(2, j%) = CDbl(Sheets("F06").Cells(riga% + j%, colonna% + 1)) * G#
                                                    Else
                                                         Beep
                                                         miavar = "CALCOLO NON ESEGUITO CAUSA ERRORE FORMATO DATI NUMERICI" + nw + nw
                                                         miavar = miavar + "ERRORE NEI DATI DEL CASO DI CARICO NUM.  " + Str(kk% - 5) + nw + nw
                                                         miavar = miavar + "VERIFICARE DATI FOGLIO F06 -->  RIGA " + Str(riga% + j%) + "   COLONNA" + Str(colonna% + 1) + nw
                                                         MsgBox (miavar)
                                                         Exit Sub
                                                    End If
                                          Next j%
                                   Case Else
                                          If qs(kk%) = 1 Then
                                                          k = 1
                                                          nnn = ns
                                                          GoSub ramocarico
                                          Else
                                              k = 1: nnn = ns: GoSub ramoscarico
                                          End If
                                          If qd(kk%) = 1 Then
                                                            k = 2
                                                            nnn = nd
                                                            GoSub ramocarico
                                          Else
                                              k = 2: nnn = nd: GoSub ramoscarico
                                          End If
                       End Select
                       At = attr(xx%)
                       Ac = din(xx%)
                       nn13 = Int(spzfin - spziniz) / passo + 1
                       If XIndex = 0 Or XIndex = 4 Then
                         If marcia$ = "AVANTI" Then
                             rr13 = 4: cc13 = ipoipo + 1
                             rr12 = 4: cc12 = 187
                         Else
                             rr13 = 4: cc13 = ipoipo + 93
                             rr12 = 100: cc12 = 187
                         End If
                         Sheets("STORE13").Cells(rr13, cc13) = Tinizio: rr13 = rr13 + 1
                       End If
                       If XIndex = 1 Then
                         If marcia$ = "AVANTI" Then
                             rr13 = 8 + nn13 * 5: cc13 = ipoipo + 1
                             rr12 = 600: cc12 = 187
                         Else
                             rr13 = 8 + nn13 * 5: cc13 = ipoipo + 93
                             rr12 = 700: cc12 = 187
                         End If
                         Sheets("STORE13").Cells(rr13, cc13) = Tinizio: rr13 = rr13 + 1
                       End If
                       If XIndex = 5 Then
                        If test% = 1 Then
                         If marcia$ = "AVANTI" Then
                             rr13 = 12 + 2 * nn13 * 5: cc13 = ipoipo + 1
                             rr12 = 200: cc12 = 187
                         Else
                             rr13 = 12 + 2 * nn13 * 5: cc13 = ipoipo + 93
                             rr12 = 300: cc12 = 187
                         End If
                        Else
                         If marcia$ = "AVANTI" Then
                             rr13 = 16 + 3 * nn13 * 5: cc13 = ipoipo + 1
                             rr12 = 400: cc12 = 187
                         Else
                             rr13 = 16 + 3 * nn13 * 5: cc13 = ipoipo + 93
                             rr12 = 500: cc12 = 187
                         End If
                        End If
                        Sheets("STORE13").Cells(rr13, cc13) = Tinizio: rr13 = rr13 + 1
                       End If
                       GoSub calcolo
                       If XIndex = 0 Or XIndex = 4 Or XIndex = 5 Or precalc% = 1 Then Exit Do
                       If XIndex = 1 Then
                         'iterazione sulla corsa tenditore
                          corsatentativo = (sviluppomax + sviluppomin) / 4
                          spostamento = corsazero - corsatentativo
                          DELTA = (corsacil - spostamento)
                          If Abs(DELTA) <= 0.01 Then
                                                    Beep
                                                    Exit Do
                          End If
                          If tentativo% = 0 Then
                                            incr = Tinizio * 0.4
                          Else
                             incr = incr / 2
                          End If
                          Tinizio = Tinizio - Sgn(DELTA) * incr
                          tentativo% = tentativo% + 1
                       End If
             Loop
             Return


calcolo:
'-------<<< ciclo di calcolo >>>---------------------------------------------
        ciclospz = 0
        GoSub zeromaxmin
        
        
        For spz = spziniz To spzfin Step passo
             Call posizione_veicoli(spz)
             GoSub lista
             GoSub zerocamp
             GoSub zeropali
             ripeti% = 0: giro% = 0: Z = 0: zz = 0: scostamento = 1
ancorata:
             Tcm(1, 0) = Tinizio / 2
             Tcm(2, 0) = Tinizio / 2
ripeti:
             zz = Z
             ripeti% = ripeti% + 1
             For k = 1 To 2
                      If marcia$ = "AVANTI" Then
                               If k = 1 Then Ksd = 1 Else Ksd = -1
                      Else
                          If k = 1 Then Ksd = -1 Else Ksd = 1
                      End If
                      If k = 1 Then nnn = ns Else nnn = nd
                      For j = 1 To nnn - 1
                                       GoSub campata
                      Next j
             Next k
             If marcia$ = "AVANTI" Then
                                Ksds = 1: Ksdd = -1
             Else
                    Ksds = -1: Ksdd = 1
             End If
             If xx% <> 1 Then
               If Ftrave > 1 Then
                     'forza sulla trave di lancio
                     Tcm(1, ns - 1) = Tcm(1, ns - 1) + Ftrave / 10 * Ksds
                     Tcm(2, nd - 1) = Tcm(2, nd - 1) + Ftrave / 10 * Ksdd
               Else
                 If pstazione > 0 And pstazione <= 1 Then
                     'attrito per deviazione fune in stazione
                     Tcm(1, ns - 1) = Tcm(1, ns - 1) + Tcm(1, ns - 1) * pstazione * attr0 / 100 * Ksds
                     Tcm(2, nd - 1) = Tcm(2, nd - 1) + Tcm(2, nd - 1) * pstazione * attr0 / 100 * Ksdd
                 End If
               End If
             End If
             If m$ = "MONTE" And tt$ = "VALLE" Then
                       If ripeti% < 4 Then GoTo ripeti Else GoTo memo
             End If
             If m$ = "VALLE" And tt$ = "VALLE" Then
                       Diften = Tcm(1, ns - 1) - Tcm(2, nd - 1)
                       If Abs(Diften) <= 0.1 Then GoTo memo
                       Tcm(1, 0) = Tcm(1, 0) - Diften / 2
                       Tcm(2, 0) = Tcm(2, 0) + Diften / 2
                       GoTo ripeti
             End If
             If m$ = "VALLE" And tt$ = "MONTE" Then
                       Diftens = Tcm(1, ns - 1) - Tinizio / 2
                       Diftend = Tcm(2, nd - 1) - Tinizio / 2
                       If Abs(Tcm(1, ns - 1) + Tcm(2, nd - 1) - Tinizio) <= 0.1 Or ripeti% > 100 Then GoTo memo
                       Tcm(1, 0) = Tcm(1, 0) - Diftens
                       Tcm(2, 0) = Tcm(2, 0) - Diftend
                       GoTo ripeti
             End If
             If m$ = "MONTE" And tt$ = "MONTE" Then
                       Diften = Tcm(1, ns - 1) + Tcm(2, nd - 1) - Tinizio
                       If Abs(Diften) <= 0.1 Then GoTo memo
                       Tcm(1, 0) = Tcm(1, 0) - Diften / 2
                       Tcm(2, 0) = Tcm(2, 0) - Diften / 2
                       GoTo ripeti
             End If

memo:
      GoSub memorizza
      If XIndex = 4 Then
            'iterazione sul contrappeso per funi ancorate
             DeltaT = Sheets("F05").Cells(22, 16)
             Sviluppo = 0
             Alleltot = 0
             Alltetot = 0
             For w% = 1 To ns - 1
                   Ci = Sqr(d(1, w%) ^ 2 + l(1, w%) ^ 2)
                   svcamp = Ci + s(1, w%)
                   Sviluppo = Sviluppo + svcamp
                   Alleltot = Alleltot + svcamp * 10 * ((Tcv(1, w%) + Tcm(1, w%)) / 2) / modulo / area
                   Alltetot = Alltetot + svcamp * KT * DeltaT
             Next w%
             For w% = 1 To nd - 1
                   Ci = Sqr(d(2, w%) ^ 2 + l(2, w%) ^ 2)
                   svcamp = Ci + s(2, w%)
                   Sviluppo = Sviluppo + svcamp
                   Alleltot = Alleltot + svcamp * 10 * ((Tcv(2, w%) + Tcm(2, w%)) / 2) / modulo / area
                   Alltetot = Alltetot + svcamp * KT * DeltaT
            Next w%
            LGH_FUNE_A_TERRA = Sviluppo - Alleltot - Alltetot
            scostamento = (LRF_FUNE_A_TERRA - LGH_FUNE_A_TERRA)
            Z = scostamento
            If Abs(scostamento) > 0.001 Then
                  If giro% = 0 Then
                                    DELTA = Tinizio / 10
                                    giro% = giro% + 1
                  Else
                      DELTA = DELTA * Z / (zz - Z)
                      giro% = giro% + 1
                  End If
                  Tinizio = Tinizio + DELTA
                  GoTo ancorata
            End If
            GoSub memorizza
            Tinizio = contr
      End If

      
             If corsacil <> 0 Then
                    'Sheets("F05").Cells(26, 5) = ipotesi$(20) + " " + Format$(corsacil, "#0.00")
             End If
             If XIndex <> 4 Then
                   'Sheets("F05").Cells(31, 1) = sviluppomax / 2
                   'Sheets("F05").Cells(31, 4) = sviluppomin / 2
             End If
             If XIndex = 1 Then
                         ppt = (contr - Tinizio) / contr * 100
                         'Sheets("F05").Cells(32, 1) = Str(ppt) + ".....Loop n-> " + Str(tentativo%)
             End If
             
             ciclospz = ciclospz + 1
             If m$ = "MONTE" Then
                                 truotas = Tcm(1, ns - 1)
                                 truotad = Tcm(2, nd - 1)
                                 If marcia$ = "AVANTI" Then
                                       Fmotrice = truotas - truotad
                                 Else
                                     Fmotrice = truotad - truotas
                                 End If
             Else
                 truotas = Tcv(1, 1)
                 truotad = Tcv(2, 1)
                 If marcia$ = "AVANTI" Then
                                 Fmotrice = truotad - truotas
                 Else
                     Fmotrice = truotas - truotad
                 End If
             End If
             
             Sheets("STORE13").Cells(rr13, cc13) = truotas: rr13 = rr13 + 1
             Sheets("STORE13").Cells(rr13, cc13) = truotad: rr13 = rr13 + 1
             Sheets("STORE13").Cells(rr13, cc13) = corsa: rr13 = rr13 + 1
             Sheets("STORE13").Cells(rr13, cc13) = spz: rr13 = rr13 + 1
             Sheets("STORE13").Cells(rr13, cc13) = Fmotrice: rr13 = rr13 + 1
             riga% = 26 + ciclospz
             If Fmotrice > Sfmax Then Sfmax = Fmotrice
             If Fmotrice < Sfmin Then Sfmin = Fmotrice
             If truotas > truotad Then aderenza = truotas / truotad Else aderenza = truotad / truotas
             If aderenza > maxscorr Then maxscorr = aderenza
             If truotas > trsmax Then trsmax = truotas
             If truotad > trdmax Then trdmax = truotad
             If truotas < trsmin Then trsmin = truotas
             If truotad < trdmin Then trdmin = truotad
             'Sheets("F05").Cells(33, 1) = maxscorr
             'Sheets("F05").Cells(33, 3) = ciclospz
             'IF ipo = 1 OR kk% = 5 THEN EXIT FOR                'impianto fermo o fune nuda (1 solo ciclo)
        Next spz
        If ciclospz > 1 Then ciclospz = ciclospz - 1
        Return
'----------------------------------------------------------------------------

memorizza:
          For k = 1 To 2
          If k = 1 Then nnn = ns Else nnn = nd
          For j = 1 To nnn - 1
          Ci = Sqr(d(k, j) ^ 2 + l(k, j) ^ 2): aalfa = Atn(d(k, j) / l(k, j))
          'tensioni,pressioni,deviazioni e attriti sui sostegni
          If XIndex <> 4 Or Abs(scostamento) <= 0.001 Then
             If Tp(k, j) > maxtp(k, j) Then maxtp(k, j) = Tp(k, j)
             If Tp(k, j) < mintp(k, j) Then mintp(k, j) = Tp(k, j)
             If Pp(k, j) > maxps(k, j) Then maxps(k, j) = Pp(k, j)
             If Pp(k, j) < minps(k, j) Then minps(k, j) = Pp(k, j)
             If Dp(k, j) > maxdev(k, j) Then maxdev(k, j) = Dp(k, j)
             If Dp(k, j) < mindev(k, j) Then mindev(k, j) = Dp(k, j)
             Sql = Abs(Abs(Pp(1, j)) - Abs(Pp(2, j)))
             If Sql > maxsq(j) Then maxsq(j) = Sql
             If Abs(attrito(k, j)) > maxattr(k, j) Then maxattr(k, j) = Abs(attrito(k, j))
             If Abs(attrito(k, j)) < minattr(k, j) Then minattr(k, j) = Abs(attrito(k, j))
             'tensioni ed angoli della fune in campata
             If Tcv(k, j) > maxtv(k, j) Then maxtv(k, j) = Tcv(k, j)
             If Tcv(k, j) < mintv(k, j) Then mintv(k, j) = Tcv(k, j)
             If Tcm(k, j) > maxtm(k, j) Then maxtm(k, j) = Tcm(k, j)
             If Tcm(k, j) < mintm(k, j) Then mintm(k, j) = Tcm(k, j)
             If Av(k, j) > maxav(k, j) Then maxav(k, j) = Av(k, j)
             If Av(k, j) < minav(k, j) Then minav(k, j) = Av(k, j)
             If Am(k, j) > maxam(k, j) Then maxam(k, j) = Am(k, j)
             If Am(k, j) < minam(k, j) Then minam(k, j) = Am(k, j)
          End If
          'frecce-sviluppi-allungamenti elastici
          Tmedia = (Tcv(k, j) + Tcm(k, j)) / 2
          Fm(k, j) = Qf * Ci ^ 2 / (8 * Tmedia): F = Fm(k, j)
          s(k, j) = 8 / 3 * Fm(k, j) ^ 2 * (Cos(aalfa)) ^ 2 / Ci
          If ncnc(k, j) > 0 And ipoipo > 1 Then
             s(k, j) = 0: Skj = 0: GoSub freccianuda
             For ii% = 1 To ncnc(k, j)
               FV(ii%) = PV(k, j) * (l(k, j) - Bo(k, j, ii%)) * Bo(k, j, ii%) / (l(k, j) * Hc(k, j))
                 If Bo(k, j, ii%) <= l(k, j) / 2 Then
                    Fvc(ii%) = FV(ii%) * l(k, j) / (2 * (l(k, j) - Bo(k, j, ii%)))
                 Else
                    Fvc(ii%) = FV(ii%) * l(k, j) / (2 * Bo(k, j, ii%))
                 End If
               Fm(k, j) = Fm(k, j) + Fvc(ii%)
             Next ii%
             For ii% = 1 To ncnc(k, j)
               Fvt(ii%) = FV(ii%)
               For III% = 1 To ii% - 1
                 tg = FV(III%) / Bo(k, j, III%)
                 Fvt(ii%) = Fvt(ii%) + tg * Bo(k, j, ii%)
               Next III%
               For III% = ii% + 1 To ncnc(k, j)
                 tg = FV(III%) / (l(k, j) - Bo(k, j, III%))
                 Fvt(ii%) = Fvt(ii%) + tg * (l(k, j) - Bo(k, j, ii%))
               Next III%
             Next ii%
             For ii% = 1 To ncnc(k, j) + 1
               If ii% = 1 Then
                 Li = l(k, j) - Bo(k, j, ii%)
                 di = Li * Tan(aalfa) - Fvt(ii%) - FF(ii%)
               End If
               If ii% > 1 And ii% <= ncnc(k, j) Then
                 Li = Bo(k, j, ii% - 1) - Bo(k, j, ii%)
                 di = ((l(k, j) - Bo(k, j, ii%)) * Tan(aalfa) - Fvt(ii%) - FF(ii%)) - ((l(k, j) - Bo(k, j, ii% - 1)) * Tan(aalfa) - Fvt(ii% - 1) - FF(ii% - 1))
               End If
               If ii% = ncnc(k, j) + 1 Then
                 Li = Bo(k, j, ii% - 1)
                 di = Li * Tan(aalfa) + Fvt(ii% - 1) + FF(ii% - 1)
               End If
               alfai = Atn(di / Li): Cii = Sqr(Li ^ 2 + di ^ 2)
               s(k, j) = s(k, j) + Cii
               Ffm = Qf * Li ^ 2 / (8 * Hc(k, j) * Cos(alfai))
               Skj = Skj + 8 / 3 * Ffm ^ 2 * (Cos(alfai)) ^ 2 / Cii
             Next ii%
             s(k, j) = s(k, j) + Skj - Sqr(l(k, j) ^ 2 + d(k, j) ^ 2)
             Svcat = 8 / 3 * Fm(k, j) ^ 2 * (Cos(aalfa)) ^ 2 / Ci
             If Svcat > s(k, j) Then s(k, j) = Svcat
          End If
          If XIndex <> 4 Or Abs(scostamento) <= 0.001 Then
              If Fm(k, j) < Fmmin(k, j) Then Fmmin(k, j) = Fm(k, j)
              If Fm(k, j) > Fmmax(k, j) Then Fmmax(k, j) = Fm(k, j)
          End If
          Allung(k, j) = (Tcv(k, j) + Tcm(k, j)) / 2 * Ci / (modulo / 10 * area)   'modulo espresso in N/mm2 e quindi diviso x 10
          Next j: Next k
          GoSub sviluppoanello
          Return


freccianuda:
           'determinazione della freccia alla ascissa generica x per
           'campate a carico uniforme con freccia in mezzaria pari ad F
           'x1,y1 coordinate appoggio valle
           'x2,y2 coordinate appoggio monte
           'alfa  angolo della corda
           'l(k,j)   lunghezza orizzontale campata j del ramo k
           'd(k,j)   dislivello campata j del ramo k
           'FF(i) = freccia alla ascissa x del carico i nella campata
           '-----------------------------------------------------------
           x1 = 0: y1 = 0: x2 = l(k, j): y2 = d(k, j)
           For ii% = 1 To ncnc(k, j)
                   X = l(k, j) - Bo(k, j, ii%)
                   y = y1 + (X - x1) * Tan(aalfa) - 4 * F / l(k, j) ^ 2 * (X - x1) * (x2 - X)
                   FF(ii%) = X * Tan(aalfa) - y
           Next ii%
           Return

sviluppoanello:
           sva = 0: ala = 0
           For k = 1 To 2
                If k = 1 Then nnn = ns Else nnn = nd
                For j = 1 To nnn - 1
                            sva = sva + s(k, j)
                            ala = ala + Allung(k, j)
                Next j
           Next k
           svtot = ala - sva
           corsa = svtot / 2
           If svtot > sviluppomax Then sviluppomax = svtot
           If svtot < sviluppomin Then sviluppomin = svtot
           Return

risultati:
         If marcia$ = "AVANTI" Then
                 If XIndex <> 1 And XIndex <> 5 And XIndex <> 2 Then
                        aaa$ = "####"
                        'tabelle di linea - formato valore -
                        rr6 = 5
                 End If
                 If XIndex = 1 Then
                        'tabelle di linea - formato valore -
                        rr6 = 2005
                 End If
                 If XIndex = 5 Then
                        If test% = 1 Then
                                        a1$ = "A": b1$ = "B": rr6 = 4005
                        Else
                            a1$ = "C": b1$ = "D": rr6 = 6005
                        End If
                 End If
                 If kk% <> 5 Then
                               Sheets("STORE06").Cells(rr6, 1 + ipoipo) = flagmed%
                               Sheets("STORE06").Cells(rr6, 92 + ipoipo) = flagmed%: rr6 = rr6 + 1
                               Sheets("STORE06").Cells(rr6, 1 + ipoipo) = ns
                               Sheets("STORE06").Cells(rr6, 92 + ipoipo) = ns: rr6 = rr6 + 1
                               Sheets("STORE06").Cells(rr6, 1 + ipoipo) = nd
                               Sheets("STORE06").Cells(rr6, 92 + ipoipo) = nd: rr6 = rr6 + 1
                 Else
                               Sheets("STORE06").Cells(rr6, 1 + ipoipo) = 9
                               Sheets("STORE06").Cells(rr6, 92 + ipoipo) = 9: rr6 = rr6 + 1
                               Sheets("STORE06").Cells(rr6, 1 + ipoipo) = ns
                               Sheets("STORE06").Cells(rr6, 92 + ipoipo) = ns: rr6 = rr6 + 1
                               Sheets("STORE06").Cells(rr6, 1 + ipoipo) = nd
                               Sheets("STORE06").Cells(rr6, 92 + ipoipo) = nd: rr6 = rr6 + 1
                 End If
                 rr666 = rr6    'riferimento inizio riga valori per ogni colonna di ipotesi
                 For k% = 1 To 2
                     If k% = 1 Then nnn = ns: cc6 = 1 + ipoipo
                     If k% = 2 Then nnn = nd: cc6 = 92 + ipoipo
                     rr6 = rr666
                     For j% = 2 To nnn
                                 GoSub sigla1
                                 GoSub registradbf
                     Next j%
                 Next k%
                 Close #5, #6
         End If
         If precalc% <> 1 Then GoSub potenza
         Return

sigla1:
     cv$ = n$(k%, j% - 1)
     cm$ = n$(k%, j%)
     sost$ = n$(k%, j%)
     Return

potenza:
       'sforzo quadratico medio e potenza
        If xx% = 1 Then Return                      '-> xx% / kk%=1  impianto fermo
        If XIndex = 0 Or XIndex = 4 Then
           If marcia$ = "AVANTI" Then
               rr13 = 4: cc13 = ipoipo + 1
           Else
               rr13 = 4: cc13 = ipoipo + 93
           End If
        End If
        If XIndex = 1 Then
           If marcia$ = "AVANTI" Then
               rr13 = 8 + nn13 * 5: cc13 = ipoipo + 1
           Else
               rr13 = 8 + nn13 * 5: cc13 = ipoipo + 93
           End If
        End If
        If XIndex = 5 Then
         If test% = 1 Then
           If marcia$ = "AVANTI" Then
               rr13 = 12 + 2 * nn13 * 5: cc13 = ipoipo + 1
           Else
               rr13 = 12 + 2 * nn13 * 5: cc13 = ipoipo + 93
           End If
         Else
           If marcia$ = "AVANTI" Then
               rr13 = 16 + 3 * nn13 * 5: cc13 = ipoipo + 1
           Else
               rr13 = 16 + 3 * nn13 * 5: cc13 = ipoipo + 93
           End If
         End If
        End If

        cc% = 0:  Ftqm = 0
        inargano = argano / 10 * Ac
        vvtt = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
        '---------------------< sforzo e potenza media >---------------------
        If ZETA = "I" Then
                dt = passo / vel
                truotas = 0
                truotas = Sheets("STORE13").Cells(rr13, cc13)
                While truotas > 0
                    truotas = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    truotad = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    corsa = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    aa1 = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    aa2 = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    Ft = truotas - truotad
                    Ftqm = Ftqm + Ft ^ 2 * dt
                    cc% = cc% + 1
                    truotas = Sheets("STORE13").Cells(rr13, cc13)
                Wend
                Ftqm = Sqr(Ftqm / (dt * cc%))
        Else
            truotas = 0
            truotas = Sheets("STORE13").Cells(rr13, cc13)
            While truotas > 0
                    truotas = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    truotad = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    corsa = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    aa1 = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    aa2 = Sheets("STORE13").Cells(rr13, cc13): rr13 = rr13 + 1
                    If m$ = "MONTE" Then
                            If marcia$ = "AVANTI" Then
                                      Ft = truotas - truotad
                            Else
                                Ft = truotad - truotas
                            End If
                    Else
                        If marcia$ = "AVANTI" Then
                                   Ft = truotad - truotas
                        Else
                            Ft = truotas - truotad
                        End If
                    End If
                    Ftqm = Ftqm + Ft
                    cc% = cc% + 1
                    truotas = Sheets("STORE13").Cells(rr13, cc13)
            Wend
            Ftqm = Ftqm / cc%
        End If
        Sfmed = Ftqm + inargano
        If Sfmed > 0 Then
                         rargmed = rendarg
        Else
            rargmed = 1 / rendarg
        End If
        Potm = Sfmed * vel / rargmed / 100
        trsmed = (trsmax + trsmin) / 2
        trdmed = (trdmax + trdmin) / 2
        If trsmed >= trdmed Then
                                maxscorr = trsmed / trdmed
        Else
            maxscorr = trdmed / trsmed
        End If
        rr12 = rr12 - 1 + ipoipo
        Sheets("STORE13").Cells(rr12, cc12) = ipoipo: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = Ftqm: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = inargano: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = Sfmed: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = rargmed: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = Potm: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = maxscorr: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = sviluppomax / 2: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = sviluppomin / 2: cc12 = cc12 + 1
        Sheets("STORE13").Cells(rr12, cc12) = trsmed + trdmed: cc12 = cc12 + 1

       '------------------< sforzo e potenza massima >----------------------
       If Abs(Sfmax) > Abs(Sfmin) Then MSforzo = Sfmax Else MSforzo = Sfmin
       MSforzo = MSforzo + inargano
       If MSforzo > 0 Then
                         rargmax = rendarg
       Else
            rargmax = 1 / rendarg
       End If
       Potmax = MSforzo * vel / rargmax / 100
       
       Sheets("STORE13").Cells(rr12, cc12) = Sfmax: cc12 = cc12 + 1
       Sheets("STORE13").Cells(rr12, cc12) = Sfmin: cc12 = cc12 + 1
       Sheets("STORE13").Cells(rr12, cc12) = MSforzo: cc12 = cc12 + 1
       Sheets("STORE13").Cells(rr12, cc12) = rargmax: cc12 = cc12 + 1
       Sheets("STORE13").Cells(rr12, cc12) = Potmax: cc12 = cc12 + 1
       Return


zerocamp:
        For k% = 1 To 2
               If k% = 1 Then nnn = ns Else nnn = nd
               For j% = 1 To nnn - 1
                       Tcv(k%, j%) = 0: Tcm(k%, j%) = 0
                       Av(k%, j%) = 0: Am(k%, j%) = 0
               Next j%
        Next k%
        Return

zeropali:
        For k% = 1 To 2
                 If k% = 1 Then nnn = ns Else nnn = nd
                 For j% = 1 To nnn
                        Tp(k%, j%) = 0: Dp(k%, j%) = 0
                        Pp(k%, j%) = 0: attrito(k%, j%) = 0
                 Next j%
        Next k%
        Return

zeromaxmin:
        maxscorr = 0
        trsmax = 0: trsmin = 99999
        trdmax = 0: trdmin = 99999
        sviluppomax = -999999: sviluppomin = 999999
        Sfmax = -999999: Sfmin = 999999
        For k% = 1 To 2
           If k% = 1 Then nnn = ns Else nnn = nd
           For j% = 1 To nnn
                  maxps(k%, j%) = -999999: minps(k%, j%) = 999999
                  Fmmin(k%, j%) = 999999: Fmmax(k%, j%) = -999999
                  maxav(k%, j%) = -999999: maxam(k%, j%) = -999999
                  maxtv(k%, j%) = -999999: maxtm(k%, j%) = -999999
                  minav(k%, j%) = 999999: minam(k%, j%) = 999999
                  mintv(k%, j%) = 999999: mintm(k%, j%) = 999999
                  maxdev(k%, j%) = -999999: mindev(k%, j%) = 999999
                  maxattr(k%, j%) = -999999: minattr(k%, j%) = 999999
                  maxtp(k%, j%) = -999999: mintp(k%, j%) = 999999
                  maxsq(j%) = -999999
           Next j%
        Next k%
        Return


'determinazione della posizione delle VETTURE IN CAMPATA
'ncnc(k,j)   = numero totale delle vetture nella campata j del ramo k
'            1<= k <= 2
'            1<= j <= n-1
'Bo(k,j,i) = distanza orizzontale della vettura i nella campata j del
'            ramo k , rispetto all'estremita di monte
'            1<= i <= 20
'alfa(k,j)   = angolo di inclinazione della corda della campata j
'            rispetto all'orizzontale

lista:
     For j = 1 To ns - 1: ncnc(1, j) = 0: alfa(1, j) = 0: Next
     For j = 1 To nd - 1: ncnc(2, j) = 0: alfa(2, j) = 0: Next
     For j = 1 To ns - 1
                      ii% = 0
                      alfa(1, j) = Atn(d(1, j) / l(1, j))
                      k = 1: GoSub conta
     Next
     For j = 1 To nd - 1
                      ii% = 0
                      alfa(2, j) = Atn(d(2, j) / l(2, j))
                      k = 2: GoSub conta
     Next
     Return
     
conta:
     For X = 1 To i(k)
     If Pr(k, X) > Prs(k, j) And Pr(k, X) < Prs(k, j + 1) Then
           ncnc(k, j) = ncnc(k, j) + 1
           ii% = ii% + 1
           Bo(k, j, ii%) = (Prs(k, j + 1) - Pr(k, X)) * Cos(alfa(k, j))
     End If
     Next X
     Return
     
     
pulisci_geoveicoli:
             Sheets("F05").Select
             Sheets("F05").Cells(1, 1).Select
             Sheets("F05").Range("T59:AF3000").Select
             Selection.ClearContents
             Sheets("F05").Cells(1, 1).Select
             Return '------------------------------< fine procedura geoveicoli>------------------


'$INCLUDE: 'campata.bi'
'determinazioni delle tensioni,angoli imbocco,frecce, pressioni
'Tk  -  Tcv(k,j) tensione a valle della campata j
'Tv  -  Tcm(k,j) tensione a monte della campata j
'       Tp(k,j)  tensione sul sostegno a valle della campata j
'       Dp(k,j)  deviazione della fune sul sostegno
'       Pp(k,j)  pressione della fune sul sostegno
'  attrito(k,j)  attrito sulle rulliere del sostegno
'av  -  av(k,j)  angolo di imbocco della fune a valle ,con l'orizzontale
'am  -  am(k,j)  angolo di imbocco della fune a monte
'Qf              peso unitario della fune nuda
'vattr           attrito sui rulli (percentuale/assoluto)
'Ac              accelerazione (+)  o decelerazione (-)
'Ci              corda della campata
'Pv(k,j)         peso delle vetture della campata j

campata:
       GoSub valoreattrito
       iterazioni = 0
       aalfa = alfa(k, j)
       aattrito = 0
       angoriz = 0
       poriz = 0
       Ci = Sqr(l(k, j) ^ 2 + d(k, j) ^ 2)
       Tk = Tcm(k, j - 1)
       a = Qf * Ci * l(k, j) / 2
       b = 0
       For X = 1 To ncnc(k, j)
             b = b + PV(k, j) * Bo(k, j, X)
       Next X
itera:
           iterazioni = iterazioni + 1
           C = Tk * Ci
           X = (a + b) / C
           epsilon = Atn(X / Sqr(1 - X ^ 2))
           beta = epsilon - aalfa
           GoSub tenscampata
           If k = 1 Then rrulli = r1(j) Else rrulli = r2(j)            'num.rulli sostegno
           If InStr(n$(k, j), "R") > 0 Then
                     Massarull = rrulli * mapp / (G# * 10)            'massa rulliera ritenuta
           Else
               Massarull = rrulli * mrit / (G# * 10)                  'massa rulliera appoggio
           End If
           Tp(k, j) = Tcm(k, j - 1) + Massarull / 2 * Ac * Ksd       'inerzia rulli

           If InStr(n$(k, j), "/C") > 0 Then
                        'attrito per deviazione orizzontale ------------------------
                        'NB: solo per calcolo con attrito %
                         p% = InStr(n$(k, j), "/C") + 2
                         angoriz = Val(mid$(n$(k, j), p%, 10)) * PI# / 180
           Else
                angoriz = 0
           End If

           If ttat$ <> "A" Then
                          Tp(k, j) = Tp(k, j) + (Abs(Pp(k, j)) + Abs(poriz)) * vattr / 2 * Ksd    'attrito % rulli + attrito % deviaz.orizz.
           Else
               Tp(k, j) = Tp(k, j) + vattr / 2 * Ksd                              'attrito assoluto rulli
           End If
           Dp(k, j) = Am(k, j - 1) + beta
           Pp(k, j) = 2 * Tp(k, j) * Sin(Dp(k, j) / 2)
           poriz = 2 * Tp(k, j) * Sin(angoriz / 2)
           If j > 1 Then
               If ttat$ <> "A" Then
                         Tk = Tp(k, j) + (Abs(Pp(k, j)) + Abs(poriz)) * vattr / 2 * Ksd + Massarull / 2 * Ac * Ksd
                         diffattr = Abs(aattrito - Abs(Pp(k, j) * vattr))
                         If diffattr >= 0.1 Then
                                        aattrito = (Abs(Pp(k, j)) + Abs(poriz)) * vattr
                                        itera$ = "S"
                         Else
                             itera$ = "N"
                         End If
               Else
                   Tk = Tp(k, j) + vattr / 2 * Ksd + Massarull / 2 * Ac * Ksd
                   aattrito = vattr
                   GoSub tenscampata
                   itera$ = "N"
               End If
           Else
               itera$ = "N"
           End If
           If itera$ = "S" Then GoTo itera

       'valori finali di campata j del ramo k
       Tcv(k, j) = Tk
       Tcm(k, j) = Tv
       If j = 1 Then
          If xx% <> 1 Then
             If Ftrave > 0 Then
                   Tcm(k, j) = Tcm(k, j) + Ftrave / 10 * Ksd
             End If
             If pstazione > 0 And pstazione <= 0.5 Then
                         Tcm(k, j) = Tcm(k, j) + Tcm(k, j) * 2 * Sin(pstazione / 2) * attr0 / 100 * Ksd
             End If
          End If
       End If
       Hc(k, j) = Hk
       Av(k, j) = beta
       Am(k, j) = Atn(nv / Hk)
       attrito(k, j) = aattrito
       Return


tenscampata:
          Hk = Tk * Cos(beta)
          Nk = Tk * Sin(beta)
          nv = Qf * Ci + PV(k, j) * ncnc(k, j) - Nk
          Tv = Sqr(nv ^ 2 + Hk ^ 2)
          Massacamp = (Qf * Ci + PV(k, j) * ncnc(k, j)) / G#          'massa fune+veicoli in campata
          Tv = Tv + Massacamp * Ac * Ksd / 10                         'tens +/- inerzia in campata
          Return

valoreattrito:
             If At = 0 Then
                          vattr = 0
                          Return
             End If
             If k = 1 Then nrulli = r1(j) Else nrulli = r2(j)
             ttat$ = UCase$(tat$(k, j))
             If Len(ttat$) > 0 Then
                                 If xx% > 3 Then
                                               vattr = atfre(k, j)
                                 Else
                                     vattr = atreg(k, j)
                                 End If
                                 If ttat$ = "A" Then
                                                  vattr = vattr * nrulli
                                 Else
                                     vattr = vattr / 100
                                 End If
             Else
                If xx% > 3 Then
                              vattr = attr1
                Else
                    vattr = attr0
                End If
                vattr = vattr / 100
             End If
             Return


nuda:  '-------------------- peso veicoli nullo : fune nuda --------------
       For k = 1 To 2
                    If k = 1 Then nnn = ns Else nnn = nd
                    For j = 1 To nnn - 1
                                      PV(k, j) = 0
                    Next j
       Next k
       Return


ramocarico:
          For j = 1 To nnn - 1
          If kk% = 1 Or kk% = 3 Then
             If d(k, j) >= 0 Then                  'inversione di carico per dislivelli negativi
                           PV(k, j) = qc * G#      'e conferma di default del check1 bottom
             Else
                 If Sheets("F05").inversione_contropendenze.value Then
                           PV(k, j) = qv * G#
                 Else
                           PV(k, j) = qc * G#
                 End If
             End If
          Else
             PV(k, j) = qc * G#
          End If
          Next
          Return

ramoscarico:
          For j = 1 To nnn - 1
          If kk% = 1 Or kk% = 3 Then
             If d(k, j) >= 0 Then                 'inversione di carico per dislivelli negativi
                           PV(k, j) = qv * G#
             Else
                 If Sheets("F05").inversione_contropendenze.value Then
                           PV(k, j) = qc * G#
                 Else
                     PV(k, j) = qv * G#
                 End If
             End If
          Else
              PV(k, j) = qv * G#
          End If
          Next
          Return


parametri:
         'parametri di linea in funzione delle ipotesi di verifica
          qs(1) = 1: qd(1) = 0
          qs(2) = 0: qd(2) = 0
          qs(3) = 0: qd(3) = 1
          qs(4) = 1: qd(4) = 1
          For uu% = 10 To 24
                If uu% < 14 Then
                         nomeipotesi$ = Sheets("F05").Cells(uu%, 1) + "  " + Sheets("F05").Cells(uu%, 2)
                Else
                         nomeipotesi$ = Sheets("F05").Cells(uu%, 1)
                End If
                ipotesi$(uu% - 9) = "  Num." + Str$(uu% - 9) + "   " + nomeipotesi$
          Next uu%
          For uu% = 4 To 9
                  modo$(uu% - 3) = Sheets("F05").Cells(7, uu%) + "  " + Sheets("F05").Cells(8, uu%)
          Next uu%
          din(1) = 0: attr(1) = 0
          din(2) = 0: attr(2) = attr0
          din(3) = acc: attr(3) = attr0
          din(4) = -dec: attr(4) = attr1
          din(5) = -fren: attr(5) = attr1
          din(6) = -rap: attr(6) = attr1
          GoSub leggiattriti
          Return

leggiattriti:
          riga% = 12
          For j% = 2 To ns - 1
                         tat$(1, j%) = Sheets("F07").Cells(riga%, 3)
                         atreg(1, j%) = Sheets("F07").Cells((riga%), 4)
                         atfre(1, j%) = Sheets("F07").Cells(riga%, 5)
                         riga% = riga% + 1
          Next j%
          riga% = 12
          For j% = 2 To nd - 1
                         tat$(2, j%) = Sheets("F07").Cells(riga%, 8)
                         atreg(2, j%) = Sheets("F07").Cells((riga%), 9)
                         atfre(2, j%) = Sheets("F07").Cells(riga%, 10)
                         riga% = riga% + 1
          Next j%
      Return

registradbf:
           If k% = 1 Then rrulli = r1(j%) Else rrulli = r2(j%)
           If flagmed% = 9 Or kk% = 5 Then
                        devmed = (maxdev(k%, j%) + mindev(k%, j%)) / 2 * 180 / PI#
                        presmed = (maxps(k%, j%) + minps(k%, j%)) / 2
                        tcvmed = (maxtv(k%, j% - 1) + mintv(k%, j% - 1)) / 2
                        tcmmed = (maxtm(k%, j% - 1) + mintm(k%, j% - 1)) / 2
                        fcmed = (Fmmin(k%, j% - 1) + Fmmax(k%, j% - 1)) / 2
                        avmed = (maxav(k%, j% - 1) + minav(k%, j% - 1)) / 2
                        ammed = (maxam(k%, j% - 1) + minam(k%, j% - 1)) / 2
                        devunit = 0: presunit = 0
                        If rrulli > 0 Then
                                         devunit = devmed / rrulli
                                         presunit = presmed / rrulli
                        End If
                        'Write #6, cv$, cm$, tcvmed, tcmmed, fcmed, avmed, ammed
                        Sheets("STORE06").Cells(rr6, cc6) = cv$: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = cm$: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = tcvmed: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = tcmmed: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = fcmed: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = avmed: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = ammed: rr6 = rr6 + 1
                        If j% < nnn Then
                                'Write #6, sost$, (maxtp(k%, j%) + mintp(k%, j%)) / 2, (maxdev(k%, j%) + mindev(k%, j%)) / 2, (maxps(k%, j%) + minps(k%, j%)) / 2, (maxattr(k%, j%) + minattr(k%, j%)) / 2
                                Sheets("STORE06").Cells(rr6, cc6) = sost$: rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = (maxtp(k%, j%) + mintp(k%, j%)) / 2: rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = (maxdev(k%, j%) + mindev(k%, j%)) / 2: rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = (maxps(k%, j%) + minps(k%, j%)) / 2: rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = (maxattr(k%, j%) + minattr(k%, j%)) / 2: rr6 = rr6 + 1
                        End If
           Else
               devmax = maxdev(k%, j%) * 180 / PI#
               presmax = maxps(k%, j%)
               devmin = mindev(k%, j%) * 180 / PI#
               presmin = minps(k%, j%)
               devmaxu = 0: devminu = 0: presmaxu = 0: presminu = 0
               If rrulli > 0 Then
                               devmaxu = devmax / rrulli
                               devminu = devmin / rrulli
                               presmaxu = presmax / rrulli
                               presminu = presmin / rrulli
               End If
                        'Write #6, cv$, cm$, maxtv(k%, j% - 1), maxtm(k%, j% - 1), Fmmax(k%, j% - 1), maxav(k%, j% - 1), maxam(k%, j% - 1)
                        Sheets("STORE06").Cells(rr6, cc6) = cv$: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = cm$: rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = maxtv(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = maxtm(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = Fmmax(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = maxav(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = maxam(k%, j% - 1): rr6 = rr6 + 1
                        'Write #6, mintv(k%, j% - 1), mintm(k%, j% - 1), Fmmin(k%, j% - 1), minav(k%, j% - 1), minam(k%, j% - 1)
                        Sheets("STORE06").Cells(rr6, cc6) = mintv(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = mintm(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = Fmmin(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = minav(k%, j% - 1): rr6 = rr6 + 1
                        Sheets("STORE06").Cells(rr6, cc6) = minam(k%, j% - 1): rr6 = rr6 + 1
                        If j% < nnn Then
                                'Write #6, sost$, maxtp(k%, j%), maxdev(k%, j%), maxps(k%, j%), maxattr(k%, j%)
                                Sheets("STORE06").Cells(rr6, cc6) = sost$: rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = maxtp(k%, j%): rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = maxdev(k%, j%): rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = maxps(k%, j%): rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = maxattr(k%, j%): rr6 = rr6 + 1
                                'Write #6, mintp(k%, j%), mindev(k%, j%), minps(k%, j%), minattr(k%, j%)
                                Sheets("STORE06").Cells(rr6, cc6) = mintp(k%, j%): rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = mindev(k%, j%): rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = minps(k%, j%): rr6 = rr6 + 1
                                Sheets("STORE06").Cells(rr6, cc6) = minattr(k%, j%): rr6 = rr6 + 1
                        End If
           End If
           Return
           
End Sub

Public Sub analisi()
        'On Local Error GoTo err_analisi
        
        Dim cekerr As String
        
        path$ = "C:\TMP_SIF\"
        nome$ = "valori"
        
        cek% = 0
        j% = ns + 4 + 5     'massimo numero di sostegni per ramo
        k% = 2              'massimo numero di rami impianto
        ReDim maxps(k%, j%), minps(k%, j%), Fmmin(k%, j%), Fmmax(k%, j%)
        ReDim maxav(k%, j%), maxam(k%, j%), maxtv(k%, j%), maxtm(k%, j%)
        ReDim minav(k%, j%), minam(k%, j%), mintv(k%, j%), mintm(k%, j%)
        ReDim maxdev(k%, j%), mindev(k%, j%), maxattr(k%, j%), minattr(k%, j%)
        ReDim maxtp(k%, j%), mintp(k%, j%), maxsq(j%)

        ReDim maxval(k%, j%), maxmon(k%, j%), minval(k%, j%), minmon(k%, j%)

        For k% = 1 To 2
        If k% = 1 Then nn = ns Else nn = nd
        For j% = 1 To nn
                  maxps(k%, j%) = -999999: minps(k%, j%) = 999999
                  Fmmin(k%, j%) = 999999: Fmmax(k%, j%) = -999999
                  maxav(k%, j%) = -999999: maxam(k%, j%) = -999999
                  maxtv(k%, j%) = -999999: maxtm(k%, j%) = -999999
                  minav(k%, j%) = 999999: minam(k%, j%) = 999999
                  mintv(k%, j%) = 999999: mintm(k%, j%) = 999999
                  maxdev(k%, j%) = -999999: mindev(k%, j%) = 999999
                  maxattr(k%, j%) = -999999: minattr(k%, j%) = 999999
                  maxtp(k%, j%) = -999999: mintp(k%, j%) = 999999
                  maxsq(j%) = -999999
                  maxval(k%, j%) = -999: maxmon(k%, j%) = -999
                  minval(k%, j%) = 999: minmon(k%, j%) = 999
        Next j%
        Next k%
        
        cekerr = ""
        aaa$ = "###"
        For w% = 1 To 15
           For u% = 1 To 6
               If SiNo%(w%, u%) = 1 Then
                                lll% = (w% - 1) * 6 + u%
                                Select Case Ipocalc
                                       Case 0
                                            rr666 = 5
                                       Case 1
                                            rr666 = 4005
                                       Case 2
                                            rr666 = 2005
                                       Case 3
                                            rr666 = 6005
                                End Select
                                rr6 = rr666: cc6 = 1 + lll%
                                flag% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                                n1% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                                n2% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                                k% = 1: fine% = 0
                                'salita
                                cc6 = 1 + lll%
                                For j% = 2 To n1%
                                       If j% = n1% Then fine% = 9
                                       GoSub leggiPMMFMM
                                Next j%
                                k% = 2: fine% = 0
                                'discesa
                                rr6 = rr666 + 3: cc6 = 92 + lll%
                                For j% = 2 To n2%
                                      If j% = n2% Then fine% = 9
                                      GoSub leggiPMMFMM
                                Next j%
                                Close
               End If
           Next u%
        Next w%
        
        Select Case Ipocalc
               Case 0
                    rr555 = 3: GoSub tabella_maxmin
               Case 1
                    rr555 = 82: GoSub tabella_maxmin
               Case 2
                    rr555 = 42: GoSub tabella_maxmin
               Case 3
                    rr555 = 122: GoSub tabella_maxmin
        End Select
        
        '---------------------<pressione massime e minime>---------------------
        If ns = nd Then
                      Select Case Ipocalc
                                       Case 0
                                            rr555 = 158
                                       Case 1
                                            rr555 = 170
                                       Case 2
                                            rr555 = 164
                                       Case 3
                                            rr555 = 176
                      End Select
                      'vale solo per impianti con ramo salita = ramo discesa
                      For j% = 2 To ns - 1
                                         maxs = Abs(maxps(1, j%))
                                         mins = Abs(minps(1, j%))
                                         If maxs < mins Then
                                                           minmins = maxs
                                                           maxmaxs = mins
                                                           mins = minmins
                                                           maxs = maxmaxs
                                         End If
                                         maxd = Abs(maxps(2, j%))
                                         mind = Abs(minps(2, j%))
                                         If maxd < mind Then
                                                           minmind = maxd
                                                           maxmaxd = mind
                                                           mind = minmind
                                                           maxd = maxmaxd
                                         End If
                                         If maxs > maxd Then maxp = maxs * Sgn(maxps(1, j%)) Else maxp = maxd * Sgn(maxps(2, j%))
                                         If mins < mind Then minp = mins * Sgn(minps(1, j%)) Else minp = mind * Sgn(minps(2, j%))
                                         rr5 = rr555: cc5 = 1 + j%
                                         Sheets("STORE05").Cells(rr5, cc5) = n$(1, j%): rr5 = rr5 + 1
                                         Sheets("STORE05").Cells(rr5, cc5) = maxp: rr5 = rr5 + 1
                                         Sheets("STORE05").Cells(rr5, cc5) = minp: rr5 = rr5 + 1
                      Next j%
                      Close
        End If
        Select Case Ipocalc
                           Case 0
                                rr555 = 185
                           Case 1
                                rr555 = 197
                           Case 2
                                rr555 = 191
                           Case 3
                                rr555 = 203
        End Select
        'ramo salita
        For j% = 2 To ns - 1
                           rr5 = rr555
                           cc5 = 1 + j%
                           Sheets("STORE05").Cells(rr5, cc5) = n$(1, j%): rr5 = rr5 + 1
                           Sheets("STORE05").Cells(rr5, cc5) = maxps(1, j%): rr5 = rr5 + 1
                           Sheets("STORE05").Cells(rr5, cc5) = minps(1, j%): rr5 = rr5 + 1
            Next j%
        Close
        'ramo discesa
        For j% = 2 To nd - 1
                           rr5 = rr555
                           cc5 = 82 + j%
                           Sheets("STORE05").Cells(rr5, cc5) = n$(2, j%): rr5 = rr5 + 1
                           Sheets("STORE05").Cells(rr5, cc5) = maxps(2, j%): rr5 = rr5 + 1
                           Sheets("STORE05").Cells(rr5, cc5) = minps(2, j%): rr5 = rr5 + 1
            Next j%
        Close

        '------------------<frecce massime e minime>---------------------------
        If ns = nd Then
                      Select Case Ipocalc
                                       Case 0
                                            rr555 = 213
                                       Case 1
                                            rr555 = 231
                                       Case 2
                                            rr555 = 222
                                       Case 3
                                            rr555 = 240
                      End Select
                      For j% = 2 To ns - 1
                                      maxs = Fmmax(1, j%)
                                      mins = Fmmin(1, j%)
                                      maxd = Fmmax(2, j%)
                                      mind = Fmmin(2, j%)
                                      If maxs > maxd Then fmax = maxs Else fmax = maxd
                                      If mins < mind Then fmin = mins Else fmin = mind
                                      rr5 = rr555
                                      cc5 = j%
                                      Sheets("STORE05").Cells(rr5, cc5) = j% - 1: rr5 = rr5 + 1
                                      Sheets("STORE05").Cells(rr5, cc5) = n$(1, j% - 1): rr5 = rr5 + 1
                                      Sheets("STORE05").Cells(rr5, cc5) = n$(1, j%): rr5 = rr5 + 1
                                      Sheets("STORE05").Cells(rr5, cc5) = fmax: rr5 = rr5 + 1
                                      Sheets("STORE05").Cells(rr5, cc5) = fmin: rr5 = rr5 + 1
                                      Sheets("STORE05").Cells(rr5, cc5) = 0: rr5 = rr5 + 1
                      Next j%
                      Close
        End If
        Select Case Ipocalc
                                       Case 0
                                            rr555 = 250
                                       Case 1
                                            rr555 = 268
                                       Case 2
                                            rr555 = 259
                                       Case 3
                                            rr555 = 277
        End Select
        For j% = 2 To ns
                       rr5 = rr555
                       cc5 = j%
                       Sheets("STORE05").Cells(rr5, cc5) = j% - 1: rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = n$(1, j% - 1): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = n$(1, j%): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = Fmmax(1, j%): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = Fmmin(1, j%): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = 0: rr5 = rr5 + 1
        Next j%
        Close
        For j% = 2 To nd
                       rr5 = rr555
                       cc5 = 80 + j%
                       Sheets("STORE05").Cells(rr5, cc5) = j% - 1: rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = n$(2, j% - 1): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = n$(2, j%): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = Fmmax(2, j%): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = Fmmin(2, j%): rr5 = rr5 + 1
                       Sheets("STORE05").Cells(rr5, cc5) = 0: rr5 = rr5 + 1
        Next j%
        Close

        'registrazione differenza angoli di imbocco campate
        Select Case Ipocalc
                           Case 0
                                rr555 = 290
                           Case 1
                                rr555 = 314
                           Case 2
                                rr555 = 302
                           Case 3
                                rr555 = 326
        End Select
        For j% = 2 To ns
                   diffv = Abs(Abs(maxval(1, j%)) - Abs(minval(1, j%)))
                   diffm = Abs(Abs(maxmon(1, j%)) - Abs(minmon(1, j%)))
                   rr5 = rr555: cc5 = j%
                   Sheets("STORE05").Cells(rr5, cc5) = j% - 1: rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = n$(1, j% - 1): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = n$(1, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = maxval(1, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = minval(1, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = maxmon(1, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = minmon(1, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = diffv: rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = diffm: rr5 = rr5 + 1
        Next j%
        For j% = 2 To nd
                   diffv = Abs(Abs(maxval(2, j%)) - Abs(minval(2, j%)))
                   diffm = Abs(Abs(maxmon(2, j%)) - Abs(minmon(2, j%)))
                   rr5 = rr555: cc5 = 80 + j%
                   Sheets("STORE05").Cells(rr5, cc5) = j% - 1: rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = n$(2, j% - 1): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = n$(2, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = maxval(2, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = minval(2, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = maxmon(2, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = minmon(2, j%): rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = diffv: rr5 = rr5 + 1
                   Sheets("STORE05").Cells(rr5, cc5) = diffm: rr5 = rr5 + 1
        Next j%
        Close

        Exit Sub

tabella_maxmin:
        rr5 = rr555
        Sheets("STORE05").Cells(rr5, 2) = n1%: rr5 = rr5 + 1
        Sheets("STORE05").Cells(rr5, 2) = n2%: rr5 = rr5 + 1
        k% = 1: fine% = 0
        For j% = 2 To n1%
                         rr5 = rr555 + 2
                         cc5 = 1 + j%
                         If j% = n1% Then fine% = 9
                         GoSub spoolPMMFMM
        Next j%
        k% = 2: rr5 = rr555 + 2
        fine% = 0
        For j% = 2 To n2%
                         rr5 = rr555 + 2
                         cc5 = 82 + j%
                         If j% = n2% Then fine% = 9
                         GoSub spoolPMMFMM
        Next j%
        Close
        Return


spoolPMMFMM:
           If k% = 1 Then rrulli = r1(j%) Else rrulli = r2(j%)
           mmaxdevu = 0: mmindevu = 0: mmaxpresu = 0: mminpresu = 0
           test$ = ""
           mmaxdev = maxdev(k%, j%)
           mmaxpres = maxps(k%, j%)
           mmindev = mindev(k%, j%)
           mminpres = minps(k%, j%)
           If Abs(mmaxpres) < Abs(mminpres) Then
                                             mmax = mminpres
                                             mmin = mmaxpres
                                             mmaxpres = mmax
                                             mminpres = mmin
           End If
           If Abs(mmaxdev) < Abs(mmindev) Then
                                            mmax = mmindev
                                            mmin = mmaxdev
                                            mmaxdev = mmax
                                            mmindev = mmin
           End If
           If rrulli = 0 Then
                             If k% = 1 And j% < n1% Then GoSub assegna_rulli
                             If k% = 2 And j% < n2% Then GoSub assegna_rulli
           End If
           If rrulli > 0 Then
             mmaxdevu = mmaxdev / rrulli
             mmindevu = mmindev / rrulli
             mmaxpresu = mmaxpres / rrulli
             mminpresu = mminpres / rrulli
             test$ = ""
             If mmaxpresu > 0 And mminpresu > 0 Then
                        'appoggi
                        If Abs(mmaxpresu) > cmaxapp / 10 Then test$ = "-NV-"
                        If Abs(mmaxdevu) > devmaxapp / PI# * 180 Then test$ = "-NV-"
             End If
             If mmaxpresu < 0 And mminpresu < 0 Then
                        'ritenute
                        If Abs(mmaxpresu) > cmaxrit / 10 Then test$ = "-NV-"
                        If Abs(mmaxdevu) > devmaxrit / PI# * 180 Then test$ = "-NV-"
             End If
             If (mmaxpresu < 0 And mminpresu > 0) Or (mmaxpresu > 0 And mminpresu < 0) Then
                        'doppio effetto
                        If Abs(mmaxpresu) > DEcarmax / 10 Then test$ = "-NV-"
                        If Abs(mmaxdevu) > devmaxrit / PI# * 180 Then test$ = "-NV-"
             End If
           End If
           If Abs(maxav(k%, j%)) < Abs(minav(k%, j%)) Then
                                         mmax = minav(k%, j%)
                                         mmin = maxav(k%, j%)
                                         maxav(k%, j%) = mmax
                                         minav(k%, j%) = mmin
           End If
           Sheets("STORE05").Cells(rr5, cc5) = n$(k%, j% - 1): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = n$(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = maxtv(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = Fmmax(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = maxav(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = maxam(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = mintv(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = Fmmin(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = minav(k%, j%): rr5 = rr5 + 1
           Sheets("STORE05").Cells(rr5, cc5) = minam(k%, j%): rr5 = rr5 + 1
           If fine% = 0 Then
               Sheets("STORE05").Cells(rr5, cc5) = n$(k%, j%): rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = maxtp(k%, j%): rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mmaxdev: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mmaxpres: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = maxattr(k%, j%): rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = rrulli: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mmaxdevu: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mmaxpresu: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = test$: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = " ": rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mintp(k%, j%): rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mmindev: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mminpres: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = minattr(k%, j%): rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = rrulli: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mmindevu: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = mminpresu: rr5 = rr5 + 1
               Sheets("STORE05").Cells(rr5, cc5) = test$: rr5 = rr5 + 1
           End If
           Return

assegna_rulli:
             'assegnazione automatica dei rulli in base ai carichi ammissibili
             If mmaxpres > 0 And mminpres > 0 Then
                               'appoggi
                                numrulli = mmaxpres / (cmaxapp / 10)
             End If
             If mmaxpres < 0 And mminpres < 0 Then
                               'ritenute
                                numrulli = Abs(mmaxpres / (cmaxrit / 10))
              End If
              If (mmaxpres > 0 And mminpres < 0) Or (mmaxpres < 0 And mminpres > 0) Then
                               'doppio effetto
                                If DEcarmax = 0 Then DEcarmax = cmaxrit
                                numrulli = Abs(mmaxpres / (DEcarmax / 10))
              End If
              rrulli = Int(numrulli / 2) * 2 + 2
              If k% = 1 Then
                            r1(j%) = rrulli
                            Sheets("F03").Cells(j% + 14, 8) = rrulli
              Else
                   r2(j%) = rrulli
                   Sheets("F03").Cells(j% + 14, 15) = rrulli
              End If
              Return

leggiPMMFMM:
           If flag% <> 9 Then
                    cv$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    cm$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxtv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxtm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    Ffmmax = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxav = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxam = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmintv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmintm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    Ffmmin = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mminav = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mminam = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    If fine% = 0 Then
                        sost$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxtp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxdev = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxps = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxattr = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmintp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmindev = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mminps = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mminattr = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    End If
                    If lll% = 2 Or lll% = 8 Or lll% = 14 Or lll% = 20 Then
                                angval = (mmaxav + mminav) / 2
                                angmon = (mmaxam + mminam) / 2
                    End If
           Else
               cv$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               cm$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxtv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxtm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               Ffmmax = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxav = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxam = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               If fine% = 0 Then
                        sost$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxtp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxdev = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxps = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxattr = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        If lll% = 2 Or lll% = 8 Or lll% = 14 Or lll% = 20 Then
                                angval = mmaxav
                                angmon = mmaxam
                        End If
               End If
           End If
           GoSub organizza
           Return

organizza:
          'tensioni,frecce,pressioni,deviazioni e attriti sui sostegni
          If mmaxtv > maxtv(k%, j%) Then maxtv(k%, j%) = mmaxtv
          If mmaxtm > maxtv(k%, j%) Then maxtv(k%, j%) = mmaxtm
          If Ffmmax > Fmmax(k%, j%) Then Fmmax(k%, j%) = Ffmmax
          If mmaxav > maxav(k%, j%) Then maxav(k%, j%) = mmaxav
          If mmaxam > maxam(k%, j%) Then maxam(k%, j%) = mmaxam
          If flag% <> 9 Then
                If mmintv < mintv(k%, j%) Then mintv(k%, j%) = mmintv
                If mmintm < mintv(k%, j%) Then mintv(k%, j%) = mmintm
                If Ffmmin < Fmmin(k%, j%) Then Fmmin(k%, j%) = Ffmmin
                If mminav < minav(k%, j%) Then minav(k%, j%) = mminav
                If mminam < minam(k%, j%) Then minam(k%, j%) = mminam
          Else
                If mmaxtv < mintv(k%, j%) Then mintv(k%, j%) = mmaxtv
                If mmaxtm < mintv(k%, j%) Then mintv(k%, j%) = mmaxtm
                If Ffmmax < Fmmin(k%, j%) Then Fmmin(k%, j%) = Ffmmax
                If mmaxav < minav(k%, j%) Then minav(k%, j%) = mmaxav
                If mmaxam < minam(k%, j%) Then minam(k%, j%) = mmaxam
          End If
          If mmaxtp > maxtp(k%, j%) Then maxtp(k%, j%) = mmaxtp
          If mmaxdev > maxdev(k%, j%) Then maxdev(k%, j%) = mmaxdev
          If mmaxps > maxps(k%, j%) Then maxps(k%, j%) = mmaxps
          If mmaxattr > maxattr(k%, j%) Then maxattr(k%, j%) = mmaxattr
          If flag% <> 9 Then
                          If mmintp < mintp(k%, j%) Then mintp(k%, j%) = mmintp
                          If mmindev < mindev(k%, j%) Then mindev(k%, j%) = mmindev
                          If mminps < minps(k%, j%) Then minps(k%, j%) = mminps
                          If mminattr < minattr(k%, j%) Then minattr(k%, j%) = mminattr
          Else
               If mmaxtp < mintp(k%, j%) Then mintp(k%, j%) = mmaxtp
               If mmaxdev < mindev(k%, j%) Then mindev(k%, j%) = mmaxdev
               If mmaxps < minps(k%, j%) Then minps(k%, j%) = mmaxps
               If mmaxattr < minattr(k%, j%) Then minattr(k%, j%) = mmaxattr
          End If
          If lll% = 2 Or lll% = 8 Or lll% = 14 Or lll% = 20 Then
               If angval > maxval(k%, j%) Then maxval(k%, j%) = angval
               If angval < minval(k%, j%) Then minval(k%, j%) = angval
               If angmon > maxmon(k%, j%) Then maxmon(k%, j%) = angmon
               If angmon < minmon(k%, j%) Then minmon(k%, j%) = angmon
          End If
          Return

Exit Sub

err_analisi:
           Beep
           MsgBox ("Error : " + cekerr + "   missing")
           Close
           Resume Next
           Exit Sub
End Sub


Public Sub vedi_linea()
        Call leggi_terreno
        'registra dati profilo e linea
        'salita
        Open "c:\TMP_SIF\transi-prf.txt" For Output As #1
        Write #1, nbs
        For j = 1 To nbs
                 Write #1, quot_terr(1, j), prog_terr(1, j)
        Next j
        For j = 1 To nbd
                 Write #1, quot_terr(2, j), prog_terr(2, j)
        Next j
        Close
        
        Open "c:\TMP_SIF\transi-lin.txt" For Output As #1
        Write #1, ns, nd
        'lettura ramo salita
        For j = 1 To ns
                 Write #1, quota(1, j), prog(1, j), n$(1, j), rulli(1, j), rulli(2, j)
        Next j
        For j = 1 To ns
                 Write #1, quota(1, j) - hv(1, j)
        Next j
        'lettura ramo discesa
        For j = 1 To nd
                 Write #1, quota(2, j), prog(2, j), n$(2, j), rulli(2, j)
        Next j
        For j = 1 To nd
                 Write #1, quota(2, j) - hv(2, j)
        Next j
        Close
        
End Sub

Public Sub codifica(tstringa)
        Dim lstringa As Integer
        lstringa = InStr(tstringa, ",")
        If lstringa > 0 Then
                  Mid(tstringa, lstringa, 1) = "."
        End If
End Sub

Public Sub tempus()
        Dim stringa, acapo, msg As String
        acapo = Chr(13) + Chr(10)
        stringa = "This instruction manual and the related software have benefited of the" + acapo
        stringa = stringa + "funding of the European Commission, Tempus Programme, Education and Culture." + acapo
        stringa = stringa + "The author is entirely responsible of the content of this manual and the software," + acapo
        stringa = stringa + "which do not necessarily reflect neither the view of the European Commission nor that of National Governments." + acapo
        Beep
        msg = MsgBox(stringa, 64, "Tempus")
End Sub

Sub Set_Ground_Profile()
On Local Error GoTo err_refresh

    'refresh grafici profilo
    
        Call leggi_terreno
        p1start = prog_terr(1, 1)
        q1start = quot_terr(1, 1)
        p1end = prog_terr(1, nbs)
        q1end = quot_terr(1, nbs) + 15
        p2start = prog_terr(2, 1)
        q2start = quot_terr(2, 1)
        p2end = prog_terr(2, nbd)
        q2end = quot_terr(2, nbd) + 15
        
        
    'libera protezione
    ActiveSheet.Unprotect ("vit210147vit")
    
    'aggiorna grafico profilo ramo SALITA---------------------------------------------------
    ActiveSheet.ChartObjects(1).Activate
    
    'ActiveChart.PlotArea.Select
    'Selection.Width = 536
    'Selection.Height = 343
    
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = q1start
        .MaximumScale = q1end                                   'pixel 536 / 343 rapporto base-altezza area grafico
        .MajorUnit = (q1end - q1start) / 10
    End With
    
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = p1start
        .MaximumScale = p1end
        .MajorUnit = (p1end - p1start) / 10
    End With
    
    Sheets("F02").Cells(1, 1).Select
    
    If nbd > 0 Then
    'aggiorna grafico profilo ramo DISCESA--------------------------------------------------
    ActiveSheet.ChartObjects(2).Activate
    
    'ActiveChart.PlotArea.Select
    'Selection.Width = 536
    'Selection.Height = 343
    
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = q2start
        .MaximumScale = q2end
        .MajorUnit = (q2end - q2start) / 10
    End With
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = p2start
        .MaximumScale = p2end
        .MajorUnit = (p2end - p2start) / 10
    End With
    End If
    
    Sheets("F02").Cells(1, 1).Select
    ActiveSheet.Protect password:="vit210147vit"
    
    
    Exit Sub
    
err_refresh:
            Beep
            MsgBox ("REFRESH : ERROR")
            Resume Next
            
End Sub
Public Sub stampa_maxmin(Ipocalc)
'On Local Error GoTo err_stampa_maxmin

           Dim area, areastampa, areaprint, verifica As String
           Dim ok As Integer
           Dim result As Double
           
           ok = 0
           'leggi da archivio i valori calcolati e stampa su foglio
           fgl = "F10"
           Sheets(fgl).Range("B30:P300").Select
           Selection.ClearContents
           Sheets(fgl).Range("B23:P26").Select
           Selection.ClearContents
           Sheets(fgl).Cells(1, 1).Select
           Sheets(fgl).Cells(30, 2).Select
           
           Call salti_pagina(fgl)
           
           Sheets(fgl).Cells(23, 2) = Sheets("F01").Cells(3, 4)
           Sheets(fgl).Cells(23, 9) = Sheets("F05").Cells(5, 11)
           Sheets(fgl).Cells(24, 9) = Sheets("F05").Cells(6, 11)
           Sheets(fgl).Cells(25, 9) = Sheets("F05").Cells(7, 11)
           Sheets(fgl).Cells(23, 13) = Sheets("F05").Cells(5, 15)
           Sheets(fgl).Cells(24, 13) = Sheets("F05").Cells(6, 15)
           Sheets(fgl).Cells(25, 13) = Sheets("F05").Cells(7, 15)
           Sheets(fgl).Cells(24, 16) = Date
           Sheets(fgl).Cells(25, 16) = Time
           
           ok = 1
           Select Case Ipocalc
                  Case 1
                       'calcolo normale
                       verifica = Sheets("F10").ipo_01.Caption
                       Sheets(fgl).Cells(24, 2) = verifica
                       rr555 = 3
                  Case 2
                       'calcolo +10%
                       verifica = Sheets("F10").ipo_02.Caption
                       Sheets(fgl).Cells(24, 2) = verifica
                       rr555 = 82
                  Case 3
                       'calcolo -10%
                       verifica = Sheets("F10").ipo_03.Caption
                       Sheets(fgl).Cells(24, 2) = verifica
                       rr555 = 122
                  Case 4
                       'calcolo -10%
                       verifica = Sheets("F10").ipo_04.Caption
                       Sheets(fgl).Cells(24, 2) = verifica
                       rr555 = 42
           End Select
           ok = 0
           
           rr% = 30
           Sheets(fgl).Cells(rr%, 2) = Sheets("F13").Cells(1, 8)
           rr% = rr% + 1
           
           Open "c:\TMP_SIF\transi-fmm.txt" For Output As #6
           'leggi la matrice ipotesi di calcolo (SiNo%(15,6)
           For numr = 6 To 20
              For numc = 7 To 12
                   kk% = numr - 5
                   xx% = numc - 6
                   ipoipo = (kk% - 1) * 6 + xx%
                   titoloipo = Sheets("F13").Cells(kk%, 1) + " : " + Sheets("F13").Cells(xx%, 5)
                   SiNo%(kk%, xx%) = 0
                   If Cells(numr, numc) = "OOO" Then SiNo%(kk%, xx%) = 1
                   If Cells(numr, numc) = "XXX" Then SiNo%(kk%, xx%) = 2
                   If SiNo%(kk%, xx%) = 1 Then
                                              Write #6, titoloipo, ipoipo
                   End If
              Next numc
           Next numr
           Write #6, "-------------------------------------", 0
           rr5 = rr555
           n1% = Sheets("STORE05").Cells(rr5, 2): rr5 = rr5 + 1
           n2% = Sheets("STORE05").Cells(rr5, 2): rr5 = rr5 + 1
           Write #6, n1%, n2%
           fine% = 0
           For j% = 2 To n1%
                          rr5 = rr555 + 2
                          cc5 = 1 + j%
                          If j% = n1% Then fine% = 1
                          GoSub leggi_maxmin
           Next j%
           
           Cells(rr%, 1).Select
           'inserisci salto pagina
           ActiveWindow.SelectedSheets.HPageBreaks.add Before:=ActiveCell
           
           Sheets(fgl).Cells(rr%, 2) = Sheets("F13").Cells(2, 8)
           rr% = rr% + 1
           fine% = 0
           For j% = 2 To n2%
                          rr5 = rr555 + 2
                          cc5 = 82 + j%
                          If j% = n2% Then fine% = 1
                          GoSub leggi_maxmin
           Next j%
           Close
           
           'imposta l'area di stampa
           area = Format(rr%, "####")
           areastampa = "A23:P" + area
           areaprint = "$A$23:$P$" + area
           Range(areastampa).Select
           ActiveSheet.PageSetup.PrintArea = areaprint
           Sheets(fgl).Cells(30, 2).Select
           
           'organizza file transito per le singole ipotesi
           For kk% = 1 To 15
              For xx% = 1 To 6
                 If SiNo%(kk%, xx%) = 1 Then
                    ipoipo = (kk% - 1) * 6 + xx%
                    Select Case Ipocalc
                           Case 1
                                 rr666 = 5
                           Case 2
                                 rr666 = 4005
                           Case 3
                                 rr666 = 6005
                           Case 4
                                 rr666 = 2005
                    End Select
                    rr6 = rr666: cc6 = 1 + ipoipo
                    flagmed% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    n1% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    n2% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    Open "c:\TMP_SIF\transi-fmm-" + Format(ipoipo, "##") + ".txt" For Output As #6
                    Write #6, n1%, n2%
                    fine% = 0
                    'salita
                    rr6 = rr666 + 3: cc6 = 1 + ipoipo
                    For j% = 2 To n1%
                         If j% = n1% Then fine% = 1
                         GoSub leggi_dati_ipotesi
                    Next j%
                    fine% = 0
                    'discesa
                    rr6 = rr666 + 3: cc6 = 92 + ipoipo
                    For j% = 2 To n2%
                         If j% = n2% Then fine% = 1
                         GoSub leggi_dati_ipotesi
                    Next j%
                    Close
                 End If
              Next xx%
           Next kk%
           
        'visualizza la linea (modulo VB6 view_line.exe)
        If Sheets("F10").test_grafico.value Then
             Open "c:\tmp_sif\prg_path" For Input As #1
             Input #1, mia_path
             Close
                   'imposta la lingua di riferimento
                   Open "c:\tmp_sif\flag_type" For Output As #1
                   If Sheets("HOME").flag_italiano Then
                                            Write #1, "IT"
                   Else
                        Write #1, "UK"
                   End If
                   Close
                   result = Shell(mia_path + "\view_line.exe", 1)
        End If
           
           
        Exit Sub
           
'---------------------------------------ROUTINE gosub/return ----------------------------------------------------

err_stampa_maxmin:
                 If ok = 0 Then
                               Beep
                               MsgBox ("PRINT SHEET F10 : ERROR TYPE " + Str(Err) + "  " + verifica)
                               Exit Sub
                 End If
                 If ok = 1 Then
                               Beep
                               MsgBox ("PRINT SHEET F10 : MISSIG VERIFY FILE" + "  " + verifica)
                               Exit Sub
                 Else
                     Resume Next
                 End If
                 Return
leggi_maxmin:
           cval = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           cmon = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           ttcc = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           ffcc = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           avv = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           amm = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           Write #6, ffcc
           GoSub scrivi_celle_1
           ttcc = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           ffcc = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           avv = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           amm = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
           Write #6, ffcc
           cval = "": cmon = ""
           GoSub scrivi_celle_1
           If fine% = 0 Then
               nsost = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               tpp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               devpp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               pspp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               attrpp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               rrulli = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               devupp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               presupp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               test$ = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               GoSub scrivi_celle_2
               nsost = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               tpp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               devpp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               pspp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               attrpp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               rrulli = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               devupp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               presupp = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               test$ = Sheets("STORE05").Cells(rr5, cc5): rr5 = rr5 + 1
               GoSub scrivi_celle_2
           End If
           Return
           
leggi_dati_ipotesi:
           cval = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           cmon = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           tvv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           tmm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           ffcc = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           avv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           amm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           Write #6, ffcc
           If flagmed% = 0 Then
                               tvv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               tmm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               ffcc = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               avv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               amm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
           End If
           Write #6, ffcc
           If fine% = 0 Then
                            nsost = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                            tpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                            devpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                            pspp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                            attrpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                            If flagmed% = 0 Then
                               tpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               devpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               pspp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                               attrpp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                            End If
           End If
           Return
           
scrivi_celle_1:
           Sheets(fgl).Cells(rr%, 2) = cval
           Sheets(fgl).Cells(rr%, 3) = cmon
           Sheets(fgl).Cells(rr%, 4) = CDbl(ttcc)
           Sheets(fgl).Cells(rr%, 5) = CDbl(ffcc)
           Sheets(fgl).Cells(rr%, 6) = CDbl(avv * 180 / PI#)
           Sheets(fgl).Cells(rr%, 7) = CDbl(amm * 180 / PI#)
           rr% = rr% + 1
           Return
scrivi_celle_2:
           Sheets(fgl).Cells(rr%, 8) = nsost
           Sheets(fgl).Cells(rr%, 9) = CDbl(tpp)
           Sheets(fgl).Cells(rr%, 10) = CDbl(devpp * 180 / PI#)
           Sheets(fgl).Cells(rr%, 11) = CDbl(pspp)
           Sheets(fgl).Cells(rr%, 12) = CDbl(attrpp)
           Sheets(fgl).Cells(rr%, 13) = CDbl(rrulli)
           Sheets(fgl).Cells(rr%, 14) = CDbl(devupp * 180 / PI#)
           Sheets(fgl).Cells(rr%, 15) = CDbl(presupp)
           Sheets(fgl).Cells(rr%, 16) = test$
           rr% = rr% + 1
           Return

End Sub

Public Sub stampa_potenze(Ipocalc, nomeipotesi$)

           Dim area, areastampa, areaprint As String
           
           'leggi da archivio i valori calcolati e stampa su foglio
           fgl = "F11"
           path$ = "C:\TMP_SIF\"
           nome$ = "valori"
           Call assegna_generali
           
           Sheets(fgl).Range("B30:Q200").Select
           Selection.ClearContents
           Sheets(fgl).Range("B23:Q26").Select
           Selection.ClearContents
           Sheets(fgl).Cells(1, 1).Select
           Sheets(fgl).Cells(30, 2).Select
           
           Call salti_pagina(fgl)
           Sheets(fgl).Cells(23, 2) = Sheets("F01").Cells(3, 4)
           Sheets(fgl).Cells(23, 9) = Sheets("F05").Cells(5, 11)
           Sheets(fgl).Cells(24, 9) = Sheets("F05").Cells(6, 11)
           Sheets(fgl).Cells(25, 9) = Sheets("F05").Cells(7, 11)
           Sheets(fgl).Cells(23, 13) = Sheets("F05").Cells(5, 15)
           Sheets(fgl).Cells(24, 13) = Sheets("F05").Cells(6, 15)
           Sheets(fgl).Cells(25, 13) = Sheets("F05").Cells(7, 15)
           Sheets(fgl).Cells(24, 16) = Date
           Sheets(fgl).Cells(25, 16) = Time
           
           rr% = 30
           
        path$ = "C:\TMP_SIF\"
        nome$ = "valori"
        
        If Sheets("F05").Option1_maxmin.value Then
                                                    KVL = "STI"
        Else
            KVL = "F"
        End If
        
        j% = 90
        ReDim numipo(j%), Ftmed(j%), inargano(j%), Sfmed(j%), rargmed(j%), Potm(j%), maxscorr(j%), svilmax(j%), svilmin(j%), Tiro(j%), dTiro(j%)
        ReDim Sfmax(j%), Sfmin(j%), MSforzo(j%), rargmax(j%), Potmax(j%)

        ReDim dnumipo(j%), dFtmed(j%), dinargano(j%), dSfmed(j%), drargmed(j%), dPotm(j%), dmaxscorr(j%), dsvilmax(j%), dsvilmin(j%)
        ReDim dSfmax(j%), dSfmin(j%), dMSforzo(j%), drargmax(j%), dPotmax(j%)
        
        aaa$ = "###"
        Select Case Ipocalc
               Case 1
                     rr1 = 4:
                     rr3 = 100:
               Case 2
                     rr1 = 200:
                     rr3 = 300:
               Case 3
                     rr1 = 400:
                     rr3 = 500:
               Case 4
                     rr1 = 600:
                     rr3 = 700:
        End Select
        
        For j% = 1 To 90
                        cc12 = 187
                        numipo(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Ftmed(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        inargano(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Sfmed(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        rargmed(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Potm(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        maxscorr(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        svilmax(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        svilmin(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Tiro(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Sfmax(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Sfmin(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        MSforzo(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        rargmax(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        Potmax(j%) = Sheets("STORE13").Cells(rr1, cc12): cc12 = cc12 + 1
                        rr1 = rr1 + 1
        Next j%
        'lettura valori per marcia avanti
        
        For j% = 1 To 90
                        cc12 = 187
                        dnumipo(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dFtmed(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dinargano(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dSfmed(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        drargmed(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dPotm(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dmaxscorr(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dsvilmax(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dsvilmin(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dTiro(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dSfmax(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dSfmin(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dMSforzo(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        drargmax(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        dPotmax(j%) = Sheets("STORE13").Cells(rr3, cc12): cc12 = cc12 + 1
                        rr3 = rr3 + 1
        Next j%
        
        
        For j% = 1 To 6 Step 1
           For ii% = 1 To 15
                w% = (ii% - 1) * 6 + j%
                If numipo(w%) > 0 Then
                            'marcia avanti
                            Sheets("F11").Cells(rr%, 2) = ">> : " + Sheets("F13").Cells(ii%, 1)
                            Sheets("F11").Cells(rr%, 6) = " : " + Sheets("F13").Cells(j%, 5)
                            Sheets("F11").Cells(rr%, 9) = CDbl(Ftmed(w%))
                            Sheets("F11").Cells(rr%, 10) = CDbl(Sfmax(w%))
                            Sheets("F11").Cells(rr%, 11) = CDbl(inargano(w%))
                            Sheets("F11").Cells(rr%, 12) = CDbl(Sfmed(w%))
                            Sheets("F11").Cells(rr%, 13) = CDbl(rargmed(w%))
                            Sheets("F11").Cells(rr%, 14) = CDbl(Potm(w%))
                            Sheets("F11").Cells(rr%, 15) = CDbl(maxscorr(w%))
                            Sheets("F11").Cells(rr%, 16) = CDbl(svilmax(w%))
                            Sheets("F11").Cells(rr%, 17) = CDbl(Tiro(w%))
                            rr% = rr% + 1
                            If KVL = "STI" Then
                              Sheets("F11").Cells(rr%, 9) = ""
                              Sheets("F11").Cells(rr%, 10) = CDbl(Sfmin(w%))
                              Sheets("F11").Cells(rr%, 11) = CDbl(inargano(w%))
                              Sheets("F11").Cells(rr%, 12) = CDbl(MSforzo(w%))
                              Sheets("F11").Cells(rr%, 13) = CDbl(rargmax(w%))
                              Sheets("F11").Cells(rr%, 14) = CDbl(Potmax(w%))
                              Sheets("F11").Cells(rr%, 15) = ""
                              Sheets("F11").Cells(rr%, 16) = CDbl(svilmin(w%))
                              Sheets("F11").Cells(rr%, 17) = ""
                              rr% = rr% + 1
                            End If
                            If dnumipo(w%) > 0 Then
                                 'marcia indietro
                                 Sheets("F11").Cells(rr%, 2) = "<< : " + Sheets("F13").Cells(ii%, 1)
                                 Sheets("F11").Cells(rr%, 6) = " : " + Sheets("F13").Cells(j%, 5)
                                 Sheets("F11").Cells(rr%, 9) = CDbl(dFtmed(w%))
                                 Sheets("F11").Cells(rr%, 10) = CDbl(dSfmax(w%))
                                 Sheets("F11").Cells(rr%, 11) = CDbl(dinargano(w%))
                                 Sheets("F11").Cells(rr%, 12) = CDbl(dSfmed(w%))
                                 Sheets("F11").Cells(rr%, 13) = CDbl(drargmed(w%))
                                 Sheets("F11").Cells(rr%, 14) = CDbl(dPotm(w%))
                                 Sheets("F11").Cells(rr%, 15) = CDbl(dmaxscorr(w%))
                                 Sheets("F11").Cells(rr%, 16) = CDbl(dsvilmax(w%))
                                 Sheets("F11").Cells(rr%, 17) = CDbl(dTiro(w%))
                                 rr% = rr% + 1
                                 If KVL = "STI" Then
                                   Sheets("F11").Cells(rr%, 9) = ""
                                   Sheets("F11").Cells(rr%, 10) = CDbl(dSfmin(w%))
                                   Sheets("F11").Cells(rr%, 11) = CDbl(dinargano(w%))
                                   Sheets("F11").Cells(rr%, 12) = CDbl(dMSforzo(w%))
                                   Sheets("F11").Cells(rr%, 13) = CDbl(drargmax(w%))
                                   Sheets("F11").Cells(rr%, 14) = CDbl(dPotmax(w%))
                                   Sheets("F11").Cells(rr%, 15) = ""
                                   Sheets("F11").Cells(rr%, 16) = CDbl(dsvilmin(w%))
                                   Sheets("F11").Cells(rr%, 17) = ""
                                   rr% = rr% + 1
                                 End If
                            End If
                    End If
              Next ii%
        Next j%
        rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(20, 1)
        Sheets("F11").Cells(rr%, 7) = contr: rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(21, 1)
        Sheets("F11").Cells(rr%, 7) = acc: rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(22, 1)
        Sheets("F11").Cells(rr%, 7) = dec: rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(23, 1)
        Sheets("F11").Cells(rr%, 7) = fren: rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(24, 1)
        Sheets("F11").Cells(rr%, 7) = rap: rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(25, 1)
        Sheets("F11").Cells(rr%, 7) = rendarg: rr% = rr% + 1
        Sheets("F11").Cells(rr%, 2) = Sheets("F13").Cells(26, 1)
        Sheets("F11").Cells(rr%, 7) = 1 / rendarg: rr% = rr% + 1
        If corsacil <> 0 Then
              Sheets("F11").Cells(rr%, 7) = corsacil: rr% = rr% + 1
        End If
           
        Sheets("F11").Cells(24, 2) = nomeipotesi$
        
        'imposta l'area di stampa
        area = Format(rr%, "####")
        areastampa = "A1:Q" + area
        areaprint = "$A$1:$Q$" + area
        Range(areastampa).Select
        ActiveSheet.PageSetup.PrintArea = areaprint
        Sheets(fgl).Cells(30, 2).Select
        
End Sub

Public Sub rel_gen()
        
On Local Error GoTo err_F12
        
   nomefoglio = "F12"
   Call elimina_dati_generali(nomefoglio)
        Call assegna_generali
        Call leggi_linea
        GoSub sviluppata
        Hveicolo = Sheets("F01").Cells(81, 5)
        Lveicolo = Sheets("F01").Cells(82, 5)
        Iveicolo = Sheets("F01").Cells(83, 5)
        Angavv = Sheets("F01").Cells(84, 5)
        Kfunpul = Sheets("F01").Cells(85, 5)
        
        'leggi la matrice ipotesi di calcolo (SiNo%(15,6)
        For numr = 10 To 24
              For numc = 4 To 9
                   kk% = numr - 9
                   xx% = numc - 3
                   SiNo%(kk%, xx%) = 0
                   If Sheets("F05").Cells(numr, numc) = "OOO" Then
                                                    SiNo%(kk%, xx%) = 1
                   End If
                   If Sheets("F05").Cells(numr, numc) = "XXX" Then SiNo%(kk%, xx%) = 2
              Next numc
        Next numr
        
        If ZETA = "F" Then
                    eg = Sheets("F05").Cells(5, 15)
                    Ev = eg
        Else
            If ZETA <> "I" Then eg = dd: Ev = eg
        End If

       rr% = 7
       cc% = 5
       
       '-------------< dati generali >----------------
       Cells(rr%, cc%) = nome$: rr% = rr% + 1
       Cells(rr%, cc%) = descrizione$: rr% = rr% + 1
       If m$ = "MONTE" And tt$ = "VALLE" Then tipstaZETA = Sheets("F13").Cells(10, 5)
       If m$ = "MONTE" And tt$ = "MONTE" Then tipstaZETA = Sheets("F13").Cells(10, 6)
       If m$ = "VALLE" And tt$ = "VALLE" Then tipstaZETA = Sheets("F13").Cells(10, 7)
       If m$ = "VALLE" And tt$ = "MONTE" Then tipstaZETA = Sheets("F13").Cells(10, 8)
       Cells(rr%, cc%) = tipstaZETA: rr% = rr% + 1
       Cells(rr%, cc%) = contr * 10: rr% = rr% + 3
       Cells(rr%, cc%) = Loavan1: rr% = rr% + 1
       Cells(rr%, cc%) = Lsavan1: rr% = rr% + 1
       Cells(rr%, cc%) = Lovm1: rr% = rr% + 1
       Cells(rr%, cc%) = Lsvm1: rr% = rr% + 1
       Cells(rr%, cc%) = sc * 2: rr% = rr% + 1
       Cells(rr%, cc%) = Ddisliv: rr% = rr% + 1
       Cells(rr%, cc%) = Ddisliv / Loavan1 * 100: rr% = rr% + 1
       Cells(rr%, cc%) = ns - 4: rr% = rr% + 1
       Cells(rr%, cc%) = sensomarcia$: rr% = rr% + 1
       Cells(rr%, cc%) = scart: rr% = rr% + 1
       Cells(rr%, cc%) = Intstaz: rr% = rr% + 1
       Cells(rr%, cc%) = nveic: rr% = rr% + 1
       If ZETA <> "F" And ZETA <> "I" Then
              'calcolo dei veicoli in stazione per impianti automatici
              'tempo di transito
              intervallo = Ev / vel
              transitoval = Sheets("F01").Cells(86, 5)
              transitomon = Sheets("F01").Cells(87, 5)
              nveicstazval = Int(transitoval / intervallo)
              nveicstazmon = Int(transitomon / intervallo)
       Else
              nveicstazval = 0
              nveicstazmon = 0
       End If
       Cells(rr%, cc%) = nveic + nveicstazval + nveicstazmon: rr% = rr% + 1
       Cells(rr%, cc%) = Ev: rr% = rr% + 1
       Cells(rr%, cc%) = Ev / vel: rr% = rr% + 1
       tsec = Lsavan1 / vel
       tmin = Int(tsec / 60): If tmin > 0 Then tsec = tsec - tmin * 60
       tpercor$ = Format$(tmin, "##") + ":" + Format$(tsec, "#0.0")
       fbuco = buco * qv * Sin(Atn(Ddisliv / Loavan1)) * 9.81
       bbuco$ = Format$(buco, "#0") + " --> F = " + Format$(fbuco, "#####0") + " N"
       Cells(rr%, cc%) = tpercor$: rr% = rr% + 1
       Cells(rr%, cc%) = vel: rr% = rr% + 1
       Cells(rr%, cc%) = port: rr% = rr% + 1
       Cells(rr%, cc%) = bbuco$: rr% = rr% + 3
       Cells(rr%, cc%) = rullier$: rr% = rr% + 1
       Cells(rr%, cc%) = diaapp: rr% = rr% + 1
       Cells(rr%, cc%) = mapp: rr% = rr% + 1
       Cells(rr%, cc%) = cmaxapp: rr% = rr% + 1
       Cells(rr%, cc%) = rullier$: rr% = rr% + 1
       Cells(rr%, cc%) = diarit: rr% = rr% + 1
       Cells(rr%, cc%) = mrit: rr% = rr% + 1
       Cells(rr%, cc%) = cmaxrit: rr% = rr% + 1
       Cells(rr%, cc%) = DEtiprul$: rr% = rr% + 1
       Cells(rr%, cc%) = DEdiamrul: rr% = rr% + 1
       Cells(rr%, cc%) = DEmassarul: rr% = rr% + 1
       Cells(rr%, cc%) = DEcarmax: rr% = rr% + 3
       Cells(rr%, cc%) = tipoveic$: rr% = rr% + 1
       Cells(rr%, cc%) = posti: rr% = rr% + 1
       Cells(rr%, cc%) = qv: rr% = rr% + 1
       Cells(rr%, cc%) = qc: rr% = rr% + 3
       Cells(rr%, cc%) = tipof$: rr% = rr% + 1
       Cells(rr%, cc%) = dia: rr% = rr% + 1
       Cells(rr%, cc%) = mf: rr% = rr% + 1
       Cells(rr%, cc%) = area: rr% = rr% + 1
       Cells(rr%, cc%) = res * 1000 / area: rr% = rr% + 1
       Cells(rr%, cc%) = res: rr% = rr% + 3
       scorrmax = 2.71828 ^ (Angavv / 180 * 3.141 * Kfunpul)
       Cells(rr%, cc%) = pstazione * attr0 / 100 + Ftrave: rr% = rr% + 1
       Cells(rr%, cc%) = argano: rr% = rr% + 1
       Cells(rr%, cc%) = rendarg: rr% = rr% + 1
       Cells(rr%, cc%) = Angavv: rr% = rr% + 1
       Cells(rr%, cc%) = Kfunpul: rr% = rr% + 1
       Cells(rr%, cc%) = scorrmax: rr% = rr% + 1
       Cells(rr%, cc%) = acc: rr% = rr% + 1
       Cells(rr%, cc%) = dec: rr% = rr% + 1
       Cells(rr%, cc%) = fren: rr% = rr% + 1
       Cells(rr%, cc%) = rap: rr% = rr% + 4

       '------------------------< seconda parte >--------------------------

       GoSub set_maxmin
       ok% = 0
       marcia$ = "S"
       aaa$ = "###"
       If XIndex = 0 Or XIndex = 4 Or XIndex = 5 Then rr1313 = 4
       If XIndex = 1 Then rr1313 = 600
       If ok% = 0 Then
                     GoSub potenze
       End If
       Close #1
       marcia$ = "D"
       If XIndex = 0 Or XIndex = 4 Or XIndex = 5 Then rr1313 = 100
       If XIndex = 1 Then rr1313 = 700
       If ok% = 0 Then
                    GoSub potenze
       End If
       Close #1
       
       For w% = 1 To 15
           For u% = 1 To 6
               If SiNo%(w%, u%) = 1 Then
                                ll% = (w% - 1) * 6 + u%
                                rr6 = 5: cc6 = 1 + ll%
                                flag% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                                n1% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                                n2% = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                                k% = 1: fine% = 0: cc6 = 1 + ll%
                                For j% = 2 To n1%
                                               If j% = n1% Then fine% = 9
                                               GoSub tensioni
                                               If j% = 2 Then tvalle1 = mmaxtv
                                               If j% = n1% Then tmonte1 = mmaxtm
                                Next j%
                                k% = 2: fine% = 0: rr6 = 8: cc6 = 92 + ll%
                                For j% = 2 To n2%
                                               If j% = n2% Then fine% = 9
                                               GoSub tensioni
                                               If j% = 2 Then tvalle2 = mmaxtv
                                               If j% = n2% Then tmonte2 = mmaxtm
                                Next j%
                                Close #1
                                If tiromaxvalle <= (tvalle1 + tvalle2) Then tiromaxvalle = (tvalle1 + tvalle2)
                                If tiromaxmonte <= (tmonte1 + tmonte2) Then tiromaxmonte = (tmonte1 + tmonte2)
               End If
           Next u%
        Next w%

          Cells(rr%, cc%) = CDbl(tmax * 10): Cells(rr%, cc% + 1) = nmmax$: rr% = rr% + 1
          Cells(rr%, cc%) = CDbl((1000 * res) / (tmax * 10)): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(tmin * 10): Cells(rr%, cc% + 1) = nmmin$: rr% = rr% + 1
          pmorsa = qc * G# * 10
          If morse > 1 Then pmorsa = 1.2 * pmorsa / morse
          Cells(rr%, cc%) = pmorsa: rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(pmorsa ^ 2 / (tmin * 10) / area / 10): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(nreg): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(nmax): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(nmin): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(normale * 10): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(rapido) * 10: rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(spontaneo * 10): rr% = rr% + 1
          alltermico = sc * KTERMICO# * 50
          Cells(rr%, cc%) = CDbl(svmax - svmin): rr% = rr% + 2
          Cells(rr%, cc%) = alltermico: rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(grsmax): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(Log(grsmax) / 3.14): rr% = rr% + 2
          Cells(rr%, cc%) = CDbl(minarullo): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(minarulliera): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(minrrullo): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(minrrulliera): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(maxapp): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(maxapp / dia / (diaapp + dia)): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(maxrit): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(-maxrit / dia / (diarit + dia) / 0.8): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(devmax * 180 / PI#): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(pendenza * 180 / PI#): rr% = rr% + 1
          Cells(rr%, cc%) = qc / morse * G# * 10 * Sin(pendenza): rr% = rr% + 1
          fventosufune = vento2 * dia * 1.1 / 1000 * Lcmax ^ 2 / 8 / (tfreccia * 10)
          If UCase$(ZETA) <> "I" Then
                         fventosuveicoli = (vento2 * areac / Ev) * Lcmax ^ 2 / 8 / (tfreccia * 10)
          Else
              fventosuveicoli = (vento2 * areac * vg) * Lcmax / 4 / (tfreccia * 10)
          End If
          frecciavento = fventosufune + fventosuveicoli
          Cells(rr%, cc%) = CDbl(frecciavento): rr% = rr% + 1
          Cells(rr%, cc%) = ncc$: rr% = rr% + 1
          spostveicolo = Hveicolo * Sin(Iveicolo / 180 * PI#) + Lveicolo / 2 * Cos(Iveicolo / 180 * PI#)
          margine = (scart / 1000 - frecciavento) - 2 * spostveicolo
          Cells(rr%, cc%) = CDbl(margine): rr% = rr% + 1
          
          summar1 = 0: summar2 = 0
          sumfusto = 0
          For j% = 1 To ns
                         summar1 = summar1 + r1(j%)
          Next j%
          For j% = 1 To nd
                        summar2 = summar2 + r2(j%)
          Next j%
          Cells(rr%, cc%) = summar1: rr% = rr% + 1
          Cells(rr%, cc%) = summar2: rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(tiromaxvalle * 10): rr% = rr% + 1
          Cells(rr%, cc%) = CDbl(tiromaxmonte * 10): rr% = rr% + 1

       Close
   
Exit Sub
       
potenze:
       For ipoipo = 1 To 90
             rr13 = rr1313 - 1 + ipoipo: cc13 = 187
             If Sheets("STORE13").Cells(rr13, cc13) = ipoipo Then
                    ii% = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Ftmed = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    inargano = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Sfmed = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    rargmed = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Potm = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    maxscorr = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    svilmax = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    svilmin = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Tiro = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Sfmax = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Sfmin = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    MSforzo = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    rargmax = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    Potmax = Sheets("STORE13").Cells(rr13, cc13): cc13 = cc13 + 1
                    If grsmax < maxscorr Then grsmax = maxscorr
                    If ii% = 2 Or ii% = 8 Or ii% = 14 Or ii% = 20 Or ii% = 26 Or ii% = 32 Or ii% = 38 Or ii% = 44 Or ii% = 50 Or ii% = 56 Then
                              If ZETA <> "F" Then
                                       If nreg < Potmax Then nreg = Potmax
                              Else
                                   If nreg < Potm Then nreg = Potm
                              End If
                              If ii% = 2 And marcia$ = "D" Then
                                      If ZETA = "F" Then
                                                  If spontaneo > Sfmed Then spontaneo = Sfmed
                                      Else
                                          If spontaneo > Sfmax Then spontaneo = Sfmax
                                          If spontaneo > Sfmin Then spontaneo = Sfmin
                                      End If
                              End If
                    Else
                        If ii% = 5 Or ii% = 11 Or ii% = 17 Or ii% = 23 Or ii% = 29 Or ii% = 35 Or ii% = 41 Or ii% = 47 Or ii% = 53 Or ii% = 59 Then
                                                  If normale > MSforzo Then normale = MSforzo
                        End If
                        If nmax < Potmax Then nmax = Potmax
                        If nmin > Potmax Then nmin = Potmax
                        If rapido > MSforzo Then rapido = MSforzo
                    End If
                    If svmax < svilmax Then svmax = svilmax
                    If svmin > svilmin Then svmin = svilmin
             End If
       Next ipoipo
       Return

set_maxmin:
       tmax = 0                'tensione massima fune
       tmin = 99999999         'tensione minima fune
       nreg = 0                'potenza a regima
       nmax = 0                'potenza massima : spunto
       nmin = 0                'potenza negativa di punta
       normale = 99999         'sforzo frenante normale
       rapido = 99999          'sforzo frenante rapido
       grsmax = 0              'scorrimento massimo
       tiromaxvalle = 0        'massimo tiro ruota a valle
       tiromaxmonte = 0        'massimo tiro ruota a monte

       minarullo = 999999
       minrrullo = 999999
       minarulliera = 999999
       minrrulliera = 999999

       maxapp = 0
       maxrit = 0
       devmax = 0
       pendenza = 0
       tfreccia = 999999

       svmax = 0
       svmin = 999999

       Return


tensioni:
           If flag% <> 9 Then
                    cv$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    cm$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxtv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxtm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    Ffmmax = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxav = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmaxam = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmintv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mmintm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    Ffmmin = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mminav = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    mminam = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    If fine% = 0 Then
                        sost$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxtp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxdev = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxps = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxattr = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmintp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmindev = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mminps = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mminattr = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                    End If
           Else
               cv$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               cm$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxtv = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxtm = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               Ffmmax = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxav = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmaxam = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
               mmintv = mmaxtv: mmintm = mmaxtm: mminam = mmaxam
               If fine% = 0 Then
                        sost$ = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxtp = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxdev = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxps = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmaxattr = Sheets("STORE06").Cells(rr6, cc6): rr6 = rr6 + 1
                        mmintp = mmaxtp: mminps = mmaxps: mmindev = mmaxdev
               End If
           End If
           If tmax < mmaxtv Then tmax = mmaxtv: nmmax$ = n$(k%, j%)
           If tmax < mmaxtm Then tmax = mmaxtm: nmmax$ = n$(k%, j%)
           If tmax < mmaxtp Then tmax = mmaxtp: nmmax$ = n$(k%, j%)
           If tmin > mmintv Then tmin = mmintv: nmmin$ = n$(k%, j%)
           If tmin > mmintm Then tmin = mmintm: nmmin$ = n$(k%, j%)
           If tmin > mmintp Then tmin = mmintp: nmmin$ = n$(k%, j%)
           If nnc = j% And k% = nr% Then
                         If tfreccia > (mmintv + mmintm) / 2 Then tfreccia = (mmintv + mmintm) / 2
           End If
           ' pressioni rulli/rulliere,deviazioni unitarie
           p1 = mminps * 10
           p2 = mmaxps * 10
           If k% = 1 And r1(j%) > 0 Then
                         pu1 = p1 / r1(j%)
                         pu2 = p2 / r1(j%)
                         du = mmaxdev / r1(j%)
           Else
               If r2(j%) > 0 Then
                                 pu1 = p1 / r2(j%)
                                 pu2 = p2 / r2(j%)
                                 du = mmaxdev / r2(j%)
               End If
           End If
           'viene omessa l'eventuale avanstazione ("A" - "a")
           If InStr(n$(k%, j%), "A") > 0 Or InStr(n$(k%, j%), "a") > 0 Then Return
           If ZETA = "F" And j% = 2 Then Return
           If ZETA = "F" And k% = 1 And j% = n1% - 1 Then Return
           If ZETA = "F" And k% = 2 And j% = n2% - 1 Then Return
           If k% = 1 And r1(j%) = 0 Then Return
           If k% = 2 And r2(j%) = 0 Then Return
           If InStr(n$(k%, j%), "W") = 0 And InStr(n$(k%, j%), "w") = 0 Then
           'non viene presa in considerazione la rulliera a doppio effetto ("W" - "w")
              If p1 > 0 Then
                        If minarulliera > p1 Then minarulliera = p1
                        If minarullo > pu1 Then minarullo = pu1
                        If maxapp < pu2 Then maxapp = pu2
              Else
               If Abs(minrrulliera) > Abs(p1) Then minrrulliera = p1
               If Abs(minrrullo) > Abs(pu1) Then minrrullo = pu1
               If Abs(maxrit) < Abs(pu2) Then maxrit = pu2
              End If
           End If
           If Abs(devmax) < Abs(du) Then devmax = du
           If Abs(pendenza) < Abs(mmaxam) Then pendenza = mmaxam
       Return

sviluppata:
        'salita
         Ddisliv = 0
         Lovm1 = 0
         Lsvm1 = 0
         Lcmax = 0
         Loavan1 = 0           'lunghezza orizz.tra avanstazioni
         Lsavan1 = 0           'lunghezza inclinata tra avantazioni
         For X = 1 To ns - 1
                         CCorda = Sqr(d(1, X) ^ 2 + l(1, X) ^ 2)
                         Ddisliv = Ddisliv + d(1, X)
                         Lovm1 = Lovm1 + l(1, X)
                         Lsvm1 = Lsvm1 + CCorda
                         If Lcmax < CCorda Then Lcmax = CCorda: nnc = X: ncc$ = n$(1, X + 1) + " - " + n$(1, X): nr% = 1
                         If X > 1 And X < ns - 1 Then
                                       Loavan1 = Loavan1 + l(1, X)
                                       Lsavan1 = Lsavan1 + CCorda
                         End If
         Next X
        'discesa
        Lovm2 = 0
        Lsvm2 = 0
        Loavan2 = 0           'lunghezza orizz.tra avanstazioni
        Lsavan2 = 0           'lunghezza inclinata tra avantazioni
        For X = 1 To nd - 1
                         CCorda = Sqr(d(2, X) ^ 2 + l(2, X) ^ 2)
                         Lovm2 = Lovm2 + l(2, X)
                         Lsvm2 = Lsvm2 + CCorda
                         If Lcmax < CCorda Then Lcmax = CCorda: nnc = X: ncc$ = n$(2, X + 1) + " - " + n$(2, X): nr% = 2
                         If X > 1 And X < nd - 1 Then
                                       Loavan2 = Loavan2 + l(2, X)
                                       Lsavan2 = Lsavan2 + CCorda
                         End If
        Next X
        sc = (Lsvm1 + Lsvm2 + PI# * Drt / 1000 / 2 + PI# * Drtrin / 1000 / 2) / 2
        Return

err_F12:
        Beep
        MsgBox ("Errore nella generazione del Report : i dati potrebbero non essere corretti")
        Exit Sub
        
End Sub

