Attribute VB_Name = "Modulo2"
Dim sorgente, destinazione As String

Public Sub traduci()

    'traduzione MENU
    Sheets("MENU").Select
    GoSub libera
    
    If Sheets("HOME").flag_italiano Then
                                        cc = 2
    Else
                                        cc = 10
    End If
    Sheets("MENU").CommandButton1.Caption = Sheets("MENU").Cells(200, cc)
    Sheets("MENU").CommandButton2.Caption = Sheets("MENU").Cells(201, cc)
    Sheets("MENU").CommandButton3.Caption = Sheets("MENU").Cells(202, cc)
    Sheets("MENU").CommandButton4.Caption = Sheets("MENU").Cells(203, cc)
    Sheets("MENU").CommandButton6.Caption = Sheets("MENU").Cells(204, cc)
    Sheets("MENU").CommandButton7.Caption = Sheets("MENU").Cells(205, cc)
    Sheets("MENU").CommandButton5.Caption = Sheets("MENU").Cells(207, cc)
    Sheets("MENU").CommandButton8.Caption = Sheets("MENU").Cells(208, cc)
    Sheets("MENU").CommandButton9.Caption = Sheets("MENU").Cells(209, cc)
    Sheets("MENU").CommandButton10.Caption = Sheets("MENU").Cells(210, cc)
    Sheets("MENU").CommandButton11.Caption = Sheets("MENU").Cells(211, cc)
    Sheets("MENU").CommandButton12.Caption = Sheets("MENU").Cells(212, cc)
    Sheets("MENU").Aggiorna_Intestazione_Tabulati.Caption = Sheets("MENU").Cells(215, cc)
    Sheets("MENU").elimina_output.Caption = Sheets("MENU").Cells(216, cc)
    Sheets("MENU").elimina_input.Caption = Sheets("MENU").Cells(217, cc)
    Sheets("MENU").Cells(33, 5) = Sheets("MENU").Cells(218, cc)
    GoSub blocca
    
    'traduzione dati generali
    Sheets("F01").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                         rr1 = 200
                         rr2 = 285
                         cc1 = 2
                         cc2 = 3
    Else
                         rr1 = 300
                         rr2 = 385
                         cc1 = 2
                         cc2 = 3
    End If
    rr = 2
    For kk% = rr1 To rr2
      For xx% = cc1 To cc2
             Sheets("F01").Cells(rr, xx%) = Sheets("F01").Cells(kk%, xx%)
      Next xx%
      rr = rr + 1
    Next kk%
    Sheets("F01").Cells(32, 7) = Sheets("F01").Cells(32, 7)
    Sheets("F01").Cells(33, 7) = Sheets("F01").Cells(33, 7)
    Sheets("F01").Cells(34, 7) = Sheets("F01").Cells(34, 7)
    rr = rr1
    cc = 7
    rr = rr + 1
    Sheets("F01").F.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").s.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").T.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").i.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").rotaz_or.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").rotaz_antior.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").argano_mon.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").argano_val.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").tensione_mon.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").tensione_val.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").tipo_tend_grav.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").tipo_tend_idra.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").in_tensione.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").ancorato_fisso.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").car_salita.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    Sheets("F01").car_salita_discesa.Caption = Sheets("F01").Cells(rr, cc): rr = rr + 1
    If Sheets("HOME").flag_italiano Then
                                       rr1 = 230
                                       rr2 = 232
                                       cc1 = 7
    Else
                         rr1 = 330
                         rr2 = 332
                         cc1 = 7
    End If
    rr = 32
    cc = 7
    For kk% = rr1 To rr2
             Sheets("F01").Cells(rr, cc) = Sheets("F01").Cells(kk%, cc1)
             rr = rr + 1
    Next kk%
    Sheets("F01").Cells(1, 1).Select
    GoSub blocca
    
    'traduzione terreno
    Sheets("F02").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                       rr1 = 1208
                                       rr2 = 1211
                                       rr3 = 1201
                                       rr4 = 1200
                                       cc1 = 2
                                       cc2 = 15
    Else
                 rr1 = 1223
                 rr2 = 1226
                 rr3 = 1216
                 rr4 = 1215
                 cc1 = 2
                 cc2 = 15
    End If
    rr = 11
    For kk% = rr1 To rr2
      For xx% = cc1 To cc2
             Sheets("F02").Cells(rr, xx%) = Sheets("F02").Cells(kk%, xx%)
      Next xx%
      rr = rr + 1
    Next kk%
    Sheets("F02").ins_rigaprof.Caption = Sheets("F02").Cells(rr3, 2)
    Sheets("F02").ins_rigaprof_d.Caption = Sheets("F02").Cells(rr3, 2): rr3 = rr3 + 1
    Sheets("F02").del_rigaprof.Caption = Sheets("F02").Cells(rr3, 2)
    Sheets("F02").del_rigaprof_d.Caption = Sheets("F02").Cells(rr3, 2): rr3 = rr3 + 1
    Sheets("F02").svuota_prof.Caption = Sheets("F02").Cells(rr3, 2)
    Sheets("F02").svuota_prof_d.Caption = Sheets("F02").Cells(rr3, 2): rr3 = rr3 + 1
    Sheets("F02").vedi_profilo_salita.Caption = Sheets("F02").Cells(rr3, 2): rr3 = rr3 + 1
    Sheets("F02").vedi_profilo_discesa.Caption = Sheets("F02").Cells(rr3, 2): rr3 = rr3 + 1
    rr = 5
    Sheets("F02").Cells(rr, 2) = Sheets("F02").Cells(rr4, 2)
    Sheets("F02").Cells(rr, 7) = Sheets("F02").Cells(rr4, 7)
    Sheets("F02").Cells(16, 1).Select
    Sheets("F02").Cells(15, 1).Select
    GoSub blocca
    
    'traduzione input linea (F03) assoluti
    Sheets("F03").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                         rr1 = 200
                         rr2 = 205
                         cc1 = 2
                         cc2 = 15
                         rr3 = 193
                         rr4 = 205
                         cc3 = 16
                         cc4 = 19
    Else
                         rr1 = 215
                         rr2 = 220
                         cc1 = 2
                         cc2 = 15
                         rr3 = 208
                         rr4 = 220
                         cc3 = 16
                         cc4 = 19
    End If
    rr = 9
    For kk% = rr1 To rr2
      For xx% = cc1 To cc2
             Sheets("F03").Cells(rr, xx%) = Sheets("F03").Cells(kk%, xx%)
      Next xx%
      rr = rr + 1
    Next kk%
    rr = 2
    For kk% = rr3 To rr4
      For xx% = cc3 To cc4
             Sheets("F03").Cells(rr, xx%) = Sheets("F03").Cells(kk%, xx%)
      Next xx%
      rr = rr + 1
    Next kk%
    If Sheets("HOME").flag_italiano Then
                           Sheets("F03").inserisci.Caption = Sheets("F03").Cells(206, 2)
                           Sheets("F03").elimina.Caption = Sheets("F03").Cells(207, 2)
                           Sheets("F03").inserisci_d.Caption = Sheets("F03").Cells(206, 2)
                           Sheets("F03").elimina_d.Caption = Sheets("F03").Cells(207, 2)
                           Sheets("F03").crea_relativi.Caption = Sheets("F03").Cells(208, 2)
    Else
                           Sheets("F03").inserisci.Caption = Sheets("F03").Cells(221, 2)
                           Sheets("F03").elimina.Caption = Sheets("F03").Cells(222, 2)
                           Sheets("F03").inserisci_d.Caption = Sheets("F03").Cells(221, 2)
                           Sheets("F03").elimina_d.Caption = Sheets("F03").Cells(222, 2)
                           Sheets("F03").crea_relativi.Caption = Sheets("F03").Cells(223, 2)
    End If
    GoSub blocca
    Sheets("F03").Cells(16, 1).Select
    Sheets("F03").Cells(15, 1).Select
    
    
    'traduzione input linea (R)
    Sheets("F04").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                         rr1 = 200
                         rr2 = 205
                         cc1 = 2
                         cc2 = 9
    Else
                         rr1 = 215
                         rr2 = 220
                         cc1 = 2
                         cc2 = 9
    End If
    rr = 9
    For kk% = rr1 To rr2
      For xx% = cc1 To cc2
             Sheets("F04").Cells(rr, xx%) = Sheets("F04").Cells(kk%, xx%)
      Next xx%
      rr = rr + 1
    Next kk%
    If Sheets("HOME").flag_italiano Then
                           Sheets("F04").Cells(6, 2) = Sheets("F04").Cells(207, 2)
                           Sheets("F04").Cells(7, 2) = Sheets("F04").Cells(208, 2)
                           Sheets("F04").elimina.Caption = Sheets("F04").Cells(209, 2)
                           Sheets("F04").inserisci.Caption = Sheets("F04").Cells(210, 2)
                           Sheets("F04").calc_assoluti.Caption = Sheets("F04").Cells(211, 2)
    Else
                           Sheets("F04").Cells(6, 2) = Sheets("F04").Cells(222, 2)
                           Sheets("F04").Cells(7, 2) = Sheets("F04").Cells(223, 2)
                           Sheets("F04").elimina.Caption = Sheets("F04").Cells(224, 2)
                           Sheets("F04").inserisci.Caption = Sheets("F04").Cells(225, 2)
                           Sheets("F04").calc_assoluti.Caption = Sheets("F04").Cells(226, 2)
    End If
    GoSub blocca
    Sheets("F04").Cells(16, 1).Select
    Sheets("F04").Cells(15, 1).Select
    
    
    
    'traduzione foglio di calcolo
    Sheets("F05").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    rr1 = 205 + nrighe: cc1 = 1
    rr2 = 209 + nrighe: cc2 = 1
    rr3 = 223 + nrighe: cc3 = 2
    rr4 = 205 + nrighe: cc4 = 4
    rr5 = 208 + nrighe: cc5 = 9
    rr6 = 202 + nrighe: cc6 = 11
    rr7 = 206 + nrighe: cc7 = 11
    rr8 = 209 + nrighe: cc8 = 11
    rr9 = 211 + nrighe: cc9 = 11
    rr10 = 216 + nrighe: cc10 = 14
    rr11 = 220 + nrighe: cc11 = 14
    Sheets("F05").Cells(6, 1) = Sheets("F05").Cells(rr1, cc1)
    rr = rr2
    For j = 10 To 24
       cc = cc2
       For k = 1 To 2
                     Sheets("F05").Cells(j, k) = Sheets("F05").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    rr = rr4
    For j = 6 To 9
       cc = cc4
       For k = 4 To 9
                     Sheets("F05").Cells(j, k) = Sheets("F05").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    rr = rr6
    For j = 3 To 7
       cc = cc6
       For k = 11 To 11
                     Sheets("F05").Cells(j, k) = Sheets("F05").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    Sheets("F05").Cells(10, 11) = Sheets("F05").Cells(rr8, cc8)
    Sheets("F05").Cells(12, 11) = Sheets("F05").Cells(rr9, cc9)
    Sheets("F05").Cells(17, 14) = Sheets("F05").Cells(rr10, cc10)
    Sheets("F05").Cells(22, 14) = Sheets("F05").Cells(rr11, cc11)
    'pulsanti
    Sheets("F05").Option1_maxmin.Caption = Sheets("F05").Cells(225 + nrighe, 1)
    Sheets("F05").Option1_med.Caption = Sheets("F05").Cells(226 + nrighe, 1)
    Sheets("F05").inversione_contropendenze.Caption = Sheets("F05").Cells(227 + nrighe, 1)
    Sheets("F05").reset_all.Caption = Sheets("F05").Cells(228 + nrighe, 1)
    Sheets("F05").calcola_normale.Caption = Sheets("F05").Cells(229 + nrighe, 1)
    Sheets("F05").calcola_idraul.Caption = Sheets("F05").Cells(230 + nrighe, 1)
    Sheets("F05").calcola_ancorata.Caption = Sheets("F05").Cells(231 + nrighe, 1)
    Sheets("F05").ceck_avanzamento.Caption = Sheets("F05").Cells(232 + nrighe, 1)
    Sheets("F05").in_esecuzione.Caption = Sheets("F05").Cells(233 + nrighe, 1)
    
    GoSub blocca
    Sheets("F05").Cells(1, 1).Select

    
    'traduzione foglio di carico campate
    Sheets("F06").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    Sheets("F06").Cells(1, 1) = Sheets("F06").Cells(200 + nrighe, 1)
    rr = 208 + nrighe
    For j = 9 To 11
       cc = 1
       For k = 1 To 22
                     Sheets("F06").Cells(j, k) = Sheets("F06").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'pulsanti
    Sheets("F06").v_carico.Caption = Sheets("F06").Cells(200 + nrighe, 4)
    Sheets("F06").v_vuoto.Caption = Sheets("F06").Cells(201 + nrighe, 4)
    Sheets("F06").v_nullo.Caption = Sheets("F06").Cells(202 + nrighe, 4)
    Sheets("F06").v_variabile.Caption = Sheets("F06").Cells(203 + nrighe, 4)
    Sheets("F06").reset_campate.Caption = Sheets("F06").Cells(204 + nrighe, 4)
    Sheets("F06").una_campata.Caption = Sheets("F06").Cells(205 + nrighe, 4)
    Sheets("F06").tutte_campate.Caption = Sheets("F06").Cells(206 + nrighe, 4)
    Sheets("F06").reset_generale_caso_carichi.Caption = Sheets("F06").Cells(207 + nrighe, 4)
    
    GoSub blocca
    Sheets("F06").Cells(12, 3).Select
    
    
    'traduzione foglio di attriti rulli
    Sheets("F07").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 15
    End If
    rr = 200 + nrighe
    For j = 1 To 11
       cc = 2
       For k = 2 To 10
                     Sheets("F07").Cells(j, k) = Sheets("F07").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'pulsanti
    Sheets("F07").attr_perc.Caption = Sheets("F07").Cells(211 + nrighe, 2)
    Sheets("F07").Attr_ass.Caption = Sheets("F07").Cells(212 + nrighe, 2)
    
    GoSub blocca
    Sheets("F07").Cells(12, 3).Select
    
    
    
    'traduzione foglio vedi tabella verifica linea (una condizione)
    Sheets("F08").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    rr = 400 + nrighe
    For j = 1 To 29
       cc = 1
       For k = 1 To 16
                     Sheets("F08").Cells(j, k) = Sheets("F08").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'pulsanti
    Sheets("F08").opt_00.Caption = Sheets("F08").Cells(429 + nrighe, 2)
    Sheets("F08").opt_01.Caption = Sheets("F08").Cells(430 + nrighe, 2)
    Sheets("F08").opt_03.Caption = Sheets("F08").Cells(431 + nrighe, 2)
    Sheets("F08").opt_02.Caption = Sheets("F08").Cells(432 + nrighe, 2)
    
    GoSub blocca
    Sheets("F08").Cells(1, 1).Select
   
    
    'traduzione foglio pompaggio in linea
    Sheets("F09").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    rr = 200 + nrighe
    For j = 1 To 26
       cc = 1
       For k = 1 To 12
                     Sheets("F09").Cells(j, k) = Sheets("F09").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'grafico
       stringa1 = Sheets("F09").Cells(227 + nrighe, 2)
       stringa2 = Sheets("F09").Cells(228 + nrighe, 2)
       stringa3 = Sheets("F09").Cells(229 + nrighe, 2)
       stringa4 = Sheets("F09").Cells(230 + nrighe, 2)
       stringa5 = Sheets("F09").Cells(231 + nrighe, 2)
       stringa6 = Sheets("F09").Cells(232 + nrighe, 2)
       stringa7 = Sheets("F09").Cells(233 + nrighe, 2)
       ActiveSheet.ChartObjects(1).Activate
       ActiveChart.ChartArea.Select
       With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = stringa1
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = stringa2
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = stringa3
       End With
       ActiveSheet.ChartObjects(2).Activate
       ActiveChart.ChartArea.Select
       With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = stringa7
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = stringa2
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = stringa3
       End With
    GoSub blocca
    Sheets("F09").Cells(1, 1).Select
    
    
    'traduzione foglio massimi e minimi
    Sheets("F10").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    rr = 400 + nrighe
    For j = 1 To 29
       cc = 1
       For k = 1 To 16
                     Sheets("F10").Cells(j, k) = Sheets("F10").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'pulsanti
    Sheets("F10").test_grafico.Caption = Sheets("F10").Cells(430 + nrighe, 2)
    Sheets("F10").ipo_01.Caption = Sheets("F10").Cells(431 + nrighe, 2)
    Sheets("F10").ipo_02.Caption = Sheets("F10").Cells(432 + nrighe, 2)
    Sheets("F10").ipo_03.Caption = Sheets("F10").Cells(433 + nrighe, 2)
    Sheets("F10").ipo_04.Caption = Sheets("F10").Cells(434 + nrighe, 2)
    Sheets("F10").rm_sd.Caption = Sheets("F10").Cells(435 + nrighe, 2)
    Sheets("F10").rm_s.Caption = Sheets("F10").Cells(436 + nrighe, 2)
    Sheets("F10").rm_d.Caption = Sheets("F10").Cells(437 + nrighe, 2)
    Sheets("F10").generazione_dxf.Caption = Sheets("F10").Cells(438 + nrighe, 2)
    GoSub blocca
    Sheets("F10").Cells(1, 1).Select
    
    
    'traduzione foglio potenze
    Sheets("F11").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    rr = 400 + nrighe
    For j = 1 To 29
       cc = 1
       For k = 1 To 17
                     Sheets("F11").Cells(j, k) = Sheets("F11").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'pulsanti
    Sheets("F11").ver_01.Caption = Sheets("F11").Cells(430 + nrighe, 2)
    Sheets("F11").ver_02.Caption = Sheets("F11").Cells(431 + nrighe, 2)
    Sheets("F11").ver_03.Caption = Sheets("F11").Cells(432 + nrighe, 2)
    Sheets("F11").ver_04.Caption = Sheets("F11").Cells(433 + nrighe, 2)
    GoSub blocca
    Sheets("F11").Cells(1, 1).Select
    
    
    'traduzione foglio dati generali
    Sheets("F12").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 200
    End If
    rr = 200 + nrighe
    For j = 7 To 109
       cc = 1
       For k = 1 To 5
                     Sheets("F12").Cells(j, k) = Sheets("F12").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    'pulsanti
    Sheets("F12").genera_report.Caption = Sheets("F12").Cells(304 + nrighe, 3)
    GoSub blocca
    Sheets("F12").Cells(1, 1).Select
    
    'traduzione foglio F13 (testi italiano-inglese)
    Sheets("F13").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 60
    End If
    rr = 60 + nrighe
    For j = 1 To 48
       cc = 1
       For k = 1 To 23
                     Sheets("F13").Cells(j, k) = Sheets("F13").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    GoSub blocca
    
    'traduzione foglio Fune Ancorata
    Sheets("F20").Select
    GoSub libera
    If Sheets("HOME").flag_italiano Then
                                        nrighe = 0
    Else
        nrighe = 50
    End If
    rr = 400 + nrighe
    For j = 1 To 10
       cc = 1
       For k = 1 To 5
                     Sheets("F20").Cells(j, k) = Sheets("F20").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    rr = 411 + nrighe
    For j = 12 To 13
       cc = 1
       For k = 1 To 11
                     Sheets("F20").Cells(j, k) = Sheets("F20").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    rr = 411 + nrighe
    For j = 12 To 13
       cc = 13
       For k = 13 To 23
                     Sheets("F20").Cells(j, k) = Sheets("F20").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    rr = 414 + nrighe
    For j = 97 To 99
       cc = 1
       For k = 1 To 3
                     Sheets("F20").Cells(j, k) = Sheets("F20").Cells(rr, cc)
                     cc = cc + 1
       Next k
       rr = rr + 1
    Next j
    Sheets("F20").Cells(1, 9) = Sheets("F20").Cells(400 + nrighe, 9)
    Sheets("F20").Cells(2, 9) = Sheets("F20").Cells(401 + nrighe, 9)
    'pulsanti
    Sheets("F20").tensione_posa.Caption = Sheets("F20").Cells(418 + nrighe, 1)
    Sheets("F20").grafico_tens_posa.Caption = Sheets("F20").Cells(419 + nrighe, 1)
    GoSub blocca
    Sheets("F20").Cells(1, 1).Select
    
    
    Close
    Sheets("MENU").Select
    Sheets("MENU").Cells(1, 1).Select
    
    Exit Sub
    
libera:
      ActiveSheet.Unprotect ("vit210147vit")
      Return
      
blocca:
      ActiveSheet.Protect password:="vit210147vit"
      Return
      
End Sub
