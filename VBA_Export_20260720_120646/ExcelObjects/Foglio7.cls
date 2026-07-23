VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




Private Sub generazione_dxf_Click()
        Call leggi_terreno
        Call leggi_linea
        Close
        'scrivi file profilo terreno
        Open "c:\tmp_sif\valori.prf" For Output As #1
        If Sheets("F10").rm_sd Or Sheets("F10").rm_s Then
                     Write #1, nbs
                     For j% = 1 To nbs
                                      npic$ = Sheets("F02").Cells(15 + j%, 3)
                                      Write #1, quot_terr(1, j%), prog_terr(1, j%), npic$
                     Next j%
        Else
                     Write #1, nbd
                     For j% = 1 To nbd
                                      npic$ = Sheets("F02").Cells(15 + j%, 7)
                                      Write #1, quot_terr(2, j%), prog_terr(2, j%), npic$
                     Next j%
        End If
        Close
        'scrivi file profilo linea
        Open "c:\tmp_sif\valori.lin" For Output As #1
        If Sheets("F10").rm_sd Or Sheets("F10").rm_s Then
                     Write #1, ns - 4
                     For j% = 1 To ns
                            rrr1 = r1(j%): rrr2 = r2(j%)
                            If Sheets("F10").rm_s Then rrr2 = rrr1
                            Write #1, quota(1, j%), prog(1, j%), n$(1, j%), hv(1, j%), rrr1, rrr2
                     Next j%
        Else
                     Write #1, nd - 4
                     For j% = 1 To nd
                            rrr1 = r2(j%): rrr2 = r2(j%)
                            Write #1, quota(2, j%), prog(2, j%), n$(2, j%), hv(2, j%), rrr1, rrr2
                     Next j%
        End If
        Close
        'scrivi file pressioni PXX ed angoli AXX per i sostegni
        Open "c:\tmp_sif\valori.PXX" For Output As #1
        Open "c:\tmp_sif\valori.AXX" For Output As #2
        If Sheets("F10").rm_sd Or Sheets("F10").rm_s Then
                     ww% = 33
                     For j% = 1 To ns - 2
                            ppmaxs = Sheets("F10").Cells(ww%, 11)
                            ppmins = Sheets("F10").Cells(ww% + 1, 11)
                            zz% = ww% + (ns - 3) * 4 + 7
                            ppmaxd = Sheets("F10").Cells(zz%, 11)
                            ppmind = Sheets("F10").Cells(zz% + 1, 11)
                            If Sheets("F10").rm_s Then
                                 'si considera solo il ramo salita
                                 ppmax = ppmaxs
                                 ppmin = ppmins
                            Else
                                 'si considerano tutti i due rami
                                 ppmax = -9999999: ppmin = 9999999
                                 If ppmax < ppmaxs Then ppmax = ppmaxs
                                 If ppmax < ppmins Then ppmax = ppmins
                                 If ppmax < ppmaxd Then ppmax = ppmaxd
                                 If ppmax < ppmind Then ppmax = ppmind
                                 If ppmin > ppmaxs Then ppmin = ppmaxs
                                 If ppmin > ppmins Then ppmin = ppmins
                                 If ppmin > ppmaxd Then ppmin = ppmaxd
                                 If ppmin > ppmind Then ppmin = ppmind
                            End If
                            alfav = (PI# / 180 * Sheets("F10").Cells(ww% - 2, 7) + PI# / 180 * Sheets("F10").Cells(ww% - 1, 7)) / 2
                            alfam = (PI# / 180 * Sheets("F10").Cells(ww% + 2, 6) + PI# / 180 * Sheets("F10").Cells(ww% + 3, 6)) / 2
                            Write #1, ppmax, ppmin
                            Write #2, alfav, alfam
                            ww% = ww% + 4
                     Next j%
        Else
                     ww% = 33 + (ns - 3) * 4 + 7
                     For j% = 1 To nd - 2
                            ppmax = Sheets("F10").Cells(ww%, 11)
                            ppmin = Sheets("F10").Cells(ww% + 1, 11)
                            alfav = (PI# / 180 * Sheets("F10").Cells(ww% - 2, 7) + PI# / 180 * Sheets("F10").Cells(ww% - 1, 7)) / 2
                            alfam = (PI# / 180 * Sheets("F10").Cells(ww% + 2, 6) + PI# / 180 * Sheets("F10").Cells(ww% + 3, 6)) / 2
                            Write #1, ppmax, ppmin
                            Write #2, alfav, alfam
                            ww% = ww% + 4
                      Next j%
        End If
        Close
        'scrivi file frecce FXX per le campate
        Open "c:\tmp_sif\valori.FXX" For Output As #1
        If Sheets("F10").rm_sd Or Sheets("F10").rm_s Then
                     ww% = 31
                     For j% = 1 To ns - 1
                            ffmax = Sheets("F10").Cells(ww%, 5)
                            ffmin = Sheets("F10").Cells(ww% + 1, 5)
                            Write #1, ffmax, ffmin
                            ww% = ww% + 4
                     Next j%
        Else
                     ww% = 31 + (ns - 3) * 4 + 7
                     For j% = 1 To nd - 1
                            ffmax = Sheets("F10").Cells(ww%, 5)
                            ffmin = Sheets("F10").Cells(ww% + 1, 5)
                            Write #1, ffmax, ffmin
                            ww% = ww% + 4
                      Next j%
        End If
        Close
        
        'imposta la lingua di riferimento
        Open "c:\tmp_sif\flag_type" For Output As #1
        If Sheets("HOME").flag_italiano Then
                                            Write #1, "IT"
        Else
             Write #1, "UK"
        End If
        Close
        
        
        Open "c:\tmp_sif\prg_path" For Input As #1
        Input #1, mia_path
        Close

        result = Shell(mia_path + "\MF_CREA_WINDXF.EXE", 1)
        
End Sub

Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub ipo_01_Click()
                          Call stampa_maxmin(1)
End Sub

Private Sub ipo_02_Click()
                          Call stampa_maxmin(2)
End Sub

Private Sub ipo_03_Click()
                          Call stampa_maxmin(3)
End Sub

Private Sub ipo_04_Click()
                          Call stampa_maxmin(4)
End Sub

Private Sub run_ACAD2008_Click()

On Local Error GoTo err_sk2008

    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close
    ActiveWorkbook.FollowHyperlink Address:=mia_path + "\" + "WinSif_MON\MF_AUTOCAD.dwg", NewWindow:=True
Exit Sub

err_sk2008:
          Beep
          testo = "NON è stato possibile aprire il Modulo MF_AUTOCAD" + Chr(10) + Chr(13)
          testo = testo + Chr(10) + Chr(13)
          testo = testo + "Aprire direttamente il file da AUTOCAD"
          MsgBox (testo)

End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Local Error GoTo err_F10

          rr = Target.Row
          cc = Target.Column
          If cc = 50 Then
              Select Case rr
                     Case 1
                           'normal test/fixed
                                 Sheets("F10").test_grafico.value = False
                                 Call ipo_01_Click
                     Case 2
                           '+10%
                                 Sheets("F10").test_grafico.value = False
                                 Call ipo_02_Click
                     Case 3
                           '-10%
                                 Sheets("F10").test_grafico.value = False
                                 Call ipo_03_Click
                     Case 4
                           'Hydraulic
                                 Sheets("F10").test_grafico.value = False
                                 Call ipo_04_Click
              End Select
          End If
              
Exit Sub

err_F10:
       Exit Sub
End Sub
