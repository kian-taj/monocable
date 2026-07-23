VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

On Local Error GoTo err_open_menu

Dim ok As Integer
Dim codice As String

    'controllo e creazione esistenza dir file transito
    ok = 1
    MkDir "C:\TMP_SIF"
    
    'lettura o creazione file per codice attivazione
    ok = 2
    Open "C:\TMP_SIF\MONOFUNI_PSWKEY" For Input As #1
    ok = 3
    Input #1, codice
    Close
    
    Open "C:\TMP_SIF\MONOFUNI_TEST_PSWKEY" For Output As #1
    Print #1, "****"
    Close
    
    Call DiskInfo
    
riprova:
    'password x MONOFUNI
    'calcolo codici di attivazione
    i1% = InStr(cek_codice_numserie, "-")
    If i1% > 0 Then
                   cd1$ = mid$(cek_codice_numserie, 1, i1% - 1)
                   cd2$ = mid$(cek_codice_numserie, i1% + 1, Len(cek_codice_numserie))
    Else
        cd1$ = cek_codice_numserie
        cd2$ = "SIF"
    End If
    c1 = 0
    For jjj% = 1 To Len(cd1$)
                           cc1 = Asc(mid$(cd1$, jjj%, 1))
                           cc1 = cc1 * jjj% * (7 + jjj%) - Asc(mid$(cd1$, jjj%, 1))
                           c1 = c1 + cc1
    Next jjj%
    c2 = 0
    For jjj% = 1 To Len(cd2$)
                          cc2 = Asc(mid$(cd2$, jjj%, 1))
                          cc2 = cc2 * jjj% * (98 + jjj% * 11) + Asc(mid$(cd2$, jjj%, 1))
                          c2 = c2 + cc2
    Next jjj%
    c3 = Int((Sqr(c1 ^ 1.8 + c2 ^ 1.9) / 2.222 * Sqr(2.88)) * 10)
    mask$ = "##########"
    cek$ = Format$(c1, mask$) + "." + Format$(c2, mask$) + "." + Format$(c3, mask$)
    ok = 0
    If cek$ = codice Then
                          Open "C:\TMP_SIF\MONOFUNI_TEST_PSWKEY" For Output As #1
                          Print #1, "activate"
                          Close
    End If
    
    If cek$ <> codice Then
                         Beep
                         nr = Chr(13) + Chr(10)
                         If Sheets("HOME").flag_italiano.value Then
                           avviso = "CODICE DI ATTIVAZIONE NON VALIDO : "
                           avviso = avviso + nr + nr
                           avviso = avviso + "CODICE MACCHINA : " + cek_codice_numserie
                           avviso = avviso + nr + nr
                           avviso = avviso + "Comunicare al distributore del programma il codice macchina per ottenere il codice di attivazione!" + nr
                         Else
                           avviso = "NON VALID CODE OF ACTIVATION : "
                           avviso = avviso + nr + nr
                           avviso = avviso + "CODE PLOTS : " + cek_codice_numserie
                           avviso = avviso + nr + nr
                           avviso = avviso + "To communicate the code to the distributor of the program plots for getting the code of activation!" + nr
                         End If
                         result = MsgBox(avviso, 48, "<<<<  INVALID  PASSWORD >>>>")
                         ok = 9
    End If
    If ok <> 0 Then
                   Beep
                   If Sheets("HOME").flag_italiano.value Then
                            testomes = "INSERIRE IL CODICE DI ATTIVAZIONE"
                   Else
                            testomes = "TO INSERT THE CODE OF ACTIVATION"
                   End If
                   codice = InputBox(testomes, , "NNNN.NNNNN.NNNNNN")
                   If codice = "" Then
                                     Windows(1).Close
                   Else
                            Open "C:\TMP_SIF\MONOFUNI_PSWKEY" For Output As #1
                            Print #1, codice
                            Close
                            GoTo riprova
                   End If
    End If
    Close
    
    Sheets("HOME").Activate
    Sheets("HOME").Cells(50, 1).Activate
    
    'leggi path_programmi (mia_path)
    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close
    
    test_ricalcolo = True
    
Exit Sub

err_open_menu:
              If ok = 7 Then
                            Beep
                            MsgBox ("DEFAULT LANGUAGE = IT")
                            Close
                            Exit Sub
              End If
              If ok = 1 Then Resume Next
              If ok = 2 Then
                            Open "C:\TMP_SIF\KEY_SIF" For Output As #1
                            Print #1, "NNNN.NNNNN.NNNNNN"
                            Close
                            Open "C:\TMP_SIF\KEY_SIF" For Input As #1
                            Resume Next
              End If
              
              
           
End Sub
