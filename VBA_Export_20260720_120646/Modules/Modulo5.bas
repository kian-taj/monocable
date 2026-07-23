Attribute VB_Name = "Modulo5"

Public nomefoglio As String
Sub elimina_store13(nomefoglio)
'
' Macro registrata il 29/06/2005 da pr33006
'
    Sheets("STORE13").Select
    Sheets("STORE13").Cells(1, 1).Select
    Range("B4:CM65000,CP4:GA65000,GE4:GS800").Select
    Selection.ClearContents
    Sheets("STORE13").Cells(1, 1).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
    
End Sub
Sub elimina_store05(nomefoglio)
'
' Macro registrata il 29/06/2005 da pr33006
'
    Sheets("STORE05").Select
    Sheets("STORE05").Cells(1, 1).Select
    Range("B3:FF32,B42:FF71,B82:FF111,B122:FF151").Select
    Selection.ClearContents
    Range("B158:FF160,B164:FF166,B170:FF172,B176:FF178").Select
    Selection.ClearContents
    Range("B185:FF187,B191:FF193,B197:FF199,B203:FF205").Select
    Selection.ClearContents
    Range("B213:FF218,B222:FF227,B231:FF236,B240:FF245").Select
    Selection.ClearContents
    Range("B250:FF255,B259:FF264,B268:FF273,B277:FF282").Select
    Selection.ClearContents
    Range("B290:FF298,B302:FF310,B314:FF322,B326:FF334").Select
    Selection.ClearContents
    Sheets("STORE05").Cells(1, 1).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
    
End Sub


Sub elimina_profilo_salita(nomefoglio)
Attribute elimina_profilo_salita.VB_Description = "Macro registrata il 29/06/2005 da pr33006"
Attribute elimina_profilo_salita.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro registrata il 29/06/2005 da pr33006
'
    Sheets("F02").Select
    Sheets("F02").Cells(15, 3).Select
    Range("C15:E1114").Select
    Selection.ClearContents
    Sheets("F02").Cells(15, 3).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
    
End Sub

Sub elimina_profilo_discesa(nomefoglio)
'
' Macro registrata il 29/06/2005 da pr33006
'
    Sheets("F02").Select
    Sheets("F02").Cells(15, 8).Select
    Range("H15:J1114").Select
    Selection.ClearContents
    Sheets("F02").Cells(15, 8).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
    
End Sub


Sub elimina_dati_generali(nomefoglio)
'
' Macro registrata il 29/06/2005 da pr33006
'
    test_ricalcolo = False
    Sheets("F12").Select
    Sheets("F12").Cells(5, 7).Select
    Range("E7:E9,E10,E13:E31,E34:E45,E48:E51,E54:E59,E62:E71,E75:E90,F75,F77,E92:E109").Select
    Selection.ClearContents
    Sheets("F12").Cells(5, 7).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
    test_ricalcolo = True
    
End Sub
Sub elimina_F01(nomefoglio)
    Sheets("F01").Select
    Sheets("F01").Cells(1, 1).Select
    Range("D3:D5").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("D7").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("D15").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("D23").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("E8:E13").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("E16:E21").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("E24:E28").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("E39:E45").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("E47:E61").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("E64:E67").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    Range("J32:J34").Select
    Selection.ClearContents
    Sheets("F01").Cells(1, 1).Select
    
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
    
End Sub

Sub elimina_F02(nomefoglio)
    Sheets("F02").Select
    Sheets("F02").Cells(15, 3).Select
    Range("C15:E1114").Select
    Selection.ClearContents
    Sheets("F02").Cells(15, 3).Select
    Range("H15:J1114").Select
    Selection.ClearContents
    Sheets("F02").Cells(15, 8).Select
    Range("L15:O1114").Select
    Selection.ClearContents
    Sheets("F02").Cells(15, 8).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(1, 1).Select
End Sub
    

Sub elimina_F03(nomefoglio)
'
' pulizia area dati assoluti linea
'
    Sheets("F03").Select
    Sheets("F03").Cells(15, 3).Select
    Range("C15:H94").Select
    Selection.ClearContents
    Sheets("F03").Cells(15, 10).Select
    Range("J15:O94").Select
    Selection.ClearContents
    Sheets("F03").Cells(15, 3).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(15, 3).Select
    
End Sub


Sub elimina_F04(nomefoglio)
'
' pulizia area dati relativi linea
'
    Sheets("F04").Select
    Sheets("F04").Cells(15, 3).Select
    Range("C15:E94").Select
    Selection.ClearContents
    Sheets("F04").Cells(15, 7).Select
    Range("G15:I94").Select
    Selection.ClearContents
    Sheets("F04").Cells(15, 3).Select
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(15, 3).Select
    
End Sub
Sub Intesta_Tabulati(foglio)
'
    nomeutente = Sheets("MENU").Cells(33, 5)
    nomeimpianto = Sheets("F01").Cells(3, 4) & Chr(10) & Sheets("F01").Cells(4, 4)
    numrelease = Sheets("HOME").Label1.Caption
    telemail = Sheets("HOME").Label2.Caption
    
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    With ActiveSheet.PageSetup
        .LeftHeader = nomeimpianto                                                'nome impianto
        .CenterHeader = ""
        .RightHeader = "&D" + " " + "&T" + " " + "pg. n° &P"
        .LeftFooter = numrelease & Chr(10) & telemail                             'versione e rif. contatto
        .CenterFooter = ""
        .RightFooter = nomeutente                                                 'nome utente
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.393700787401575)
        .TopMargin = Application.InchesToPoints(1.18110236220472)
        .BottomMargin = Application.InchesToPoints(0.905511811023622)
        .HeaderMargin = Application.InchesToPoints(0.47244094488189)
        .FooterMargin = Application.InchesToPoints(0.354330708661417)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    
End Sub

Sub elimina_fogli_calcoli(nomefoglio)
       
       Sheets("F05").Select
       Sheets("F05").Cells(1, 1).Select
       Sheets("F05").v_passo.Text = 0
       Sheets("F05").offset_inizio.Text = 0
       Sheets("F05").offset_fine.Text = 0
       Sheets("F05").Range("O3:O7").Select
       Selection.ClearContents
       Sheets("F05").Cells(1, 1).Select
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
       
       
       Sheets("F09").Select
       Sheets("F09").Cells(1, 1).Select
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
       
       Sheets("F20").Select
       Sheets("F20").Range("A14:D94").Select
       Selection.ClearContents
       Sheets("F20").Cells(1, 1).Select
       Sheets("F20").Range("M14:P94").Select
       Selection.ClearContents
       Sheets("F20").Cells(30, 2).Select
       Sheets("F20").Cells(1, 1).Select
       
       Sheets(nomefoglio).Select
       Sheets(nomefoglio).Cells(1, 1).Select


End Sub


Sub elimina_F06_F07(nomefoglio)
'
' pulizia area dati assoluti linea
'
    Sheets("F06").Select
    Sheets("F06").Cells(15, 3).Select
    Range("A12:V92").Select
    Selection.ClearContents
    Sheets("F06").Cells(1, 1).Select
    Sheets("F06").Text4.Text = "0"
    Sheets("F07").Select
    Sheets("F07").Cells(15, 3).Select
    Range("B12:E92,G12:J92").Select
    Selection.ClearContents
    Sheets("F07").Cells(1, 1).Select
    
    Sheets(nomefoglio).Select
    Sheets(nomefoglio).Cells(15, 3).Select
    
End Sub


Sub elimina_PosaFPT(nomefoglio)

          Sheets("F20").Select
          Sheets("F20").Cells(1, 1).Select
          Sheets("F20").Range("A14:D94,M14:P94,F14").Select
          Sheets("F20").Range("F14").Activate
          Selection.ClearContents
          Sheets("F20").Cells(1, 1).Select
          Sheets(nomefoglio).Select
          Sheets(nomefoglio).Cells(15, 3).Select

End Sub
Public Sub psw_activate()
On Local Error GoTo set_password

Dim NFGL(30) As String
Dim i, testok As Integer

i = 0: testok = 0
For Each ws In Worksheets
                         i = i + 1
                         NFGL(i) = ws.Name
Next ws

'applica la password di protezione
For j = 1 To i
             Sheets(NFGL(j)).Select
             Sheets(NFGL(j)).Cells(1, 1).Select
             ActiveSheet.Protect password:="vit210147vit"
Next j
Sheets("HOME").Select
Sheets("HOME").Cells(1, 1).Select

If testok = 0 Then
              Beep
              MsgBox ("ALL THE SHEETS WHERE PROTECT ")
Else
              Beep
              MsgBox ("NOT ALL THE SHEETS WHERE PROTECT ")
End If

Exit Sub

set_password:
            If NFGL(j) = "F13" Then Resume Next
            Beep
            testok = 1
            Resume Next
            
End Sub


Public Sub psw_remove()
On Local Error GoTo unlock_password

Dim NFGL(30) As String
Dim i, testok As Integer

i = 0: testok = 0
For Each ws In Worksheets
                         i = i + 1
                         NFGL(i) = ws.Name
Next ws

'applica la password di protezione
For j = 1 To i
             Sheets(NFGL(j)).Select
             Sheets(NFGL(j)).Cells(1, 1).Select
             ActiveSheet.Unprotect ("vit210147vit")
Next j
Sheets("HOME").Select
Sheets("HOME").Cells(1, 1).Select

If testok = 0 Then
              Beep
              MsgBox ("ALL THE SHEETS WHERE UNPROTECT ")
Else
              Beep
              MsgBox ("NOT ALL THE SHEETS WHERE UNPROTECT ")
End If


Exit Sub

unlock_password:
            If NFGL(j) = "F13" Then Resume Next
            Beep
            testok = 1
            Resume Next
            
End Sub
