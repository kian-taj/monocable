VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub reset_generale_caso_carichi_Click()
    
    Sheets("F06").Select
    Sheets("F06").Cells(15, 3).Select
    Range("C12:V92").Select
    Selection.ClearContents
    Sheets("F06").Cells(1, 1).Select
    Sheets("F06").Text4.Text = "0"

End Sub

Private Sub text2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = Text2.Text
        Call virgola_to_punto(tstringa)
        Text2.Text = tstringa
End Sub
Private Sub text3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = Text3.Text
        Call virgola_to_punto(tstringa)
        Text3.Text = tstringa
End Sub
Private Sub text4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = Text4.Text
        Call virgola_to_punto(tstringa)
        Text4.Text = tstringa
End Sub
Private Sub Text1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        tstringa = text1.Text
        Call virgola_to_punto(tstringa)
        text1.Text = tstringa
End Sub

Private Sub Worksheet_Activate()

Call leggi_linea
riga% = 12
If ns = nd Then
   For j% = 2 To ns
        Cells(riga%, 1) = n$(1, j% - 1)
        Cells(riga%, 2) = n$(1, j%)
        riga% = riga% + 1
   Next j%
Else

End If

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
       Dim cnum, rnum, cramo As Integer
       Dim vcar As Single
       
       Call leggi_linea
       rnum = Target.Row
       cnum = Target.Column
       If rnum >= 12 And rnum <= 92 Then
          If cnum >= 3 And cnum <= 22 Then
              'controllo compatibilità valori numerici nelle caselle di testo
              If Not IsNumeric(Sheets("F06").text1.Text) Then
                      Beep
                      MsgBox ("Errore: formato numerico del valore del peso veicolo carico NON VALIDO : controlla !!!")
              Exit Sub
              End If
              If Not IsNumeric(Text2.Text) Then
                      Beep
                      MsgBox ("Errore: formato numerico del valore del peso veicolo vuoto NON VALIDO : controlla !!!")
              Exit Sub
              End If
              If Not IsNumeric(Text3.Text) Then
                      Beep
                      MsgBox ("Errore: formato numerico del valore di zero peso veicolo NON VALIDO : controlla !!!")
              Exit Sub
              End If
              If Not IsNumeric(Text4.Text) Then
                      Beep
                      MsgBox ("Errore: formato numerico del valore del peso veicolo imposto NON VALIDO : controlla !!!")
              Exit Sub
              End If
              'assegnazione valori
              If v_carico.value Then vcar = CDbl(text1.Text)
              If v_vuoto.value Then vcar = CDbl(Text2.Text)
              If v_nullo.value Then vcar = CDbl(Text3.Text)
              If v_variabile.value Then vcar = CDbl(Text4.Text)
              If una_campata.value Then
                    Cells(rnum, cnum) = vcar
              Else
                    riga% = 12
                    If Cells(11, cnum) = "Salita" Then cramo = ns Else cramo = nd
                    For j% = 1 To cramo - 1
                        If tutte_campate.value Then
                                    Cells(riga%, cnum) = vcar
                        Else
                            Cells(riga, cnum) = ""
                        End If
                        riga% = riga% + 1
                    Next j%
              End If
              Cells(rnum + 1, cnum).Select
          End If
       End If
       
End Sub
