VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Aggiorna_Intestazione_Tabulati_Click()
' intesta tutti i fogli per la stampa -------------------------------------------
On Local Error GoTo err_intesta

Dim NFGL(30) As String
Dim i As Integer

i = 1: ok = 0
NFGL(i) = "HOME": i = i + 1
NFGL(i) = "MENU": i = i + 1
For j = 1 To 9
               NFGL(i) = "F0" & Format(j, "##"): i = i + 1
Next j
For j = 10 To 12
               NFGL(i) = "F" & Format(j, "##"): i = i + 1
Next j
NFGL(i) = "F20"

'applica la password di protezione
For j = 1 To i
             Sheets(NFGL(j)).Select
             Sheets(NFGL(j)).Cells(1, 1).Select
             Call Intesta_Tabulati(NFGL(j))
Next j
Beep
If ok = 0 Then
             MsgBox ("ALL the sheets were registered")
Else
     MsgBox ("NOT ALL of the sheets were registered")
End If

Sheets("MENU").Select
Sheets("MENU").Cells(1, 1).Select
               
Exit Sub

err_intesta:
           Beep
           ok = 1
           Resume Next
           
End Sub

Private Sub CommandButton1_Click()
         Sheets("F01").Select
         Sheets("F01").Cells(1, 1).Select
End Sub

Private Sub CommandButton10_Click()
         Sheets("F10").Select
         Sheets("F10").Cells(1, 1).Select
End Sub

Private Sub CommandButton11_Click()
         Sheets("F11").Select
         Sheets("F11").Cells(1, 1).Select
End Sub

Private Sub CommandButton12_Click()
         Sheets("F12").Select
         Sheets("F12").Cells(1, 1).Select
End Sub

Private Sub CommandButton2_Click()
         Sheets("F02").Select
         Sheets("F02").Cells(1, 1).Select
End Sub

Private Sub CommandButton3_Click()
         Sheets("F03").Select
         Sheets("F03").Cells(1, 1).Select
End Sub

Private Sub CommandButton4_Click()
         Sheets("F04").Select
         Sheets("F04").Cells(1, 1).Select
End Sub

Private Sub CommandButton5_Click()
         Sheets("F05").Select
         Sheets("F05").Cells(1, 1).Select
End Sub

Private Sub CommandButton6_Click()
         Sheets("F06").Select
         Sheets("F06").Cells(1, 1).Select
End Sub

Private Sub CommandButton7_Click()
         Sheets("F07").Select
         Sheets("F07").Cells(1, 1).Select
End Sub

Private Sub CommandButton8_Click()
         Sheets("F08").Select
         Sheets("F08").Cells(1, 1).Select
End Sub

Private Sub CommandButton9_Click()
         Sheets("F09").Select
         Sheets("F09").Cells(1, 1).Select
End Sub

Private Sub elimina_input_Click()
        Dim nr, avviso As String
        nr = Chr(13) + Chr(10)
        avviso = "ELIMINAZIONE DI TUTTI I DATI IMMESSI" + nr + nr
        avviso = avviso + "SEI SICURO ?"
        result = MsgBox(avviso, vbYesNo + vbCritical)
        If result = 6 Then
           nomefoglio = "MENU"
           Call elimina_F01(nomefoglio)
           Call elimina_F02(nomefoglio)
           Call elimina_F03(nomefoglio)
           Call elimina_F04(nomefoglio)
        End If
End Sub

Private Sub elimina_output_Click()
        Dim nr, avviso As String
        nr = Chr(13) + Chr(10)
        avviso = "ELIMINAZIONE DI TUTTI I DATI ELABORATI" + nr + nr
        avviso = avviso + "SEI SICURO ?"
        result = MsgBox(avviso, vbYesNo + vbCritical)
        If result = 6 Then
           nomefoglio = "MENU"
           Call elimina_dati_generali(nomefoglio)
           Call elimina_fogli_calcoli(nomefoglio)
           Call elimina_F06_F07(nomefoglio)
           Call elimina_PosaFPT(nomefoglio)
        End If
        
End Sub


Private Sub Worksheet_Activate()
        Sheets("MENU").Cells(1, 1).Select
End Sub

