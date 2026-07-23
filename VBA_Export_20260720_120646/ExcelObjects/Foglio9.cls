VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub


Private Sub ver_01_Click()
           nomeipotesi$ = Sheets("F11").ver_01.Caption
           Call stampa_potenze(1, nomeipotesi$)
End Sub

Private Sub ver_02_Click()
           nomeipotesi$ = Sheets("F11").ver_02.Caption
           Call stampa_potenze(2, nomeipotesi$)
End Sub

Private Sub ver_03_Click()
           nomeipotesi$ = Sheets("F11").ver_03.Caption
           Call stampa_potenze(3, nomeipotesi$)
End Sub

Private Sub ver_04_Click()
           nomeipotesi$ = Sheets("F11").ver_04.Caption
           Call stampa_potenze(4, nomeipotesi$)
End Sub

