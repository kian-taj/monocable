VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub assegna_attriti_salita_Click()
      Call leggi_linea
      riga% = 12
      For j% = 2 To ns - 1
            Sheets("F07").Cells(riga%, 2) = n$(1, j%)
            If Sheets("F07").attr_perc.value Then Sheets("F07").Cells(riga%, 3) = "%" Else Sheets("F07").Cells(riga%, 3) = "A"
            Sheets("F07").Cells(riga%, 4) = Sheets("F07").Cells(4, 9)
            Sheets("F07").Cells(riga%, 5) = Sheets("F07").Cells(5, 9)
            riga% = riga% + 1
      Next j%
End Sub

Private Sub Assegna_attrito_discesa_Click()
      Call leggi_linea
      riga% = 12
      For j% = 2 To nd - 1
            Sheets("F07").Cells(riga%, 7) = n$(2, j%)
            If Sheets("F07").attr_perc.value Then Sheets("F07").Cells(riga%, 8) = "%" Else Sheets("F07").Cells(riga%, 8) = "A"
            Sheets("F07").Cells(riga%, 9) = Sheets("F07").Cells(4, 9)
            Sheets("F07").Cells(riga%, 10) = Sheets("F07").Cells(5, 9)
            riga% = riga% + 1
      Next j%
End Sub

Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
      riga% = Target.Row
      colonna% = Target.Column
      If colonna% = 2 And riga% > 11 And riga% < 91 Then
            If Sheets("F07").attr_perc.value Then Sheets("F07").Cells(riga%, 3) = "%" Else Sheets("F07").Cells(riga%, 3) = "A"
            Sheets("F07").Cells(riga%, 4) = Sheets("F07").Cells(4, 9)
            Sheets("F07").Cells(riga%, 5) = Sheets("F07").Cells(5, 9)
            Sheets("F07").Cells(1, 1).Select
      End If
      If colonna% = 7 And riga% > 11 And riga% < 91 Then
            If Sheets("F07").attr_perc.value Then Sheets("F07").Cells(riga%, 8) = "%" Else Sheets("F07").Cells(riga%, 8) = "A"
            Sheets("F07").Cells(riga%, 9) = Sheets("F07").Cells(4, 9)
            Sheets("F07").Cells(riga%, 10) = Sheets("F07").Cells(5, 9)
            Sheets("F07").Cells(1, 1).Select
      End If
      
End Sub

