VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub fpt_setup_Click()
Dim j, kk, rr, cc1, cc2 As Integer
Dim testo As Variant

j = Sheets("F01").Fune_Portante.ListIndex
If j > 0 Then
      testo = Sheets("ROPE_1").Cells(10 + j, 2)
      If IsNumeric(testo) Then
                              'scelta fune
                              rr = 10 + j
                              For kk = rr To 10 Step -1
                                                      testo = Sheets("ROPE_1").Cells(kk, 2)
                                                      If Not IsNumeric(testo) Then Exit For
                              Next kk
                              If Sheets("F01").fpt_1770 Then cc1 = 6: cc2 = 9
                              If Sheets("F01").fpt_1960 Then cc1 = 7: cc2 = 10
                              If Sheets("F01").fpt_2160 Then cc1 = 8: cc2 = 11
                              'assegna i valori
                              Sheets("F01").Cells(7, 4) = testo
                              Sheets("F01").Cells(8, 5) = Sheets("ROPE_1").Cells(rr, 2)
                              Sheets("F01").Cells(9, 5) = Sheets("ROPE_1").Cells(rr, 3)
                              Sheets("F01").Cells(10, 5) = Sheets("ROPE_1").Cells(rr, 4)
                              Sheets("F01").Cells(11, 5) = Sheets("ROPE_1").Cells(rr, 5)
                              Sheets("F01").Cells(12, 5) = Sheets("ROPE_1").Cells(rr, cc1)
                              Sheets("F01").Cells(12, 7) = Sheets("ROPE_1").Cells(rr, cc2)
                              
      End If
End If
End Sub

Private Sub fs_setup_Click()
Dim j, kk, rr, cc1, cc2 As Integer
Dim testo As Variant

j = Sheets("F01").Fune_Segnalazione.ListIndex
If j > 0 Then
      testo = Sheets("ROPE_3").Cells(10 + j, 2)
      If IsNumeric(testo) Then
                              'scelta fune
                              rr = 10 + j
                              For kk = rr To 10 Step -1
                                                      testo = Sheets("ROPE_3").Cells(kk, 2)
                                                      If Not IsNumeric(testo) Then Exit For
                              Next kk
                              cc1 = 7: cc2 = 10
                              'assegna i valori
                              Sheets("F01").Cells(23, 4) = testo
                              Sheets("F01").Cells(24, 5) = Sheets("ROPE_3").Cells(rr, 2)
                              Sheets("F01").Cells(25, 5) = Sheets("ROPE_3").Cells(rr, 3)
                              Sheets("F01").Cells(26, 5) = Sheets("ROPE_3").Cells(rr, 4)
                              Sheets("F01").Cells(28, 5) = Sheets("ROPE_3").Cells(rr, cc1)
                              Sheets("F01").Cells(28, 7) = Sheets("ROPE_3").Cells(rr, cc2)
                              
      End If
End If

End Sub

Private Sub ft_setup_Click()
Dim j, kk, rr, cc1, cc2 As Integer
Dim testo As Variant

j = Sheets("F01").Fune_Tenditrice.ListIndex
If j > 0 Then
      testo = Sheets("ROPE_2").Cells(10 + j, 2)
      If IsNumeric(testo) Then
                              'scelta fune
                              rr = 10 + j
                              For kk = rr To 10 Step -1
                                                      testo = Sheets("ROPE_2").Cells(kk, 2)
                                                      If Not IsNumeric(testo) Then Exit For
                              Next kk
                              cc1 = 7: cc2 = 10
                              'assegna i valori
                              Sheets("F01").Cells(15, 4) = testo
                              Sheets("F01").Cells(16, 5) = Sheets("ROPE_2").Cells(rr, 2)
                              Sheets("F01").Cells(17, 5) = Sheets("ROPE_2").Cells(rr, 3)
                              Sheets("F01").Cells(18, 5) = Sheets("ROPE_2").Cells(rr, 4)
                              Sheets("F01").Cells(19, 5) = Sheets("ROPE_2").Cells(rr, 5)
                              Sheets("F01").Cells(20, 5) = Sheets("ROPE_2").Cells(rr, cc1)
                              Sheets("F01").Cells(20, 7) = Sheets("ROPE_2").Cells(rr, cc2)
                              
      End If
End If

End Sub



Private Sub Image1_Click()
    Sheets("MENU").Select
    Sheets("MENU").Cells(50, 1).Select
End Sub


