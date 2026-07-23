VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub del_rigaprof_Click()

      Dim j, k As Integer
      
      If cek_ins_del Then
              For j = numriga To 1113
                     Cells(j, numcolonna + 1) = Cells(j + 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j + 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j + 1, numcolonna + 3)
                     'mezzecoste
                     Cells(j, 12) = Cells(j + 1, 12)
                     Cells(j, 13) = Cells(j + 1, 13)
              Next j
      Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo1 = Sheets("F02").Cells(1213, 2)
          Else
                testo1 = Sheets("F02").Cells(1228, 2)
          End If
          MsgBox (testo1)
      End If

End Sub


Private Sub del_rigaprof_d_Click()

      Dim j, k As Integer
      If cek_ins_del_d Then
              For j = numriga To 1113
                     Cells(j, numcolonna + 1) = Cells(j + 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j + 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j + 1, numcolonna + 3)
                     'mezzecoste
                     Cells(j, 14) = Cells(j + 1, 14)
                     Cells(j, 15) = Cells(j + 1, 15)
              Next j
      Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo1 = Sheets("F02").Cells(1213, 2)
          Else
                testo1 = Sheets("F02").Cells(1228, 2)
          End If
          MsgBox (testo1)
      End If

End Sub

Private Sub Image1_Click()
        Sheets("MENU").Select
        Sheets("MENU").Cells(50, 1).Select
End Sub

Private Sub ins_rigaprof_Click()
      
      Dim j, k As Integer
      
      If cek_ins_del Then
              For j = 1114 To numriga + 1 Step -1
                     Cells(j, numcolonna + 1) = Cells(j - 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j - 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j - 1, numcolonna + 3)
                     'mezzecoste
                     Cells(j, 12) = Cells(j - 1, 12)
                     Cells(j, 13) = Cells(j - 1, 13)
              Next j
              Cells(numriga, numcolonna + 1) = ""
              Cells(numriga, numcolonna + 2) = ""
              Cells(numriga, numcolonna + 3) = ""
              'mezzecoste
              Cells(numriga, 12) = ""
              Cells(numriga, 13) = ""
      Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo1 = Sheets("F02").Cells(1213, 2)
          Else
                testo1 = Sheets("F02").Cells(1228, 2)
          End If
          MsgBox (testo1)
      End If

End Sub

Private Sub ins_rigaprof_d_Click()

      Dim j, k As Integer
      
      If cek_ins_del_d Then
              For j = 1114 To numriga + 1 Step -1
                     Cells(j, numcolonna + 1) = Cells(j - 1, numcolonna + 1)
                     Cells(j, numcolonna + 2) = Cells(j - 1, numcolonna + 2)
                     Cells(j, numcolonna + 3) = Cells(j - 1, numcolonna + 3)
                     'mezzecoste
                     Cells(j, 14) = Cells(j - 1, 14)
                     Cells(j, 15) = Cells(j - 1, 15)
              Next j
              Cells(numriga, numcolonna + 1) = ""
              Cells(numriga, numcolonna + 2) = ""
              Cells(numriga, numcolonna + 3) = ""
              'mezzecoste
              Cells(numriga, 14) = ""
              Cells(numriga, 15) = ""
      Else
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo1 = Sheets("F02").Cells(1213, 2)
          Else
                testo1 = Sheets("F02").Cells(1228, 2)
          End If
          MsgBox (testo1)
      End If

End Sub

Private Sub Svuota_prof_Click()
         Dim result As Integer
         
         If Sheets("HOME").flag_italiano.value Then
                testomes = Sheets("F02").Cells(1214, 2)
         Else
                testomes = Sheets("F02").Cells(1229, 2)
         End If
         result = MsgBox(testomes, 52, "Warning")
         If result <> 6 Then Exit Sub
         nomefoglio = "F02"
         Call elimina_profilo_salita(nomefoglio)
End Sub

Private Sub svuota_prof_d_Click()
         Dim result As Integer
         
         If Sheets("HOME").flag_italiano.value Then
                testomes = Sheets("F02").Cells(1214, 7)
         Else
                testomes = Sheets("F02").Cells(1229, 7)
         End If
         result = MsgBox(testomes, 52, "Warning")
         If result <> 6 Then Exit Sub
         nomefoglio = "F02"
         Call elimina_profilo_discesa(nomefoglio)
End Sub

Private Sub vedi_profilo_discesa_Click()
        Call Set_Ground_Profile
        Sheets("F02").Cells(57, 23).Select
End Sub

Private Sub vedi_profilo_salita_Click()
        Call Set_Ground_Profile
        Sheets("F02").Cells(15, 23).Select
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
       
       On Local Error GoTo err_range
       Dim testo1, testo2, test3, testo4 As String
       
       numsost = Target.value
       numriga = Val(Target.Row)
       numcolonna = Val(Target.Column)
       If Sheets("HOME").flag_italiano.value Then
                                                testo1 = Sheets("F02").Cells(1206, 2)
                                                testo2 = Sheets("F02").Cells(1207, 2)
       Else
                testo1 = Sheets("F02").Cells(1221, 2)
                testo2 = Sheets("F02").Cells(1222, 2)
       End If
       
       If (numriga > 14 And numriga < 1113) And (numcolonna = 2 Or numcolonna = 7) Then
           If numcolonna = 2 Then
             test_riga.Caption = "< Riga =" + Str(numsost) + " > " + testo1
             cek_ins_del = True
             test_riga_d.Caption = testo2
           Else
             test_riga_d.Caption = "< Riga =" + Str(numsost) + " > " + testo1
             cek_ins_del_d = True
             test_riga.Caption = testo2
           End If
       Else
            test_riga.Caption = testo2
            test_riga_d.Caption = testo2
            cek_ins_del = False
            cek_ins_del_d = False
       End If
       Exit Sub
       
err_range:
          Beep
          If Sheets("HOME").flag_italiano Then
                                              testo1 = Sheets("F02").Cells(1212, 2)
          Else
                testo1 = Sheets("F02").Cells(1227, 2)
          End If
          MsgBox (testo1)
          Exit Sub
End Sub
