VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub flag_inglese_Click()


    'leggi path_programmi (mia_path)
    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close
    'scrivi lingua
    Open "c:\tmp_sif\flag_type" For Output As #1
    Print #1, "UK"
    Close
    Call traduci
    
End Sub

Private Sub flag_italiano_Click()

    'leggi path_programmi (mia_path)
    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close
    'scrivi lingua
    Open "c:\tmp_sif\flag_type" For Output As #1
    Print #1, "IT"
    Close
    Call traduci

End Sub



Private Sub Label1_Click()
   On Local Error GoTo err_chiave
                         Dim chiave As String
                         'rimozione password fogli
                         chiave = "*"
                         Beep
                         chiave = InputBox("INSERIRE LA PASSWORD DI RIMOZIONE", "SHEETS UNPROTECT MODE")
                         If chiave = "vit210147vit" Then
                                                        Call psw_remove
                         Else
                               Beep
                               MsgBox ("PROTEZIONE NON RIMOSSA")
                         End If
Exit Sub

err_chiave:
          Beep
          End

End Sub

Private Sub Worksheet_Activate()
        Sheets("HOME").Cells(1, 1).Select
End Sub


