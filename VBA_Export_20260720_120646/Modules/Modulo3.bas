Attribute VB_Name = "Modulo3"
Sub Vai_a_F00()
Attribute Vai_a_F00.VB_Description = "Macro registrata il 26/08/02 da *"
Attribute Vai_a_F00.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Vai_a_F00 Macro
' Macro registrata il 26/08/02 da *
'

'
    Sheets("MENU").Select
    Range("B2").Select
End Sub
Sub Vai_a_HOME()
Attribute Vai_a_HOME.VB_Description = "Macro registrata il 26/08/02 da *"
Attribute Vai_a_HOME.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Vai_a_HOME Macro
' Macro registrata il 26/08/02 da *
'

'
    Sheets("HOME").Select
    Range("B1").Select
End Sub

Sub vai_a_menu()
    'leggi path_programmi (mia_path)
    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close

    Workbooks.Open FileName:=mia_path + "\" + "MENU.xls"
    'Windows("MENU.xls").Activate
    'ActiveWindow.WindowState = xlMaximized
    Windows(2).Close
End Sub

