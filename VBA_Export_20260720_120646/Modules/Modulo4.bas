Attribute VB_Name = "Modulo4"
Sub Info_Tempus_Clic()
         ActiveWorkbook.FollowHyperlink Address:=mia_path + "\" + "tempus.html", NewWindow:=True
End Sub

Sub help_web_Clic()
    Close
    'leggi path_programmi (mia_path)
    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close

    If Sheets("HOME").flag_italiano.value Then
              ActiveWorkbook.FollowHyperlink Address:=mia_path + "\" + "help_sif\index.htm", NewWindow:=True
    Else
              ActiveWorkbook.FollowHyperlink Address:=mia_path + "\" + "help_sif_uk\index.htm", NewWindow:=True
    End If
End Sub


Sub help_home()
             nomefoglio = "HOME.htm"
             Call help_fogli(nomefoglio)
End Sub

Sub help_menu()
             nomefoglio = "menu.htm"
             Call help_fogli(nomefoglio)
End Sub

Sub help_F01()
             nomefoglio = "F01.htm"
             Call help_fogli(nomefoglio)
End Sub

Sub help_F02()
             nomefoglio = "F02.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F03()
             nomefoglio = "F03.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F04()
             nomefoglio = "F04.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F05()
             nomefoglio = "F05.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F06()
             nomefoglio = "F06.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F07()
             nomefoglio = "F07.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F08()
             nomefoglio = "F08.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F09()
             nomefoglio = "F09.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F10()
             nomefoglio = "F10.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F11()
             nomefoglio = "F11.htm"
             Call help_fogli(nomefoglio)
End Sub
Sub help_F12()
             nomefoglio = "F12.htm"
             Call help_fogli(nomefoglio)
End Sub

Sub help_fogli(nomefoglio)
    Close
    'leggi path_programmi (mia_path)
    Open "c:\tmp_sif\prg_path" For Input As #1
    Input #1, mia_path
    Close

    If Sheets("HOME").flag_italiano.value Then
              ActiveWorkbook.FollowHyperlink Address:=mia_path + "\" + "help_sif\pagine\monofuni_" + nomefoglio, NewWindow:=True
    Else
              ActiveWorkbook.FollowHyperlink Address:=mia_path + "\" + "help_sif_uk\pagine\monofuni_" + nomefoglio, NewWindow:=True
    End If
End Sub


