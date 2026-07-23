VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} idraulic_input 
   Caption         =   "Verifica della linea per guasto al tenditore idraulico : impostazione dei parametri di riferimento"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   OleObjectBlob   =   "idraulic_input.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "idraulic_input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub start_verifica_Click()
       XIndex = 2
       Ipocalc = 2
       idraulic_input.Hide
       
       Call Sheets("F05").prepara_calcolo
       Call esegui_verifica(XIndex)
       
       'genera tabella max_min
       Sheets("F10").Select
       Call stampa_maxmin(4)
       Sheets("F05").Select
       'genera tabella potenze
       Sheets("F11").Select
       nomeipotesi$ = Sheets("F11").ver_04.Caption
       Call stampa_potenze(4, nomeipotesi$)
       Sheets("F05").Select
       'genera foglio dati generali
       Sheets("F12").Select
       Call rel_gen
       Sheets("F05").Select
       
       Unload idraulic_input
End Sub

Private Sub stop_verifica_Click()
          Unload idraulic_input
End Sub


Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
      If KeyCode = 13 Then
                         cella_num = 3
                         Call modifica_parametri(cella_num)
      End If
End Sub


Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
      If KeyCode = 13 Then
                         cella_num = 2
                         Call modifica_parametri(cella_num)
      End If
End Sub


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
      If KeyCode = 13 Then
                         cella_num = 1
                         Call modifica_parametri(cella_num)
      End If
End Sub

Private Sub UserForm_Initialize()
          idraulic_input.Caption = Sheets("F13").Cells(30, 1)
          Label1.Caption = Sheets("F13").Cells(31, 1)
          Label2.Caption = Sheets("F13").Cells(32, 1)
          Label3.Caption = Sheets("F13").Cells(33, 1)
          Label4.Caption = Sheets("F13").Cells(34, 1)
          Label5.Caption = Sheets("F13").Cells(35, 1)
          Label6.Caption = Sheets("F13").Cells(36, 1)
          Label7.Caption = Sheets("F13").Cells(37, 1)
          Label8.Caption = Sheets("F13").Cells(38, 1)
          Label9.Caption = Sheets("F13").Cells(39, 1)
          Label10.Caption = Sheets("F13").Cells(40, 1)
          Label11.Caption = Sheets("F13").Cells(41, 1)
          Label12.Caption = Sheets("F13").Cells(42, 1)
          Opt1.Caption = Sheets("F13").Cells(43, 1)
          Opt2.Caption = Sheets("F13").Cells(44, 1)
          start_verifica.Caption = Sheets("F13").Cells(45, 1)
          stop_verifica.Caption = Sheets("F13").Cells(46, 1)
End Sub

Public Sub modifica_parametri(cella_num)

           Select Case cella_num
                 Case 1
                        corsaz(8, 1) = KTERMICO# * lanello / 2 * Val(idraulic_input.TextBox1.Text)
                        tstringa = Format(corsaz(8, 1), "###0.00")
                        Call codifica(tstringa)
                        Label34.Caption = tstringa
                 Case 2
                        tstringa = idraulic_input.TextBox2.Text
                        Call codifica(tstringa)
                        idraulic_input.TextBox2.Text = tstringa
                        corsaz(9, 1) = Val(idraulic_input.TextBox2.Text)
                 Case 3
                        tstringa = idraulic_input.TextBox3.Text
                        Call codifica(tstringa)
                        idraulic_input.TextBox3.Text = tstringa
                        corsaz(10, 1) = Val(idraulic_input.TextBox3.Text)
           End Select

           If ZETA <> "F" Then
                       corsaz(11, 1) = corsaz(5, 1) + corsaz(6, 1) + corsaz(7, 1) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                       corsaz(12, 1) = corsaz(5, 2) + corsaz(6, 2) + corsaz(7, 2) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                       corsaz(13, 1) = corsaz(5, 3) + corsaz(6, 3) + corsaz(7, 3) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
           Else
                       corsaz(11, 1) = corsaz(6, 1) + corsaz(7, 1) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                       corsaz(12, 1) = corsaz(6, 2) + corsaz(7, 2) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
                       corsaz(13, 1) = corsaz(6, 3) + corsaz(7, 3) + corsaz(8, 1) + corsaz(9, 1) + corsaz(10, 1)
           End If
           maxcorsa = 0
           If corsaz(11, 1) > maxcorsa Then maxcorsa = corsaz(11, 1)
           If corsaz(12, 1) > maxcorsa Then maxcorsa = corsaz(12, 1)
           If corsaz(13, 1) > maxcorsa Then maxcorsa = corsaz(13, 1)
           corsaz(14, 1) = Int(corsaz(11, 1) * 10) / 10 + 0.1

           tstringa = Format(corsaz(11, 1), "###0.00")
           Call codifica(tstringa)
           idraulic_input.Label35.Caption = tstringa
           
           tstringa = Format(corsaz(12, 1), "###0.00")
           Call codifica(tstringa)
           idraulic_input.Label36.Caption = tstringa
           
           tstringa = Format(corsaz(13, 1), "###0.00")
           Call codifica(tstringa)
           idraulic_input.Label37.Caption = tstringa
           
           tstringa = Format(corsaz(14, 1), "###0.00")
           Call codifica(tstringa)
           idraulic_input.TextBox4.Text = tstringa

End Sub
