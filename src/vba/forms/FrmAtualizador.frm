VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAtualizador 
   Caption         =   "Atualizador"
   ClientHeight    =   2712
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3900
   OleObjectBlob   =   "FrmAtualizador.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmAtualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_atualizar_Click()
    
    Dim Inicio As Date
    
    Inicio = Now
    
    If IsDate(Me.txtData.Text) Then
        Program.Atualizar Me.txtData.Text
    Else
        Me.txtData.SelStart = 0
        Me.txtData.SelLength = Len(Me.txtData.Text)
        Exit Sub
    End If
    
    MsgBox "Dados atualizados com sucesso. " & vbNewLine & "Tempo de execução : " & Format(Now - Inicio, "hh:mm:ss"), vbExclamation, ""
    
    Unload Me

End Sub

Private Sub UserForm_Initialize()
    With Me
        .txtData.Text = Format(Date, "dd/mm/yyyy")
        .txtData.SelStart = 0
        .txtData.SelLength = 2
        .Label3.Caption = "MIS - " & Year(Date)
    End With
End Sub
