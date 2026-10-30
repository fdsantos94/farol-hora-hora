VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(cancel As Boolean)

    On Error Resume Next
    
    ' Registra o logout do relatório
    Call RegistroLog.Log(0)
    
    If UCase(Environ$("Username")) <> "BI" Then
        ' Registra o logo das sheets do relatório
        Call RegistroLog.RegistarLogsSheets
        
        Set wsLogs = ThisWorkbook.Sheets(wsNomeLogs)
    
        wsLogs.Range("A2:F" & Rows.Count).ClearContents
    
    End If

End Sub

Private Sub Workbook_Open()

    Application.ScreenUpdating = False
    
    ' Registra o login do relatório
    Call RegistroLog.Log(1)
    
    If UCase(Environ$("Username")) <> "BI" Then
        ' Cria a base de logs das sheets
        Call RegistroLog.RecriarBaseLogsSheets
        
    End If
           
    ' fecha os campos da tabela dinamica
    Application.CommandBars("PivotTable Field List").Visible = False
    
    ActiveWorkbook.RefreshAll
    
    Capa.Select
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    
    'coleta o nome quando uma nova sheet for criada
    UltimaPlanilha = Sh.Name
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    'valida se é uma nova planilha, se for, no primeiro momento não precisa coletar o log
    If Sh.Name = UltimaPlanilha Then
    
        UltimaPlanilha = "" 'limpa a variável
        Exit Sub
        
    End If

    'não coleta o log da capa e da sheet controle de logs
    If Sh.Name <> wsNomeLogs And Sh.Name <> "Capa" Then
    
        Call RegistroLog.ColetarRegistrosLogSheets
        
    End If
    
End Sub


