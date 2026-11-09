Attribute VB_Name = "Program"
Option Explicit
Dim Inicio As Date

Sub Atualizar(Optional ByVal Data As String = "default")
    
    ' Desativa eventos e atualiza��es para melhor processamento das informa��es
    With Application
        .DisplayFullScreen = False 'para aparecer o status bar
        .Wait (Now + TimeValue("0:00:01")) 'para dar tempo de refletir o fullscreem e pegar o status bar
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .StatusBar = True
    End With
    
    Data = IIf(Data = "default", Date - 1, Data)
    
    'Call Program.VisibleTrue
         
    Call Program.Sincronizar(Data)
    
    
    ' Contabiliza o tempo de execu��o da etapa anterior e mostra a pr�xima etapa
    Application.StatusBar = "fn_FarolVozSkills atualizada em: " & Format(Now - Inicio, "hh:mm:ss") & " | Atualizando paineis..."
    Debug.Print ("fn_FarolVozSkills atualizada em: " & Format(Now - Inicio, "hh:mm:ss"))
    Inicio = Now
    
    'Olhar se o dash vai atualizar quando mudar o m�s
    ActiveWorkbook.RefreshAll
    
    Call Program.OcultarLinhasFarol("Farol")
    
    
    ' Ativa eventos e atualiza��es
    With Application
        .Calculation = xlCalculationAutomatic
        .DisplayFullScreen = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .StatusBar = False
    End With

End Sub

Sub Sincronizar(Optional ByVal Data As String = "default")
    
    Dim conn As New Connection
    Dim rec As Object
    Dim sql As String
    Dim Mes As Integer
    Dim Ano As Integer
    Dim DataInicio As Date
    Dim DataFim As Date
    Dim DataAuxiliar As Date
    
    ' Repita esse c�digo para execu��o de cada etapa de atualiza��o do seu relat�rio
    Application.StatusBar = "Realizando Conex�o com o Banco de Dados"
    'Debug.Print (Format(Now - Inicio, "hh:mm:ss"))
    Inicio = Now
    
    With conn
        .Tipo = cnnSqlServer
        .Server = "SE0000BD1994\INSTANCE1994"
        .Base = "BANCO_MIS"
        .Trusted = True
        .CommandTimeOut = 120
        .Abrir
    End With
    
    'Data = IIf(Data = "default", Date - 1, Data)
    DataInicio = DateSerial(Year(Data), Month(Data), 1)
    DataFim = DateSerial(Year(Data), Month(Data) + 1, 0)
    DataAuxiliar = Date - 1
    Mes = Month(Data)
    Ano = Year(Data)
    
    ' Contabiliza o tempo de execu��o da etapa anterior e mostra a pr�xima etapa
    Application.StatusBar = "Conex�o realizada em: " & Format(Now - Inicio, "hh:mm:ss") & " | Atualizando vw_GlossarioUsuariosExcecao..."
    Debug.Print ("Conex�o realizada em: " & Format(Now - Inicio, "hh:mm:ss"))
    Inicio = Now
    
    '-------------------------- Base Gloss�rio -------------------------
    
    sql = "set dateformat dmy "
    sql = sql + "select * "
    sql = sql + "from BANCO_MIS.dbo.vw_GlossarioUsuariosExcecao "
    
    Set rec = conn.ExecuteReader(sql)
    
    If Not rec Is Nothing Then
        With Glossario
            .Range("AD2:AD" & .Rows.Count).ClearContents
            .Range("AD2").CopyFromRecordset rec
        End With
    End If
    
    ' Contabiliza o tempo de execu��o da etapa anterior e mostra a pr�xima etapa
    Application.StatusBar = "vw_GlossarioUsuariosExcecao atualizada em: " & Format(Now - Inicio, "hh:mm:ss") & " | Atualizando fn_SumarioVoz..."
    Debug.Print ("vw_GlossarioUsuariosExcecao em: " & Format(Now - Inicio, "hh:mm:ss"))
    Inicio = Now
    
    '-------------------------- Base Hsplit -------------------------
    
    sql = "Set dateformat dmy "
    'sql = sql + "declare @DataInicio date = '" & DataInicio & "', @DataFim date = '" & DataFim & "' "
    sql = sql + "select * "
    sql = sql + "from BANCO_MIS.Sac.fn_FarolVoz ('" & Data & "')" '(@DataInicio) "
    
    Set rec = conn.ExecuteReader(sql)
    
    If Not rec Is Nothing Then
        
        With BaseCMS
            '.Range("A2:Aw" & .Rows.Count).ClearContents
            If .Range("Tb_Base")(1) <> "" Then .Range("Tb_Base").Delete
            .Range("A2").CopyFromRecordset rec
        End With
    
    End If
    
    ' Contabiliza o tempo de execu��o da etapa anterior e mostra a pr�xima etapa
    Application.StatusBar = "fn_FarolVoz atualizada em: " & Format(Now - Inicio, "hh:mm:ss") & " | Atualizando fn_FarolTemposOperador..."
    Debug.Print ("fn_FarolVoz atualizada em: " & Format(Now - Inicio, "hh:mm:ss"))
    Inicio = Now
    
    '-------------------------- Base Hagent -------------------------
    
    sql = "Set dateformat dmy "
    'sql = sql + "declare @DataInicio date = '" & DataInicio & "', @DataFim date = '" & DataFim & "' "
    sql = sql + "select * "
    sql = sql + "from BANCO_MIS.Sac.fn_FarolVozTemposOperador ('" & Data & "')"

    Set rec = conn.ExecuteReader(sql)

    If Not rec Is Nothing Then

        With BaseOperadorCMS
            ''.Range("A2:Aw" & .Rows.Count).ClearContents
            If .Range("Tb_BaseOperador")(1) <> "" Then .Range("Tb_BaseOperador").Delete
            .Range("A2").CopyFromRecordset rec
        End With

    End If

    ' Contabiliza o tempo de execu��o da etapa anterior e mostra a pr�xima etapa
    Application.StatusBar = "fn_FarolVozTemposOperador atualizada em: " & Format(Now - Inicio, "hh:mm:ss") & " | Atualizando fn_FarolVozSkills..."
    Debug.Print ("fn_FarolVozTemposOperador atualizada em: " & Format(Now - Inicio, "hh:mm:ss"))
    Inicio = Now
    
    
    '-------------------------- Base Skills -------------------------
    
    sql = "Set dateformat dmy "
    'sql = sql + "declare @DataInicio date = '" & DataInicio & "', @DataFim date = '" & DataFim & "' "
    sql = sql + "select * "
    sql = sql + "from BANCO_MIS.Sac.fn_FarolVozSkills ('" & Data & "') "

    Set rec = conn.ExecuteReader(sql)

    If Not rec Is Nothing Then

        With BaseCMSSKILLS
            ''.Range("A2:Aw" & .Rows.Count).ClearContents
            If .Range("Tb_BaseSKILLS")(1) <> "" Then .Range("Tb_BaseSKILLS").Delete
            .Range("A2").CopyFromRecordset rec
        End With

    End If
    
    conn.Fechar

End Sub
 
Function fn_PermissaoBases() As Boolean
    
    'Essa fun��o serve para dar permiss�o as bases do relat�rio
    'Apenas usuarios chaves devem ter esse acesso
    'Normalmente � algum focal ou planejamento
    
    Select Case UCase(Environ$("username"))
        
        Case "FDSANTOS", "BI"
             
             fn_PermissaoBases = True
             
        Case Else
        
            fn_PermissaoBases = False
            
    End Select

End Function

Sub ChamarFormAtualizacao()
Attribute ChamarFormAtualizacao.VB_ProcData.VB_Invoke_Func = "A\n14"

    FrmAtualizador.Show

End Sub

Private Sub VisibleTrue()
    
    Application.ScreenUpdating = False
    
    
    Base.Visible = xlSheetVisible
    BaseSKILLS.Visible = xlSheetVisible
    BaseOperadorCMS.Visible = xlSheetVisible
    
    ConfigMenu.Visible = xlSheetVisible
    'ConfigDash.Visible = xlSheetVisible
    'Dashboard.Visible = xlSheetVisible
    
    Glossario.Visible = xlSheetVisible
    Versionamento.Visible = xlSheetVisible
    
    ActiveWindow.DisplayWorkbookTabs = True
    
    Application.ScreenUpdating = True

End Sub

Private Sub VisibleFalse()
    
    Application.ScreenUpdating = False
    
    Base.Visible = xlSheetVeryHidden
    BaseSKILLS.Visible = xlSheetVeryHidden
    BaseOperadorCMS.Visible = xlSheetVeryHidden
    
    ConfigMenu.Visible = xlSheetVeryHidden
    'ConfigDash.Visible = xlSheetVeryHidden
    'Dashboard.Visible = xlSheetVeryHidden
    
    Glossario.Visible = xlSheetVeryHidden
    Versionamento.Visible = xlSheetVeryHidden
    
    Application.ScreenUpdating = True

End Sub

Private Sub VisibleEnvio()
    
    If UCase(Environ$("username")) = "BI" Then
        
        Sheets("Envio").Visible = True
        
    End If

End Sub

Private Sub OcultarLinhasFarol(Sheet As String)

    Dim xRg As Range
    Dim Ocultar As Range
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(Sheet)
    
    'ws.Select
    ws.Calculate
    
    Application.ScreenUpdating = False

    ws.Range("C11:C100").EntireRow.Hidden = False

    For Each xRg In ws.Range("C11:C41")
        If xRg.value = "" Then
            If Ocultar Is Nothing Then
                Set Ocultar = xRg
            Else
                Set Ocultar = Union(Ocultar, xRg)
            End If
        End If
    Next xRg

    If Not Ocultar Is Nothing Then Ocultar.EntireRow.Hidden = True

    Application.ScreenUpdating = True
    
End Sub

