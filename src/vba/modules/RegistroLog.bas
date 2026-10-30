Attribute VB_Name = "RegistroLog"
Public wsLogs As Worksheet
Public Const wsNomeLogs As String = "ControleLogsSheets"
Public Const CodMis As Integer = 9999
Public UltimaPlanilha As String

Private IdRelatorio As Integer, DataLog As Date, UserName As String, Computer As String, UserDomain As String, Status As String

Sub Log(ByVal Evento As Integer)
    
    '-----------------------------------------------------------------------------------------
    'como utilizar coloca na pasta de trabalho
    '
    'Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '
    '   RegistroLog.Log(0)
    '
    'End Sub
    '
    'Private Sub Workbook_Open()
    '
    '   RegistroLog.Log(1)
    '
    'End Sub
    
    '-----------------------------------------------------------------------------------------
    
    'pegar o id do relatorio na tabela db_portal_mis.dbo.tb_relatorios
    
    IdRelatorio = CodMis
    
    '-----------------------------------------------------------------------------------------
    
    If Evento = 1 Then
        Status = "Login"
    Else
        Status = "Logout"
    End If
    
    DataLog = Format(Now, "yyyy-mm-dd hh:mm:ss")
    UserName = Environ$("username")
    Computer = Environ$("ComputerName")
    UserDomain = Environ$("UserDomain")
    
    Call InsertLog(IdRelatorio, DataLog, UserName, Computer, UserDomain, Status)

End Sub

Private Sub InsertLog(ByVal IdRelatorio As Integer, DataLog As Date, UserName As String, Computer As String, UserDomain As String, ByVal Status As String)
    
    'Conexão Banco de Dados
    Dim conn As New Connection
    Dim sql As String
    Dim rec As Object
    Dim primeiralinha As Integer
    Dim ultimalinha As Integer
            
    On Error GoTo sair
    
    With conn
        .Tipo = cnnSqlServer
        .Server = "SE0000BD1994\INSTANCE1994"
        .Base = "BANCO_MIS"
        .User = "user123"
        .Password = "senha123"
        .Abrir
        .CommandTimeOut = 90
    End With
    
    sql = ""
    sql = sql + " set dateformat dmy "
    sql = sql + " insert into LogonLogoffRelatoriosMis values ('" & IdRelatorio & "', '" & DataLog & "', '" & UserName & " ', '" & Computer & "', '" & UserDomain & "', '" & Status & "') "
    
    Set rec = conn.ExecuteReader(sql)
    
    conn.Fechar
    
sair:
        Exit Sub

End Sub

Sub RecriarBaseLogsSheets()

    Application.DisplayAlerts = False
    
    ' Se já existir, deleta
    On Error Resume Next
    ThisWorkbook.Sheets(wsNomeLogs).Visible = xlSheetVisible
    ThisWorkbook.Sheets(wsNomeLogs).Delete
    On Error GoTo 0
    
    ' Cria novamente
    Set wsLogs = ThisWorkbook.Sheets.Add
    wsLogs.Name = wsNomeLogs
    
    ' Cabeçalhos
    With wsLogs
        .Range("A1").value = "Data"
        .Range("B1").value = "Horário"
        .Range("C1").value = "Máquina"
        .Range("D1").value = "Usuário"
        .Range("E1").value = "SheetAcessada"
        .Range("F1").value = "IdRelatorio"
        '.Visible = xlSheetHidden
        .Visible = xlSheetVeryHidden
    End With
    
    Application.DisplayAlerts = True

End Sub

Sub ColetarRegistrosLogSheets()
   
    Dim DataLog As String
    Dim HoraAcessada As Date
    Dim Usuario As String
    Dim Maquina As String
    Dim NomePlanilha As String
    
    Dim linha As Integer
    Dim i As Integer
    Dim rng As Range
    
    DataLog = Format(Now, "yyyy-mm-dd")
    HoraAcessada = Format(Now, "hh:mm:ss")
    Usuario = Environ$("username")
    Maquina = Environ$("ComputerName")
    NomePlanilha = ActiveSheet.Name
    
    Set wsLogs = ThisWorkbook.Sheets(wsNomeLogs)
    
    linha = wsLogs.Range("A1").CurrentRegion.Rows.Count + 1
    
    ' Verifica se já existe a planilha registrada na coluna E
    Set rng = wsLogs.Range("E:E").Find(What:=NomePlanilha, LookAt:=xlWhole, MatchCase:=False)
    
    If Not rng Is Nothing Then
        'já existe ? não registra novamente
        Exit Sub
    End If
    
    With wsLogs
        .Range("A" & linha).value = DataLog
        .Range("B" & linha).value = HoraAcessada
        .Range("C" & linha).value = Maquina
        .Range("D" & linha).value = Usuario
        .Range("E" & linha).value = NomePlanilha
        .Range("F" & linha).value = CodMis
    End With
    
    wsLogs.Cells.Find What:="", LookAt:=xlPart
    
End Sub

Sub RegistarLogsSheets()

    Dim conn As New Connection
    Dim rec As Object
    Dim sql As String
    
    Dim linha As Integer
    Dim i As Integer
    
    On Error GoTo sair

    ' Conexão Padrão
    With conn
        .Tipo = cnnSqlServer
        .Server = "SE0000BD1994\INSTANCE1994"
        .Base = "BANCO_MIS"
        .User = "user123"
        .Password = "senha123"
        .Abrir
        .CommandTimeOut = 90
    End With
    
    ' Planilha Logs
    Set wsLogs = ThisWorkbook.Sheets(wsNomeLogs)
    
    ' Total registros
    linha = wsLogs.Range("A1").CurrentRegion.Rows.Count
    
    ' Loop para popular o log em banco
    For i = 2 To linha
    
        sql = "set dateformat ymd insert into MIS.Interno.LogAcessoSheets values('" _
                & Format(wsLogs.Range("A" & i).value, "yyyy-mm-dd") & "','" _
                & Format(wsLogs.Range("B" & i).value, "hh:mm:ss") & "','" _
                & wsLogs.Range("C" & i).value & "','" _
                & wsLogs.Range("D" & i).value & "','" _
                & wsLogs.Range("E" & i).value & "'," _
                & wsLogs.Range("F" & i).value & ")"
        
        Set rec = conn.ExecuteReader(sql)
    
    Next
    
    conn.Fechar

sair:
    Exit Sub
    
End Sub

