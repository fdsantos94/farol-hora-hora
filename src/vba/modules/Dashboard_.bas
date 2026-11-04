Attribute VB_Name = "Dashboard_"
Option Explicit

Sub Btn_Dashboard()
    
    Dim i As Integer
    
    Application.ScreenUpdating = False
        
    'Dashboard.Activate
    
    With Dashboard
        .Activate
        .Range("A1:AC45").Select
    End With
    
    For i = 1 To 2
        Call ConfigPlanilha.ModoApresentacaoOn
    Next
    
    Dashboard.Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub NavegarMenus()

    Dim BotaoClicado As String
    
    ' Captura dinamicamente o nome exato do Shape/Botão que acionou o clique
    On Error Resume Next
    BotaoClicado = Application.Caller
    On Error GoTo 0
    
    If BotaoClicado = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    
    With ActiveSheet
        ' 1. Reset de estado: oculta todos os grupos de elementos em uma única chamada vetorial
        .Shapes.Range(Array( _
            "ConjuntoOnePage", _
            "ConjuntoVolumetriaComparativo", _
            "ConjuntoVolumetriaUnico", _
            "ConjuntoNs", _
            "SubMenuVolumetria", _
            "SubMenuVolumetria2", _
            "GrafNS1", _
            "GrafNS2", _
            "SubMenuTma", _
            "ConjuntoTelaTmaIntradia", _
            "ConjuntoTelaTmaAderencia" _
        )).Visible = False
        
        ' 2. Exibição cirúrgica: exibe apenas os componentes pertencentes ao botão clicado
        Select Case BotaoClicado
        
            Case "BtMenuOnePage" ' Adapte para o nome real do seu Shape no Excel
            
                Call Dashboard_.FormatarMenus(BotaoClicado)
                
                .Shapes.Range(Array("ConjuntoOnePage")).Visible = True
                
                
                
                
            Case "BtMenuVolumetria"
            
                Call Dashboard_.FormatarMenus(BotaoClicado)
            
                .Shapes.Range(Array("ConjuntoVolumetriaComparativo", "ConjuntoVolumetriaComparativo" _
                                    , "SubMenuVolumetria")).Visible = True
                
                Call Dashboard_.FormatarMenus("BtVolComparativo")
                
                
            Case "BtVolUnico"
            
                .Shapes.Range(Array("ConjuntoVolumetriaUnico", "SubMenuVolumetria")).Visible = True
                
                Call Dashboard_.FormatarMenus("BtVolUnico")
                
            Case "BtVolComparativo"
            
                Call Dashboard_.FormatarMenus("BtVolComparativo")
                
                .Shapes.Range(Array("ConjuntoVolumetriaComparativo", "SubMenuVolumetria")).Visible = True
                
            Case "BtVolNs"
            
                .Shapes.Range(Array("ConjuntoNs", "SubMenuVolumetria", "SubMenuVolumetria2", "GrafNS1")).Visible = True
                
                Call Dashboard_.FormatarMenus("BtVolNs")

                Call Dashboard_.FormatarMenus("BtVolAderencia")
            
            Case "BtVolAderencia"
            
                Call Dashboard_.FormatarMenus("BtVolAderencia")
                
                .Shapes.Range(Array("ConjuntoNs", "SubMenuVolumetria", "SubMenuVolumetria2", "GrafNS1")).Visible = True

            Case "BtVolDetalhamento"
            
                Call Dashboard_.FormatarMenus("BtVolDetalhamento")
                
                .Shapes.Range(Array("ConjuntoNs", "SubMenuVolumetria", "SubMenuVolumetria2", "GrafNS2")).Visible = True
                                
                                
                                
                                
                                
            Case "BtMenuTma"
            
                Call Dashboard_.FormatarMenus(BotaoClicado)
                
                .Shapes.Range(Array("ConjuntoTelaTmaIntradia", "SubMenuTma")).Visible = True
                
                Call Dashboard_.FormatarMenus("BtTmaIntradia")
            
            Case "BtTmaIntradia"
            
                Call Dashboard_.FormatarMenus("BtTmaIntradia")
                
                .Shapes.Range(Array("ConjuntoTelaTmaIntradia", "SubMenuTma")).Visible = True
                              
            Case "BtTmaAderencia"
            
                .Shapes.Range(Array("SubMenuTma", "ConjuntoTelaTmaAderencia")).Visible = True
                
                Call Dashboard_.FormatarMenus("BtTmaAderencia")
                
                
                
                
                
            Case "BtMenuTempos"
            
                Call Dashboard_.FormatarMenus(BotaoClicado)
                
                                   
            ' Adicione novos Cases à medida que o painel expandir
            Case Else
                ' Comportamento neutro para chamadas não registradas
        End Select
        
        .Range("A1").Select
    End With
    
    Application.ScreenUpdating = True
    
End Sub

' ==============================================================================
' CONTROLADOR DE NAVEGAÇÃO MULTI-MENU (UNIFICADO)
' ==============================================================================
Private Sub FormatarMenus(Optional ByVal nomeMenu As String = "")
    
    On Error GoTo TratarErro
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------------------------------
    ' 1. RESOLUÇÃO DO BOTÃO DISPARADOR
    ' --------------------------------------------------------------------------
    Dim botaoAlvo As String
    If Len(Trim$(nomeMenu)) = 0 Then
        botaoAlvo = CStr(Application.Caller)
    Else
        botaoAlvo = nomeMenu
    End If
    
    ' --------------------------------------------------------------------------
    ' 2. DEFINIÇÃO DOS GRUPOS E SELEÇÃO DO CONJUNTO AFETADO
    ' --------------------------------------------------------------------------
    Dim menuPrincipal As Variant
    Dim subMenuEsquerdo As Variant
    Dim subMenuDireito As Variant
    Dim subMenuTma As Variant
    
    menuPrincipal = Array("BtMenuOnePage", "BtMenuVolumetria", "BtMenuTma", "BtMenuTempos")
    subMenuEsquerdo = Array("BtVolComparativo", "BtVolUnico", "BtVolNs")
    subMenuDireito = Array("BtVolAderencia", "BtVolDetalhamento")
    subMenuTma = Array("BtTmaIntradia", "BtTmaAderencia")
    
    Dim grupoAtual As Variant
    Dim ehEstiloPrincipal As Boolean
    
    If fn_ExisteNoArray(botaoAlvo, menuPrincipal) Then
        grupoAtual = menuPrincipal
        ehEstiloPrincipal = True
    ElseIf fn_ExisteNoArray(botaoAlvo, subMenuEsquerdo) Then
        grupoAtual = subMenuEsquerdo
        ehEstiloPrincipal = False
    ElseIf fn_ExisteNoArray(botaoAlvo, subMenuDireito) Then
        grupoAtual = subMenuDireito
        ehEstiloPrincipal = False
    ElseIf fn_ExisteNoArray(botaoAlvo, subMenuTma) Then
        grupoAtual = subMenuTma
        ehEstiloPrincipal = False
    Else
        GoTo SairDaMacro
    End If
    
    ' --------------------------------------------------------------------------
    ' 3. PALETA DE CORES GLOBAL (CONSTANTES COM RGB CALCULADO)
    ' --------------------------------------------------------------------------
    Dim corLaranja As Long, corBranca As Long, corCinza As Long
    corLaranja = RGB(255, 107, 0)
    corBranca = RGB(255, 255, 255)
    corCinza = RGB(128, 128, 128)
    
    ' --------------------------------------------------------------------------
    ' 4. LOOP DE APLICAÇÃO DE DESIGN ATIVO/INATIVO
    ' --------------------------------------------------------------------------
    Dim itemBotao As Variant
    Dim shpBotao As Shape
    Dim ehBotaoClicado As Boolean
    
    For Each itemBotao In grupoAtual
        Set shpBotao = Dashboard.Shapes(CStr(itemBotao))
        ehBotaoClicado = (shpBotao.Name = botaoAlvo)
        
        With shpBotao
            .Line.Visible = msoFalse
            .TextFrame2.TextRange.Font.Bold = msoTrue
            
            If ehBotaoClicado Then
                ' --- ESTADO ATIVO (ON) ---
                .Fill.Visible = msoTrue
                .Fill.Solid
                .Fill.Transparency = 0#
                .Shadow.Visible = msoTrue
                .Shadow.Type = msoShadow21
                
                If ehEstiloPrincipal Then
                    .Fill.ForeColor.RGB = corLaranja
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corBranca
                Else
                    .Fill.ForeColor.RGB = corBranca
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corLaranja
                End If
            Else
                ' --- ESTADO INATIVO (OFF) ---
                .Shadow.Visible = msoFalse
                
                If ehEstiloPrincipal Then
                    .Fill.Visible = msoTrue
                    .Fill.Solid
                    .Fill.ForeColor.RGB = corBranca
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corCinza
                Else
                    .Fill.Visible = msoFalse
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corCinza
                End If
            End If
        End With
    Next itemBotao

SairDaMacro:
    On Error Resume Next
    Set shpBotao = Nothing
    Dashboard.Range("A1").Select
    Application.ScreenUpdating = True
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao processar a navegação de menus: " & Err.Description, vbCritical, "Erro no Dashboard"
    Resume SairDaMacro
End Sub
' ==============================================================================
' FUNÇÃO AUXILIAR: VERIFICA A PRESENÇA DE UM TEXTO EM UM VETOR (CASE-INSENSITIVE)
' ==============================================================================
Private Function fn_ExisteNoArray(ByVal itemProcurado As String, ByVal vetorAlvo As Variant) As Boolean
    
    If Not IsArray(vetorAlvo) Then Exit Function
    
    Dim elemento As Variant
    
    For Each elemento In vetorAlvo
        If StrComp(CStr(elemento), itemProcurado, vbTextCompare) = 0 Then
            fn_ExisteNoArray = True
            Exit Function
        End If
    Next elemento
    
    fn_ExisteNoArray = False
    
End Function
