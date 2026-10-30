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
Private Sub FormatarMenus(Optional menu As Variant)
    
    On Error GoTo TratarErro
    Application.ScreenUpdating = False
    
    ' --------------------------------------------------------------------------
    ' 1. DEFINIÇÃO DOS GRUPOS DE BOTÕES (Mapeie o nome das suas Shapes aqui)
    ' --------------------------------------------------------------------------
    Dim MenuPrincipal As Variant
    Dim SubMenuEsquerdo As Variant
    Dim SubMenuDireito As Variant
    Dim SubMenuTma As Variant
    
    ' MENU I: Navegação Superior Principal
    MenuPrincipal = Array("BtMenuOnePage", "BtMenuVolumetria", "BtMenuTma", "BtMenuTempos")
    
    ' MENU II: Submenu Lateral Esquerdo
    SubMenuEsquerdo = Array("BtVolComparativo", "BtVolUnico", "BtVolNs")    ', "SubMenuTma", "BtTmaIntradia", "BtTmaAderencia")
    
    ' MENU III: Submenu Lateral Direito
    SubMenuDireito = Array("BtVolAderencia", "BtVolDetalhamento")
    
    SubMenuTma = Array("BtTmaIntradia", "BtTmaAderencia")
    
    ' --------------------------------------------------------------------------
    ' 2. PALETA DE CORES GLOBAL
    ' --------------------------------------------------------------------------
    Dim CorLaranja As Long, CorBranca As Long, CorCinza As Long
    
    CorLaranja = RGB(255, 107, 0)   ' Laranja Destaque
    CorBranca = RGB(255, 255, 255)  ' Branco Puro
    CorCinza = RGB(128, 128, 128)   ' Cinza Neutro
    
    ' --------------------------------------------------------------------------
    ' 3. IDENTIFICAÇÃO DO BOTÃO CLICADO E DO GRUPO PERTENCENTE
    ' --------------------------------------------------------------------------
    Dim NomeBotaoClicado As String
    Dim GrupoAtual As Variant
    Dim EstiloPrincipal As Boolean
    
    If menu = "" Then
        NomeBotaoClicado = Application.Caller
    Else
        NomeBotaoClicado = menu
    End If
    
    ' Verifica a qual grupo o botão clicado pertence
    If IsInArray(NomeBotaoClicado, MenuPrincipal) Then
        GrupoAtual = MenuPrincipal
        EstiloPrincipal = True  ' Aplica regras do Menu I (Fundo Laranja)
        
    ElseIf IsInArray(NomeBotaoClicado, SubMenuEsquerdo) Then
        GrupoAtual = SubMenuEsquerdo
        EstiloPrincipal = False ' Aplica regras dos Submenus (Fundo Branco + Transparente)
        
    ElseIf IsInArray(NomeBotaoClicado, SubMenuDireito) Then
        GrupoAtual = SubMenuDireito
        EstiloPrincipal = False ' Aplica regras dos Submenus (Fundo Branco + Transparente)
        
    ElseIf IsInArray(NomeBotaoClicado, SubMenuTma) Then
        GrupoAtual = SubMenuTma
        EstiloPrincipal = False ' Aplica regras dos Submenus (Fundo Branco + Transparente)
            
            
    Else
        ' Caso a forma clicada não esteja mapeada nos vetores
        GoTo SairDaMacro
    End If
    
    ' --------------------------------------------------------------------------
    ' 4. LOOP DE ESTILIZAÇÃO DO GRUPO AFETADO
    ' --------------------------------------------------------------------------
    Dim ItemBotao As Variant
    Dim Shp As Shape
    
    For Each ItemBotao In GrupoAtual
        Set Shp = Dashboard.Shapes(CStr(ItemBotao))
        
        With Shp
            .Line.Visible = msoFalse ' Remove contornos para manter o design clean
            
            If .Name = NomeBotaoClicado Then
                ' ==============================================================
                ' ESTADO ATIVO (ON)
                ' ==============================================================
                .Fill.Visible = msoTrue
                .Fill.Solid
                .Fill.Transparency = 0#
                .Shadow.Visible = msoTrue
                .Shadow.Type = msoShadow21 ' Sombra de elevação suave
                
                If EstiloPrincipal Then
                    ' MENU I ATIVO: Fundo Laranja + Texto Branco
                    .Fill.ForeColor.RGB = CorLaranja
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CorBranca
                Else
                    ' MENU II/III ATIVO: Fundo Branco + Texto Laranja
                    .Fill.ForeColor.RGB = CorBranca
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CorLaranja
                End If
                
                .TextFrame2.TextRange.Font.Bold = msoTrue
                
            Else
                ' ==============================================================
                ' ESTADO INATIVO (OFF)
                ' ==============================================================
                .Shadow.Visible = msoFalse
                
                If EstiloPrincipal Then
                    ' MENU I INATIVO: Fundo Branco + Texto Cinza
                    .Fill.Visible = msoTrue
                    .Fill.Solid
                    .Fill.ForeColor.RGB = CorBranca
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CorCinza
                Else
                    ' MENU II/III INATIVO: Fundo Transparente + Texto Cinza
                    .Fill.Visible = msoFalse
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CorCinza
                End If
                
                .TextFrame2.TextRange.Font.Bold = msoTrue
                
            End If
        End With
    Next ItemBotao

SairDaMacro:
    Set Shp = Nothing
    'ActiveSheet.Shapes(NomeBotaoClicado).Select
    ActiveSheet.Range("A1").Select
    Application.ScreenUpdating = True
    Exit Sub

TratarErro:
    MsgBox "Erro ao processar a navegação de menus: " & Err.Description, vbCritical, "Erro no Dashboard"
    Resume SairDaMacro
End Sub

' ==============================================================================
' FUNÇÃO AUXILIAR: VERIFICA SE UMA SHAPE PERTENCE AO VETOR DO MENU
' ==============================================================================
Private Function IsInArray(valToBeFound As String, arr As Variant) As Boolean
    
    Dim element As Variant
    
    For Each element In arr
        If LCase(CStr(element)) = LCase(valToBeFound) Then
            IsInArray = True
            Exit Function
        End If
    Next element
    
    IsInArray = False
    
End Function


