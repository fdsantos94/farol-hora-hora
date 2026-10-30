VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Farol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim ControleUpdateTbd As Boolean

Sub SelecionarDimensao(ByVal NomeBotao As String)

    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    ControleUpdateTbd = True
    Dim NomeIndicador As String
    
    ' Padrão para quando rodar em delivery
    If NomeBotao = "" Then
        NomeBotao = "Bt_Produto"
    End If
        
    'Remoção dos campos da tabela dinamica
    With Farol.PivotTables("Tbd_Analitico").PivotFields
        ' Colunas a serem ocultas (tem que ser todas em que algum momento passam como linhas)
        .Item("Skill").Orientation = xlHidden
        .Item("Intervalo").Orientation = xlHidden
        .Item("Turno").Orientation = xlHidden
        .Item("DataGrafico").Orientation = xlHidden
        .Item("DiaDaSemana").Orientation = xlHidden
        .Item("Quinzena").Orientation = xlHidden
        .Item("Semana").Orientation = xlHidden
    End With
    
    'reexibe, no módulo de program é oculto para printar e enviar por e-mail
    Farol.Range("C11:C100").EntireRow.Hidden = False
''
''    Coluna hc, ela é oculta por ter memória de calculo diferente
''    Farol.Range("V:V").EntireColumn.Hidden = False 'exibe
''    Farol.Range("X:X").EntireColumn.Hidden = False 'exibe
''
''    Farol.Range("W:W").EntireColumn.Hidden = True 'oculta
''    Farol.Range("Y:Y").EntireColumn.Hidden = True 'oculta
            
    ' Identifica o botão selecionado e aplica as configurações pertinentes a cada um
    Select Case NomeBotao
    
        Case "Bt_Data"

            ' Coluna (s) a ser exibida
            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("DataGrafico").Orientation = xlRowField
                .Position = 1
                .ShowDetail = False
            End With
            
            Farol.Range("C5").value = "Data"
    
        Case "Bt_DiaSemana"

            ' Coluna (s) a ser exibida
            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("DiaDaSemana").Orientation = xlRowField
                .Position = 1
                .ShowDetail = False
            End With
            
            Farol.Range("C5").value = "Dia Da Semana"
            
        Case "Bt_Quinzena"

            ' Coluna (s) a ser exibida
            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("Quinzena").Orientation = xlRowField
                .Position = 1
                
                .Item("DataGrafico").Orientation = xlRowField
                .Position = 2
            End With
            
            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Farol.PivotTables("Tbd_Analitico").PivotFields.Item("Quinzena").ShowDetail = False
            Farol.PivotTables("Tbd_Analitico").PivotFields.Item("DataGrafico").ShowDetail = False
            
            Farol.Range("C5").value = "Quinzena >> Data"
            
        Case "Bt_Semana"

            ' Coluna (s) a ser exibida
            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("Semana").Orientation = xlRowField
                .Position = 1
                
                .Item("DiaDaSemana").Orientation = xlRowField
                .Position = 2
            End With
            
            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Farol.PivotTables("Tbd_Analitico").PivotFields.Item("Semana").ShowDetail = False
            Farol.PivotTables("Tbd_Analitico").PivotFields.Item("DiaDaSemana").ShowDetail = False
            
            Farol.Range("C5").value = "Semana >> Dia da Semana"
            
        Case "Bt_Intervalo"

            ' Coluna (s) a ser exibida
            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("Intervalo").Orientation = xlRowField
                .Position = 1
                .ShowDetail = False
                
            End With
              
'            'Coluna hc, ela é oculta por ter memória de calculo diferente
'            Farol.Range("V:V").EntireColumn.Hidden = True 'oculta
'            Farol.Range("X:X").EntireColumn.Hidden = True 'oculta
'
'            Farol.Range("W:W").EntireColumn.Hidden = False 'exibe
'            Farol.Range("Y:Y").EntireColumn.Hidden = False 'exibe
            
            Farol.Range("C5").value = "Intervalo"
            
        Case "Bt_Turno"

            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("Turno").Orientation = xlRowField
                .Position = 1

                .Item("Intervalo").Orientation = xlRowField
                .Position = 2
            End With

            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Farol.PivotTables("Tbd_Analitico").PivotFields.Item("Turno").ShowDetail = False
            Farol.PivotTables("Tbd_Analitico").PivotFields.Item("Intervalo").ShowDetail = False
            
            Farol.Range("C5").value = "Turno >> Intervalo"
            
        Case "Bt_Skill"
        
            ' Coluna (s) a ser exibida
            With Farol.PivotTables("Tbd_Analitico").PivotFields
                .Item("Skill").Orientation = xlRowField
                .Position = 1
                .ShowDetail = False
            End With
       
            Farol.Range("C5").value = "Skill"
            
    End Select
    
'    If NomeIndicador = "" Then
'        NomeIndicador = Right(NomeBotao, Len(NomeBotao) - 3) 'Ignora os 3 caracteres a esquerda (Bt_)
'    End If
'
    ' Renomeia a coluna conforme a dimensão setada e autoajusta as colunas
    With Farol.Range("C5")
        '.value = NomeIndicador + " "
        .CurrentRegion.EntireColumn.AutoFit
    End With
'
'    'oculta devido ao autoajusta acima
'    Select Case NomeBotao
'        Case "Bt_Intervalo"
'            Farol.Range("O:O").EntireColumn.Hidden = True 'oculta
'        Case Else
'            Farol.Range("P:P").EntireColumn.Hidden = True 'oculta
'    End Select
    
    Range("A1").Select
    
    ControleUpdateTbd = False
        
    Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)

    ' Evita o redimensionamento indesejado de botões/segmentações acima da tabela
    ' Se deixar direto na tabela dinamica (ajustar automaticamente a largura...) por algum motivo estava redimensionando as segmentações e botões que ficam acima da tabela
    ' Dessa forma, garanto o autoajuste sem ter problemas com os botões acima da tabela
    If ControleUpdateTbd = False Then
        If Target.Name = "" Then
            Farol.Range("C9").CurrentRegion.EntireColumn.AutoFit
        End If
    End If

End Sub

Private Sub OcultarFiltro()

    Application.ScreenUpdating = False
    
    With ActiveSheet
        .Shapes.Range(Array("FundoTransparente")).Visible = False
        .Shapes.Range(Array("ConjuntoFiltrosSumario")).Visible = False
        .Range("A1").Select
    End With

    
    Application.ScreenUpdating = True

End Sub

Private Sub MostrarFiltro()

    Application.ScreenUpdating = False
    
    With ActiveSheet
        .Shapes.Range(Array("FundoTransparente")).Visible = True
        .Shapes.Range(Array("ConjuntoFiltrosSumario")).Visible = True
        .Range("A1").Select
    End With
       
    Application.ScreenUpdating = True
    
End Sub

Private Sub BlocoIndicadores()
    
    Application.ScreenUpdating = False
    
    Dim Botao As String
    Dim Dimensao As String
    
    Botao = Farol.Shapes(Application.Caller).Name
    Dimensao = Farol.Range("C5").value
    
    ' Configuração os botões off
    Farol.Shapes.Range(Array("TxtTodos", "TxtVolumetria", "TxtTempos", "TxtPercTempos")).Select
    
    With Selection.ShapeRange
        .Item(Botao).Fill.Transparency = 0#
        '.Fill.ForeColor.RGB = RGB(0, 0, 0) 'cor de fundo
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(127, 127, 127) 'cor da fonte
        .TextFrame2.TextRange.Font.Bold = msoFalse 'negrito
        '.Line.ForeColor.RGB = RGB(242, 242, 242) 'cor da linha
        .Fill.Visible = msoFalse ' sem preenchimento
        '.Shadow.Type = msoShadow21 'msoShadow25 'estilo da sombra
        '.Shadow.Visible = msoFalse
    End With
    
    ' Configuração do botão on
    Farol.Shapes.Item(Botao).Select
    
    With Selection.ShapeRange
        '.Fill.ForeColor.RGB = RGB(45, 64, 89) 'cor de fundo
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(214, 4, 61) 'cor da fonte
        .TextFrame2.TextRange.Font.Bold = msoTrue 'negrito
        '.Line.ForeColor.RGB = RGB(255, 255, 255) 'cor da linha
        .Fill.Visible = msoFalse ' sem preenchimento
        '.Shadow.Type = msoShadow21 'estilo da sombra
        '.Shadow.Visible = msoFalse
        '.Fill.Transparency = 0#
    End With
    
    Select Case Botao
    
        Case "TxtTodos"
        
            Farol.Range("D:BA").EntireColumn.Hidden = False 'exibe
'
'            'controle para exibir o calculo de HC correto de acordo com a dimensao
'            If Dimensao = "Intervalo" Then
'                Farol.Range("V:V").EntireColumn.Hidden = True 'oculta
'                Farol.Range("X:X").EntireColumn.Hidden = True 'oculta
'
'                Farol.Range("W:W").EntireColumn.Hidden = False 'exibe
'                Farol.Range("Y:Y").EntireColumn.Hidden = False 'exibe
'            Else
'                'Coluna hc, ela é oculta por ter memória de calculo diferente
'                Farol.Range("V:V").EntireColumn.Hidden = False 'exibe
'                Farol.Range("X:X").EntireColumn.Hidden = False 'exibe
'
'                Farol.Range("W:W").EntireColumn.Hidden = True 'oculta
'                Farol.Range("Y:Y").EntireColumn.Hidden = True 'oculta
'            End If
        
        Case "TxtVolumetria"
        
            Farol.Range("D:U").EntireColumn.Hidden = False 'exibe
            Farol.Range("V:BA").EntireColumn.Hidden = True  'oculta
            
            ' oculta informações do forecast
            'Farol.Range("E:E, F:F, G:G, M:M, O:O, R:R").EntireColumn.Hidden = True
                
        Case "TxtTempos"
        
            Farol.Range("V:AL").EntireColumn.Hidden = False 'exibe
            Farol.Range("D:U").EntireColumn.Hidden = True  'oculta
            Farol.Range("AM:BA").EntireColumn.Hidden = True  'oculta

'            'controle para exibir o calculo de HC correto de acordo com a dimensao
'            If Dimensao = "Intervalo" Then
'                Farol.Range("V:V").EntireColumn.Hidden = True 'oculta
'                Farol.Range("X:X").EntireColumn.Hidden = True 'oculta
'
'                Farol.Range("W:W").EntireColumn.Hidden = False 'exibe
'                Farol.Range("Y:Y").EntireColumn.Hidden = False 'exibe
'            Else
'                'Coluna hc, ela é oculta por ter memória de calculo diferente
'                Farol.Range("V:V").EntireColumn.Hidden = False 'exibe
'                Farol.Range("X:X").EntireColumn.Hidden = False 'exibe
'
'                Farol.Range("W:W").EntireColumn.Hidden = True 'oculta
'                Farol.Range("Y:Y").EntireColumn.Hidden = True 'oculta
'            End If
            
        Case "TxtPercTempos"
        
            Farol.Range("AM:BA").EntireColumn.Hidden = False 'exibe
            Farol.Range("D:AL").EntireColumn.Hidden = True  'oculta
            'Farol.Range("D:N").EntireColumn.Hidden = True  'oculta

    End Select
    
    Farol.Range("A1").Select
    
    Application.ScreenUpdating = True

End Sub
