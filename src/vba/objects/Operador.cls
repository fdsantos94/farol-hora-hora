VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Operador"
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
    With Operador.PivotTables("Tbd_Operador").PivotFields
        ' Colunas a serem ocultas (tem que ser todas em que algum momento passam como linhas)
        .Item("Nome").Orientation = xlHidden
        .Item("Supervisor").Orientation = xlHidden
        .Item("Regional").Orientation = xlHidden
        .Item("DataGrafico").Orientation = xlHidden
        .Item("Semana").Orientation = xlHidden
        .Item("DiaDaSemana").Orientation = xlHidden
        .Item("Quinzena").Orientation = xlHidden
        .Item("Intervalo").Orientation = xlHidden
        .Item("Turno").Orientation = xlHidden
    End With
    
    ' Identifica o botão selecionado e aplica as configurações pertinentes a cada um
    Select Case NomeBotao
    
        Case "Bt_Operador"

            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("Nome").Orientation = xlRowField
                .Position = 1
                .ShowDetail = False
            End With
            
            Operador.Range("C5").value = "Nome"
            
        Case "Bt_Regional"

            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("Regional").Orientation = xlRowField
                .Position = 1
                
                .Item("Nome").Orientation = xlRowField
                .Position = 2
            End With

            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Regional").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = False
            
            Operador.Range("C5").value = "Regional"
            
        Case "Bt_Equipe"
        
            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                '.Item("Coordenador").Orientation = xlRowField
                '.Position = 1

                .Item("Supervisor").Orientation = xlRowField
                .Position = 1
            
                .Item("Nome").Orientation = xlRowField
                .Position = 2
                
                End With

            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            'Operador.PivotTables("Tbd_Operador").PivotFields.Item("Coordenador").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Supervisor").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = False
            
            Operador.Range("C5").value = "Coordenador >> Supervisor >> Operador"
            
        Case "Bt_Data"

            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("DataGrafico").Orientation = xlRowField
                .Position = 1
                
                .Item("Nome").Orientation = xlRowField
                .Position = 2
                
            End With
            
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("DataGrafico").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = False
            
            Operador.Range("C5").value = "Data >> Operador"
    
        Case "Bt_DiaSemana"

            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("DiaDaSemana").Orientation = xlRowField
                .Position = 1

                .Item("Nome").Orientation = xlRowField
                .Position = 2
                
            End With
            
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("DiaDaSemana").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = False
            
            Operador.Range("C5").value = "Dia Da Semana"
            
        Case "Bt_Quinzena"

            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("Quinzena").Orientation = xlRowField
                .Position = 1
                
                .Item("Nome").Orientation = xlRowField
                .Position = 2
            End With
            
            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Quinzena").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = False
            
            Operador.Range("C5").value = "Quinzena >> Data"
            
        Case "Bt_Semana"

            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("Semana").Orientation = xlRowField
                .Position = 1
                
                .Item("Nome").Orientation = xlRowField
                .Position = 2
            End With
            
            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Semana").ShowDetail = False
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = False
            
            Operador.Range("C5").value = "Semana >> Dia da Semana"
            
        Case "Bt_Intervalo"

            ' Coluna (s) a ser exibida
            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("Intervalo").Orientation = xlRowField
                .Position = 1
                
                .Item("Nome").Orientation = xlRowField
                .Position = 2
                
            End With
              
'            'Coluna hc, ela é oculta por ter memória de calculo diferente
'            Operador.Range("V:V").EntireColumn.Hidden = True 'oculta
'            Operador.Range("X:X").EntireColumn.Hidden = True 'oculta
'
'            Operador.Range("W:W").EntireColumn.Hidden = False 'exibe
'            Operador.Range("Y:Y").EntireColumn.Hidden = False 'exibe
            
            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Intervalo").ShowDetail = True
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = True
            
            Operador.Range("C5").value = "Operador >> Intervalo"
            
        Case "Bt_Turno"

            With Operador.PivotTables("Tbd_Operador").PivotFields
                .Item("Nome").Orientation = xlRowField
                .Position = 1
                
                .Item("Turno").Orientation = xlRowField
                .Position = 2

            End With

            ' Necessário fazer fora para recolher os campos
            ' Em tempo de execução ele não recolhe
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Nome").ShowDetail = True
            Operador.PivotTables("Tbd_Operador").PivotFields.Item("Turno").ShowDetail = True
            
            Operador.Range("C5").value = "Operador >> Turno"
            
    End Select
    
'    If NomeIndicador = "" Then
'        NomeIndicador = Right(NomeBotao, Len(NomeBotao) - 3) 'Ignora os 3 caracteres a esquerda (Bt_)
'    End If
'
    ' Renomeia a coluna conforme a dimensão setada e autoajusta as colunas
    With Operador.Range("C5")
        '.value = NomeIndicador + " "
        .CurrentRegion.EntireColumn.AutoFit
    End With
    
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
            Operador.Range("C9").CurrentRegion.EntireColumn.AutoFit
        End If
    End If

End Sub

Private Sub OcultarFiltro()

    Application.ScreenUpdating = False
    
    With ActiveSheet
        .Shapes.Range(Array("FundoTransparente")).Visible = False
        .Shapes.Range(Array("ConjuntoFiltrosOperador")).Visible = False
        .Range("A1").Select
    End With

    
    Application.ScreenUpdating = True

End Sub

Private Sub MostrarFiltro()

    Application.ScreenUpdating = False
    
    With ActiveSheet
        .Shapes.Range(Array("FundoTransparente")).Visible = True
        .Shapes.Range(Array("ConjuntoFiltrosOperador")).Visible = True
        .Range("A1").Select
    End With
       
    Application.ScreenUpdating = True
    
End Sub

Private Sub BlocoIndicadores()
    
    Application.ScreenUpdating = False
    
    Dim Botao As String

    Botao = Operador.Shapes(Application.Caller).Name
    
    ' Configuração os botões off
    Operador.Shapes.Range(Array("TxtTodos", "TxtTempos", "TxtPercTempos")).Select
    
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
    Operador.Shapes.Item(Botao).Select
    
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
        
            Operador.Range("D:AD").EntireColumn.Hidden = False 'exibe
        
                
        Case "TxtTempos"
        
            Operador.Range("F:R").EntireColumn.Hidden = False 'exibe
            Operador.Range("S:AD").EntireColumn.Hidden = True  'oculta
            Operador.Range("D:E").EntireColumn.Hidden = True  'oculta

        Case "TxtPercTempos"
        
            Operador.Range("S:AC").EntireColumn.Hidden = False 'exibe
            Operador.Range("D:R").EntireColumn.Hidden = True  'oculta

    End Select
    
    Operador.Range("A1").Select
    
    Application.ScreenUpdating = True

End Sub
