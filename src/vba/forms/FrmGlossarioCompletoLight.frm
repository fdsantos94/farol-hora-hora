VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmGlossarioCompletoLight 
   Caption         =   "MIS | Management Information System"
   ClientHeight    =   9432.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9636.001
   OleObjectBlob   =   "FrmGlossarioCompletoLight.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmGlossarioCompletoLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListBox1_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    
    Dim valor As String
    Dim ItemSelecionado As Integer
    
    ' Texto da coluna da linha selecionada
    valor = Me.ListBox1.value
    
    ' Nº indice da coluna da linha selecionada
    ItemSelecionado = Me.ListBox1.ListIndex
    
    ' Definição de valores para o segundo FrmGlossarioIndicador
    With FrmGlossarioDetalheLight
        .TxtIndicador.value = Me.ListBox1.List(ItemSelecionado, 0)
        .TxtConceito.value = Me.ListBox1.List(ItemSelecionado, 2)
        .TxtCalculo.value = Me.ListBox1.List(ItemSelecionado, 1)
    End With
    
    FrmGlossarioDetalheLight.Show

End Sub

Private Sub TxtPesquisa_Change()
    
    'condicional para verificar se o usuario está na lista de problemas
    'caso esteja, retornará true e aplicará as configurações
    'caso contrário, manterá a configuração original
    If FrmGlossarioConfiguracoes.fn_ican = True Then
        ' Validação para efeito do texto exemplo
        If TxtPesquisa.Text = Empty Then
        
            LbPesquisa.Caption = "     Procure um indicador aqui..."
        
        Else
        
            LbPesquisa.Caption = Empty
        
        End If
        
    Else
    
        ' Validação para efeito do texto exemplo
        If TxtPesquisa.Text = Empty Then
        
            LbPesquisa.Caption = "       Procure um indicador aqui..."
        
        Else
        
            LbPesquisa.Caption = Empty
        
        End If
        
    End If
    
    ' Local onde é direcionado o texto pesquisado
    Glossario.Range("F2").value = TxtPesquisa.value
    
    Call FrmGlossarioConfiguracoes.Filtro_Avancado_Indicador
    
    'TxtMoradorRegistros.Value = Glossario.Range("W:W").CurrentRegion.Rows.Count - 1

End Sub

Private Sub UserForm_Initialize()
    
    Dim linha           As Integer
    Dim Base            As Range
    Dim NomePlanilha    As String
    Dim Coluna1         As Integer
    Dim Coluna2         As Integer
    Dim Coluna3         As Integer
    Dim Usuario         As Boolean
    
    'condicional para verificar se o usuario está na lista de problemas
    'caso esteja, retornará true e aplicará as configurações
    'caso contrário, manterá a configuração original
    If FrmGlossarioConfiguracoes.fn_ican = True Then
    
        ' Define dimensões do form principal
        With Me
            .Top = 0
            .Height = 408
            .Width = 397.8
        End With
        
        ' Definição dos botões do form
        With Me
            .TxtPesquisa.Height = 25.5
            .TxtPesquisa.Left = 6
            .TxtPesquisa.Top = 36
            .TxtPesquisa.Width = 168
            .TxtPesquisa.BackStyle = fmBackStyleTransparent
            
            .LbPesquisa.Height = 29.25
            .LbPesquisa.Left = 0
            .LbPesquisa.Top = 38 '36
            .LbPesquisa.Width = 186
            .LbPesquisa.BackStyle = fmBackStyleTransparent
            .LbPesquisa.Caption = "     Procure um indicador aqui..."
            .LbPesquisa.Font.Size = 11
            
            .ListBox1.Height = 316.3
            .ListBox1.Left = 12
            .ListBox1.Top = 66
            .ListBox1.Width = 363
        End With
    
    Else
    
        ' Define configurações das caixas de busca
        With Me
            .LbPesquisa.BackStyle = fmBackStyleTransparent
            .LbPesquisa.Caption = "       Procure um indicador aqui..."
            .TxtPesquisa.BackStyle = fmBackStyleTransparent
        End With
    
    End If
        
    ' Pega o nome da planilha que contem os dados do glossário
    NomePlanilha = "'" & Glossario.Name & "'!"
    
    ' Conta o total de linhas
    linha = Glossario.Range("A1").CurrentRegion.Rows.Count
    
    ' Define o intervalo das bases
    Set Base = Glossario.Range(Glossario.Cells(2, 1), Glossario.Cells(linha, 3))
    
    ' Auto ajusta as colunas
    Glossario.Columns("A:C").AutoFit
    
    ' Pega a largura de cada coluna + um acrescimo para melhor visibilidade no form
    Coluna1 = Glossario.Columns("A:A").Width + 10
    Coluna2 = Glossario.Columns("B:B").Width + 10
    Coluna3 = Glossario.Columns("C:C").Width + 10
    
    ' Configura a list box do FrmGlossario
    With Me
        .ListBox1.ColumnCount = 2
        .ListBox1.RowSource = NomePlanilha & Base.Address 'Pega o nome da planilha e o endereço da base preenchida
        .ListBox1.Font.Size = 10
        .ListBox1.Font.Name = "Arial Nova Light"
        .ListBox1.ColumnWidths = Coluna1 & "pt;" & Coluna2 & "pt;" & Coluna3 & "pt" '"30 pt;150 pt;150 pt;
        .ListBox1.ColumnHeads = True
        .ListBox1.ForeColor = RGB(24, 23, 23)
    End With

End Sub
