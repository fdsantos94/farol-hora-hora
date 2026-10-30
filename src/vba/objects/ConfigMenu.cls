VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Application.EnableEvents = True

    ' Definição e associação de valor das variáveis
    ' Cada variável é usada para controlar os menus de cada sheet analitica
    Dim NomeBotaoTabulacao As Range
    Dim NomeBotaoOperador As Range
    Dim NomeBotaoSkills As Range
'    Dim NomeBotaoSumarioReceptivo As Range
'    Dim NomeBotaoSumarioChat As Range

    Set NomeBotaoTabulacao = ConfigMenu.Range("E3")
    Set NomeBotaoOperador = ConfigMenu.Range("J3")
    Set NomeBotaoSkills = ConfigMenu.Range("O3")
'    Set NomeBotaoSumarioReceptivo = ConfigMenu.Range("T3")
'    Set NomeBotaoSumarioChat = ConfigMenu.Range("Y3")

    Select Case True

        Case Not Intersect(Target, NomeBotaoTabulacao) Is Nothing

            Call Farol.SelecionarDimensao(NomeBotaoTabulacao)
            'Call Analitico.Atualizar(NomeBotaoTabulacao)

        Case Not Intersect(Target, NomeBotaoOperador) Is Nothing

            Call Operador.SelecionarDimensao(NomeBotaoOperador)

        Case Not Intersect(Target, NomeBotaoSkills) Is Nothing

            Call Skills.SelecionarDimensao(NomeBotaoSkills)
'
'        Case Not Intersect(Target, NomeBotaoSumarioReceptivo) Is Nothing
'
'            Call Sumario.SelecionarDimensao(NomeBotaoSumarioReceptivo)
'
'        Case Not Intersect(Target, NomeBotaoSumarioChat) Is Nothing
'
'            Call SumarioChat.SelecionarDimensao(NomeBotaoSumarioChat)

    End Select

End Sub



