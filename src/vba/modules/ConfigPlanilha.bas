Attribute VB_Name = "ConfigPlanilha"
Option Explicit

Sub ModoApresentacaoOn()
    
    With Application
        '.ScreenUpdating = False                               ' desativa a atualização de tela (evita piscadas)
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"   ' oculta a Faixa de Opções (Ribbon)
        .DisplayFullScreen = True                              ' ativa o modo tela cheia do Excel
        .DisplayFormulaBar = False                             ' oculta a barra de fórmulas
        .ActiveWindow.Zoom = True                              ' ajusta o zoom automaticamente para caber na tela
        '.OnKey "{Escape}", ""                                 ' desativa a tecla ESC (impede o usuário de sair de certos modos)
        .CommandBars("PivotTable Field List").Visible = False  ' fecha os campos da tabela dinamica
    End With
    
    With ActiveWindow
        .DisplayHeadings = False                               ' oculta cabeçalhos de linhas e colunas
        .DisplayGridlines = False                              ' oculta linhas de grade
        .DisplayHorizontalScrollBar = True                     ' exibe a barra de rolagem horizontal
        .DisplayVerticalScrollBar = True                       ' exibe a barra de rolagem vertical
        .DisplayWorkbookTabs = False                           ' oculta as abas das planilhas
        '.Zoom = True                                          ' zoom automático da janela (mesmo efeito do acima)
    End With
    
    'Application.ScreenUpdating = True                         ' ativa a atualização de tela (haverá piscadas)

End Sub

Sub ModoApresentacaoOff()
Attribute ModoApresentacaoOff.VB_ProcData.VB_Invoke_Func = "E\n14"
    
    With Application
        .ScreenUpdating = False                                 ' desativa a atualização de tela (evita piscadas)
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"     ' mostra a faixa de opções
        .DisplayFullScreen = False                              ' sai da tela cheia
        .DisplayFormulaBar = True                               ' mostra a barra de fórmulas
        '.ActiveWindow.Zoom = 100                               ' volta para zoom padrão (ajuste se quiser)
        .OnKey "{Escape}"                                       ' restaura tecla ESC se você quiser atribuir algo depois
        .CommandBars("PivotTable Field List").Visible = False   ' fecha os campos da tabela dinamica
    End With
    
    With ActiveWindow
        '.DisplayHeadings = True                                ' mostra cabeçalhos das linhas e colunas
        '.DisplayGridlines = True                               ' oculta linhas de grade
        .DisplayHorizontalScrollBar = True                      ' deixa a barra horizontal visível
        .DisplayVerticalScrollBar = True                        ' deixa a barra vertical visível
        .DisplayWorkbookTabs = True                             ' mostra as abas das planilhas
        '.Zoom = 100                                            ' zoom padrão opcional
    End With

    Application.ScreenUpdating = True                           ' ativa a atualização de tela (haverá piscadas)
    
End Sub

Sub ModoApresentacaoSemiOff()
    
    With Application
        .ScreenUpdating = False                                 ' desativa a atualização de tela (evita piscadas)
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"    ' mostra a faixa de opções
        .DisplayFullScreen = False                              ' sai da tela cheia
        .DisplayFormulaBar = True                               ' mostra a barra de fórmulas
        '.ActiveWindow.Zoom = 100                               ' volta para zoom padrão (ajuste se quiser)
        .OnKey "{Escape}"                                       ' restaura tecla ESC se você quiser atribuir algo depois
        .CommandBars("PivotTable Field List").Visible = False   ' fecha os campos da tabela dinamica
    End With
    
    With ActiveWindow
        '.DisplayHeadings = True                                ' mostra cabeçalhos das linhas e colunas
        '.DisplayGridlines = True                               ' oculta linhas de grade
        .DisplayHorizontalScrollBar = True                      ' deixa a barra horizontal visível
        .DisplayVerticalScrollBar = True                        ' deixa a barra vertical visível
        '.DisplayWorkbookTabs = True                            ' mostra as abas das planilhas
        '.Zoom = 100                                            ' zoom padrão opcional
    End With

    Application.ScreenUpdating = True                           ' ativa a atualização de tela (haverá piscadas)
    
End Sub
