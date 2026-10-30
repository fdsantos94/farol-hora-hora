VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Deactivate()

   Dim i As Integer
   
    For i = 1 To 2
        'Call Dashboard_.ConfigPlanilha.ModoApresentacaoOff
        Call ConfigPlanilha.ModoApresentacaoSemiOff
    Next

End Sub
Private Sub OcultarFiltro()

    Application.ScreenUpdating = False
    
    With ActiveSheet
        .Shapes.Range(Array("FundoTransparente")).Visible = False
        .Shapes.Range(Array("ConjuntoFiltrosDashboard")).Visible = False
        .Shapes.Range(Array("TxtSelecionarFiltro")).Visible = True
        .Range("A1").Select
    End With

    
    Application.ScreenUpdating = True

End Sub

Private Sub MostrarFiltro()

    Application.ScreenUpdating = False
    
    With ActiveSheet
        .Shapes.Range(Array("FundoTransparente")).Visible = True
        .Shapes.Range(Array("ConjuntoFiltrosDashboard")).Visible = True
        .Shapes.Range(Array("TxtSelecionarFiltro")).Visible = False
        .Range("A1").Select
    End With
       
    Application.ScreenUpdating = True
    
End Sub



