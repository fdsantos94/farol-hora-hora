VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Capa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    
    'Application.ScreenUpdating = False
    
    With Capa
        .Select
        .Range("A1:W42").Select
    End With
    
    'For i = 1 To 2
        Call ConfigPlanilha.ModoApresentacaoOn
    'Next
    
    Capa.Range("A1").Select
    
    'Application.ScreenUpdating = True
    
End Sub

Private Sub Worksheet_Deactivate()

   Dim i As Integer

    For i = 1 To 2
        'Call Dashboard_.ConfigPlanilha.ModoApresentacaoOff
        Call ConfigPlanilha.ModoApresentacaoSemiOff
    Next

End Sub

