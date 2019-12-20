VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ImpressaoDeConhecimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click


    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub cmdImprimir_Click()

If Me.txtInicio > Me.txtTermino Then
    MsgBox "O Valor de inicio não pode ser maior que o Valor de Terminio do conhecimento.", vbOKOnly + vbCritical
    Me.txtInicio.SetFocus
ElseIf Me.txtTermino = "" Then
    MsgBox "O Valor de Terminio do conhecimento não pode ser vasio.", vbOKOnly + vbCritical
    Me.txtTermino.SetFocus
Else
    Conhecimentos Val(Me.txtInicio), Val(Me.txtTermino)
End If

End Sub
