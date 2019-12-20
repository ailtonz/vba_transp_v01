VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Conhecimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub FreteValor_Exit(Cancel As Integer)
    Me.TotalDaPrestacao = Me.FreteValor + Me.Outros
End Sub

Private Sub Seguro_Exit(Cancel As Integer)
    Me.Outros = Me.Transp_TotalMercadoria * Me.Seguro / 100
End Sub

Private Sub Transp_ValMercadoria_01_Exit(Cancel As Integer)
    Me.Transp_TotalMercadoria = Me.Transp_ValMercadoria_01 + Transp_ValMercadoria_02 + Transp_ValMercadoria_03 + Transp_ValMercadoria_04
End Sub

Private Sub Transp_ValMercadoria_02_Exit(Cancel As Integer)
    Me.Transp_TotalMercadoria = Me.Transp_ValMercadoria_01 + Transp_ValMercadoria_02 + Transp_ValMercadoria_03 + Transp_ValMercadoria_04
End Sub

Private Sub Transp_ValMercadoria_03_Exit(Cancel As Integer)
    Me.Transp_TotalMercadoria = Me.Transp_ValMercadoria_01 + Transp_ValMercadoria_02 + Transp_ValMercadoria_03 + Transp_ValMercadoria_04
End Sub

Private Sub Transp_ValMercadoria_04_Exit(Cancel As Integer)
    Me.Transp_TotalMercadoria = Me.Transp_ValMercadoria_01 + Transp_ValMercadoria_02 + Transp_ValMercadoria_03 + Transp_ValMercadoria_04
End Sub

