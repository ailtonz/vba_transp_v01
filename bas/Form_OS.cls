VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ImprimirOS_Click()
' Este código foi criado pelo Assistente de Botão de Comando.
On Error GoTo Err_ImprimirOS_Click

    Dim strNomeDoc As String

    strNomeDoc = "OrdensDeServicos"
    ' Imprime a OS usando o Filtro da OS Atual.
    DoCmd.OpenReport strNomeDoc, acViewNormal, "Filtro da OS Atual para Impressão"

Exit_ImprimirOS_Click:
    Exit Sub

Err_ImprimirOS_Click:
    ' Se a ação foi cancelada pelo usuário, não exibe uma mensagem de erro.
    Const conErrDoCmdCancelled = 2501
    If (Err = conErrDoCmdCancelled) Then
        Resume Exit_ImprimirOS_Click
    Else
        MsgBox Err.Description
        Resume Exit_ImprimirOS_Click
    End If

End Sub

Private Sub ExibirCadastroDeClientes_Click()
On Error GoTo Err_ExibirCadastroDeClientes_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Clientes"
    DoCmd.OpenForm stDocName, , , "n_CódigoCliente = " & Me.CódigoCliente

Exit_ExibirCadastroDeClientes_Click:
    Exit Sub

Err_ExibirCadastroDeClientes_Click:
    MsgBox Err.Description
    Resume Exit_ExibirCadastroDeClientes_Click
    
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

