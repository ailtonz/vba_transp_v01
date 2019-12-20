VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AbrirDadosCobran�a_Click()
On Error GoTo Err_AbrirDadosCobran�a_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Clientes"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_AbrirDadosCobran�a_Click:
    Exit Sub

Err_AbrirDadosCobran�a_Click:
    MsgBox Err.Description
    Resume Exit_AbrirDadosCobran�a_Click
    
End Sub
Private Sub DadosDeCobran�aDeClientes_Click()
On Error GoTo Err_DadosDeCobran�aDeClientes_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Dados de Cobran�a de Clientes"
    
    stLinkCriteria = "[n_C�digoCliente]=" & Me![n_C�digoCliente]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_DadosDeCobran�aDeClientes_Click:
    Exit Sub

Err_DadosDeCobran�aDeClientes_Click:
    MsgBox Err.Description
    Resume Exit_DadosDeCobran�aDeClientes_Click
    
End Sub
Sub Combina��o191_AfterUpdate()
    ' Encontrar o registro que coincide com o controle.
    Me.RecordsetClone.FindFirst "[c_Cliente] = '" & Me![Combina��o191] & "'"
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub ImprimirFichaDoCliente_Click()
' Este c�digo foi criado pelo Assistente de Bot�o de Comando.
On Error GoTo Err_ImprimirFichaDoCliente_Click

    Dim strNomeDoc As String

    strNomeDoc = "FichaDeCliente"
    ' Imprime a Ficha usando o Filtro do Cliente Atual.
    DoCmd.OpenReport strNomeDoc, acViewNormal, "Filtro do Cliente Atual para Impress�o"

Exit_ImprimirFichaDoCliente_Click:
    Exit Sub

Err_ImprimirFichaDoCliente_Click:
    ' Se a a��o foi cancelada pelo usu�rio, n�o exibe uma mensagem de erro.
    Const conErrDoCmdCancelled = 2501
    If (Err = conErrDoCmdCancelled) Then
        Resume Exit_ImprimirFichaDoCliente_Click
    Else
        MsgBox Err.Description
        Resume Exit_ImprimirFichaDoCliente_Click
    End If

End Sub

Private Sub VisualizarListaDeClientes_Click()
On Error GoTo Err_VisualizarListaDeClientes_Click

    Dim stDocName As String

    stDocName = "ListaAlfabaticaDeClientes"
    DoCmd.OpenReport stDocName, acPreview

Exit_VisualizarListaDeClientes_Click:
    Exit Sub

Err_VisualizarListaDeClientes_Click:
    MsgBox Err.Description
    Resume Exit_VisualizarListaDeClientes_Click
    
End Sub
Private Sub ImprimirListaDeClientes_Click()
On Error GoTo Err_ImprimirListaDeClientes_Click

    Dim stDocName As String

    stDocName = "ListaAlfabaticaDeClientes"
    DoCmd.OpenReport stDocName, acNormal

Exit_ImprimirListaDeClientes_Click:
    Exit Sub

Err_ImprimirListaDeClientes_Click:
    MsgBox Err.Description
    Resume Exit_ImprimirListaDeClientes_Click
    
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