Attribute VB_Name = "ADM"
Option Compare Database
Option Explicit

Public strTabela As String
Public fLote        ' Arquivo texto
Public Lote         ' Variável do arquivo de lote
Public Caminho      ' Caminho da aplicação

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Sub Conhecimentos(Inicio As Integer, Termino As Integer)
On Error GoTo Conhecimentos_err

'Banco de dados corrente
Dim dbBase As Database

'Conhecimentos
Dim rConhecimentos As DAO.Recordset
Dim sConhecimentos As String

'Remetente
Dim rRemetente As DAO.Recordset
Dim sRemetente As String

'Destinatário
Dim rDestinatario As DAO.Recordset
Dim sDestinatario As String

'Consignatário
Dim rConsignatario As DAO.Recordset
Dim sConsignatario As String

'Redespacho
Dim rRedespacho As DAO.Recordset
Dim sRedespacho As String

    '----------------
    'Seleção de dados
    '----------------
    
    sConhecimentos = "Select * from Conhecimentos WHERE Conhecimento Between " & Inicio & " And " & Termino & " Order by Conhecimento"
    sRemetente = "Select * from Clientes"
    sDestinatario = "Select * from Clientes"
    sConsignatario = "Select * from Clientes"
    sRedespacho = "Select * from Clientes"

    '---------------------------
    'Setar tabelas selecionadas
    '---------------------------
    
    Set dbBase = CurrentDb
    Set rConhecimentos = dbBase.OpenRecordset(sConhecimentos)
    Set rRemetente = dbBase.OpenRecordset(sRemetente)
    Set rDestinatario = dbBase.OpenRecordset(sDestinatario)
    Set rConsignatario = dbBase.OpenRecordset(sConsignatario)
    Set rRedespacho = dbBase.OpenRecordset(sRedespacho)

    '---------------------
    'Arquivo de lote
    '---------------------

    Caminho = Application.CurrentProject.Path
    
    If Not Dir(Caminho & "\Lote.txt") = "" Then Kill Caminho & "\Lote.txt"
    
    Set fLote = CreateObject("Scripting.FileSystemObject")
    Set Lote = fLote.CreateTextFile(Caminho & "\Lote.txt", True)
    
    '----------------------------
    'Configuração de formulário
    '----------------------------
    
     Prt -1, 250, 152  'Tamanho de formulario
'     Prt -4, 1, " "   'NORMAL
'     Prt -4, 2, " "   'NEGRITO
'     Prt -4, 3, " "   'ITALICO
    

'---------------------
'Dados do conhecimento
'---------------------

While Not rConhecimentos.EOF
    
    '------------------
    'Cabeçalho
    '------------------

    Prt 17, 130, rConhecimentos.Fields("ETC_RTB")
    Prt 27, 124, rConhecimentos.Fields("CFOP")
    Prt 31, 128, Format(rConhecimentos.Fields("Emissao"), "dd")
    Prt 31, 154, UCase(Format(rConhecimentos.Fields("Emissao"), "mmmm"))
    Prt 31, 179, Format(rConhecimentos.Fields("Emissao"), "yyyy")
   
    
    '---------------------------
    'Localiza dados do Remetente
    '---------------------------

    rRemetente.MoveLast
    rRemetente.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Remetente") & "'"

    If rConhecimentos.Fields("Remetente") <> "" Then
       Prt 39, 17, UCase(rConhecimentos.Fields("Remetente"))
       Prt 43, 17, UCase(rRemetente.Fields("c_Endereço"))
       Prt 48, 17, UCase(rRemetente.Fields("c_Cidade"))
       Prt 48, 59, UCase(rRemetente.Fields("c_Estado"))
       Prt 48, 79, rRemetente.Fields("c_CEP")
       Prt 52, 17, rRemetente.Fields("c_CNPJ")
       Prt 52, 79, rRemetente.Fields("c_InscriçãoEstadual")
    End If

    '------------------------------
    'Localiza dados do Destinatário
    '------------------------------

    rDestinatario.MoveLast
    rDestinatario.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Destinatario") & "'"

    If rConhecimentos.Fields("Destinatario") <> "" Then
       Prt 39, 129, UCase(rConhecimentos.Fields("Destinatario"))
       Prt 43, 129, UCase(rDestinatario.Fields("c_Endereço"))
       Prt 48, 129, UCase(rDestinatario.Fields("c_Cidade"))
       Prt 48, 211, UCase(rDestinatario.Fields("c_Estado"))
       Prt 52, 129, rDestinatario.Fields("c_CNPJ")
       Prt 52, 184, rDestinatario.Fields("c_InscriçãoEstadual")
    End If

    '-------------------------------
    'Localiza dados do Consignatário
    '-------------------------------

    rConsignatario.MoveLast
    rConsignatario.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Consignatario") & "'"

    If rConhecimentos.Fields("Consignatario") <> "" Then
       Prt 60, 22, UCase(rConhecimentos.Fields("Consignatario"))
       Prt 64, 17, UCase(rConsignatario.Fields("c_Endereço"))
       Prt 68, 17, UCase(rConsignatario.Fields("c_Cidade"))
       Prt 68, 99, UCase(rConsignatario.Fields("c_Estado"))
    End If

    If rConhecimentos.Fields("FreteConsignatario") = "A Pagar" Then
       Prt 73, 28, "X"
    ElseIf rConhecimentos.Fields("FreteConsignatario") = "Pago" Then
       Prt 73, 82, "X"
    End If

    Prt 78, 17, rConhecimentos.Fields("Distancia")

    '----------------------------
    'Localiza dados do Redespacho
    '----------------------------

    rRedespacho.MoveLast
    rRedespacho.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Redespacho") & "'"

    If rConhecimentos.Fields("FreteRedespacho") = "Pago" Then
       Prt 60, 148, "X"
    ElseIf rConhecimentos.Fields("FreteRedespacho") = "A Pagar" Then
       Prt 60, 179, "X"
    End If

    If rConhecimentos.Fields("Redespacho") <> "" Then
       Prt 66, 129, UCase(rConhecimentos.Fields("Redespacho"))
       Prt 70, 129, UCase(rRedespacho.Fields("c_Endereço"))
       Prt 74, 129, UCase(rRedespacho.Fields("c_Cidade"))
       Prt 74, 209, UCase(rRedespacho.Fields("c_Estado"))
       Prt 78, 129, rRedespacho.Fields("c_CNPJ")
    End If

    '-------------------
    'Coleta / Entrega
    '-------------------

    Prt 87, 181, UCase(rConhecimentos.Fields("LocColeta"))
    Prt 97, 181, UCase(rConhecimentos.Fields("LocEntrega01"))
    Prt 101, 181, UCase(rConhecimentos.Fields("LocEntrega02"))
    Prt 105, 181, UCase(rConhecimentos.Fields("LocEntrega03"))

    '------------------------------------
    'Mercadoria Transportada (Diz conter)
    '------------------------------------

    Prt 93, 3, UCase(rConhecimentos.Fields("Transp_NatCarga_01"))
    Prt 93, 23, rConhecimentos.Fields("Transp_Quantidade_01")
    Prt 93, 39, UCase(rConhecimentos.Fields("Transp_Especie_01"))
    Prt 93, 55, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_01"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_01"), "##,##0.00")
    Prt 93, 90, rConhecimentos.Fields("Transp_NotaFiscal_01")
    Prt 93, 122, Space(20 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_01"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_01"), "##,##0.00")

    Prt 97, 3, UCase(rConhecimentos.Fields("Transp_NatCarga_02"))
    Prt 97, 23, rConhecimentos.Fields("Transp_Quantidade_02")
    Prt 97, 39, UCase(rConhecimentos.Fields("Transp_Especie_02"))
    Prt 97, 55, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_02"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_02"), "##,##0.00")
    Prt 97, 90, rConhecimentos.Fields("Transp_NotaFiscal_02")
    Prt 97, 122, Space(20 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_02"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_02"), "##,##0.00")

    Prt 101, 3, UCase(rConhecimentos.Fields("Transp_NatCarga_03"))
    Prt 101, 23, rConhecimentos.Fields("Transp_Quantidade_03")
    Prt 101, 39, UCase(rConhecimentos.Fields("Transp_Especie_03"))
    Prt 101, 55, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_03"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_03"), "##,##0.00")
    Prt 101, 90, rConhecimentos.Fields("Transp_NotaFiscal_03")
    Prt 101, 122, Space(20 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_03"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_03"), "##,##0.00")

    Prt 105, 3, UCase(rConhecimentos.Fields("Transp_NatCarga_04"))
    Prt 105, 23, rConhecimentos.Fields("Transp_Quantidade_04")
    Prt 105, 39, UCase(rConhecimentos.Fields("Transp_Especie_04"))
    Prt 105, 55, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_04"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_04"), "##,##0.00")
    Prt 105, 90, rConhecimentos.Fields("Transp_NotaFiscal_04")
    Prt 105, 122, Space(20 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_04"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_04"), "##,##0.00")

    Prt 93, 154, UCase(rConhecimentos.Fields("Veiculo_Placa_1"))
    Prt 97, 154, UCase(rConhecimentos.Fields("Veiculo_Placa_2"))
    Prt 101, 154, UCase(rConhecimentos.Fields("Veiculo_Motorista"))

    '-------------------
    'Composição do frete
    '-------------------

    Prt 120, 4, rConhecimentos.Fields("FretePesoVolume")
    Prt 120, 27, Space(10 - Len(Format(rConhecimentos.Fields("FreteValor"), "##,##0.00"))) & Format(rConhecimentos.Fields("FreteValor"), "##,##0.00")
    Prt 120, 50, rConhecimentos.Fields("Sec_Cat_Desp_Pedagio")
    Prt 120, 71, rConhecimentos.Fields("Outros")
    Prt 120, 94, rConhecimentos.Fields("Suframa")
    Prt 120, 114, Space(10 - Len(Format(rConhecimentos.Fields("TxEmergencia"), "##,##0.00"))) & Format(rConhecimentos.Fields("TxEmergencia"), "##,##0.00")
    Prt 120, 136, Space(10 - Len(Format(rConhecimentos.Fields("TotalDaPrestacao"), "##,##0.00"))) & Format(rConhecimentos.Fields("TotalDaPrestacao"), "##,##0.00")
    Prt 120, 164, Space(10 - Len(Format(rConhecimentos.Fields("BaseDeCalculo"), "##,##0.00"))) & Format(rConhecimentos.Fields("BaseDeCalculo"), "##,##0.00")
    Prt 120, 189, rConhecimentos.Fields("Aliquota")
    Prt 120, 204, rConhecimentos.Fields("ICMS")

    Prt 127, 9, UCase(rConhecimentos.Fields("Observacoes01"))
    Prt 132, 9, UCase(rConhecimentos.Fields("Observacoes02"))
    Prt 137, 9, UCase(rConhecimentos.Fields("Observacoes03"))
    Prt 142, 9, UCase(rConhecimentos.Fields("Observacoes04"))
    Prt 145, 9, UCase(rConhecimentos.Fields("Observacoes05"))

    '------------------
    'Pulo de pagina
    '------------------

    Prt -3, -3, "NewPage"

    '------------------
    'Prox. Conhecimento
    '------------------

    rConhecimentos.MoveNext

Wend

Lote.Close
'MsgBox "ATENÇÃO: O arquivo foi gerado!", vbInformation
ImpressaoDeLote

Conhecimentos_Fim:

    Set rConhecimentos = Nothing
    Set rRemetente = Nothing
    Set rDestinatario = Nothing
    Set rConsignatario = Nothing
    Set rRedespacho = Nothing

    Set dbBase = Nothing
    
    Exit Sub
Conhecimentos_err:
    MsgBox Err.Description
    Resume Conhecimentos_Fim


End Sub

Sub Prt(Linha, Coluna, Texto)

Dim mSup As Long
Dim mEsq As Long

mEsq = 0
mSup = -5

DoEvents
If Val(Linha) < 0 Then
   Lote.Write Linha & ";" & Coluna & ";" & Texto & Chr(13)
Else
   Lote.Write Linha + mSup & ";" & Coluna + mEsq & ";" & Texto & Chr(13)
End If

End Sub

Sub ImpressaoDeLote()

Dim RetVal

ChDir Application.CurrentProject.Path

Caminho = Application.CurrentProject.Path

If Not Dir(Caminho & "\Impressao.exe") = "" Then
    RetVal = Shell(Caminho & "\Impressao.exe", 6)
Else
    MsgBox "O programa de impressão não foi encontrado!", vbCritical + vbOKOnly, "CANDIMAR - Gerenciamento de Serviços"
    
End If

End Sub


'Public Sub Conhecimentos(Inicio As Integer, Termino As Integer)
''''' ORIGINAL EPSON FX
''Banco de dados corrente
'Dim dbBase As Database
'
''Conhecimentos
'Dim rConhecimentos As DAO.Recordset
'Dim sConhecimentos As String
'
''Remetente
'Dim rRemetente As DAO.Recordset
'Dim sRemetente As String
'
''Destinatário
'Dim rDestinatario As DAO.Recordset
'Dim sDestinatario As String
'
''Consignatário
'Dim rConsignatario As DAO.Recordset
'Dim sConsignatario As String
'
''Redespacho
'Dim rRedespacho As DAO.Recordset
'Dim sRedespacho As String
'
'    '----------------
'    'Seleção de dados
'    '----------------
'
'    sConhecimentos = "Select * from Conhecimentos WHERE Conhecimento Between " & Inicio & " And " & Termino & " Order by Conhecimento"
'    sRemetente = "Select * from Clientes"
'    sDestinatario = "Select * from Clientes"
'    sConsignatario = "Select * from Clientes"
'    sRedespacho = "Select * from Clientes"
'
'    '---------------------------
'    'Setar tabelas selecionadas
'    '---------------------------
'
'    Set dbBase = CurrentDb
'    Set rConhecimentos = dbBase.OpenRecordset(sConhecimentos)
'    Set rRemetente = dbBase.OpenRecordset(sRemetente)
'    Set rDestinatario = dbBase.OpenRecordset(sDestinatario)
'    Set rConsignatario = dbBase.OpenRecordset(sConsignatario)
'    Set rRedespacho = dbBase.OpenRecordset(sRedespacho)
'
'    '---------------------
'    'Arquivo de lote
'    '---------------------
'
'    Caminho = Application.CurrentProject.Path
'
'    If Not Dir(Caminho & "\Lote.txt") = "" Then Kill Caminho & "\Lote.txt"
'
'    Set fLote = CreateObject("Scripting.FileSystemObject")
'    Set Lote = fLote.CreateTextFile(Caminho & "\Lote.txt", True)
'
'    '----------------------------
'    'Configuração de formulário
'    '----------------------------
'
'     Prt -1, 250, 152 'Tamanho de formulario
''     Prt -4, 1, " "   'NORMAL
''     Prt -4, 2, " "   'NEGRITO
''     Prt -4, 3, " "   'ITALICO
'
'
''---------------------
''Dados do conhecimento
''---------------------
'
'While Not rConhecimentos.EOF
'
'    '------------------
'    'Cabeçalho
'    '------------------
'
'    Prt 17, 130, rConhecimentos.Fields("ETC_RTB")
'    Prt 27, 130, rConhecimentos.Fields("CFOP")
'    Prt 31, 135, Format(rConhecimentos.Fields("Emissao"), "dd")
'    Prt 31, 160, UCase(Format(rConhecimentos.Fields("Emissao"), "mmmm"))
'    Prt 31, 185, Format(rConhecimentos.Fields("Emissao"), "yyyy")
'
'
'    '---------------------------
'    'Localiza dados do Remetente
'    '---------------------------
'
'    rRemetente.MoveLast
'    rRemetente.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Remetente") & "'"
'
'    If rConhecimentos.Fields("Remetente") <> "" Then
'       Prt 39, 23, UCase(rConhecimentos.Fields("Remetente"))
'       Prt 43, 23, UCase(rRemetente.Fields("c_Endereço"))
'       Prt 48, 23, UCase(rRemetente.Fields("c_Cidade"))
'       Prt 48, 65, UCase(rRemetente.Fields("c_Estado"))
'       Prt 48, 85, rRemetente.Fields("c_CEP")
'       Prt 52, 23, rRemetente.Fields("c_CNPJ")
'       Prt 52, 85, rRemetente.Fields("c_InscriçãoEstadual")
'    End If
'
'    '------------------------------
'    'Localiza dados do Destinatário
'    '------------------------------
'
'    rDestinatario.MoveLast
'    rDestinatario.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Destinatario") & "'"
'
'    If rConhecimentos.Fields("Destinatario") <> "" Then
'       Prt 39, 135, UCase(rConhecimentos.Fields("Destinatario"))
'       Prt 43, 135, UCase(rDestinatario.Fields("c_Endereço"))
'       Prt 48, 135, UCase(rDestinatario.Fields("c_Cidade"))
'       Prt 48, 217, UCase(rDestinatario.Fields("c_Estado"))
'       Prt 52, 135, rDestinatario.Fields("c_CNPJ")
'       Prt 52, 190, rDestinatario.Fields("c_InscriçãoEstadual")
'    End If
'
'    '-------------------------------
'    'Localiza dados do Consignatário
'    '-------------------------------
'
'    rConsignatario.MoveLast
'    rConsignatario.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Consignatario") & "'"
'
'    If rConhecimentos.Fields("Consignatario") <> "" Then
'       Prt 60, 28, UCase(rConhecimentos.Fields("Consignatario"))
'       Prt 64, 23, UCase(rConsignatario.Fields("c_Endereço"))
'       Prt 68, 23, UCase(rConsignatario.Fields("c_Cidade"))
'       Prt 68, 105, UCase(rConsignatario.Fields("c_Estado"))
'    End If
'
'    If rConhecimentos.Fields("FreteConsignatario") = "A Pagar" Then
'       Prt 73, 34, "X"
'    ElseIf rConhecimentos.Fields("FreteConsignatario") = "Pago" Then
'       Prt 73, 88, "X"
'    End If
'
'    Prt 78, 23, rConhecimentos.Fields("Distancia")
'
'    '----------------------------
'    'Localiza dados do Redespacho
'    '----------------------------
'
'    rRedespacho.MoveLast
'    rRedespacho.FindFirst "c_Cliente = '" & rConhecimentos.Fields("Redespacho") & "'"
'
'    If rConhecimentos.Fields("FreteRedespacho") = "Pago" Then
'       Prt 60, 154, "X"
'    ElseIf rConhecimentos.Fields("FreteRedespacho") = "A Pagar" Then
'       Prt 60, 185, "X"
'    End If
'
'    If rConhecimentos.Fields("Redespacho") <> "" Then
'       Prt 66, 135, UCase(rConhecimentos.Fields("Redespacho"))
'       Prt 70, 135, UCase(rRedespacho.Fields("c_Endereço"))
'       Prt 74, 135, UCase(rRedespacho.Fields("c_Cidade"))
'       Prt 74, 215, UCase(rRedespacho.Fields("c_Estado"))
'       Prt 78, 135, rRedespacho.Fields("c_CNPJ")
'    End If
'
'    '-------------------
'    'Coleta / Entrega
'    '-------------------
'
'    Prt 87, 187, UCase(rConhecimentos.Fields("LocColeta"))
'    Prt 97, 187, UCase(rConhecimentos.Fields("LocEntrega01"))
'    Prt 101, 187, UCase(rConhecimentos.Fields("LocEntrega02"))
'    Prt 105, 187, UCase(rConhecimentos.Fields("LocEntrega03"))
'
'    '------------------------------------
'    'Mercadoria Transportada (Diz conter)
'    '------------------------------------
'
'    Prt 93, 8, UCase(rConhecimentos.Fields("Transp_NatCarga_01"))
'    Prt 93, 29, rConhecimentos.Fields("Transp_Quantidade_01")
'    Prt 93, 45, UCase(rConhecimentos.Fields("Transp_Especie_01"))
'    Prt 93, 61, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_01"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_01"), "##,##0.00")
'    Prt 93, 96, rConhecimentos.Fields("Transp_NotaFiscal_01")
'    Prt 93, 128, Space(10 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_01"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_01"), "##,##0.00")
'
'    Prt 97, 8, UCase(rConhecimentos.Fields("Transp_NatCarga_02"))
'    Prt 97, 29, rConhecimentos.Fields("Transp_Quantidade_02")
'    Prt 97, 45, UCase(rConhecimentos.Fields("Transp_Especie_02"))
'    Prt 97, 61, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_02"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_02"), "##,##0.00")
'    Prt 97, 96, rConhecimentos.Fields("Transp_NotaFiscal_02")
'    Prt 97, 128, Space(10 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_02"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_02"), "##,##0.00")
'
'    Prt 101, 8, UCase(rConhecimentos.Fields("Transp_NatCarga_03"))
'    Prt 101, 29, rConhecimentos.Fields("Transp_Quantidade_03")
'    Prt 101, 45, UCase(rConhecimentos.Fields("Transp_Especie_03"))
'    Prt 101, 61, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_03"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_03"), "##,##0.00")
'    Prt 101, 96, rConhecimentos.Fields("Transp_NotaFiscal_03")
'    Prt 101, 128, Space(10 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_03"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_03"), "##,##0.00")
'
'    Prt 105, 8, UCase(rConhecimentos.Fields("Transp_NatCarga_04"))
'    Prt 105, 29, rConhecimentos.Fields("Transp_Quantidade_04")
'    Prt 105, 45, UCase(rConhecimentos.Fields("Transp_Especie_04"))
'    Prt 105, 61, Space(10 - Len(Format(rConhecimentos.Fields("Transp_Peso_04"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_Peso_04"), "##,##0.00")
'    Prt 105, 96, rConhecimentos.Fields("Transp_NotaFiscal_04")
'    Prt 105, 128, Space(10 - Len(Format(rConhecimentos.Fields("Transp_ValMercadoria_04"), "##,##0.00"))) & Format(rConhecimentos.Fields("Transp_ValMercadoria_04"), "##,##0.00")
'
'    Prt 93, 160, UCase(rConhecimentos.Fields("Veiculo_Placa_1"))
'    Prt 97, 160, UCase(rConhecimentos.Fields("Veiculo_Placa_2"))
'    Prt 101, 160, UCase(rConhecimentos.Fields("Veiculo_Motorista"))
'
'    '-------------------
'    'Composição do frete
'    '-------------------
'
'    Prt 120, 10, rConhecimentos.Fields("FretePesoVolume")
'    Prt 120, 33, Space(10 - Len(Format(rConhecimentos.Fields("FreteValor"), "##,##0.00"))) & Format(rConhecimentos.Fields("FreteValor"), "##,##0.00")
'    Prt 120, 56, rConhecimentos.Fields("Sec_Cat_Desp_Pedagio")
'    Prt 120, 77, rConhecimentos.Fields("Outros")
'    Prt 120, 100, rConhecimentos.Fields("Suframa")
'    Prt 120, 120, Space(10 - Len(Format(rConhecimentos.Fields("TxEmergencia"), "##,##0.00"))) & Format(rConhecimentos.Fields("TxEmergencia"), "##,##0.00")
'    Prt 120, 140, Space(10 - Len(Format(rConhecimentos.Fields("TotalDaPrestacao"), "##,##0.00"))) & Format(rConhecimentos.Fields("TotalDaPrestacao"), "##,##0.00")
'    Prt 120, 170, Space(10 - Len(Format(rConhecimentos.Fields("BaseDeCalculo"), "##,##0.00"))) & Format(rConhecimentos.Fields("BaseDeCalculo"), "##,##0.00")
'    Prt 120, 195, rConhecimentos.Fields("Aliquota")
'    Prt 120, 210, rConhecimentos.Fields("ICMS")
'
'    Prt 127, 15, UCase(rConhecimentos.Fields("Observacoes01"))
'    Prt 132, 15, UCase(rConhecimentos.Fields("Observacoes02"))
'    Prt 137, 15, UCase(rConhecimentos.Fields("Observacoes03"))
'    Prt 142, 15, UCase(rConhecimentos.Fields("Observacoes04"))
'    Prt 145, 15, UCase(rConhecimentos.Fields("Observacoes05"))
'
'    '------------------
'    'Pulo de pagina
'    '------------------
'
'    Prt -3, -3, "NewPage"
'
'    '------------------
'    'Prox. Conhecimento
'    '------------------
'
'    rConhecimentos.MoveNext
'
'Wend
'
'Lote.Close
''MsgBox "ATENÇÃO: O arquivo foi gerado!", vbInformation
'ImpressaoDeLote
'
'End Sub
