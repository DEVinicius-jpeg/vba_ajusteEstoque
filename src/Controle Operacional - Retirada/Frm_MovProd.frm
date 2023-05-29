VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_MovProd 
   Caption         =   "Movimentação de Estoque"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14025
   OleObjectBlob   =   "Frm_MovProd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_MovProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private xformat() As New ClassFormat

Private Sub btAdicionar_Click()
    Call inserirProdutoLista
    lbregistros.Caption = lvProdutos.ListItems.Count
End Sub

Private Sub btBaixarDav_Click()
    Frm_PesqDav.Show
End Sub

Private Sub btEditarMov_Click()
    Dim rs As Recordset
    Dim sqlDelete, sqlUpdate As String
    Dim idMov As LongLong
    Dim qtdUpdate As Integer
    
    idMov = Application.InputBox("Informe o ID da movimentação que deseja excluir:", "Excluir Movimentação", Type:=1)
    
    If idMov = 0 Then Exit Sub
    
    sqlSelect = "SELECT e.ES_ID, e.ES_LOG_NOVO, e.US_LOGIN FROM ESTOQUE e WHERE e.ES_ID = " & idMov & "  AND e.EM_ID IN (10000011, 10000012) AND e.ES_DATA_MOVIMENTO >= '2022-12-27';"
    
    conexaoFireBird
        Set rs = getRecordset(sqlSelect)
        If rs.BOF = True Then
            mensagemInformacao "Não foi encontrado uma movimentação com o ID fornecido."
        Else
            If rs.Fields("US_LOGIN") <> usuarioLogin Then
                mensagemInformacao "A movimentação de ID: " & idMov & " não foi realizada pelo seu usuário."
                fechaBancoDeDados
                Exit Sub
            ElseIf DateDiff("s", rs.Fields("ES_LOG_NOVO"), Now) < 1200 Then
                qtdUpdate = Application.InputBox("Informe a quantidade para qual deseja editar:", "Editar Movimentação", Type:=1)
                If idMov = 0 Then
                    fechaBancoDeDados
                    Exit Sub
                End If
                sqlUpdate = "UPDATE ESTOQUE e SET e.ES_QUANTIDADE = " & qtdUpdate & " WHERE e.ES_ID = " & idMov & ";"
                Set rs = getRecordset(sqlUpdate)
                fechaBancoDeDados
                mensagemInformacao "A movimentação de ID: " & idMov & " foi atualizada com sucesso!"
                ThisWorkbook.RefreshAll
                Exit Sub
            Else
                mensagemInformacao "A movimentação de ID: " & idMov & " já excedeu o limite de tempo, de 20 minutos, para edição."
            End If
        End If
        
    fechaBancoDeDados
End Sub

Private Sub btExcluirMov_Click()
    Dim rs As Recordset
    Dim sqlDelete, sqlSelect As String
    Dim idMov As LongLong
    
    idMov = Application.InputBox("Informe o ID da movimentação que deseja excluir:", "Excluir Movimentação", Type:=1)
    
    If idMov = 0 Then Exit Sub
    
    sqlSelect = "SELECT e.ES_ID, e.ES_LOG_NOVO, e.US_LOGIN FROM ESTOQUE e WHERE e.ES_ID = " & idMov & "  AND e.EM_ID IN (10000011, 10000012) AND e.ES_DATA_MOVIMENTO >= '2022-12-27';"
    
    conexaoFireBird
        Set rs = getRecordset(sqlSelect)
        If rs.BOF = True Then
            mensagemInformacao "Não foi encontrado uma movimentação com o ID fornecido."
        Else
            If rs.Fields("US_LOGIN") <> usuarioLogin Then
                mensagemInformacao "A movimentação de ID: " & idMov & " não foi realizada pelo seu usuário."
                fechaBancoDeDados
                Exit Sub
            ElseIf DateDiff("s", rs.Fields("ES_LOG_NOVO"), Now) < 1200 Then
                sqlDelete = "DELETE FROM ESTOQUE e WHERE e.ES_ID = " & idMov & ";"
                Set rs = getRecordset(sqlDelete)
                fechaBancoDeDados
                mensagemInformacao "A movimentação de ID: " & idMov & " foi excluída com sucesso!"
                ThisWorkbook.RefreshAll
                Exit Sub
            Else
                mensagemInformacao "A movimentação de ID: " & idMov & " já excedeu o limite de tempo, de 20 minutos, para exclusão."
            End If
        End If
        
    fechaBancoDeDados
End Sub

Private Sub UserForm_terminate()
    Application.Visible = True

End Sub

Private Sub btMovEstoque_Click()
    Dim rs As Recordset
    Dim sqlInsert As String
    Dim i As Integer
    
    With lvProdutos
        If .ListItems.Count > 0 Then
            conexaoFireBird
                For i = 1 To .ListItems.Count
                    If .ListItems(i).SubItems(2) = 0 Then
                        sqlInsert = "INSERT INTO ESTOQUE (PD_ID, ES_QUANTIDADE, EM_ID, EL_ID, ES_DATA_MOVIMENTO, ES_LOTE, US_LOGIN, ES_CUSTO, ES_RASTREABILIDADE, ES_TIPO) VALUES " _
                                  & "(" & .ListItems(i) & ", " & .ListItems(i).SubItems(4) & ", 10000012, 10000003, '" & Format(Date, "yyyy/MM/dd") & "', '', '" & usuarioLogin & "', 0, 0, 0);"
                        Set rs = getRecordset(sqlInsert)
                    Else
                        sqlInsert = "INSERT INTO ESTOQUE (PD_ID, ES_QUANTIDADE, EM_ID, EL_ID, ES_DATA_MOVIMENTO, ES_LOTE, US_LOGIN, ES_CUSTO, ES_RASTREABILIDADE, ES_TIPO) VALUES " _
                                  & "(" & .ListItems(i) & ", " & -Abs(.ListItems(i).SubItems(4)) & ", 10000011, 10000003, '" & Format(Date, "yyyy/MM/dd") & "', '', '" & usuarioLogin & "', 0, 0, 0);"
                        Set rs = getRecordset(sqlInsert)
                    End If
                    mensagemInformacao "Movimentação realizada com sucesso!!!" & vbNewLine & "Produto: " & Trim(.ListItems(i).SubItems(1)) & vbNewLine & "Movimento: " & .ListItems(i).SubItems(3) & vbNewLine & "Quantidade: " & .ListItems(i).SubItems(4) & vbNewLine & "Exiba as movimentações para visualizar."
                Next
            fechaBancoDeDados
            .ListItems.Clear
            lbregistros.Caption = lvProdutos.ListItems.Count
            ThisWorkbook.RefreshAll
        Else
            mensagemInformacao "A lista de produtos esta vazia."
        End If
    End With
    
End Sub

Private Sub lvProdutos_KeyDown(KeyCode As Integer, ByVal Shift As Integer) ' Evento que exclui um item selecionado da lista de produtos ao teclar o backspace
    If KeyCode = VBA.vbKeyBack Then
        lvProdutos.ListItems.Remove (lvProdutos.SelectedItem.Index)
        lbregistros.Caption = lvProdutos.ListItems.Count
    End If
End Sub

Private Sub UserForm_Initialize()
    Call carregaListViewProdutos
    Call comboboxTipoMov
    Call chamarFormat
    
    Me.tbCodigo.SetFocus
    
End Sub

Private Sub tbCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = VBA.vbKeyReturn Or KeyCode = VBA.vbKeyTab Then
        
        Dim rs As ADODB.Recordset
        Dim querySelect As String
        
        On Error GoTo erro
        conexaoFireBird
        
        querySelect = "SELECT p.PD_NOME FROM PRODUTO p WHERE p.PD_ID = " & Me.tbCodigo.Value & " AND p.PL_ID = 10001;"
                
        Set rs = getRecordset(querySelect)
        
        If rs.BOF = False Then
            Me.tbNomeProduto.Value = rs.Fields("PD_NOME")
            fechaBancoDeDados
        Else
erro:
            fechaBancoDeDados
            mensagemInformacao "Código de produto inválido!"
            Me.tbNomeProduto.Value = Empty
            Me.tbCodigo.SetFocus
        End If
    End If
End Sub

Private Sub btExibirMov_Click()
    If Application.Visible = False Then
        ThisWorkbook.RefreshAll
        Application.Visible = True
    Else
        Application.Visible = False
    End If
    
End Sub

Private Sub carregaListViewProdutos() 'ListView dos produtos
    With lvProdutos
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="Código", Width:=70, Alignment:=0
        .ColumnHeaders.Add Text:="Nome do Produto ", Width:=276, Alignment:=0
        .ColumnHeaders.Add Text:="", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="Tipo de Movimentação", Width:=136, Alignment:=0
        .ColumnHeaders.Add Text:="Quantidade", Width:=90, Alignment:=0
        
    End With

End Sub

Sub comboboxTipoMov()
    Me.cbTipoMovimentacao.AddItem "AJUSTE DE SALDO - ENTRADA", 0
    Me.cbTipoMovimentacao.AddItem "AJUSTE DE SALDO - SAIDA", 1
End Sub

Private Sub tbQuantidade_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) ' Evento para quando o usuario teclar enter enviar as inf do produto para a listView de produtos

    If KeyCode = VBA.vbKeyReturn Then
        Call inserirProdutoLista
    End If
    
    lbregistros.Caption = lvProdutos.ListItems.Count
    
End Sub

Sub inserirProdutoLista()
    Dim linha As Integer
    
    If Me.tbQuantidade.Value <> "" And Me.tbQuantidade.Value <> 0 And Me.tbCodigo.Value <> "" And Me.cbTipoMovimentacao.Value <> "" And Me.tbNomeProduto.Value <> "" Then
        linha = lvProdutos.ListItems.Count + 1
        With lvProdutos
            lvProdutos.ListItems.Add = Me.tbCodigo.Text
            lvProdutos.ListItems(linha).SubItems(1) = Me.tbNomeProduto.Text
            lvProdutos.ListItems(linha).SubItems(2) = Me.cbTipoMovimentacao.ListIndex
            lvProdutos.ListItems(linha).SubItems(3) = Me.cbTipoMovimentacao.Value
            lvProdutos.ListItems(linha).SubItems(4) = Me.tbQuantidade.Value
                    
            Me.tbCodigo.Text = Empty
            Me.tbNomeProduto.Text = Empty
            Me.cbTipoMovimentacao.Text = Empty
            Me.tbQuantidade.Text = Empty

            Me.tbCodigo.SetFocus
        End With
    Else
        Me.tbQuantidade.SetFocus
        mensagemInformacao "As informações do produto está Invalida, tente novamente!"
    End If
End Sub

Private Sub cbTipoMovimentacao_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub chamarFormat() 'Rotina que confere todos os objetos do fomulário para colocar a mascara de entrada de acordo com o objeto

    Dim i As Integer
    Dim cont As Integer
    
    cont = Me.Controls.Count - 1
    
    ReDim xformat(0 To cont)
        For i = 0 To cont
            Select Case Me.Controls(i).Tag
                Case Is = "numero" ' Mascara de apenas numeros
                    Set xformat(i).toNumero = Me.Controls(i)
            End Select
        Next
End Sub
