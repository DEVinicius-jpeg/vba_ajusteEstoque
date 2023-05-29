VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_BaixaDav 
   Caption         =   "Baixar Dav"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   OleObjectBlob   =   "Frm_BaixaDav.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_BaixaDav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdID As LongLong
Private xformat() As New ClassFormat

Private Sub btCargaTotal_Click()
    Dim contador As Integer
    Dim linha As Integer
    
    linha = 1
    
    With lvPedidoItem
        contador = .ListItems.Count
        While linha <= contador
            .ListItems(linha).SubItems(6) = .ListItems(linha).SubItems(3) - .ListItems(linha).SubItems(4)
            linha = linha + 1
        Wend
    End With
    
End Sub

Private Sub btEnviar_Click()
    Dim rs As Recordset
    Dim sqlUpdate, sqlInsert, nomeProduto As String
    Dim qtdCarregamento, qtdPedido, qtdExpedicao, controle, linha, status  As Integer
    Dim codigo, idItem As LongLong
    Dim carregamentoRealizado As Boolean
    
    linha = 1
    carregamentoRealizado = False
    
    With Me.lvPedidoItem
        controle = .ListItems.Count
        conexaoFireBird
        While linha <= controle
        
            qtdCarregamento = CInt(.ListItems(linha).SubItems(6))
            
            If qtdCarregamento > 0 Then
            
                nomeProduto = .ListItems(linha).SubItems(1)
                qtdPedido = CInt(.ListItems(linha).SubItems(3))
                qtdExpedicao = CInt(.ListItems(linha).SubItems(4))
                codigo = .ListItems(linha)
                idItem = .ListItems(linha).SubItems(5)
            
                sqlInsert = "INSERT INTO ESTOQUE (PD_ID, ES_QUANTIDADE, EM_ID, EL_ID, ES_DATA_MOVIMENTO, ES_LOTE, US_LOGIN, ES_CUSTO, ES_RASTREABILIDADE, ES_TIPO) VALUES " _
                          & "(" & codigo & ", " & -Abs(qtdCarregamento) & ", 10000002, 10000003, '" & Format(Date, "yyyy/MM/dd") & "', '', '" & usuarioLogin & "', 0, 0, 0);"
                          
                If qtdCarregamento + qtdExpedicao = qtdPedido Then
                    status = 3
                Else
                    status = 2
                End If
                
                sqlUpdate = "UPDATE PEDIDO_ITEM pi2 SET pi2.PEI_QUANTIDADE_SALDO_EXP = " & qtdCarregamento + qtdExpedicao & " , pi2.PEI_STATUS_EXP = " & status & ", pi2.PEI_DATA_ENTREGA_DAV = '" & Format(Date, "yyyy/MM/dd") & "', pi2.US_LOGIN = '" & usuarioLogin & "' WHERE pi2.PEI_ID = " & idItem & ";"
                
                Set rs = getRecordset(sqlInsert)
                Set rs = getRecordset(sqlUpdate)
                
                carregamentoRealizado = True
                
                mensagemInformacao "Carregamento realizado com sucesso!!!" & vbNewLine & "Produto: " & nomeProduto & vbNewLine & "Status: " & statusProduto(status) & vbNewLine & "Quantidade: " & qtdCarregamento
                
            End If
            
            linha = linha + 1
        Wend
        fechaBancoDeDados
    End With
    
    If carregamentoRealizado = True Then Call geraListViewPedidoItem
    
End Sub

Private Sub btImprimir_Click()

    On Error GoTo 1
    
    ActiveWorkbook.FollowHyperlink "http://?id=" & Trim(lvPedidoItem.ListItems(1).SubItems(5)) & "", NewWindow:=True
    
    Exit Sub
1:     MsgBox err.Description
End Sub

Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = VBA.vbKeyReturn Then
        With lvPedidoItem.SelectedItem
            If Me.txtEdit.Value <= (.ListSubItems(3) - .ListSubItems(4)) Then
                .ListSubItems(6).Text = Me.txtEdit.Value
            Else
                mensagemInformacao "Quantidade inválida, tente novamente!"
            End If
        End With
        Me.Frame1.Visible = False
        Me.txtEdit.Value = ""
    End If
End Sub

Private Sub UserForm_Initialize()

    Me.tbData.Value = Frm_PesqDav.lvPedidos.SelectedItem
    Me.tbDav.Value = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(1)
    Me.tbCliente.Value = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(2)
    pdID = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(3)
    Me.tbVendedor.Value = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(5)
    Me.tbCidade.Value = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(6)
    Me.tbCNPJ.Value = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(4)
    Me.tbBairro.Value = Frm_PesqDav.lvPedidos.SelectedItem.ListSubItems(7)
    
    Unload Frm_PesqDav
    Call geraListViewPedidoItem
    Call chamarFormat
End Sub

Sub geraListViewPedidoItem()

    Dim rs As Recordset
    Dim sqlSelect As String
    Dim linha As Integer
    
    sqlSelect = "SELECT p.PD_ID, pi2.PEI_ID, p.PD_NOME, pi2.PEI_STATUS_EXP, pi2.PEI_QUANTIDADE, pi2.PEI_QUANTIDADE_SALDO_EXP, pi2.PEI_DATA_ENTREGA_DAV FROM PEDIDO_ITEM pi2 LEFT JOIN PRODUTO p ON pi2.PD_ID = p.PD_ID WHERE pi2.PEI_NOTA_ID = " & pdID & ";"
    
    lvPedidoItem.ColumnHeaders.Clear
    lvPedidoItem.ListItems.Clear
    
    With lvPedidoItem
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="Código", Width:=55, Alignment:=0
        .ColumnHeaders.Add Text:="Nome Produto", Width:=220, Alignment:=0
        .ColumnHeaders.Add Text:="Status", Width:=65, Alignment:=0
        .ColumnHeaders.Add Text:="Qtd. Pedido", Width:=80, Alignment:=0
        .ColumnHeaders.Add Text:="Qtd. Retirado", Width:=80, Alignment:=0
        .ColumnHeaders.Add Text:="id", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="Qtd. Carregamento", Width:=95, Alignment:=0
        .ColumnHeaders.Add Text:="Dt. Entrega", Width:=60, Alignment:=0
        
        linha = .ListItems.Count + 1
        
        conexaoFireBird
            
            Set rs = getRecordset(sqlSelect)
            
            Do Until rs.EOF
                .ListItems.Add = rs.Fields("PD_ID")
                .ListItems(linha).SubItems(1) = Trim(rs.Fields("PD_NOME"))
                .ListItems(linha).SubItems(2) = statusProduto(rs.Fields("PEI_STATUS_EXP"))
                .ListItems(linha).SubItems(3) = rs.Fields("PEI_QUANTIDADE")
                .ListItems(linha).SubItems(4) = rs.Fields("PEI_QUANTIDADE_SALDO_EXP")
                .ListItems(linha).SubItems(5) = rs.Fields("PEI_ID")
                .ListItems(linha).SubItems(6) = 0
                If IsNull(rs.Fields(6)) Then
                    .ListItems(linha).SubItems(7) = ""
                Else
                    .ListItems(linha).SubItems(7) = rs.Fields("PEI_DATA_ENTREGA_DAV")
                End If

                linha = .ListItems.Count + 1
                rs.MoveNext
            Loop
            
        fechaBancoDeDados
    End With
        
End Sub

Private Sub lvPedidoItem_DblClick()
    Dim i As Integer
    If lvPedidoItem.SelectedItem.SubItems(3) > lvPedidoItem.SelectedItem.SubItems(4) Then
        If Not lvPedidoItem.SelectedItem Is Nothing Then
            i = 7
            With Frame1
                
                .Visible = True
                .Top = (lvPedidoItem.SelectedItem.Top + lvPedidoItem.Top) + 3.5
                .Left = (lvPedidoItem.ColumnHeaders(i).Left + lvPedidoItem.Left) + 3
                .Width = lvPedidoItem.ColumnHeaders(i).Width
                .Height = lvPedidoItem.SelectedItem.Height
                
                .ZOrder msoBringToFront
            End With
            
            With txtEdit
                
                .Visible = True
                .Text = lvPedidoItem.SelectedItem.SubItems(6)
                .SetFocus
                .SelStart = 0
                .Left = 0
                .Top = 0
                .Width = lvPedidoItem.ColumnHeaders(i).Width
                .Height = lvPedidoItem.SelectedItem.Height
                .SelLength = Len(.Text)
            End With
        End If
    End If
End Sub
Private Sub lvPedidoItem_Click()
    Me.Frame1.Visible = False
    Me.txtEdit.Value = ""
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
