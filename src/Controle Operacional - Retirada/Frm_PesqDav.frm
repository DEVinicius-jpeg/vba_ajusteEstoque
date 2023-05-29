VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_PesqDav 
   Caption         =   "Pedidos"
   ClientHeight    =   8940.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "Frm_PesqDav.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_PesqDav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim querySelect As String

Private Sub btSelecionar_Click()
    Frm_BaixaDav.Show
End Sub

Private Sub UserForm_activate()
    
    querySelect = "SELECT DISTINCT pi2.PEI_DATA, CAST(pi2.PEI_DAV AS NUMERIC) AS DAV, pi2.PEI_CLIENTE_NOME, pi2.PEI_ENTREGA_CAIXA, pi2.PEI_NOTA_ID, pi2.PEI_CPF_CNPJ, pi2.PEI_VENDEDOR_NOME, pi2.PEI_CIDADE, pi2.PEI_BAIRRO FROM PEDIDO_ITEM pi2 WHERE PEI_DATA >= '2022-12-28' AND pi2.PEI_STATUS_EXP NOT IN (3,4) AND pi2.PEI_ENTREGA_CAIXA <> 2;"
    
    Call carregaListViewPedidos(querySelect)
    Call geraComboBoxTipo
    
End Sub

Function carregaListViewPedidos(ByVal sql As String)

    Dim rs As Recordset
    Dim linha As Integer
    Dim Inprocess As Boolean
    
    On Error GoTo error
    
    lvPedidos.ColumnHeaders.Clear
    lvPedidos.ListItems.Clear
    
    Inprocess = False
    
    conexaoFireBird
    
    Set rs = getRecordset(sql)
    
    Inprocess = True
    
    With lvPedidos
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .MultiSelect = False
        .LabelEdit = lvwManual
        
        .ColumnHeaders.Add Text:="Data", Width:=87, Alignment:=0
        .ColumnHeaders.Add Text:="Dav", Width:=65, Alignment:=0
        .ColumnHeaders.Add Text:="Nome do Cliente", Width:=230, Alignment:=0
        .ColumnHeaders.Add Text:="id", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="cnpj", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="vendedor", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="cidade", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="bairro", Width:=0, Alignment:=0
        .ColumnHeaders.Add Text:="Tipo", Width:=69, Alignment:=0
        
        linha = lvPedidos.ListItems.Count + 1
        
        Do Until rs.EOF
            .ListItems.Add = rs.Fields("PEI_DATA")
            .ListItems(linha).SubItems(1) = rs.Fields("DAV")
            .ListItems(linha).SubItems(2) = Trim(rs.Fields("PEI_CLIENTE_NOME"))
            .ListItems(linha).SubItems(3) = rs.Fields("PEI_NOTA_ID")
            .ListItems(linha).SubItems(4) = rs.Fields("PEI_CPF_CNPJ")
            .ListItems(linha).SubItems(5) = rs.Fields("PEI_VENDEDOR_NOME")
            .ListItems(linha).SubItems(6) = rs.Fields("PEI_CIDADE")
            .ListItems(linha).SubItems(7) = rs.Fields("PEI_BAIRRO")
            .ListItems(linha).SubItems(8) = tipoPedido(rs.Fields("PEI_ENTREGA_CAIXA"))
    
            linha = lvPedidos.ListItems.Count + 1
            rs.MoveNext
        Loop
    End With
    
    lbregistros.Caption = lvPedidos.ListItems.Count
    
Function_Exit:
    fechaBancoDeDados
    Exit Function
    
error:
    If Inprocess = True Then
        Call carregaListViewPedidos(querySelect)
    End If
    
    mensagemErro err.Description
    Resume Function_Exit

End Function

Sub geraComboBoxTipo()
    Me.cbTipo.AddItem "", 0
    Me.cbTipo.AddItem "CAIXA", 1
    Me.cbTipo.AddItem "ENTREGA", 2
    
    Me.cbTipo.Value = Me.cbTipo.List(1)

End Sub

Sub pesquisar()
    
    Dim sqlSelect As String
    
    If Me.cbTipo.ListIndex = 0 Then
        mensagemInformacao ""
        Exit Sub
    Else
        sqlSelect = "SELECT DISTINCT pi2.PEI_DATA, CAST(pi2.PEI_DAV AS NUMERIC) AS DAV, pi2.PEI_CLIENTE_NOME, pi2.PEI_ENTREGA_CAIXA, pi2.PEI_NOTA_ID, pi2.PEI_CPF_CNPJ, pi2.PEI_VENDEDOR_NOME, pi2.PEI_CIDADE, pi2.PEI_BAIRRO FROM PEDIDO_ITEM pi2 WHERE " _
                & " PEI_DATA >= '2022-12-28' AND pi2.PEI_STATUS_EXP NOT IN (3,4) AND pi2.PEI_ENTREGA_CAIXA IN( " & Me.cbTipo.ListIndex & ") AND pi2.PEI_CLIENTE_NOME LIKE '%" & Me.tbNomeCliente.Text & "%' AND pi2.PEI_DAV LIKE '%" & Me.tbDav.Text & "%';"
    
        carregaListViewPedidos (sqlSelect)
    End If
    
End Sub

Private Sub btPesquisar_Click()
    Call pesquisar
End Sub

Private Sub tbDav_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = VBA.vbKeyReturn Then Call pesquisar
End Sub

Private Sub tbNomeCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = VBA.vbKeyReturn Then Call pesquisar
End Sub

Private Sub cbTipo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub















