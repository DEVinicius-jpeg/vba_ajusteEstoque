Attribute VB_Name = "MdlFuncao"
Public Function statusProduto(ByVal status As Integer) As String
    Select Case status
        Case Is = 1
            statusProduto = "PENDENTE"
        Case Is = 2
            statusProduto = "PARCIAL"
        Case Is = 3
            statusProduto = "ENCERRADO"
        Case Is = 4
            statusProduto = "CANCELADO"
    End Select
End Function

Public Function tipoPedido(ByVal tipo As Integer) As String
    Select Case tipo
        Case Is = 1
            tipoPedido = "CAIXA"
        Case Is = 2
            tipoPedido = "ENTREGA"
    End Select
End Function
