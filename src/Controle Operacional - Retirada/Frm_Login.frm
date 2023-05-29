VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Login 
   Caption         =   "Login"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "Frm_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public validaLogin As Boolean

Private Sub btEntrar_Click()
    Dim rs As Recordset
    Dim sqlSelect As String
    
    validaLogin = False
    
    sqlSelect = "SELECT u.US_LOGIN, u.US_SENHA FROM USUARIO u;"
    If Me.tbLogin.Value <> "" And Me.tbSenha.Value <> "" Then
        conexaoFireBird
            Set rs = getRecordset(sqlSelect)
            Do While rs.EOF = False
                If rs.Fields("US_LOGIN") = UCase(Me.tbLogin.Value) Then
                    If rs.Fields("US_SENHA") = Me.tbSenha.Value Then
                        usuarioLogin = rs.Fields("US_LOGIN")
                        fechaBancoDeDados
                        validaLogin = True
                        Unload Frm_Login
                        Frm_MovProd.Show False
                        Exit Sub
                    Else
                        mensagemInformacao "A senha incorreta, tente novamente!!"
                        fechaBancoDeDados
                        Exit Sub
                    End If
                End If
                rs.MoveNext
            Loop
        fechaBancoDeDados
    
        mensagemInformacao "Usuário não encontrado"
    End If
    
End Sub
Private Sub chExibir_Click()
    '
    If Me.chExibir.Value = True Then
        Me.tbSenha.PasswordChar = ""
    Else
        Me.tbSenha.PasswordChar = "*"
    End If
    
End Sub

Private Sub UserForm_terminate()
    If validaLogin = False Then Application.Visible = True
End Sub
