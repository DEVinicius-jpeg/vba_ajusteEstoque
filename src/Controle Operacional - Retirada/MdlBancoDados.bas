Attribute VB_Name = "MdlBancoDados"
Option Explicit
Option Private Module

Public usuarioLogin

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset



Public Sub conexaoFireBird()

    On Error GoTo erro
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    With cn
        .connectionString = "DRIVER=Firebird/Interbase(r) driver;UID=;PWD=;DBNAME="
        .Mode = adModeReadWrite
        .Open
    
    End With
    
    rs.ActiveConnection = cn
    
    Exit Sub

erro:
    mensagemErro err.Description
    
    fechaBancoDeDados

End Sub
Public Sub fechaBancoDeDados()

    On Error Resume Next
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
End Sub

Function getRecordset(ByVal sql As String, Optional ByVal cursorType As CursorTypeEnum = adOpenForwardOnly, _
    Optional ByVal lockType As LockTypeEnum = adLockReadOnly) As Recordset

    If rs.State = adStateOpen Then
    
        rs.Close
        
    End If
        
    rs.Open sql, cn, cursorType, lockType
    
    Set getRecordset = rs
    
End Function
