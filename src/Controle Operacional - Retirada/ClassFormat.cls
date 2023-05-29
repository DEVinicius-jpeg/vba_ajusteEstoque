VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents toNumero As MSForms.TextBox
Attribute toNumero.VB_VarHelpID = -1

Private Sub toNumero_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With toNumero
        Select Case KeyAscii
            Case 48 To 57
            Case Else: KeyAscii = 0
        End Select
    End With
End Sub
