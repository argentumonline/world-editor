VERSION 5.00
Begin VB.Form frmErrors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Errores"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtErrors 
      Enabled         =   0   'False
      Height          =   2025
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   7905
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("WorldEditor.Form.Helpers")
Option Explicit
Public Sub AddError(message As String)
    txtErrors.Text = txtErrors.Text & message & vbCrLf
End Sub

Public Sub ClearErrors()
    txtErrors.Text = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub
Private Sub txtErrors_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Public Function HasErrors() As Boolean
    HasErrors = LenB(txtErrors.Text) > 0
End Function

