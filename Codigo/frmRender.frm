VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmRender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renderizado"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox tmpPic 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   2280
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtSizeY 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "3200"
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   120
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtSizeX 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "3200"
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Alto:"
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ancho:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public formatPic As eFormatPic

Private Sub cmdAceptar_Click()
Dim SizeX As Long
Dim SizeY As Long

If Not IsNumeric(txtSizeX.Text) Then
    MsgBox "El ancho es inválido."
    Exit Sub
End If
If Not IsNumeric(txtSizeY.Text) Then
    MsgBox "El alto es inválido."
    Exit Sub
End If

SizeX = txtSizeX.Text
SizeY = txtSizeY.Text


Call MapCapture(formatPic, SizeX, SizeY)
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub txtSizeX_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii))) And _
    (KeyAscii <> 8) And _
    (KeyAscii <> 44) And _
    (KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub txtSizeY_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii))) And _
    (KeyAscii <> 8) And _
    (KeyAscii <> 44) And _
    (KeyAscii <> 46) Then KeyAscii = 0
End Sub




