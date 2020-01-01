VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRenderAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renderizar todos los mapas"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSizeX 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "64"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtSizeY 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "64"
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox tmpPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2280
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pgbProgressTotal 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblEstadoTotal 
      Alignment       =   2  'Center
      Caption         =   "0/1"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ancho:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Alto:"
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmRenderAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("WorldEditor.Form")
'*************************************************
'Author: Anagrama
'Last modified: 13/08/2016
'Maneja la generacion de renders de todos los mapas en la carpeta /Mapas/.
'*************************************************
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

    Call RenderAllMaps(formatPic, SizeX, SizeY)
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



