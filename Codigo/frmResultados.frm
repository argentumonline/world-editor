VERSION 5.00
Begin VB.Form frmResultados 
   Caption         =   "Resultados de la última tarea"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtResultados 
      Height          =   8535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   9375
   End
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   9000
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   873
      Caption         =   "Cerrar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9480
      Y1              =   8880
      Y2              =   8880
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    txtResultados.Text = ""
    Unload Me
End Sub

Public Sub AgregarLinea(ByRef Texto As String)
    txtResultados.Text = txtResultados.Text & vbCrLf & Texto
End Sub

Private Sub Form_Load()

End Sub
