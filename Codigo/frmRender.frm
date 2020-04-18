VERSION 5.00
Object = "{97FD4A65-A045-4F5C-8C6C-262505F7C013}#6.0#0"; "Argentum.ocx"
Begin VB.Form frmRender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renderizado"
   ClientHeight    =   13515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   25620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   901
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1708
   StartUpPosition =   3  'Windows Default
   Begin WorldEditor.UcRenderOptions renderOption 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   2835
      _extentx        =   5001
      _extenty        =   4948
   End
   Begin ArgentumOCX.MyPicture slave 
      CausesValidation=   0   'False
      Height          =   2715
      Left            =   2970
      TabIndex        =   2
      Top             =   60
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   4789
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3720
      Width           =   1275
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("WorldEditor.Form")
Option Explicit
Private WithEvents exporter As clsMapExport
Attribute exporter.VB_VarHelpID = -1

Public formatPic As eFormatPic

Private Sub cmdAceptar_Click()
    Set exporter = New clsMapExport
    Call Me.renderOption.ConfigureExporter(exporter)
    Call exporter.SetPicture(slave)
    
    
    Call exporter.Capture
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set exporter = Nothing
End Sub

Private Sub exporter_OnCaptured()
    Unload Me
End Sub
