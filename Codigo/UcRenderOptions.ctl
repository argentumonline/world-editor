VERSION 5.00
Begin VB.UserControl UcRenderOptions 
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   2730
   ScaleWidth      =   2865
   Begin VB.Frame frameDraw 
      Caption         =   "Draw"
      Height          =   2025
      Left            =   30
      TabIndex        =   9
      Top             =   630
      Width           =   1455
      Begin VB.CheckBox chkNpcs 
         Caption         =   "Npcs"
         Height          =   315
         Left            =   270
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "Objects"
         Height          =   315
         Left            =   270
         TabIndex        =   14
         Top             =   1380
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkLayer4 
         Caption         =   "Layer 4"
         Height          =   315
         Left            =   270
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkLayer3 
         Caption         =   "Layer 3"
         Height          =   315
         Left            =   270
         TabIndex        =   12
         Top             =   780
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkLayer2 
         Caption         =   "Layer 2"
         Height          =   315
         Left            =   270
         TabIndex        =   11
         Top             =   510
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkFloor 
         Caption         =   "Floor"
         Height          =   315
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Frame FrameSize 
      Caption         =   "Size"
      Height          =   555
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   2745
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "100"
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   810
         TabIndex        =   6
         Text            =   "100"
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height"
         Height          =   225
         Left            =   1530
         TabIndex        =   7
         Top             =   210
         Width           =   645
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width"
         Height          =   225
         Left            =   300
         TabIndex        =   5
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.Frame FrameFormat 
      Caption         =   "Format"
      Height          =   1605
      Left            =   1620
      TabIndex        =   0
      Top             =   630
      Width           =   1155
      Begin VB.OptionButton optJpg 
         Caption         =   "Jpg"
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   1200
         Width           =   885
      End
      Begin VB.OptionButton optBmp 
         Caption         =   "Bmp"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   810
         Width           =   885
      End
      Begin VB.OptionButton optPng 
         Caption         =   "Png"
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   885
      End
   End
End
Attribute VB_Name = "UcRenderOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("WorldEditor.UserControls")
Option Explicit

Public Sub ConfigureExporter(ByRef exporter As clsMapExport)

    Call exporter.SetOptions(GetOptions())

End Sub

Private Function GetOptions() As MapExportOptions

'TODO make validations

GetOptions.Width = txtWidth.Text
GetOptions.Height = txtHeight.Text

GetOptions.floor = IIf(chkFloor.Value = vbChecked, True, False)
GetOptions.layer2 = IIf(chkLayer2.Value = vbChecked, True, False)
GetOptions.layer3 = IIf(chkLayer3.Value = vbChecked, True, False)
GetOptions.layer4 = IIf(chkLayer4.Value = vbChecked, True, False)
GetOptions.objects = IIf(chkObjects.Value = vbChecked, True, False)
GetOptions.npcs = IIf(chkNpcs.Value = vbChecked, True, False)

If optPng.Value Then
    GetOptions.format = png
ElseIf optBmp.Value Then
    GetOptions.format = bmp
Else
    GetOptions.format = jpg
End If

End Function

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii))) And _
                                       (KeyAscii <> 8) And _
                                       (KeyAscii <> 44) And _
                                       (KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If (Not IsNumeric(Chr$(KeyAscii))) And _
                                       (KeyAscii <> 8) And _
                                       (KeyAscii <> 44) And _
                                       (KeyAscii <> 46) Then KeyAscii = 0
End Sub


