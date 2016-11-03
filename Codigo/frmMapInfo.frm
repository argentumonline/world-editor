VERSION 5.00
Begin VB.Form frmMapInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información del Mapa"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMismoBando 
      Caption         =   "Mismo Bando"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox chkTierra 
      Caption         =   "Mapa Tierra"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CheckBox chkRoboNpc 
      Caption         =   "NPC Compartidos"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CheckBox chkResu 
      Caption         =   "Resu sin efecto"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox chkOcultar 
      Caption         =   "Ocultar sin efecto"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CheckBox chkInvocar 
      Caption         =   "Invocar sin efecto"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CheckBox chkInmo 
      Caption         =   "Inmo sin efecto"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox chkInvi 
      Caption         =   "Invi sin efecto"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtMapVersion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Text            =   "0"
      Top             =   480
      Width           =   2655
   End
   Begin WorldEditor.lvButtons_H cmdMusica 
      Height          =   330
      Left            =   3600
      TabIndex        =   15
      Top             =   810
      Width           =   735
      _extentx        =   1296
      _extenty        =   582
      caption         =   "&Más"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":628A
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
      _extentx        =   2990
      _extenty        =   661
      caption         =   "&Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":62B6
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.ComboBox txtMapRestringir 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":62E2
      Left            =   1680
      List            =   "frmMapInfo.frx":62F5
      TabIndex        =   12
      Text            =   "NO"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CheckBox chkMapPK 
      Caption         =   "PK (inseguro)"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox txtMapTerreno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":631C
      Left            =   1680
      List            =   "frmMapInfo.frx":6329
      TabIndex        =   10
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ComboBox txtMapZona 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":6346
      Left            =   1680
      List            =   "frmMapInfo.frx":6353
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtMapMusica 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "0"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtMapNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Text            =   "Nuevo Mapa"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox chkMapBackup 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox chkMapNoEncriptarMP 
      Caption         =   "No Encriptar MP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox chkMapMagiaSinEfecto 
      Caption         =   "Magia Sin Efecto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Versión del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   135
      X2              =   4315
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label5 
      Caption         =   "Restringir:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Terreno:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Zona:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Musica:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4315
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "frmMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private Sub chkInmo_Click()
MapInfo.InmovilizarSinEfecto = chkInmo.Value
MapInfo.Changed = 1
End Sub

Private Sub chkInvi_Click()
MapInfo.InviSinEfecto = chkInvi.Value
MapInfo.Changed = 1
End Sub

Private Sub chkInvocar_Click()
MapInfo.InvocarSinEfecto = chkInvocar.Value
MapInfo.Changed = 1
End Sub

Private Sub chkMapBackup_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.BackUp = chkMapBackup.Value
MapInfo.Changed = 1
End Sub

Private Sub chkMapMagiaSinEfecto_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.MagiaSinEfecto = chkMapMagiaSinEfecto.Value
MapInfo.Changed = 1
End Sub

Private Sub chkMapNoEncriptarMP_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.NoEncriptarMP = chkMapNoEncriptarMP.Value
MapInfo.Changed = 1
End Sub

Private Sub chkMapPK_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.PK = chkMapPK.Value
MapInfo.Changed = 1
End Sub

Private Sub chkMismoBando_Click()
MapInfo.MismoBando = chkMismoBando.Value
MapInfo.Changed = 1
End Sub

Private Sub chkOcultar_Click()
MapInfo.OcultarSinEfecto = chkOcultar.Value
MapInfo.Changed = 1
End Sub

Private Sub chkResu_Click()
MapInfo.ResuSinEfecto = chkResu.Value
MapInfo.Changed = 1
End Sub

Private Sub chkRoboNpc_Click()
MapInfo.RoboNpcsPermitido = chkRoboNpc.Value
MapInfo.Changed = 1
End Sub

Private Sub chkTierra_Click()
MapInfo.MapaTierra = chkTierra.Value
MapInfo.Changed = 1
End Sub

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

Private Sub cmdMusica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMusica.Show vbModeless
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub txtMapMusica_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Music = txtMapMusica.Text
frmMain.lblMapMusica.Caption = MapInfo.Music
MapInfo.Changed = 1
End Sub

Private Sub txtMapVersion_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
MapInfo.MapVersion = txtMapVersion.Text
frmMain.lblMapVersion.Caption = MapInfo.MapVersion
MapInfo.Changed = 1
End Sub

Private Sub txtMapNombre_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.name = txtMapNombre.Text
frmMain.lblMapNombre.Caption = MapInfo.name
MapInfo.Changed = 1
End Sub

Private Sub txtMapRestringir_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapRestringir_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Restringir = txtMapRestringir.Text
MapInfo.Changed = 1
End Sub

Private Sub txtMapTerreno_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapTerreno_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Terreno = txtMapTerreno.Text
MapInfo.Changed = 1
End Sub

Private Sub txtMapZona_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapZona_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Zona = txtMapZona.Text
MapInfo.Changed = 1
End Sub
