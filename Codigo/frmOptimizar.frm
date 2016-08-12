VERSION 5.00
Begin VB.Form frmOptimizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizar Mapa"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "frmOptimizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBloquearArbolesEtc 
      Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkMapearArbolesEtc 
      Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTodoBordes 
      Caption         =   "Quitar NPCs, Objetos y Translados en los Bordes Exteriores"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigTrans 
      Caption         =   "Quitar Trigger's en Translados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigBloq 
      Caption         =   "Quitar Trigger's Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrans 
      Caption         =   "Quitar Translados Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin WorldEditor.lvButtons_H cOptimizar 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Caption         =   "&Optimizar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin WorldEditor.lvButtons_H cCancelar 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      Caption         =   "&Cancelar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Optimizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

' Quita Translados Bloqueados
' Quita Trigger's Bloqueados
' Quita Trigger's en Translados
' Quita NPCs, Objetos y Translados en los Bordes Exteriores
' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa

modEdicion.Deshacer_Add "Aplicar Optimizacion del Mapa" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        With MapData(X, Y)
            ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
            If (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) And (chkQuitarTodoBordes.value = 1) Then
                'Quitar NPCs
                If .NPCIndex > 0 Then
                    EraseChar .CharIndex
                    .NPCIndex = 0
                End If
                
                ' Quitar Objetos
                .OBJInfo.objindex = 0
                .OBJInfo.Amount = 0
                .ObjGrh.GrhIndex = 0
                ' Quitar Translados
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
                ' Quitar Triggers
                .Trigger = 0
            End If
            
            ' ** Quitar Translados y Triggers en Bloqueo
            If (.Blocked = 1) Then
                If (.TileExit.Map > 0) And (chkQuitarTrans.value = 1) Then ' Quita Translado Bloqueado
                    .TileExit.Map = 0
                    .TileExit.Y = 0
                    .TileExit.X = 0
                ElseIf (.Trigger > 0) And (chkQuitarTrigBloq.value = 1) Then ' Quita Trigger Bloqueado
                    .Trigger = 0
                End If
            End If
            
            ' ** Quitar Triggers en Translado
            If (.TileExit.Map > 0) And (chkQuitarTrigTrans.value = 1) Then
                If (.Trigger > 0) Then ' Quita Trigger en Translado
                    .Trigger = 0
                End If
            End If
            
            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
            If (.OBJInfo.objindex > 0) And ((chkMapearArbolesEtc.value = 1) Or (chkBloquearArbolesEtc.value = 1)) Then
                Select Case ObjData(.OBJInfo.objindex).ObjType
                    Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                        If (.Graphic(3).GrhIndex <> .ObjGrh.GrhIndex) And (chkMapearArbolesEtc.value = 1) Then .Graphic(3) = .ObjGrh
                        If (chkBloquearArbolesEtc.value = 1) And (.Blocked = 0) Then .Blocked = 1
                End Select
            End If
            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        End With
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub cCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Unload Me
End Sub

Private Sub cOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Call Optimizar
End Sub
