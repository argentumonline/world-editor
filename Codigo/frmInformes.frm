VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmInformes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12315
   Icon            =   "frmInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstObjResults 
      DragMode        =   1  'Automatic
      Height          =   3735
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtInfo 
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin WorldEditor.lvButtons_H cmdObjetos 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      Caption         =   "&Objetos"
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
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "&Cerrar"
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
   Begin WorldEditor.lvButtons_H cmdTranslados 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   4200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      Caption         =   "&Translados"
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
   Begin WorldEditor.lvButtons_H cmdNPCs 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   4200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      Caption         =   "&NPCs/Hostiles"
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
   Begin WorldEditor.lvButtons_H cmdArboles 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   4200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      Caption         =   "&Recursos"
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
   Begin WorldEditor.lvButtons_H cmdBorrarObjetos 
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   3960
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   873
      Caption         =   "&Eliminar Objetos Seleccionados"
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
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6480
      Y1              =   4815
      Y2              =   4815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6480
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6480
      Y1              =   4070
      Y2              =   4070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6480
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("WorldEditor.Form.Tools")
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

Private Sub cmdArboles_Click()
    Call SetObjRemoveVisibility(False)
    Call InformeRecursos
End Sub

Private Sub cmdBorrarObjetos_Click()
    Call BorrarObjetosSeleccionados
End Sub

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

''
'   Genera el informe de Recursos (Arboles, Yacimientos, Cardúmenes
'

Private Sub InformeRecursos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

txtInfo.Text = "Informe de Recursos (X,Y)"

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).OBJInfo.objindex > 0 Then
            If IsResource(ObjData(MapData(X, Y).OBJInfo.objindex).ObjType) Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & MapData(X, Y).OBJInfo.Amount & " del Objeto " & MapData(X, Y).OBJInfo.objindex & " - " & ObjData(MapData(X, Y).OBJInfo.objindex).Name
            End If
        End If
    Next X
Next Y

End Sub

''
'   Genera el informe de Objetos
'

Private Sub ActalizarObjetos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    Dim Y As Integer
    Dim X As Integer
    Dim Count As Integer
    Dim lstItem As listItem

    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    Call SetObjRemoveVisibility(True)
    txtInfo.Text = "Informe de Objetos (X,Y)"
    
    lstObjResults.ListItems.Clear
    With lstObjResults
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "", 300
        .ColumnHeaders.Add 2, , "X-Y", 800
        .ColumnHeaders.Add 3, , "Amount", 1000
        .ColumnHeaders.Add 4, , "Objeto", .Width
    End With
    
    Count = 1

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).OBJInfo.objindex > 0 Then
                If Not IsResource(ObjData(MapData(X, Y).OBJInfo.objindex).ObjType) Then
                
                    txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & MapData(X, Y).OBJInfo.Amount & " " & ObjData(MapData(X, Y).OBJInfo.objindex).name
                    With lstObjResults
                        Set lstItem = .ListItems.Add(, , Count)
                        lstItem.SubItems(1) = X & "-" & Y
                        lstItem.SubItems(2) = MapData(X, Y).OBJInfo.Amount
                        lstItem.SubItems(3) = ObjData(MapData(X, Y).OBJInfo.objindex).name & " (" & MapData(X, Y).OBJInfo.objindex & ")"
                    End With
                    
                    Count = Count + 1
                End If
            End If
        Next X
    Next Y
    
    lstObjResults.Refresh

End Sub

''
'   Genera el informe de NPCs
'

Private Sub ActalizarNPCs()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer
Dim NPCIndex As Integer

If Not MapaCargado Then
    Exit Sub
End If

txtInfo.Text = "Informe de NPCs/Hostiles (X,Y)"

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        NPCIndex = MapData(X, Y).NPCIndex
        
        If NPCIndex > 0 Then
            If NpcData(NPCIndex).Hostile Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & NpcData(NPCIndex).Name & " (Hostil)"
            Else
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & NpcData(NPCIndex).Name
            End If
        End If
    Next X
Next Y

End Sub

''
'   Genera el informe de Translados
'

Private Sub ActalizarTranslados()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

txtInfo.Text = "Informe de Translados (X,Y)"

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).TileExit.Map > 0 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " nos traslada a la posición " & MapData(X, Y).TileExit.X & "," & MapData(X, Y).TileExit.Y & " del Mapa " & MapData(X, Y).TileExit.Map
                If ((X < 20 And MapData(X, Y).TileExit.X < 20) Or (X > 80 And MapData(X, Y).TileExit.X > 80)) And (X <> MapData(X, Y).TileExit.X) Then
                    txtInfo.Text = txtInfo.Text & " (X sospechoso)"
                End If
                If ((Y < 20 And MapData(X, Y).TileExit.Y < 20) Or (Y > 80 And MapData(X, Y).TileExit.Y > 80)) And (Y <> MapData(X, Y).TileExit.Y) Then
                    txtInfo.Text = txtInfo.Text & " (Y sospechoso)"
                End If
            End If
    Next X
Next Y

End Sub

Private Sub cmdNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call SetObjRemoveVisibility(False)
Call ActalizarNPCs
End Sub

Private Sub cmdObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call SetObjRemoveVisibility(True)
Call ActalizarObjetos
End Sub

Private Sub cmdTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call SetObjRemoveVisibility(False)
Call ActalizarTranslados
End Sub


Private Function IsResource(ByVal ObjType As Integer) As Boolean
    IsResource = ObjType = 4 Or ObjType = 22 Or ObjType = 45
End Function

Private Sub AddObjectToList(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal objindex As Integer)
On Error GoTo ErrHandler
    
    Dim listItem As MSComctlLib.listItem
    
    Set listItem = lstObjResults.ListItems.Add(index, "", X & "-" & Y & " - " & MapData(X, Y).OBJInfo.Amount & " - " & ObjData(MapData(X, Y).OBJInfo.objindex).name)
    Call listItem.ListSubItems.Add(1, "", X & "-" & Y)
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number
    
End Sub


Private Sub SetObjRemoveVisibility(ByVal Visible As Boolean)
    If Visible Then
        Me.Width = 12400
        Call cmdBorrarObjetos.SetFocus
    Else
        Me.Width = 6720
        Call Me.SetFocus
    End If
End Sub

Private Sub BorrarObjetosSeleccionados()
On Error GoTo ErrHandler
    Dim I As Integer
    
    Dim X As Integer
    Dim Y As Integer
    Dim TxtCoords() As String
    
    If MsgBox("¿Deseas eliminar los objetos de las coordenadas seleccionadas?", vbYesNo) = vbYes Then
        For I = 1 To lstObjResults.ListItems.Count
            If lstObjResults.ListItems.Item(I).Checked Then
                TxtCoords = Split(lstObjResults.ListItems.Item(I).ListSubItems(1).Text, "-")
                X = Int(TxtCoords(0))
                Y = Int(TxtCoords(1))
                Call QuitarObjeto(X, Y, True)
            End If
        Next I
        
        Call ActalizarObjetos
    End If
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number
End Sub

