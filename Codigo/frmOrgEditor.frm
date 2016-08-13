VERSION 5.00
Begin VB.Form frmOrgEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Organizacion de Mapas"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   811
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   9840
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   9840
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Left            =   10680
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdUpSize 
      Caption         =   "Actualizar Tamaño"
      Height          =   255
      Left            =   9840
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtAlto 
      Height          =   285
      Left            =   10560
      TabIndex        =   6
      Text            =   "10"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtAncho 
      Height          =   285
      Left            =   10560
      TabIndex        =   5
      Text            =   "10"
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox mapPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9600
      Left            =   0
      ScaleHeight     =   636
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin VB.Label Label5 
      Caption         =   "Usar las flechitas para mover el mapa si es mas grande que 10 de ancho o de alto."
      Height          =   735
      Left            =   9720
      TabIndex        =   12
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Mapa:"
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Alto:"
      Height          =   255
      Left            =   9720
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Ancho:"
      Height          =   255
      Left            =   9720
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   9720
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmOrgEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'Author: Anagrama
'Last modified: 13/08/2016
'Editor de la organizacion de mapas en el MapOrg.dat.
'*************************************************

Option Explicit

Private MapOrg(100, 100) As Integer
Private MapWidth As Byte
Private MapHeight As Byte
Private SelectedX As Byte
Private SelectedY As Byte
Private OffsetX As Byte
Private OffsetY As Byte

Private Function LoadGeneralMapOrg() As Boolean
    Dim FilePath As String
    Dim X As Byte
    Dim y As Byte

    FilePath = App.path & "\Init\MapOrg.dat"
    If Not FileExist(FilePath, vbArchive) Then Exit Function
    
    MapWidth = Val(GetVar(FilePath, "General", "MapWidth"))
    MapHeight = Val(GetVar(FilePath, "General", "MapWidth"))
    
    If MapWidth = 0 Or MapHeight = 0 Then Exit Function
    
    For X = 1 To MapWidth
        For y = 1 To MapHeight
            MapOrg(X, y) = Val(GetVar(FilePath, "General", X & "-" & y))
        Next y
    Next X
    
    txtAncho.Text = MapWidth
    txtAlto.Text = MapHeight
    LoadGeneralMapOrg = True
End Function

Private Function LoadDungeonMapOrg() As Boolean
    Dim FilePath As String
    Dim X As Byte
    Dim y As Byte

    FilePath = App.path & "\Init\MapOrg.dat"
    If Not FileExist(FilePath, vbArchive) Then Exit Function
    
    MapWidth = Val(GetVar(FilePath, "Dungeon", "MapWidth"))
    MapHeight = Val(GetVar(FilePath, "Dungeon", "MapWidth"))
    
    If MapWidth = 0 Or MapHeight = 0 Then Exit Function
    
    For X = 1 To MapWidth
        For y = 1 To MapHeight
            MapOrg(X, y) = Val(GetVar(FilePath, "Dungeon", X & "-" & y))
        Next y
    Next X
    LoadDungeonMapOrg = True
End Function

Private Sub cmbTipo_Click()
    If cmbTipo.List(cmbTipo.ListIndex) = "General" Then
        Call ClearMapOrg
        If Not LoadGeneralMapOrg Then
            MapWidth = 10
            MapHeight = 10
        End If
    ElseIf cmbTipo.List(cmbTipo.ListIndex) = "Dungeon" Then
        Call ClearMapOrg
        If Not LoadDungeonMapOrg Then
            MapWidth = 10
            MapHeight = 10
        End If
    End If
    
    txtAncho.Text = MapWidth
    txtAlto.Text = MapHeight
    Call DrawMapOrg
End Sub

Private Sub cmdCambiar_Click()
    If SelectedX <= 0 Or SelectedX > MapWidth Or SelectedY <= 0 Or SelectedY > MapHeight Then Exit Sub
    If (Not FileExist(App.path & "\Renders\" & txtMap.Text & ".bmp", vbArchive)) And txtMap.Text <> 0 Then Exit Sub
    
    MapOrg(SelectedX, SelectedY) = txtMap.Text
    Call DrawMapOrg
End Sub

Private Sub cmdGuardar_Click()
    Dim FilePath As String
    Dim X As Byte
    Dim y As Byte
    
    FilePath = App.path & "\Init\MapOrg.dat"
    
    If cmbTipo.List(cmbTipo.ListIndex) = "General" Then
        Call WriteVar(FilePath, "General", "MapWidth", str(MapWidth))
        Call WriteVar(FilePath, "General", "MapHeight", str(MapHeight))
        For X = 1 To MapWidth
            For y = 1 To MapWidth
                Call WriteVar(FilePath, "General", X & "-" & y, str(MapOrg(X, y)))
            Next y
        Next X
    ElseIf cmbTipo.List(cmbTipo.ListIndex) = "Dungeon" Then
        Call WriteVar(FilePath, "Dungeon", "MapWidth", str(MapWidth))
        Call WriteVar(FilePath, "Dungeon", "MapHeight", str(MapHeight))
        For X = 1 To MapWidth
            For y = 1 To MapWidth
                Call WriteVar(FilePath, "Dungeon", X & "-" & y, str(MapOrg(X, y)))
            Next y
        Next X
    End If
End Sub

Private Sub cmdUpSize_Click()
    MapWidth = Val(txtAncho.Text)
    MapHeight = Val(txtAlto.Text)
    Call DrawMapOrg
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37 'left
            If OffsetX = 0 Then Exit Sub
            OffsetX = OffsetX - 1
        Case 38 'up
            If OffsetY = 0 Then Exit Sub
            OffsetY = OffsetY - 1
        Case 39 'right
            If 10 + OffsetX >= MapWidth Then Exit Sub
            OffsetX = OffsetX + 1
        Case 40 'down
            If 10 + OffsetY >= MapHeight Then Exit Sub
            OffsetY = OffsetY + 1
        Case Else: Exit Sub
    End Select
    
    Call DrawMapOrg
End Sub

Private Sub Form_Load()
    cmbTipo.Clear
    cmbTipo.AddItem "General"
    cmbTipo.AddItem "Dungeon"
    cmbTipo.ListIndex = 0
    
    Me.KeyPreview = True
    If Not LoadGeneralMapOrg Then
        MapWidth = 10
        MapHeight = 10
    End If
    
    Call DrawMapOrg
End Sub

Private Sub mapPic_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If X < mapPic.Left Or X > mapPic.Width Or y < mapPic.Top Or y > mapPic.Height Or X = 0 Or y = 0 Then Exit Sub
    
    SelectedX = Int(X / 64) + 1 + OffsetX
    SelectedY = Int(y / 64) + 1 + OffsetY
    
    If SelectedX > MapWidth Or SelectedY > MapHeight Then Exit Sub
    
    txtMap.Text = MapOrg(SelectedX, SelectedY)
End Sub

Private Sub txtAncho_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii))) And _
    (KeyAscii <> 8) And _
    (KeyAscii <> 44) And _
    (KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub txtAlto_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii))) And _
    (KeyAscii <> 8) And _
    (KeyAscii <> 44) And _
    (KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub txtMap_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii))) And _
    (KeyAscii <> 8) And _
    (KeyAscii <> 44) And _
    (KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub DrawMapOrg()
    Dim X As Byte
    Dim y As Byte
    
    mapPic.Cls
    
    For X = 1 To MapWidth
        For y = 1 To MapWidth
            If MapOrg(X + OffsetX, y + OffsetY) > 0 Then
                mapPic.PaintPicture LoadPicture(App.path & "\Renders\" & MapOrg(X + OffsetX, y + OffsetY) & ".bmp"), (X - 1) * 64, (y - 1) * 64
                mapPic.CurrentX = (X - 1) * 64 + 12
                mapPic.CurrentY = (y - 1) * 64 + 12
                mapPic.Print MapOrg(X + OffsetX, y + OffsetY)
            End If
        Next y
    Next X
End Sub

Private Sub ClearMapOrg()
    Dim X As Byte
    Dim y As Byte
    
    For X = 1 To MapWidth
        For y = 1 To MapWidth
            MapOrg(X, y) = 0
        Next y
    Next X
End Sub
