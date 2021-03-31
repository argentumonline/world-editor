VERSION 5.00
Begin VB.Form frmFKEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar FK.Ind"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCambiarNumMaps 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtNumMaps 
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Text            =   "300"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Recargar Archivo"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar Archivo"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CheckBox chkLluvia 
      Caption         =   "Llueve"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ListBox lstMaps 
      Height          =   6300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad de Mapas:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblMap 
      Caption         =   "1"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Mapa:"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmFKEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("WorldEditor.Form")
Option Explicit

Private bLluvia() As Byte

Public Sub LoadFK()
    Dim N As Integer
    Dim I As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open DirIndex & "fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For I = 1 To Nu
        Get #N, , bLluvia(I)
    Next I
    
    Close #N
End Sub

Public Sub SaveFK()
    Dim N As Integer
    Dim I As Long
    Dim Nu As Integer
    
    Nu = UBound(bLluvia)
    
    N = FreeFile
    Open DirIndex & "fk.ind" For Binary As #N
    
    Put #N, , MiCabecera
    
    Put #N, , Nu
    
    For I = 1 To Nu
        Put #N, , bLluvia(I)
    Next I
    
    Close #N
End Sub

Private Sub chkLluvia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstMaps.ListIndex < 0 Then Exit Sub
    
    bLluvia(lstMaps.ListIndex + 1) = chkLluvia.Value
    lstMaps.List(lstMaps.ListIndex) = "Mapa " & lstMaps.ListIndex + 1 & " - " & IIf(bLluvia(lstMaps.ListIndex + 1) = 1, "Llueve", "No Llueve")
    'Call ShowMapList(UBound(bLluvia))
End Sub

Private Sub cmdCambiarNumMaps_Click()
    Dim I As Integer
    
    ReDim Preserve bLluvia(1 To txtNumMaps.Text) As Byte
    
    Call ShowMapList(txtNumMaps.Text)
    
End Sub

Private Sub cmdCargar_Click()
    Dim I As Integer
    
    Call LoadFK
    
    txtNumMaps.Text = UBound(bLluvia)
    
    Call ShowMapList(UBound(bLluvia))
End Sub

Private Sub cmdGuardar_Click()
    Call SaveFK
End Sub

Private Sub ShowMapList(ByVal NumMaps As Integer)
    Dim I As Integer
    
    lstMaps.Clear
    
    For I = 1 To NumMaps
        lstMaps.AddItem "Mapa " & I & " - " & IIf(bLluvia(I) = 1, "Llueve", "No Llueve")
    Next I

End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    Call LoadFK
    
    txtNumMaps.Text = UBound(bLluvia)
    
    Call ShowMapList(UBound(bLluvia))
End Sub

Private Sub lstMaps_Click()
    lblMap.Caption = lstMaps.ListIndex + 1
    chkLluvia.Value = bLluvia(lstMaps.ListIndex + 1)
End Sub
