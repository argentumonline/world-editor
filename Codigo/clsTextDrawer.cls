VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("WorldEditor.Class")
Option Explicit

Private Type RECTO ' xD
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, _
    ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, _
    ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, _
    ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InvalidateRect Lib "user32" ( _
    ByVal hwnd As Long, lpRect As RECTO, ByVal bErase As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_PAINT As Long = &HF

Private Const TEXT_TRANSPARENT As Long = 1
Private Const TEXT_OPAQUE As Long = 2

Private Const FW_DONTCARE As Long = 0
Private Const FW_THIN  As Long = 100
Private Const FW_EXTRALIGHT  As Long = 200
Private Const FW_ULTRALIGHT  As Long = 200
Private Const FW_LIGHT  As Long = 300
Private Const FW_NORMAL As Long = 400
Private Const FW_REGULAR As Long = 400
Private Const FW_MEDIUM  As Long = 500
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_DEMIBOLD As Long = 600
Private Const FW_BOLD  As Long = 700
Private Const FW_EXTRABOLD  As Long = 800
Private Const FW_ULTRABOLD As Long = 800
Private Const FW_BLACK  As Long = 900
Private Const FW_HEAVY  As Long = 900

Private Const ANSI_CHARSET As Long = 0
Private Const DEFAULT_CHARSET As Long = 1
Private Const SYMBOL_CHARSET As Long = 2
Private Const SHIFTJIS_CHARSET As Long = 128
Private Const OEM_CHARSET  As Long = 255

Private Const OUT_TT_ONLY_PRECIS  As Long = 7
Private Const CLIP_DEFAULT_PRECIS  As Long = 0
Private Const CLIP_LH_ANGLES  As Long = &H10
Private Const PROOF_QUALITY  As Long = 2
Private Const TRUETYPE_FONTTYPE As Long = &H4

Private Const TEXTBUFFER_SIZE As Long = 1024 * 2

Private Type TextInfo
    X As Long
    Y As Long
    color As Long
    centered As Boolean
    Text As String
End Type

Private Type size
    cx As Long
    cy As Long
End Type

Private TextBuffer(TEXTBUFFER_SIZE - 1) As TextInfo
Private textCount As Long

Private lngFont As Long
Private OldFont As Long

Private Const MAX_COLORS As Byte = 255

Private DirectDraw As DirectDraw7
Private useVideoMemory As Boolean
Private TextSurfaces(MAX_COLORS - 1) As Long
Private Surface(MAX_COLORS - 1) As DirectDrawSurface7
Private CantColors As Long
Private surfacesize As size
Private letters As String
Private lettersrect(255) As RECT
Private font As Integer
Public Function InitText(ByRef DD As DirectDraw7, ByVal videoMemory As Boolean)
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

  
    font = wGL_Graphic_Renderer.Create_Font(LoadBytes("FONT/Primary.ttf"))
    
Dim ret As size
Dim tsize As Integer
Dim I As Long

tsize = 0
letters = ""

For I = 32 To 255
    letters = letters & Chr$(I)
    
    lettersrect(I).Left = tsize
    lettersrect(I).Right = tsize + ret.cx
    lettersrect(I).Top = 0
    lettersrect(I).Bottom = 13
    
    tsize = tsize + ret.cx
    If I = 126 Then I = 160
Next I

CantColors = 1
End Function


Public Sub DrawText(ByVal X As Long, ByVal Y As Long, ByRef Text As String, ByVal color As Long, ByRef Surface As DirectDrawSurface7)
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

End Sub

Public Function DrawTextToDC(ByVal hDC As Long, Optional ByVal Shadow As Boolean, Optional ByVal FontSize As Long = 13) As Long

End Function

Public Function AddText(ByVal xPos As Integer, ByVal yPos As Integer, ByVal fontcolor As Long, ByRef strText As String, Optional ByVal centered As Boolean = False) As Long
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

If LenB(strText) > 0 Then
    If textCount < TEXTBUFFER_SIZE Then
        With TextBuffer(textCount)
            .X = xPos
            .Y = yPos
            .color = fontcolor
            .Text = strText
            .centered = centered
        End With
        textCount = textCount + 1
        AddText = textCount
    Else
        AddText = -1
    End If
End If

End Function

Private Sub Class_Terminate()
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

Dim I As Long

For I = 0 To CantColors
    Set Surface(I) = Nothing
Next I

End Sub
