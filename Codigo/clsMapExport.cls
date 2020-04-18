VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("WorldEditor.Class")
Option Explicit
Public Event OnCaptured()

Private WithEvents Picture As MyPicture
Attribute Picture.VB_VarHelpID = -1
Private format As eFormatPic
Private device As Integer
Private formatExtension(3) As String
Private options As MapExportOptions

Friend Sub SetOptions(ByRef newOptions As MapExportOptions)
    options = newOptions
End Sub
Public Property Let OutputFormat(ByVal newFormat As eFormatPic)
    format = newFormat
End Property

Public Property Get OutputFormat() As eFormatPic
    OutputFormat = format
End Property

Public Sub SetPicture(ByRef slavePicture As MyPicture)
    Set Picture = slavePicture
    
    Picture.Width = options.Width
    Picture.Height = options.Height
    
    device = wGL_Graphic.Create_Device_From_Display(Picture.hwnd, options.Width, options.Height)
End Sub

Public Sub Capture()
    If device <> -1 Then
        Call Invalidate(Picture.hwnd)
    End If
End Sub

Private Sub Class_Initialize()
    device = -1
    Set Picture = Nothing
    format = eFormatPic.bmp
    
    formatExtension(eFormatPic.bmp) = "bmp"
    formatExtension(eFormatPic.jpg) = "jpg"
    formatExtension(eFormatPic.png) = "png"
    
End Sub

Public Sub Save()
    Call wGL_Graphic.Capture(Picture.hwnd, App.path & "\Renders\" & NumMap_Save & "." & formatExtension(options.format))
End Sub

Private Sub picture_Paint()
    Call wGL_Graphic.Use_Device(device)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, &H0, 1#, 0)
    Call wGL_Graphic_Renderer.Update_Projection(&H0, 3200, 3200)
    
    Dim Drawable    As Integer
    Dim DrawableX   As Integer
    Dim DrawableY   As Integer
    
    Dim X           As Integer
    Dim Y           As Integer
    
    Dim MinY As Integer
    Dim MaxY As Integer
    Dim MinX As Integer
    Dim MaxX As Integer
    MinY = 1
    MaxY = 100
    MinX = 1
    MaxX = 100
    
    
    If options.floor Then
        For Y = MinY To MaxY
            DrawableY = (Y - 1) * TilePixelHeight
            For X = MinX To MaxX
                DrawableX = (X - 1) * TilePixelWidth
                Call DrawGrh(MapData(X, Y).Graphic(1), DrawableX, DrawableY, GetDepth(1, X, Y), 0, 1)
            Next X
        Next Y
    End If
    
    For Drawable = 0 To (g_Swarm.Query(MinX, MinY, MaxX, MaxY) - 1)
        X = g_Swarm.Query_X(Drawable)
        Y = g_Swarm.Query_Y(Drawable)

        DrawableX = (X - 1) * TilePixelWidth
        DrawableY = (Y - 1) * TilePixelHeight

        Select Case (g_Swarm.Query_Layer(Drawable))
            Case 1
                If options.layer2 Then
                    Call DrawGrh(MapData(X, Y).Graphic(2), DrawableX, DrawableY, GetDepth(2, X, Y), 1, 1)
                End If
            Case 2
                If options.layer3 Then
                    Call DrawGrh(MapData(X, Y).Graphic(3), DrawableX, DrawableY, GetDepth(3, X, Y, 2), 1, 1, , , , True)
                End If
            Case 3
                If options.layer4 Then
                    Call DrawGrh(MapData(X, Y).Graphic(4), DrawableX, DrawableY, GetDepth(4, X, Y), 1, 1)
                End If
            Case 4
                If options.objects Then
                    Call DrawGrh(MapData(X, Y).ObjGrh, DrawableX, DrawableY, GetDepth(3, X, Y, 1), 1, 1, , , , True)
                End If
            Case 5
                If options.npcs Then
                    Call CharRender(MapData(X, Y).CharIndex, DrawableX, DrawableY)
                End If
            Case modEdicion.BLOCK_LAYER
                If options.blocks Then
                    Call DrawGrhIndex(modEdicion.BlockGrhIndex, DrawableX, DrawableY, GetDepth(modEdicion.BLOCK_LAYER, X, Y, 1), True)
                End If
            Case modEdicion.EXIT_LAYER
                If options.exits Then
                    Call DrawGrhIndex(modEdicion.ExitGrhIndex, DrawableX, DrawableY, GetDepth(modEdicion.EXIT_LAYER, X, Y, 1), True)
                End If
            Case modEdicion.TRIGGER_LAYER
                If options.triggers Then
                    Call Draw_Text(FuentesJuego.Talk.id, FuentesJuego.Talk.Tamanio, DrawableX + 16, DrawableY + 16, GetDepth(modEdicion.TRIGGER_LAYER, X, Y, 1), FuentesJuego.Talk.color, FONT_ALIGNMENT_MIDDLE Or FONT_ALIGNMENT_CENTER, CStr(MapData(X, Y).Trigger))
                End If
        End Select
    Next Drawable
    
    Call wGL_Graphic_Renderer.Flush
    Call wGL_Graphic.Commit
    Call Save
    RaiseEvent OnCaptured
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set Picture = Nothing
    Call wGL_Graphic.Destroy_Device(device)
    device = -1
End Sub

