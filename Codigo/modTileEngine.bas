Attribute VB_Name = "modTileEngine"
Option Explicit
Public Const ENGINE_SPEED As Single = 0.018
Public Type tFuente
    id As Integer
    Tamanio As Long
    color As Long
End Type

Public timerEngine As Currency

Public Type tFuentesJuego
    FuenteBase As tFuente
    
    'Nicks
    NickCriminal As tFuente
    NickCiudadano As tFuente
    NickAtacable As tFuente
    NickConcilio As tFuente
    NickConsejo As tFuente
    NickDios As tFuente
    NickSemidios As tFuente
    NickConsejero As tFuente
    NickAdmins As tFuente
    NickRolemasters As tFuente
    NickNpcs As tFuente
    
    'General
    Talk As tFuente
    Fight As tFuente
    Warning As tFuente
    Info As tFuente
    InfoBold As tFuente
    Execution As tFuente
    Party As tFuente
    Poison As tFuente
    Guild As tFuente
    Server As tFuente
    GuildMsg As tFuente
    Centinela As tFuente
    GMSG As tFuente
    
    ConsejoVesA As tFuente
    ConcilioVesA As tFuente
    
    Inventarios As tFuente

End Type

Public FuentesJuego As tFuentesJuego
''''''''''''''''''''''''''''''''''''''''''''''''
''' WGL (TEMPORALLY)
'''''''''''''''''''''''''''''''''''''''''''''''

Private g_Material(0 To 65535) As Integer
Private g_Technique_1 As Integer
Private g_Technique_2 As Integer
Public g_Swarm As New wGL_Temp_Swarm


Private Type PostEffectUniform
    Effect  As wGL_Uniform
End Type

Private g_Post_Effect_Device    As Integer
Private g_Post_Effect_Material  As Integer
Private g_Post_Effect_Technique As Integer
Private g_Post_Effect_Uniform   As PostEffectUniform
Private g_Rain_Material         As Integer
Private GFX_PATH As String
Public Function GetImageFromNum(ByVal fileNum As Long) As Byte()
On Error GoTo ErrHandler
  
    Dim InfoHead As INFOHEADER
    GFX_PATH = "graphics"
    If Get_InfoHeader(App.path & "\" & GFX_PATH & "\", fileNum & ".BMP", InfoHead) Then
        Call Extract_File(App.path & "\" & GFX_PATH & "\", InfoHead, GetImageFromNum)
    Else
        If Get_InfoHeader(App.path & "\" & GFX_PATH & "\", fileNum & ".PNG", InfoHead) Then
            Call Extract_File(App.path & "\" & GFX_PATH & "\", InfoHead, GetImageFromNum)
        Else
            GoTo ErrHandler
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetImageFromNum de ModDDEX.bas")
End Function

Public Function CreateFont(ByVal id As Integer, ByVal Tamanio As Long, ByVal color As Long) As tFuente
    
    CreateFont.id = id
    CreateFont.Tamanio = Tamanio
    CreateFont.color = color
    
End Function

Public Sub LoadFontDescription()
On Error GoTo ErrHandler
    
    Dim Font As Integer
    Font = wGL_Graphic_Renderer.Create_Font(LoadBytes("FONT/Primary.ttf"))


    ' RGBA
    FuentesJuego.FuenteBase = CreateFont(Font, 14, &HFFFFFFFF)
    FuentesJuego.NickCriminal = CreateFont(Font, 14, &HFFFF0000)
    FuentesJuego.NickCiudadano = CreateFont(Font, 14, &HFF0080FF)
    FuentesJuego.NickAtacable = CreateFont(Font, 14, &HFFB332FF)
    FuentesJuego.NickAdmins = CreateFont(Font, 14, &HFFFFFFFF)
    FuentesJuego.NickDios = CreateFont(Font, 14, &HFFFAFA96)
    FuentesJuego.NickSemidios = CreateFont(Font, 14, &HFF1EFF30)
    FuentesJuego.NickConsejero = CreateFont(Font, 14, &HFF1E9630)
    FuentesJuego.NickAdmins = CreateFont(Font, 14, &HFFB4B4B4)
    FuentesJuego.NickConcilio = CreateFont(Font, 14, &HFFFF3200)
    FuentesJuego.NickConsejo = CreateFont(Font, 14, &HFF0C3FF)
    FuentesJuego.NickNpcs = CreateFont(Font, 14, &HFFB6A951)
    
    FuentesJuego.Talk = CreateFont(Font, 14, &HFFFFFFFF)
    FuentesJuego.Fight = CreateFont(Font, 14, &HFFFF0000)
    FuentesJuego.Warning = CreateFont(Font, 14, &HFF2033E9)
    FuentesJuego.Info = CreateFont(Font, 14, &HFF41BE9C)
    FuentesJuego.InfoBold = CreateFont(Font, 14, &HFF31BE9C)
    FuentesJuego.Execution = CreateFont(Font, 14, &HFF828282)
    FuentesJuego.Party = CreateFont(Font, 14, &HFFFFB4FF)
    FuentesJuego.Poison = CreateFont(Font, 14, &HFF00FF00)
    
    FuentesJuego.Guild = CreateFont(Font, 14, &HFFFFFFFF)
    FuentesJuego.Server = CreateFont(Font, 14, &HFF00B900)
    FuentesJuego.GuildMsg = CreateFont(Font, 14, &HFFFFC71B)
    
    FuentesJuego.ConsejoVesA = CreateFont(Font, 14, &HFF00C8FF)
    FuentesJuego.ConcilioVesA = CreateFont(Font, 14, &HFFFF3200)
    
    FuentesJuego.Centinela = CreateFont(Font, 14, &HFF00FF00)
    FuentesJuego.GMSG = CreateFont(Font, 14, &HFFFFFFFF)

    FuentesJuego.Inventarios = CreateFont(Font, 12, &HFFFFFFFF)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadFontDescription de Mod_TileEngine.bas")
End Sub

Public Sub Draw(ByRef destination As wGL_Rectangle, ByRef source As wGL_Rectangle, ByVal Depth As Single, ByVal Angle As Single, ByVal color As Long, ByVal Graphic As Long, ByVal Alpha As Boolean)
 
    If (g_Material(Graphic) = 0) Then
        g_Material(Graphic) = wGL_Graphic_Renderer.Create_Material
        
        Call wGL_Graphic_Renderer.Update_Material_Texture(g_Material(Graphic), &H0, wGL_Graphic.Create_Texture_From_Image(GetImageFromNum(Graphic)))
    End If
    
    If (Alpha) Then
        Call wGL_Graphic_Renderer.Draw(destination, source, Depth, Angle, color, g_Material(Graphic), g_Technique_2)
    Else
        Call wGL_Graphic_Renderer.Draw(destination, source, Depth, Angle, color, g_Material(Graphic), g_Technique_1)
    End If
    
End Sub

Public Function GetCharacterDimension(ByVal CharIndex As Integer, ByRef RangeX As Single, ByRef RangeY As Single)
    Dim I As Long
    
    Dim BestRange As Long
            
    With CharList(CharIndex)
    
        ' Try to calculate the best width and height using all four direction of the entity's body
        If (.iBody <> 0) Then
            For I = 1 To 4
                If (GrhData(.Body.Walk(I).grhIndex).TileWidth > RangeX) Then
                    RangeX = GrhData(.Body.Walk(I).grhIndex).TileWidth
                End If
                If (GrhData(.Body.Walk(I).grhIndex).TileHeight > RangeY) Then
                    RangeY = GrhData(.Body.Walk(I).grhIndex).TileHeight
                End If
            Next I
        End If
                
        ' Try to calculate the best width and height using all four direction of the entity's body
        If (.iHead <> 0) Then

            For I = 1 To 4
                If (GrhData(.Head.Head(I).grhIndex).TileWidth > RangeX) Then
                    RangeX = GrhData(.Head.Head(I).grhIndex).TileWidth
                End If
            Next I
            For I = 1 To 4
                If (GrhData(.Head.Head(I).grhIndex).TileHeight > BestRange) Then
                    BestRange = GrhData(.Head.Head(I).grhIndex).TileHeight
                End If
            Next I

            RangeY = RangeY + BestRange
        End If
        
    End With


End Function

Public Function GetDepth(ByVal Layer As Single, Optional ByVal X As Single = 1, Optional ByVal Y As Single = 1, Optional ByVal Z As Single = 1) As Single

    GetDepth = -1# + (Layer * 0.1) + ((Y - 1) * 0.001) + ((X - 1) * 0.00001) + ((Z - 1) * 0.000001)
    
End Function

Public Function LoadBytes(ByVal FileName As String) As Byte()

    Open App.path + "\" + FileName For Binary Access Read Lock Read As #1

        ReDim LoadBytes(LOF(1) - 1)
    
        Get #1, , LoadBytes

    Close #1
    
End Function

Public Function ARGB(Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte) As Long
    If Alpha > 127 Then
       ARGB = ((Alpha - 128) * &H1000000 Or &H80000000) Or Blue Or (Green * &H100&) Or (Red * &H10000)
    Else
       ARGB = (Alpha * &H1000000) Or Blue Or (Green * &H100&) Or (Red * &H10000)
    End If
End Function

Public Sub TempClearForm()
           
    Call wGL_Graphic.Use_Device(&H0)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, &H0, 1#, 0)

End Sub

Public Sub SetTileBuffer(ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer)
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    
        MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    Dim surfaceDesc As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error Resume Next
    Set DirectX = New DirectX7
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

    
    '****** INIT DirectDraw ******
    ' Create the root DirectDraw object
    Set DirectDraw = DirectX.DirectDrawCreate("")
    
    If Err Then
        MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

    
 Dim Mode As Long
    Mode = MODE_COMPATIBLE ' MODE_SYNCHRONISED MODE_COMPATIBLE
    
    If (wGL_Graphic.Create_Driver(DRIVER_DIRECT3D9, Mode, frmMain.picMain.hwnd, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight)) = False Then
        MsgBox "No se pudo encontrar d3d9.dll. Esto puede deberse a que tu sistema operativo no es compatible, o que alguna de las librerías no está correctamente instalada o actualizada. " _
               & "Contacta a Soporte para más información."
        End
    End If

    ' TEMPORALLY (New lib's version will remove all of this)
    
    g_Technique_1 = wGL_Graphic_Renderer.Create_Technique
    Call wGL_Graphic_Renderer.Update_Technique_Program(g_Technique_1, wGL_Graphic.Create_Program(LoadBytes("Shader\Basic.vs"), LoadBytes("Shader\Basic-1.fs")))

    Dim Descriptor As wGL_Graphic_Descriptor
    Descriptor.Depth = COMPARISON_LESS_EQUAL
    Descriptor.Depth_Mask = True
    Descriptor.Mask_Red = True: Descriptor.Mask_Green = True: Descriptor.Mask_Blue = True: Descriptor.Mask_Alpha = True
    Descriptor.Stencil_Mask = &HFF
    Call wGL_Graphic_Renderer.Update_Technique_Descriptor(g_Technique_1, Descriptor)
    
    Dim Sampler As wGL_Graphic_Sampler
    Sampler.Address_X = SAMPLER_ADDRESS_WRAP
    Sampler.Address_Y = SAMPLER_ADDRESS_WRAP
    Call wGL_Graphic_Renderer.Update_Technique_Sampler(g_Technique_1, 0, Sampler)
    
    g_Technique_2 = wGL_Graphic_Renderer.Create_Technique
    Call wGL_Graphic_Renderer.Update_Technique_Program(g_Technique_2, wGL_Graphic.Create_Program(LoadBytes("Shader\Basic.vs"), LoadBytes("Shader\Basic-2.fs")))
        Call wGL_Graphic_Renderer.Update_Technique_Sampler(g_Technique_2, 0, Sampler)
    
    Descriptor.Blend_Color_Source = BLEND_FACTOR_SRC_ALPHA
    Descriptor.Blend_Color_Destination = BLEND_FACTOR_ONE_MINUS_SRC_ALPHA
    Descriptor.Depth_Mask = False
    Call wGL_Graphic_Renderer.Update_Technique_Descriptor(g_Technique_2, Descriptor)
    
    Dim Texture As Integer
    Texture = wGL_Graphic.Create_Texture(FORMAT_BGRA8, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight, &H0, True, True)
    g_Post_Effect_Material = wGL_Graphic_Renderer.Create_Material
    Call wGL_Graphic_Renderer.Update_Material_Texture(g_Post_Effect_Material, 0, Texture)

    g_Post_Effect_Device = wGL_Graphic.Create_Device(Texture, 0, 0, 0, wGL_Graphic.Create_Texture(FORMAT_D24S8, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight, 0, False, True))
    
    g_Post_Effect_Technique = wGL_Graphic_Renderer.Create_Technique
    Call wGL_Graphic_Renderer.Update_Technique_Program(g_Post_Effect_Technique, wGL_Graphic.Create_Program(LoadBytes("Shader\Effect.vs"), LoadBytes("Shader\Effect.fs")))
    
    g_Rain_Material = wGL_Graphic_Renderer.Create_Material
    Call wGL_Graphic_Renderer.Update_Material_Texture(g_Rain_Material, 0, wGL_Graphic.Create_Texture_From_Image(GetImageFromNum(15168)))
    
    Call LoadFontDescription
    'Load graphic data into memory
    modIndices.CargarIndicesDeGraficos

    frmCargando.X.Caption = "Iniciando Control de Superficies..."

    'Wave Sound
    Set DirectSound = DirectX.DirectSoundCreate("")
    DirectSound.SetCooperativeLevel setDisplayFormhWnd, DSSCL_PRIORITY
    LastSoundBufferUsed = 1
    
    InitTileEngine = True
    Call TempClearForm
End Function
Public Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    If UserMoving Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False
            End If
        End If
        
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = False
            End If
        End If
    End If
    
                    Call wGL_Graphic.Use_Device(&H0)
                    Call wGL_Graphic_Renderer.Update_Projection(&H0, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight)

                        g_Post_Effect_Uniform.Effect.X = 0
                    
                    
                    Call wGL_Graphic.Use_Uniform(&H0, False, g_Post_Effect_Uniform, 1)
    
                    Dim destination As wGL_Rectangle, source As wGL_Rectangle
                    destination.X1 = 0#: destination.X2 = frmMain.picMain.ScaleWidth: destination.Y1 = 0#: destination.Y2 = frmMain.picMain.ScaleHeight
                    source.X1 = 0#: source.X2 = 1#: source.Y1 = 0#: source.Y2 = 1#
                    Call wGL_Graphic_Renderer.Draw(destination, source, 0#, 0#, -1, g_Post_Effect_Material, g_Post_Effect_Technique)
                        
                    Call wGL_Graphic_Renderer.Flush

                        '****** Update screen ******
                    Call wGL_Graphic.Use_Device(g_Post_Effect_Device)
                    
                    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, &H0, 1#, 0)
                    Call wGL_Graphic_Renderer.Update_Projection(&H0, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight)
                    
                    
                    Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
                        
                    'Call Dialogos.Render
            
                    'Call DialogosClanes.Draw(FuentesJuego.Guild)
                    'Call DibujarCartel
                    
                    'If (bRain And bLluvia(UserMap)) Then
                        'Call DrawRain
                    'End If
 
                    Call wGL_Graphic_Renderer.Flush
    Call wGL_Graphic.Commit
    'FPS update
    If fpsLastCheck + 1000 < DirectX.TickCount Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        fpsLastCheck = DirectX.TickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
    
    'Get timing info
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
        Dim Elapsed As Currency
        Elapsed = GetElapsedTime()
        timerTicksPerFrame = Elapsed * ENGINE_SPEED
        timerEngine = timerEngine + Elapsed
End Sub

Public Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal OffsetX As Integer, ByVal OffsetY As Integer)
    Dim ScreenMinY  As Integer  'Start Y pos on current screen
    Dim ScreenMaxY  As Integer  'End Y pos on current screen
    Dim ScreenMinX  As Integer  'Start X pos on current screen
    Dim ScreenMaxX  As Integer  'End X pos on current screen
    Dim MinY        As Integer  'Start Y pos on current map
    Dim MaxY        As Integer  'End Y pos on current map
    Dim MinX        As Integer  'Start X pos on current map
    Dim MaxX        As Integer  'End X pos on current map
    Dim X           As Integer
    Dim Y           As Integer
    Dim Drawable    As Integer
    Dim DrawableX   As Integer
    Dim DrawableY   As Integer
    
    'Calculate ceiling alpha
    Dim Alpha As Long
    Alpha = IIf(bTecho, &H60FFFFFF, -1)
    
    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth
    
    'Figure out Ends and Starts of map
    MinY = ScreenMinY
    MaxY = ScreenMaxY
    MinX = ScreenMinX
    MaxX = ScreenMaxX
    
    If OffsetY < 0 Then
        MaxY = MaxY + 1
    ElseIf OffsetY > 0 Then
        MinY = MinY - 1
    End If
    If OffsetX < 0 Then
        MaxX = MaxX + 1
    ElseIf OffsetX > 0 Then
        MinX = MinX - 1
    End If
    
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

    For Y = MinY To MaxY
        DrawableY = (Y - ScreenMinY) * TilePixelHeight + OffsetY
    
        For X = MinX To MaxX
            DrawableX = (X - ScreenMinX) * TilePixelWidth + OffsetX
        
            Call DrawGrh(MapData(X, Y).Graphic(1), DrawableX, DrawableY, GetDepth(1, X, Y), 0, 1)
            
        Next X
    Next Y
    
    If bSelectSup Then
        DrawableY = (SobreY - ScreenMinY) * TilePixelHeight + OffsetY
        DrawableX = (SobreX - ScreenMinX) * TilePixelWidth + OffsetX
        If MosaicoChecked Then
            Call DrawGrh(CurrentGrh(((X + DespX) Mod mAncho) + 1, ((Y + DespY) Mod MAlto) + 1), DrawableX, DrawableY, GetDepth(CurLayer + 1, X, Y), 0, 1)
        Else
            Call DrawGrh(CurrentGrh(0), DrawableX, DrawableY, GetDepth(CurLayer + 1, X, Y), 0, 1)
        End If
    End If
        
    For Drawable = 0 To (g_Swarm.Query(MinX, MinY, MaxX, MaxY) - 1)
        X = g_Swarm.Query_X(Drawable)
        Y = g_Swarm.Query_Y(Drawable)

        DrawableX = (X - ScreenMinX) * TilePixelWidth + OffsetX
        DrawableY = (Y - ScreenMinY) * TilePixelHeight + OffsetY

        Select Case (g_Swarm.Query_Layer(Drawable))
            Case 1
            If bVerCapa(2) Then
                Call DrawGrh(MapData(X, Y).Graphic(2), DrawableX, DrawableY, GetDepth(2, X, Y), 1, 1)
            End If
            Case 2
                If bVerCapa(3) Then
                    Call DrawGrh(MapData(X, Y).Graphic(3), DrawableX, DrawableY, GetDepth(3, X, Y, 2), 1, 1, , , , True)
                End If
            Case 3
                If bVerCapa(4) Then
                    Call DrawGrh(MapData(X, Y).Graphic(4), DrawableX, DrawableY, GetDepth(4, X, Y), 1, 1)
                End If
            Case 4
                If bVerObjetos Then
                    Call DrawGrh(MapData(X, Y).ObjGrh, DrawableX, DrawableY, GetDepth(3, X, Y, 1), 1, 1, , , , True)
                End If
            Case 5
                If bVerNpcs Then
                    Call CharRender(MapData(X, Y).CharIndex, DrawableX, DrawableY)
                End If
        End Select
    Next Drawable
End Sub


Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 25/05/2011 (Amraphen)
'Draw char's to screen without offcentering them
'16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
'25/05/2011: Amraphen - Agregado movimiento de armas al golpear.
'***************************************************
On Error GoTo ErrHandler
  
    Dim moved As Boolean
    Dim attacked As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Dim I As Byte
    Dim LastIndex As Byte
    Dim TextOffsetY As Integer
    
    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame

                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If

        
            attacked = False
            
        'If done moving stop animation
        If Not moved Then
            .Body.Walk(.Heading).Started = 0
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
            'Draw Body
            If .Body.Walk(.Heading).grhIndex Then _
                Call DrawGrh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 2), 1, 1, , 0, , True)
            
            'Draw Head
            If .Head.Head(.Heading).grhIndex > 0 Then
                If .Head.Head(.Heading).grhIndex Then
                    Call DrawGrh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, GetDepth(3, .Pos.X, .Pos.Y, 3), 1, 0, , , , True)
                End If
                
                'Draw Helmet
                'If .Casco.Head(.Heading).grhIndex Then
                    'Call DrawGrh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, GetDepth(3, .Pos.X, .Pos.Y, 4), 1, 0, , , , True)
                'End If
                
                'Draw Weapon
                'If .Arma.WeaponWalk(.Heading).grhIndex Then _
                    'Call DrawGrh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 6), 1, 1, , 0, , True)
                
                'Draw Shield
                'If .Escudo.ShieldWalk(.Heading).grhIndex Then _
                    'Call DrawGrh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 5), 1, 1, , 0, , True)
                
                Dim fuente As tFuente
                ' Set a default font.
                fuente = FuentesJuego.FuenteBase
        End If
        
        ' Set chat text offsets
        'TextOffsetY = GetChatOverheadTextOffset(CharIndex, PixelOffsetY, TilePixelHeight)
        
        'Update dialogs
        'Call Dialogos.UpdateDialogPos(PixelOffsetX + TilePixelWidth \ 2, TextOffsetY, 0#, CharIndex)
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CharRender de TileEngine.bas")
End Sub

Public Sub DrawGrh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Z As Single, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal color As Long = -1, Optional ByVal killAtEnd As Byte = 1, Optional ByVal Angle As Integer = 0, Optional ByVal Alpha As Boolean = False)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    If Grh.grhIndex = 0 Then
     Exit Sub
    End If
    Dim CurrentGrhIndex As Integer
    Dim CurrentFrame    As Integer

    If Animate Then
        If Grh.Started = 1 Then
            CurrentFrame = ((timerEngine - Grh.FrameCounter) * GrhData(Grh.grhIndex).NumFrames / Grh.Speed)
            
            If CurrentFrame > GrhData(Grh.grhIndex).NumFrames Then
                CurrentFrame = (CurrentFrame Mod GrhData(Grh.grhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                        If killAtEnd Then Exit Sub
                    End If
                Else
                    Grh.FrameCounter = timerEngine
                End If
            End If
        End If
    End If
    If (CurrentFrame = 0) Then CurrentFrame = 1
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhIndex).Frames(CurrentFrame)

    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                    
        Dim source As wGL_Rectangle, destination As wGL_Rectangle
    
        source.X1 = .S0
        source.Y1 = .T0
        source.X2 = .S1
        source.Y2 = .T1
        destination.X1 = X
        destination.Y1 = Y
        destination.X2 = X + .pixelWidth
        destination.Y2 = Y + .pixelHeight
        
        Call Draw(destination, source, Z, Angle, color, .fileNum, Alpha)
 
    End With

End Sub

Sub DrawGrhIndex(ByVal grhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Z As Single, ByVal Center As Byte, Optional ByVal color As Long = -1, Optional ByVal Angle As Integer = 0)

    With GrhData(grhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                    
        Dim source As wGL_Rectangle, destination As wGL_Rectangle
    
        source.X1 = .S0
        source.Y1 = .T0
        source.X2 = .S1
        source.Y2 = .T1
        destination.X1 = X
        destination.Y1 = Y
        destination.X2 = X + .pixelWidth
        destination.Y2 = Y + .pixelHeight
        
        Call Draw(destination, source, Z, Angle, color, .fileNum, False)
 
    End With
  
End Sub
