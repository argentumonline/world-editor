Attribute VB_Name = "modDirectDraw"
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

''
' modDirectDraw
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Option Explicit

Public bTecho       As Boolean 'hay techo?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Function LoadWavetoDSBuffer(ByRef DS As DirectSound, ByRef DSB As DirectSoundBuffer, ByRef sfile As String) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
    
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(sfile, bufferDesc, waveFormat)
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    LoadWavetoDSBuffer = True
End Function

Function DeInitTileEngine() As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
Dim loopc As Integer

'****** Clear DirectX objects ******

Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tx As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tx = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tx < MinXBorder Or tx > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tx
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tx As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tx = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With CharList(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
    Dim RangeX As Single, RangeY As Single
        Call GetCharacterDimension(CharIndex, RangeX, RangeY)
        Call g_Swarm.Insert(5, CharIndex, X, Y, RangeX, RangeY)
    bRefreshRadar = True ' GS
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With CharList(CharIndex)
        .Active = 0
        
        .Moving = 0
        .Pos.X = 0
        .Pos.Y = 0
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    If CharIndex = 0 Then Exit Sub
    
    CharList(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    Call g_Swarm.Remove(5, CharIndex, 0, 0, 0, 0)
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1

    bRefreshRadar = True ' GS
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal grhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.grhIndex = grhIndex
    
    If grhIndex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.grhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.grhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.grhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = X + addX
        nY = Y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
        Call g_Swarm.Move(CharIndex, nX, nY)
    End With
    
    'areas viejos
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    If CharIndex <> UserCharIndex Then
    '        Call EraseChar(CharIndex)
    '    End If
    'End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
    End With

    bRefreshRadar = True ' GS
End Sub

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    
    loopc = 1
    Do While (CharList(loopc).Active = 1) And (loopc <= UBound(CharList))
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If Not InMapLegalBounds(X, Y) Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, Y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, Y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function

Function InMapLegalBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If (X < MinXBorder) Or (X > MaxXBorder) Or (Y < MinYBorder) Or (Y > MaxYBorder) Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Public Sub DDrawGrhtoSurface(ByRef Surface As DirectDrawSurface7, ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Grh.grhIndex = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.grhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhIndex).Frames(Grh.FrameCounter)
    
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
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Surface.BltFast(X, Y, SurfaceDB.Surface(.fileNum), SourceRect, DDBLTFAST_WAIT)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurri� un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripci�n del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Public Sub DDrawTransGrhIndextoSurface(ByRef Surface As DirectDrawSurface7, ByVal grhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte)
    Dim SourceRect As RECT
    
    If grhIndex = 0 Then Exit Sub
    
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
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Surface.BltFast(X, Y, SurfaceDB.Surface(.fileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
End Sub

Public Sub DDrawTransGrhtoSurface(ByRef Surface As DirectDrawSurface7, ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    Dim ddsdDest As DDSURFACEDESC2
    
On Error GoTo error

    If Grh.grhIndex = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.grhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhIndex).Frames(Grh.FrameCounter)
    
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
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Surface.BltFast(X, Y, SurfaceDB.Surface(.fileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurri� un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripci�n del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub DrawBackBufferSurface()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Sub DrawGrhtoHdc(ByVal hDC As Long, ByVal grhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    'Call SurfaceDB.Surface(GrhData(grhIndex).fileNum).BltToDC(hDC, SourceRect, destRect)
End Sub

Sub PlayWaveDS(ByRef File As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), File) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub

Sub RenderScreenOld(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 by GS
'*************************************************
Dim Y           As Long     'Keeps track of where on map we are
Dim X           As Long     'Keeps track of where on map we are
Dim ScreenMinY  As Integer  'Start Y pos on current screen
Dim ScreenMaxY  As Integer  'End Y pos on current screen
Dim ScreenMinX  As Integer  'Start X pos on current screen
Dim ScreenMaxX  As Integer  'End X pos on current screen
Dim MinY        As Integer  'Start Y pos on current map
Dim MaxY        As Integer  'End Y pos on current map
Dim MinX        As Integer  'Start X pos on current map
Dim MaxX        As Integer  'End X pos on current map
Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
Dim ScreenXOffset   As Integer
Dim ScreenYOffset   As Integer
Dim minXOffset  As Integer
Dim minYOffset  As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim Grh         As Grh      'Temp Grh for show tile and blocked
                    
    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth
    
    MinY = ScreenMinY - TileBufferSize
    MaxY = ScreenMaxY + TileBufferSize
    MinX = ScreenMinX - TileBufferSize
    MaxX = ScreenMaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < YMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize
    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize
    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If ScreenMinY > YMinMapSize Then
        ScreenMinY = ScreenMinY - 1
    Else
        ScreenYOffset = (YMinMapSize - ScreenMinY) + 1
        ScreenMinY = YMinMapSize
    End If
    
    If ScreenMaxY < YMaxMapSize Then
        ScreenMaxY = ScreenMaxY + 1
    ElseIf ScreenMaxY > YMaxMapSize Then
        ScreenMaxY = YMaxMapSize
    End If
    
    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenXOffset = (XMinMapSize - ScreenMinX) + 1
        ScreenMinX = XMinMapSize
    End If
    
    If ScreenMaxX < XMaxMapSize Then
        ScreenMaxX = ScreenMaxX + 1
    ElseIf ScreenMaxX > XMaxMapSize Then
        ScreenMaxX = XMaxMapSize
    End If
    
    'Draw floor layer
    ScreenY = ScreenYOffset
    For Y = ScreenMinY To ScreenMaxY
        ScreenX = ScreenXOffset
        For X = ScreenMinX To ScreenMaxX
            
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).grhIndex <> 0 Then
                Call DDrawGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(1), _
                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                    0, 1)
            End If
                
            If bSelectSup Then
                If CurLayer = 1 Then
                    If X = SobreX And Y = SobreY Then
                        If MosaicoChecked Then
                            Call DDrawGrhtoSurface(BackBufferSurface, CurrentGrh(((X + DespX) Mod mAncho) + 1, ((Y + DespY) Mod MAlto) + 1), _
                                (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                                (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                                0, 1)
                        Else
                            Call DDrawGrhtoSurface(BackBufferSurface, CurrentGrh(0), _
                                (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                                (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                                0, 1)
                        End If
                    End If
                End If
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next X
        
        'Increment ScreenY
        ScreenY = ScreenY + 1
    Next Y
        
    If bVerCapa(2) Then
        'Draw floor layer 2
        ScreenY = minYOffset
        For Y = MinY To MaxY
            ScreenX = minXOffset
            For X = MinX To MaxX
                
                'Layer 2 **********************************
                If MapData(X, Y).Graphic(2).grhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(2), _
                            (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                            (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                            1, 1)
                End If
                
                If bSelectSup Then
                    If CurLayer = 2 Then
                        If (X = SobreX) And (Y = SobreY) Then
                            If MosaicoChecked Then
                                Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(((X + DespX) Mod mAncho) + 1, ((Y + DespY) Mod MAlto) + 1), _
                                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                    1, 1)
                            Else
                                Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(0), _
                                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                    1, 1)
                            End If
                        End If
                    End If
                End If
                '******************************************
                
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    Else
        If bSelectSup Then
            If CurLayer = 2 Then
                X = SobreX
                Y = SobreY
                ScreenX = (X - MinX) + minXOffset
                ScreenY = (Y - MinY) + minYOffset
                
                If MosaicoChecked Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(((X + DespX) Mod mAncho) + 1, ((Y + DespY) Mod MAlto) + 1), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
                Else
                    Call DDrawTransGrhtoSurface(BackBufferSurface, CurrentGrh(0), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
                End If
            End If
        End If
    End If
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For Y = MinY To MaxY
        ScreenX = minXOffset
        For X = MinX To MaxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            With MapData(X, Y)
                'Object Layer **********************************
                If (.ObjGrh.grhIndex <> 0) And bVerObjetos Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '***********************************************
                
                
                'Char layer ************************************
                If (.CharIndex <> 0) And bVerNpcs Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If (.Graphic(3).grhIndex <> 0) And bVerCapa(3) Then
                    'Draw
                    Call DDrawTransGrhtoSurface(BackBufferSurface, .Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '************************************************
            End With
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    Grh.FrameCounter = 1
    Grh.Started = 0
    
    If bVerCapa(4) Then
        'Draw layer 4
        ScreenY = minYOffset
        For Y = MinY To MaxY
            ScreenX = minXOffset
            For X = MinX To MaxX
                With MapData(X, Y)
                    'Layer 4 **********************************
                    If .Graphic(4).grhIndex <> 0 Then
                        'Draw
                        Call DDrawTransGrhtoSurface(BackBufferSurface, .Graphic(4), _
                            (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                            (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                            1, 1)
                    End If
                    '**********************************
                End With
                
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    'Draw trans, bloqs, triggers and select tiles
    ScreenY = ScreenYOffset
    For Y = ScreenMinY To ScreenMaxY
        ScreenX = ScreenXOffset
        For X = ScreenMinX To ScreenMaxX
            With MapData(X, Y)
                PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX
                PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY
                
                '**********************************
                If (.TileExit.Map <> 0) And bTranslados Then
                    Grh.grhIndex = 3
                    
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 0)
                End If
                
                'Show blocked tiles
                If (.Blocked = 1) And bBloqs Then
                    Grh.grhIndex = 4
                    
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 0)
                    
                    'BackBufferSurface.SetFillColor vbRed
                    
                    'Call BackBufferSurface.DrawBox( _
                        PixelOffsetXTemp + TilePixelWidth \ 2, _
                        PixelOffsetYTemp + TilePixelHeight \ 2, _
                        (PixelOffsetXTemp + 5) + TilePixelWidth \ 2, _
                        (PixelOffsetYTemp + 5) + TilePixelHeight \ 2)
                End If
                
                If bTriggers Then
                    Call TextDrawer.AddText(PixelOffsetXTemp + TilePixelWidth \ 2, PixelOffsetYTemp + TilePixelHeight \ 2, vbRed, str(.Trigger), True)
                End If
                
                If .Select Then
                    BackBufferSurface.SetForeColor vbGreen
                    BackBufferSurface.SetFillStyle 1
                    BackBufferSurface.DrawBox PixelOffsetXTemp, PixelOffsetYTemp, PixelOffsetXTemp + TilePixelWidth, PixelOffsetYTemp + TilePixelHeight
                End If
                '******************************************
            
                ScreenX = ScreenX + 1
            End With
        Next X
        
        'Increment ScreenY
        ScreenY = ScreenY + 1
    Next Y
    
    Dim DC As Long
    
    DC = BackBufferSurface.GetDC
    
    Call TextDrawer.DrawTextToDC(DC)
    Call BackBufferSurface.ReleaseDC(DC)
End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 16/09/2010 (Zama)
'Draw char's to screen without offcentering them
'16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    
    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        'Draw Body
        If .Body.Walk(.Heading).grhIndex Then _
            Call DDrawTransGrhtoSurface(BackBufferSurface, .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
    
        'Draw Head
        If .Head.Head(.Heading).grhIndex Then _
            Call DDrawTransGrhtoSurface(BackBufferSurface, .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
    End With
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long)
    If LenB(strText) > 0 Then
        'TextDrawer.DrawText lngXPos - 2, lngYPos - 1, strText, vbBlack, BackBufferSurface
        TextDrawer.DrawText lngXPos, lngYPos, strText, lngColor, BackBufferSurface
    End If
End Sub

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
