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
