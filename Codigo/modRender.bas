Attribute VB_Name = "modRender"
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
' modRender
'
' @author Torres Patricio (Pato)
' @version 1.0.0
' @date 20110312

Option Explicit

Public Const FullMapRenderX As Integer = 3200
Public Const FullMapRenderY As Integer = 3200

Public Enum eFormatPic
    bmp
    png
    jpg
End Enum

Public Sub RenderAllMaps(ByRef format As eFormatPic, ByVal SizeX As Long, ByVal SizeY As Long)
'*************************************************
'Author: Anagrama
'Last modified:12/08/2016
'12/08/2016: Anagrama - Genera una captura de cada mapa en la carpeta de mapas.
'*************************************************
    Dim FileCount As String
    Dim File() As String
    Dim FilePath As String
    Dim Extension As String
    Dim num As Integer
    Dim NumFiles As Integer
    
    FilePath = App.path & "\Mapas\"
    Extension = "*.map"
    
    FileCount = Dir$(FilePath & Extension)
    Do While Len(FileCount)
        NumFiles = NumFiles + 1
        ReDim Preserve File(1 To NumFiles) As String
        File(UBound(File)) = FileCount
        FileCount = Dir$
    Loop
    
    frmRenderAll.pgbProgressTotal.Value = 0
    frmRenderAll.pgbProgressTotal.max = NumFiles
    frmRenderAll.lblEstadoTotal = "0/" & NumFiles
    
    For num = 1 To UBound(File)
        Call modMapIO.NuevoMapa
        modMapIO.AbrirMapa FilePath & File(num), MapData
        Call MapCapture(format, SizeX, SizeY, 1)
        frmRenderAll.pgbProgressTotal.Value = frmRenderAll.pgbProgressTotal.Value + 1
        frmRenderAll.lblEstadoTotal = frmRenderAll.pgbProgressTotal.Value & "/" & NumFiles
    Next num
    
End Sub

Public Sub MapCapture(ByRef format As eFormatPic, ByVal SizeX As Long, ByVal SizeY As Long, Optional ByVal RenderAll As Byte = 0)
'*************************************************
'Author: Torres Patricio(Pato)
'Last modified:12/03/11
'12/08/2016: Anagrama - Modificado para generar tamaños inferiores sin distorcionarse.
'                       Cambiado el nombre de la carpeta destino de Screenshots a Renders.
'                       Ahora guarda el nombre del archivo en vez del nombre del mapa.
'                       Agregada distincion al capturar 1 o todos los mapas.
'*************************************************
Dim Y           As Long     'Keeps track of where on map we are
Dim X           As Long     'Keeps track of where on map we are
Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
Dim ScreenXOffset   As Integer
Dim ScreenYOffset   As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim Grh         As Grh      'Temp Grh for show tile and blocked
Dim renderSurface As DirectDrawSurface7
Dim surfaceDesc As DDSURFACEDESC2
Dim srcRect As RECT
Dim destRect As RECT
Dim MyMinX As Byte
Dim MyMaxX As Byte
Dim MyMinY As Byte
Dim MyMaxY As Byte

    With surfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        If ClientSetup.bUseVideo Then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End If
        .lHeight = 3200 '32(Tamaño del pixel)*100(Ancho en pixeles)*100(Alto en pixeles)
        .lWidth = 3200
        
        Set renderSurface = DirectDraw.CreateSurface(surfaceDesc)
    End With

    With srcRect
        .Right = 3200
        .Bottom = 3200
    End With

    If RenderAll = 0 Then
        frmRender.pgbProgress.Value = 0
        frmRender.pgbProgress.max = 50000
        MyMinX = XMinMapSize
        MyMaxX = XMaxMapSize
        MyMinY = YMinMapSize
        MyMaxY = YMaxMapSize
    Else
        frmRenderAll.pgbProgress.Value = 0
        frmRenderAll.pgbProgress.max = 5
        MyMinX = 9
        MyMaxX = 92
        MyMinY = 7
        MyMaxY = 94
        srcRect.Bottom = 87 * 32
        srcRect.Right = 83 * 32
    End If

    Call renderSurface.BltColorFill(srcRect, 0)
    
    If RenderAll = 1 Then
        frmRenderAll.pgbProgress.Value = frmRenderAll.pgbProgress.Value + 1
        frmRenderAll.lblEstado.Caption = "Renderizado de primer capa - 20%"
    End If
    'Draw floor layer
    For Y = MyMinY To MyMaxY
        For X = MyMinX To MyMaxX
            If RenderAll = 0 Then
                'Layer 1 **********************************
                If MapData(X, Y).Graphic(1).grhIndex <> 0 Then
                    Call DDrawGrhtoSurface(renderSurface, MapData(X, Y).Graphic(1), _
                        (X - 1) * TilePixelWidth, _
                        (Y - 1) * TilePixelHeight, _
                        0, 1)
                End If
                '******************************************
                frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1
                frmRender.lblEstado.Caption = "Renderizado de primer capa - " & (Y - 1) + (X / 100) & "%"
            Else
                'Layer 1 **********************************
                If MapData(X, Y).Graphic(1).grhIndex <> 0 Then
                    Call DDrawGrhtoSurface(renderSurface, MapData(X, Y).Graphic(1), _
                        (X - 9) * TilePixelWidth, _
                        (Y - 7) * TilePixelHeight, _
                        0, 1)
                End If
                '******************************************
            End If
            DoEvents
        Next X
    Next Y

    If RenderAll = 1 Then
        frmRenderAll.pgbProgress.Value = frmRenderAll.pgbProgress.Value + 1
        frmRenderAll.lblEstado.Caption = "Renderizado de segunda capa - 40%"
    End If
    
    'Draw floor layer 2
    For Y = MyMinY To MyMaxY
        For X = MyMinX To MyMaxX
            If RenderAll = 0 Then
                'Layer 2 **********************************
                If (MapData(X, Y).Graphic(2).grhIndex <> 0) And bVerCapa(2) Then
                    Call DDrawTransGrhtoSurface(renderSurface, MapData(X, Y).Graphic(2), _
                            (X - 1) * TilePixelWidth, _
                            (Y - 1) * TilePixelHeight, _
                            1, 1)
                End If
                '******************************************
                frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1
                frmRender.lblEstado = "Renderizado de segunda capa - " & (Y - 1) + (X / 100) & "%"
            Else
                'Layer 2 **********************************
                If (MapData(X, Y).Graphic(2).grhIndex <> 0) And bVerCapa(2) Then
                    Call DDrawTransGrhtoSurface(renderSurface, MapData(X, Y).Graphic(2), _
                            (X - 9) * TilePixelWidth, _
                            (Y - 7) * TilePixelHeight, _
                            1, 1)
                End If
                '******************************************
            End If
            DoEvents
        Next X
    Next Y
    
    If RenderAll = 1 Then
        frmRenderAll.pgbProgress.Value = frmRenderAll.pgbProgress.Value + 1
        frmRenderAll.lblEstado.Caption = "Renderizado de objetos y tercera capa - 60%"
    End If
    
    'Draw Transparent Layers
    For Y = MyMinY To MyMaxY
        For X = MyMinX To MyMaxX
            If RenderAll = 0 Then
                PixelOffsetXTemp = (X - 1) * TilePixelWidth
                PixelOffsetYTemp = (Y - 1) * TilePixelHeight
            Else
                PixelOffsetXTemp = (X - 9) * TilePixelWidth
                PixelOffsetYTemp = (Y - 7) * TilePixelHeight
            End If
            
            With MapData(X, Y)
                'Object Layer **********************************
                If (.ObjGrh.grhIndex <> 0) And bVerObjetos Then
                    Call DDrawTransGrhtoSurface(renderSurface, .ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '***********************************************
                
                'Layer 3 *****************************************
                If (.Graphic(3).grhIndex <> 0) And bVerCapa(3) Then
                    'Draw
                    Call DDrawTransGrhtoSurface(renderSurface, .Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '************************************************
                
                If RenderAll = 0 Then
                    frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1
                    frmRender.lblEstado.Caption = "Renderizado de objetos y tercer capa - " & (Y - 1) + (X / 100) & "%"
                End If
                DoEvents
            End With
        Next X
    Next Y
    
    Grh.FrameCounter = 1
    Grh.Started = 0
    
    If RenderAll = 1 Then
        frmRenderAll.pgbProgress.Value = frmRenderAll.pgbProgress.Value + 1
        frmRenderAll.lblEstado.Caption = "Renderizado de cuarta capa - 80%"
    End If
    
    'Draw layer 4
    For Y = MyMinY To MyMaxY
        For X = MyMinX To MyMaxX
            With MapData(X, Y)
                If RenderAll = 0 Then
                    'Layer 4 **********************************
                    If (.Graphic(4).grhIndex <> 0) And bVerCapa(4) Then
                        'Draw
                        Call DDrawTransGrhtoSurface(renderSurface, .Graphic(4), _
                            (X - 1) * TilePixelWidth, _
                            (Y - 1) * TilePixelHeight, _
                            1, 1)
                    End If
                    '**********************************
                    frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1
                    frmRender.lblEstado.Caption = "Renderizado de cuarta capa - " & (Y - 1) + (X / 100) & "%"
                Else
                    'Layer 4 **********************************
                    If (.Graphic(4).grhIndex <> 0) And bVerCapa(4) Then
                        'Draw
                        Call DDrawTransGrhtoSurface(renderSurface, .Graphic(4), _
                            (X - 9) * TilePixelWidth, _
                            (Y - 7) * TilePixelHeight, _
                            1, 1)
                    End If
                    '**********************************
                End If
                DoEvents
            End With
        Next X
    Next Y
    
    If RenderAll = 0 Then
        'Draw trans, bloqs, triggers and select tiles
        For Y = MyMinY To MyMaxY
            For X = MyMinX To MyMaxX
                With MapData(X, Y)
                    PixelOffsetXTemp = (X - 1) * TilePixelWidth
                    PixelOffsetYTemp = (Y - 1) * TilePixelHeight
                    
                    '**********************************
                    If (.TileExit.Map <> 0) And bTranslados Then
                        Grh.grhIndex = 3
                        
                        Call DDrawTransGrhtoSurface(renderSurface, Grh, _
                            PixelOffsetXTemp, _
                            PixelOffsetYTemp, _
                            1, 0)
                    End If
                    
                    'Show blocked tiles
                    If (.Blocked = 1) And bBloqs Then
                        Grh.grhIndex = 4
                        
                        Call DDrawTransGrhtoSurface(renderSurface, Grh, _
                            PixelOffsetXTemp, _
                            PixelOffsetYTemp, _
                            1, 0)
                    End If
                    '******************************************
                    
                    frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1
                    frmRender.lblEstado.Caption = "Renderizado de translados y bloqueos - " & (Y - 1) + (X / 100) & "%"
                    DoEvents
                End With
            Next X
        Next Y
    End If

    destRect.Right = srcRect.Right
    destRect.Bottom = srcRect.Bottom
    
    frmRenderAll.tmpPic.Width = srcRect.Right
    frmRenderAll.tmpPic.Height = srcRect.Bottom

    frmRenderAll.picMap.Width = SizeX
    frmRenderAll.picMap.Height = SizeY

    Call renderSurface.BltToDC(frmRenderAll.tmpPic.hDC, srcRect, destRect)

    frmRenderAll.tmpPic.Picture = frmRenderAll.tmpPic.image
    
    Dim token As Long
    token = InitGDIPlus
    frmRenderAll.picMap = Resize(frmRenderAll.tmpPic.Picture.handle, frmRenderAll.tmpPic.Picture.type, frmRenderAll.picMap.ScaleWidth, frmRenderAll.picMap.ScaleHeight, , False)
    FreeGDIPlus token

    If Not FileExist(App.path & "\Renders", vbDirectory) Then MkDir (App.path & "\Renders")
    
    Select Case format
        Case eFormatPic.bmp
            Call SavePicture(frmRenderAll.picMap.image, App.path & "\Renders\" & NumMap_Save & ".bmp")
            
        Case eFormatPic.png
            Call StartUpGDIPlus(GdiPlusVersion)
            Call SavePictureAsPNG(frmRenderAll.picMap.Picture, App.path & "\Renders\" & NumMap_Save & ".png")
            Call ShutdownGDIPlus
            
        Case eFormatPic.jpg
            Call StartUpGDIPlus(GdiPlusVersion)
            Call SavePictureAsJPG(frmRenderAll.picMap.Picture, App.path & "\Renders\" & NumMap_Save & ".jpg")
            Call ShutdownGDIPlus
    End Select
End Sub
