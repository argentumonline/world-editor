Attribute VB_Name = "modRender"
'@Folder("WorldEditor.Modules.Render")
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

Public Type MapExportOptions
   Width As Integer
   Height As Integer
   
   floor As Boolean
   layer2 As Boolean
   layer3 As Boolean
   layer4 As Boolean
   objects As Boolean
   npcs As Boolean
   format As eFormatPic

End Type

Public Sub RenderAllMaps(ByRef format As eFormatPic, ByVal SizeX As Long, ByVal SizeY As Long)
'*************************************************
'Author: Anagrama
'Last modified:12/08/2016
'12/08/2016: Anagrama - Genera una captura de cada mapa en la carpeta de mapas.
'*************************************************
    Dim FileCount As String
    Dim file() As String
    Dim FilePath As String
    Dim Extension As String
    Dim num As Integer
    Dim NumFiles As Integer
    
    FilePath = App.path & "\Mapas\"
    Extension = "*.map"
    
    FileCount = Dir$(FilePath & Extension)
    Do While Len(FileCount)
        NumFiles = NumFiles + 1
        ReDim Preserve file(1 To NumFiles) As String
        file(UBound(file)) = FileCount
        FileCount = Dir$
    Loop
    
    frmRenderAll.pgbProgressTotal.Value = 0
    frmRenderAll.pgbProgressTotal.max = NumFiles
    frmRenderAll.lblEstadoTotal = "0/" & NumFiles
    
    For num = 1 To UBound(file)
        Call modMapIO.NuevoMapa
        modMapIO.AbrirMapa FilePath & file(num), MapData
        Call MapCapture(format, SizeX, SizeY, 1)
        frmRenderAll.pgbProgressTotal.Value = frmRenderAll.pgbProgressTotal.Value + 1
        frmRenderAll.lblEstadoTotal = frmRenderAll.pgbProgressTotal.Value & "/" & NumFiles
    Next num
    
End Sub

Public Sub MapCapture(ByRef format As eFormatPic, ByVal SizeX As Long, ByVal SizeY As Long, Optional ByVal RenderAll As Byte = 0)

End Sub
