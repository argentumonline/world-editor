Attribute VB_Name = "modPaneles"
'@Folder("WorldEditor.Modules")
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
' modPaneles
'
' @remarks Funciones referentes a los Paneles de Funcion
' @author gshaxor@gmail.com
' @version 0.3.28
' @date 20060530

Option Explicit
Public Enum PanelsTypes
    Surfaces
    Exits
    Blocks
    NPC
    NPCHostile
    Objects
    Triggers
    Count
End Enum


Private Type PanelType
    panelFrame As Frame
    btn As lvButtons_H
    usePreview As Boolean
End Type

Private device As Integer
Private Pic As MyPicture
Private currentPanel As PanelsTypes
Private panels(PanelsTypes.Count) As PanelType

Private Sub SetFrameAs(ByRef f As Frame, pType As PanelsTypes, btn As lvButtons_H, usePreview As Boolean)
    With panels(pType)
        Set .panelFrame = f
        Set .btn = btn
        .usePreview = usePreview
    End With
End Sub

Public Sub InitPanelModule(holder As MyPicture)
    Dim i As Integer
    
    Set Pic = holder
    device = wGL_Graphic.Create_Device_From_Display(Pic.hwnd, Pic.ScaleWidth, Pic.ScaleHeight)
        
    Call Invalidate(Pic.hwnd)
    
    Call SetFrameAs(frmMain.frameSurface, Surfaces, frmMain.SelectPanel(Surfaces), True)
    Call SetFrameAs(frmMain.frameBlock, Blocks, frmMain.SelectPanel(Blocks), False)
    Call SetFrameAs(frmMain.frameExit, Exits, frmMain.SelectPanel(Exits), False)
    Call SetFrameAs(frmMain.frameNpc, NPC, frmMain.SelectPanel(NPC), True)
    Call SetFrameAs(frmMain.frameNpcH, NPCHostile, frmMain.SelectPanel(NPCHostile), True)
    Call SetFrameAs(frmMain.frameObject, Objects, frmMain.SelectPanel(Objects), True)
    Call SetFrameAs(frmMain.frameTriggers, Triggers, frmMain.SelectPanel(Triggers), True)
    
    For i = 0 To PanelsTypes.Count - 1
         panels(i).panelFrame.Visible = False
    Next
End Sub
Public Sub SetPanel(selectedPanel As PanelsTypes)
    With panels(currentPanel)
        .panelFrame.Visible = False
        .btn.Value = False
    End With
    With panels(selectedPanel)
        .panelFrame.Visible = True
        .btn.Value = True
         frmMain.PreviewGrh.Visible = .usePreview
         If .usePreview Then
            frmMain.StatTxt.Top = (frmMain.PreviewGrh.Top + frmMain.PreviewGrh.Height) + 2
            frmMain.StatTxt.Height = (frmMain.ScaleHeight - frmMain.StatTxt.Top) - 2
        Else
            frmMain.StatTxt.Top = frmMain.PreviewGrh.Top + 2
            frmMain.StatTxt.Height = (frmMain.ScaleHeight - frmMain.StatTxt.Top) - 2
         End If
    End With
    currentPanel = selectedPanel
End Sub

Public Sub DestroyPanelModule()
    wGL_Graphic.Destroy_Device (device)
End Sub

Public Sub Render()
    Call wGL_Graphic.Use_Device(device)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, &H0, 1#, 0)
    Call wGL_Graphic_Renderer.Update_Projection(&H0, Pic.ScaleWidth, Pic.ScaleHeight)
    
    
    If MosaicoChecked Then
        Dim X As Integer, Y As Integer
        For X = 1 To mAncho
            For Y = 1 To MAlto
                If CurrentGrh(X, Y).grhIndex Then
                    With GrhData(CurrentGrh(X, Y).grhIndex)
                        'TODO fix for grh size
                        Call DrawGrhIndex(.Frames(1), (X * 32) - 32, (Y * 32) - 32, -1#, 0)
                    End With
                End If
            Next Y
        Next X
    Else
        If CurrentGrh(0).grhIndex > 0 Then
            With GrhData(CurrentGrh(0).grhIndex)
                'Call DrawGrhIndex(.Frames(1), 0, 0, -1#, 0)
                Call DrawGrhIndexWithLimit(.Frames(1), 0, 0, -1#, Pic.ScaleWidth, Pic.ScaleHeight)
            End With
        End If
    End If

    Call wGL_Graphic_Renderer.Flush
End Sub
''
' Muestra los controles que componen a la funcion seleccionada del Panel
'
' @param Numero Especifica el numero de Funcion
' @param Ver Especifica si se va a ver o no
' @param Normal Inidica que ahi que volver todo No visible

''
' Filtra del Listado de Elementos de una Funcion
'
' @param Numero Indica la funcion a Filtrar

Public Sub Filtrar(ByVal Numero As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

    Dim vDatos As String
    Dim i As Long
    Dim Filtro As String
    
    If frmMain.cFiltro(Numero).ListCount > 5 Then
        frmMain.cFiltro(Numero).RemoveItem 0
    End If
    
    frmMain.cFiltro(Numero).AddItem frmMain.cFiltro(Numero).Text
    frmMain.lListado(Numero).Clear
        
    Filtro = frmMain.cFiltro(Numero).Text
    
    Select Case Numero
        Case 0 ' superficie
            For i = 0 To MaxSup
                vDatos = SupData(i).name
                
                If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                    frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                End If
            Next i
            
        Case 1 ' NPCs
            For i = 1 To NumNPCs
                If Not NpcData(i).Hostile Then
                    vDatos = NpcData(i).name
                    
                    If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                        frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                    End If
                End If
            Next i
        Case 2 ' NPCs Hostiles
            For i = 1 To NumNPCs
                If NpcData(i).Hostile Then
                    vDatos = NpcData(i).name
                    
                    If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                        frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                    End If
                End If
            Next i
            
        Case 3 ' Objetos
            For i = 1 To NumOBJs
                vDatos = ObjData(i).name
                
                If (LenB(Filtro) = 0) Or (InStr(1, UCase$(vDatos), UCase$(Filtro))) Then
                    frmMain.lListado(Numero).AddItem vDatos & " - #" & i
                End If
            Next i
    End Select
End Sub

Public Function DameGrhIndex(ByVal GrhIn As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

DameGrhIndex = SupData(GrhIn).Grh
End Function

Public Sub ActualizarMosaico()
If MosaicoChecked Then
    mAncho = Val(frmConfigSup.mAncho)
    MAlto = Val(frmConfigSup.mLargo)
    
    ReDim CurrentGrh(1 To mAncho, 1 To MAlto) As Grh
Else
    ReDim CurrentGrh(0) As Grh
End If

Call fPreviewGrh(frmMain.cGrh.Text)
Call Render
End Sub

Public Sub fPreviewGrh(ByVal GrhIn As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 22/05/06
'*************************************************
Dim X As Byte
Dim Y As Byte

If Val(GrhIn) < 1 Then
    frmMain.cGrh.Text = UBound(GrhData)
    Exit Sub
End If

If Val(GrhIn) > UBound(GrhData) Then
    frmMain.cGrh.Text = 1
    Exit Sub
End If

If MosaicoChecked Then
    For Y = 1 To MAlto
        For X = 1 To mAncho
            'Change CurrentGrh
            If Not fullyBlack(GrhIn) Then
                InitGrh CurrentGrh(X, Y), GrhIn
            Else
                InitGrh CurrentGrh(X, Y), 0
            End If
            
            GrhIn = GrhIn + 1
        Next X
    Next Y
Else
    If Not fullyBlack(GrhIn) Then
        InitGrh CurrentGrh(0), GrhIn
    Else
        InitGrh CurrentGrh(0), 0
    End If
End If
End Sub
