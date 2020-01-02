VERSION 5.00
Object = "{97FD4A65-A045-4F5C-8C6C-262505F7C013}#6.0#0"; "Argentum.ocx"
Begin VB.Form frmPalett 
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1036
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll 
      Height          =   8235
      LargeChange     =   32
      Left            =   13020
      SmallChange     =   16
      TabIndex        =   1
      Top             =   30
      Width           =   405
   End
   Begin ArgentumOCX.MyPicture pic 
      CausesValidation=   0   'False
      Height          =   8625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   15214
   End
End
Attribute VB_Name = "frmPalett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("WorldEditor.Form.Tools")
Option Explicit
Private device As Integer
Private Type GrhMap
    grhIndex As Integer
    Name As String
    RECT As RECT
End Type

Private Map() As GrhMap
Private selected As Integer
Private OffsetY As Integer
Private totalHeigth As Integer

Private Sub Form_Initialize()
    ReDim Map(UBound(SupData))
    
    Dim i As Integer

    For i = 0 To UBound(SupData)
       Map(i).grhIndex = SupData(i).Grh
       Map(i).Name = SupData(i).Name
    Next i
    
    Call GenerateMap

    device = wGL_Graphic.Create_Device_From_Display(pic.hwnd, pic.ScaleWidth * 3, pic.ScaleHeight * 3)
    Invalidate pic.hwnd
End Sub

Private Sub GenerateMap()
 Dim drawX As Integer
    Dim drawY As Integer
    Dim bestY As Integer
    Dim i As Integer
    
    For i = 0 To UBound(Map)
        With GrhData(Map(i).grhIndex)
            
            Map(i).RECT.Left = drawX
            Map(i).RECT.Top = drawY
            Map(i).RECT.Right = drawX + .pixelWidth
            Map(i).RECT.Bottom = drawY + .pixelHeight
            
            
            If .pixelHeight > bestY Then
                bestY = .pixelHeight
            End If
            drawX = drawX + .pixelWidth
            If drawX > pic.ScaleWidth Then
                drawX = 0
                drawY = drawY + bestY
                bestY = 0
            End If
        End With
    Next
    
    totalHeigth = drawY
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub Form_Resize()
    pic.Width = Me.ScaleWidth - Me.VScroll.Width
    pic.Height = Me.ScaleHeight
    VScroll.Left = pic.Width
    VScroll.Height = Me.ScaleHeight
    
    Call GenerateMap
    VScroll.max = totalHeigth
    VScroll.LargeChange = pic.Height - 32
    Pic_Paint
    
End Sub

Private Sub Form_Terminate()
   Call wGL_Graphic.Destroy_Device(device)
End Sub

Private Sub pic_DblClick()
    ReDim Map(UBound(SupData))
    
    Dim i As Integer

    Dim max As Integer

    For i = 0 To UBound(SupData)
        With SupData(i)
            If InStr(1, .Name, "piso", vbTextCompare) Then
                Map(max).grhIndex = SupData(i).Grh
                Map(max).Name = SupData(i).Name
                max = max + 1
            End If
        End With
    Next i
    
    If max = 0 Then
        max = 1
    End If
    
    ReDim Preserve Map(max - 1)
    selected = 0
    Call GenerateMap
    Pic_Paint
    selected = 0
    VScroll.max = totalHeigth
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
VScroll.SetFocus
    Dim i As Integer
    Y = Y - OffsetY
    For i = 0 To UBound(Map)
        With Map(i)
            If X > .RECT.Left And X < .RECT.Right And Y > .RECT.Top And Y < .RECT.Bottom Then
                If selected <> i Then
                    selected = i
                    Me.Caption = .Name & "(" & .grhIndex & ")"
                    Pic_Paint
                End If
                Exit Sub
            End If
            
        End With
    Next
End Sub
Private Sub Pic_Paint()
    Dim drawX As Integer
    Dim drawY As Integer
    Dim bestY As Integer
    Dim i As Integer

    
    Call wGL_Graphic.Use_Device(device)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, &H0, 1#, 0)
    Call wGL_Graphic_Renderer.Update_Projection(&H0, pic.ScaleWidth, pic.ScaleHeight)
    
    For i = 0 To UBound(Map)
        With Map(i)
            Call DrawGrhIndex(.grhIndex, .RECT.Left, .RECT.Top + OffsetY, -1#, 0)
        End With
    Next
    
    With Map(selected)
        Call modPrimitives.DrawBox(.RECT.Left, .RECT.Top + OffsetY, .RECT.Right, .RECT.Bottom + OffsetY, &H60FFFFFF)
    End With
    

    Call wGL_Graphic_Renderer.Flush
End Sub

Private Sub VScroll_Change()
    OffsetY = -VScroll.Value
    Pic_Paint
End Sub
