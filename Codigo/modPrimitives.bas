Attribute VB_Name = "modPrimitives"
'@Folder("WorldEditor.Modules.Render")
Option Explicit

Private primitiveTecnique As Integer

Public Sub InitPrimitivesModule()
    primitiveTecnique = wGL_Graphic_Renderer.Create_Technique

    Call wGL_Graphic_Renderer.Update_Technique_Program(primitiveTecnique, wGL_Graphic.Create_Program(LoadBytes("Shader\Effect.vs"), LoadBytes("Shader\Shader1.fs")))

    Dim Descriptor As wGL_Graphic_Descriptor
    
    Descriptor.Depth = COMPARISON_LESS_EQUAL
    Descriptor.Mask_Red = True: Descriptor.Mask_Green = True: Descriptor.Mask_Blue = True: Descriptor.Mask_Alpha = True
    
    Descriptor.Blend_Color_Source = BLEND_FACTOR_SRC_ALPHA
    Descriptor.Blend_Color_Destination = BLEND_FACTOR_ONE
    
    Descriptor.Depth_Mask = False
    Call wGL_Graphic_Renderer.Update_Technique_Descriptor(primitiveTecnique, Descriptor)
    
End Sub
Public Sub DrawBox(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, color As Long)
                    Dim source As wGL_Rectangle, destination As wGL_Rectangle
    
        destination.X1 = X1
        destination.Y1 = Y1
        destination.X2 = X2
        destination.Y2 = Y2
        Call wGL_Graphic_Renderer.Draw(destination, source, GetDepth(8), 0, color, 0, primitiveTecnique)
End Sub
