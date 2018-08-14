Attribute VB_Name = "PathXORStack"
Option Explicit

Public Type XORStackElement
    x0 As Single
    y0 As Single
    x1 As Single
    y1 As Single
    color As Long
    dw As Integer
End Type

Public Const XORStackSize = 200

Public Type XORStackType
    Enabled As Boolean
    ElementCount As Long
    Element() As XORStackElement
End Type

Public XORStack As XORStackType

Public Sub OpenXORStack()
    ReDim XORStack.Element(XORStackSize)
    
    XORStack.Enabled = True
    XORStack.ElementCount = 0
    
    FrmMain.PicPath.DrawMode = 7
End Sub

Public Sub PushXORStack(x0 As Single, y0 As Single, x1 As Single, y1 As Single, color As Long, dw As Integer)
    XORStack.ElementCount = XORStack.ElementCount + 1
    If XORStack.ElementCount > XORStackSize Then
        ReDim Preserve XORStack.Element(XORStack.ElementCount)
    End If
    
    XORStack.Element(XORStack.ElementCount).x0 = x0
    XORStack.Element(XORStack.ElementCount).y0 = y0
    XORStack.Element(XORStack.ElementCount).x1 = x1
    XORStack.Element(XORStack.ElementCount).y1 = y1
    XORStack.Element(XORStack.ElementCount).color = color 'or xor(color)
    XORStack.Element(XORStack.ElementCount).dw = dw
End Sub

Public Sub PopAllXORStack()
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single, color As Long, dw As Integer
    Dim i As Long
    
    If XORStack.Enabled = True Then
        FrmMain.PicPath.DrawMode = 7
        For i = XORStack.ElementCount To 1 Step -1
            x0 = XORStack.Element(i).x0
            y0 = XORStack.Element(i).y0
            x1 = XORStack.Element(i).x1
            y1 = XORStack.Element(i).y1
            color = XORStack.Element(i).color
            dw = XORStack.Element(i).dw
            
            If x0 <> -99999 Then
                LLine x0, y0, x1, y1, color, dw
            Else
                FrmMain.PicPath.PSet (x1, y1), color
            End If
        Next
        
        XORStack.ElementCount = 0
    Else
        OpenXORStack
    End If
End Sub

Public Sub CloseXORStack()
    ReDim XORStack.Element(0)
    
    XORStack.Enabled = False
    XORStack.ElementCount = 0
    
    FrmMain.PicPath.DrawMode = 13
End Sub
