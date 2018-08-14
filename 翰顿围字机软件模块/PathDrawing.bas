Attribute VB_Name = "PathDrawing"
Option Explicit

'Public Const WorkingMode = 1  '0-Lines & Points, 1-Points Only
'Public Const ToolMax = 17 - WorkingMode * 10
'Public Const LayerMax = 7 - WorkingMode * 2
Public Const Tool_Max = 23
Public Const Layer_Max = 7

Public ToolMax As Integer
Public LayerMax As Integer

Public ColorMode As Integer
Public ColorMode1 As Integer

Enum ToolType
    None
    MoveCanvas
    SelectByBox
    SetPoint
    SetSegment
    SetBox
    SetCircle
    SetCircle_3p
    SetEllipse
    SetSPLine
    BreakSegment
    BreakSegment_2
    RoundCornerByPoint
    RoundCornerByPoint_2
    EditElement
    MoveElement
    MoveElement_Point
    RotateElement
    MirrorElement
    CopyElement
    DeleteElement
    DeleteElement_Point
    ConnectTwoPoints
    Reverse
    ZoomIn
    ZoomOut
    SetReferenceLine
    SetReferencePoint
    StartPoint
    StopPoint
    MakeHole1
    MakeHole2
    MakeHole3
    PieceArray
    PieceArrange
    Unit
    Seperate
    ConvertToSegments
    MeasureDistance
    MeasureScale
    
    CreateCircle
    CreateArc
    CreateSector
    CreateEllipse
    CreateSquare
    CreatePolygon
    CreateCurvePolygon
    CreateTrapezoid
    CreateTrapezoid1
    CreateTriangle
    CreateTriangle1
    CreateTriangle2
    CreateRectangle
    CreateParallelRectangle
    Create5PointStar
    CreateMultiPointStar
End Enum

Enum CatchPointMode
    Normal
    Alone
    NotStartPoint
    NotEndPoint
    NotArcPoint
    NotArcCenter
    OutputStartPoint
    OutputEndPoint
End Enum

Enum CatchedPointType
    AlonePoint
    SegmentStartPoint
    SegmentInnerPoint
    SegmentEndPoint
    ArcStartPoint
    ArcEndPoint
End Enum

Enum CatchedElementType
    ElementNone
    ElementPoint
    ElementSegment
    ElementArc
    ElementSPLine
    
    PointAlone
    PointWithSegment0
    PointWithSegment1
    PointWithBothSegments
    PointOfArcPoint0
    PointOfArcPoint1
    PointOfArcPointm
    PointOfSPLinePoint0
    PointOfSPLinePoint1
    PointOfSPLinePointm
    
    ArcCenterPoint
    ArcEdge
End Enum

Enum ShowPositionMode
    OnlyStautsBar
    OnlyControlBar
    BothStatusAndControlBar
End Enum

Public Type Title_Value
    t As String
    v As Double
End Type

Public Const Pi = 3.14159265358979
Public Const PI2 = Pi * 2
Public Const PI_2 = Pi / 2
Public Const PI_180 = Pi / 180

Public PixelWidth As Double

Public ScrollFactorXY As Integer
Public ScrollFactorZ As Integer

Public ViewMinX As Double
Public ViewMaxX As Double
Public ViewMinY As Double
Public ViewMaxY As Double
Public ViewMargin As Double

Public MainGridX As Double
Public MainGridY As Double
Public SubGridX As Double
Public SubGridY As Double

Public AuxLineEnabled As Boolean
Public AuxLineVisible As Boolean
Public AuxXLineCount As Integer
Public AuxYLineCount As Integer
Public AuxXLine() As Double
Public AuxYLine() As Double

Public UserMinX As Double
Public UserMaxX As Double
Public UserMinY As Double
Public UserMaxY As Double
Public UserMinZ As Double
Public UserMaxZ As Double

Public UserOrgX As Double
Public UserOrgY As Double
Public UserOrgZ As Double

Public PathMinX As Integer
Public PathMaxX As Integer
Public PathMinY As Integer
Public PathMaxY As Integer

Public ZoomFactor As Double
Public ShiftX As Double
Public ShiftY As Double

Public CurTool As Integer
Public CurToolStep As Single

Public PointSize As Integer
Public TrapWidth As Integer, utw As Double
Public HVTrapWidth As Integer
Public CornerR As Double
Public ArcStepFactor As Integer
Public MinPathStep As Double

Public CanShowCursorReferenceLines As Boolean

Public ToolName(Tool_Max) As String
Public ToolTask(Tool_Max) As Integer

Public NormalColor As Long
Public HighLightColor As Long
Public XORColor As Long

Public LayerColor(Layer_Max) As Long
Public LayerZValue(Layer_Max) As Double
Public CurLayer As Long

Public ShowPoints As Boolean
Public ShowDirection As Boolean

Public KeepPlatformData As Boolean

Public TempSPline As Path_SPLine
Public CurOutputPointIndex As Long
Public DataChanged As Boolean
Public DirectionChanged As Boolean

Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Sub InitParameter()
    Dim d As Integer
    
    ViewMinX = 0
    ViewMaxX = Device_UserSize(1)
    ViewMinY = 0
    ViewMaxY = Device_UserSize(2)
    ViewMargin = 0.03
    
    If KeepPlatformData = True Then
        Exit Sub
    End If
    KeepPlatformData = True
    
    UserOrgX = 0
    UserOrgY = 0
    UserOrgZ = 0
    
    If ViewMaxX > 4000 Or ViewMaxY > 4000 Then
        d = 10
    ElseIf ViewMaxX > 2000 Or ViewMaxY > 2000 Then
        d = 5
    ElseIf ViewMaxX > 1000 Or ViewMaxY > 1000 Then
        d = 2
    Else
        d = 1
    End If

    MainGridX = 100 * d
    MainGridY = 100 * d
    SubGridX = 5 * d
    SubGridY = 5 * d
    
    PointSize = 5 'pixels
    TrapWidth = 8
    HVTrapWidth = 12
    
    CornerR = 20
    ArcStepFactor = 100
    MinPathStep = 1
    SPLine_SegmentBetweenPoints = 20
        
    ReadUserParameter
    
    AuxLineEnabled = True
    AuxLineVisible = True
    
End Sub

Sub ShowToolBox(ByVal left As Integer, ByVal top As Integer)
    Dim ToolIconWidth As Integer, ToolIconHeight As Integer
    Dim i As Integer
    
    If Device_Mode = 0 Then
        ToolName(0) = "平移"
        ToolTask(0) = ToolType.MoveCanvas
        ToolName(1) = "样条"
        ToolTask(1) = ToolType.SetSPLine '.SelectByBox
        ToolName(2) = "设点"
        ToolTask(2) = ToolType.SetPoint
        ToolName(3) = "直线"
        ToolTask(3) = ToolType.SetSegment
        ToolName(4) = "圆CR"
        ToolTask(4) = ToolType.SetCircle
        ToolName(5) = "圆3P"
        ToolTask(5) = ToolType.SetCircle_3p
        ToolName(6) = "椭圆"
        ToolTask(6) = ToolType.SetEllipse
        ToolName(7) = "矩形"
        ToolTask(7) = ToolType.SetBox
        ToolName(8) = "圆角"
        ToolTask(8) = ToolType.RoundCornerByPoint
        ToolName(9) = "倒角"
        ToolTask(9) = ToolType.RoundCornerByPoint_2
        ToolName(10) = "设断点"
        ToolTask(10) = ToolType.BreakSegment
        ToolName(11) = "设中点"
        ToolTask(11) = ToolType.BreakSegment_2
        ToolName(12) = "编辑"
        ToolTask(12) = ToolType.EditElement '.ConnectTwoPoints
        ToolName(13) = "移动"
        ToolTask(13) = ToolType.MoveElement
        'ToolName(14) = "旋转"
        'ToolTask(14) = ToolType.RotateElement
        ToolName(14) = "镜像"
        ToolTask(14) = ToolType.MirrorElement
        ToolName(15) = "复制"
        ToolTask(15) = ToolType.CopyElement '.ConnectTwoPoints
        ToolName(16) = "删点"
        ToolTask(16) = ToolType.DeleteElement_Point
        ToolName(17) = "删线"
        ToolTask(17) = ToolType.DeleteElement
        ToolName(18) = "线段化"
        ToolTask(18) = ToolType.ConvertToSegments
        'ToolName(19) = "设起点"
        'ToolTask(19) = ToolType.StartPoint
        ToolName(19) = "反向"
        ToolTask(19) = ToolType.Reverse
        'ToolName(21) = "桥位"
        'ToolTask(21) = ToolType.MakeHole1
        'ToolName(22) = "直切"
        'ToolTask(22) = ToolType.MakeHole2
        'ToolName(23) = "鹰嘴"
        'ToolTask(23) = ToolType.MakeHole3
        ToolName(20) = "放大"
        ToolTask(20) = ToolType.ZoomIn
        ToolName(21) = "缩小"
        ToolTask(21) = ToolType.ZoomOut
        ToolName(22) = "测距离"
        ToolTask(22) = ToolType.MeasureDistance
        ToolName(23) = "测范围"
        ToolTask(23) = ToolType.MeasureScale
        
        'ToolName(18) = "辅线"
        'ToolTask(18) = ToolType.SetReferenceLine
        'ToolName(19) = "辅点"
        'ToolTask(19) = ToolType.SetReferencePoint
    End If
    
    ToolIconWidth = 60
    ToolIconHeight = 24
    
    FrmMain.FraToolBox.left = left
    FrmMain.FraToolBox.top = top
    FrmMain.FraToolBox.Width = 2 * ToolIconWidth + 7
    FrmMain.FraToolBox.Height = (ToolIconHeight + 2) * ((ToolMax + 1) / 2) + 11
    
    On Error Resume Next
    
    For i = 0 To ToolMax
        If i > 0 Then
            Load FrmMain.CmdTool(i)
            FrmMain.CmdTool(i).Visible = True
        End If
        
        'frmmain.CmdTool(i).Picture = frmmain.ImgLstTool.ListImages(i + 1).Picture
        FrmMain.CmdTool(i).caption = ToolName(i)
        If i Mod 2 = 0 Then
            FrmMain.CmdTool(i).Move 3 * Screen.TwipsPerPixelX, (9 + Int(i / 2) * (ToolIconHeight + 2)) * Screen.TwipsPerPixelY, ToolIconWidth * Screen.TwipsPerPixelX, ToolIconHeight * Screen.TwipsPerPixelY
        Else
            FrmMain.CmdTool(i).Move (3 + ToolIconWidth) * Screen.TwipsPerPixelX, (9 + Int(i / 2) * (ToolIconHeight + 2)) * Screen.TwipsPerPixelY, ToolIconWidth * Screen.TwipsPerPixelX, ToolIconHeight * Screen.TwipsPerPixelY
        End If
    Next
    For i = ToolMax + 1 To Tool_Max
        FrmMain.CmdTool(i).Visible = False
    Next
    
    FrmMain.FraEdit.left = left
    FrmMain.FraEdit.top = FrmMain.FraToolBox.top + FrmMain.FraToolBox.Height + 5
    FrmMain.FraEdit.Width = FrmMain.FraToolBox.Width
    FrmMain.FraEdit.Visible = False
End Sub

Sub ShowEditData(caption As String, n As Long, t_v() As Title_Value, ByVal tag_val As Long)
    Dim i As Long, tx As Double, ty As Double, w As Double, w1 As Double, l As Double
    
    tx = Screen.TwipsPerPixelX
    ty = Screen.TwipsPerPixelY
    
    For i = FrmMain.TxtEdit.Count - 1 To 1 Step -1
        Unload FrmMain.TxtEdit(i)
        Unload FrmMain.LblEdit(i)
    Next
    
    FrmMain.LblEdit(0).AutoSize = True
    FrmMain.LblEdit(0).Alignment = 2
    
    w = 0
    For i = 0 To n - 1
        FrmMain.LblEdit(0).caption = t_v(i).t
        w1 = FrmMain.LblEdit(0).Width
        If w1 > w Then w = w1
    Next
    l = Max(36 * tx, w + 10 * tx)
    
    For i = 1 To n - 1
        Load FrmMain.LblEdit(i)
        Load FrmMain.TxtEdit(i)
    Next
    
    For i = 0 To n - 1
        FrmMain.LblEdit(i).Move 7 * tx, (i * 22 + 22) * ty
        FrmMain.TxtEdit(i).Move l, (i * 22 + 22) * ty, 82 * tx - l
        
        FrmMain.LblEdit(i).caption = t_v(i).t
        FrmMain.TxtEdit(i).Text = str(t_v(i).v)
        
        FrmMain.LblEdit(i).Visible = True
        FrmMain.TxtEdit(i).Visible = True
    Next
    
    If tag_val = ToolType.PieceArray Then
        FrmMain.ChkEdit(0).Move 9 * tx, ((n + 1) * 22 + 10) * ty
        n = n + 2
        FrmMain.ChkEdit(0).Visible = True
        FrmMain.CmdEdit.caption = "建立阵列"
        
    ElseIf tag_val >= ToolType.CreateCircle Then
        FrmMain.ChkEdit(0).Visible = False
        FrmMain.CmdEdit.caption = "创建"
        FrmMain.CmdEdit.Enabled = True
        
    Else
        FrmMain.ChkEdit(0).Visible = False
        FrmMain.CmdEdit.caption = "数据更新"
        FrmMain.CmdEdit.Enabled = True
    End If
    FrmMain.CmdEdit.Move 10 * tx, (n * 22 + 30) * ty
    FrmMain.FraEdit.Height = (n * 22 + 62)
    FrmMain.FraEdit.caption = caption
    FrmMain.FraEdit.Visible = True
    FrmMain.TxtEdit(0).SetFocus
    FrmMain.TxtEdit(0).SelStart = Len(FrmMain.TxtEdit(0).Text)
    
    FrmMain.FraEdit.Tag = ""
    FrmMain.CmdEdit.Tag = str(tag_val)
End Sub

Sub SnapUXUY(ByVal ux0 As Double, ByVal uy0 As Double, ByRef ux As Double, ByRef uy As Double)
    Dim d As Double, i As Integer, X As Double, Y As Double, dX As Double, dy As Double, kx As Integer, ky As Integer
    
    ux = ux0
    kx = 0
    uy = uy0
    ky = 0
    
    If AuxLineEnabled = True Then
    
        d = TrapWidth * (UserMaxX - UserMinX) / PathMaxX
        
        If AuxXLineCount > 0 Then
            For i = 0 To AuxXLineCount - 1
                If i = 0 Then
                    dX = Abs(ux0 - AuxXLine(i))
                    ux = AuxXLine(i)
                ElseIf Abs(ux0 - AuxXLine(i)) < dX Then
                    dX = Abs(ux0 - AuxXLine(i))
                    ux = AuxXLine(i)
                End If
            Next
            If dX > d Then
                ux = ux0
            Else
                kx = 1
            End If
        End If
        
        If AuxYLineCount > 0 Then
            For i = 0 To AuxYLineCount - 1
                If i = 0 Then
                    dy = Abs(uy0 - AuxYLine(i))
                    uy = AuxYLine(i)
                ElseIf Abs(uy0 - AuxYLine(i)) < dy Then
                    dy = Abs(uy0 - AuxYLine(i))
                    uy = AuxYLine(i)
                End If
            Next
            If dy > d Then
                uy = uy0
            Else
                ky = 1
            End If
        End If
    End If
        
    'If FrmMain.ChkSnapGrid.Value = 0 Then
    If FrmMain.Toolbar1.Buttons(20).value = tbrUnpressed Then
        If PixelWidth > 0.25 Then
            d = 2
        ElseIf PixelWidth > 0.2 Then
            d = 4
        ElseIf PixelWidth > 0.1 Then
            d = 5
        ElseIf PixelWidth > 0.05 Then
            d = 10
        ElseIf PixelWidth > 0.025 Then
            d = 20
        ElseIf PixelWidth > 0.02 Then
            d = 40
        Else
            d = 50
        End If
        
        If kx = 0 Then ux = Round(d * ux0, 0) / d
        If ky = 0 Then uy = Round(d * uy0, 0) / d
    Else
        If kx = 0 And SubGridX > 0 Then
            For X = ViewMinX To ViewMaxX Step SubGridX
                If X = ViewMinX Then
                    dX = Abs(ux0 - X)
                    ux = X
                ElseIf Abs(ux0 - X) < dX Then
                    dX = Abs(ux0 - X)
                    ux = X
                End If
            Next
            If dX > SubGridX / 2 Then
                ux = ux0
            End If
        End If
        
        If ky = 0 And SubGridY > 0 Then
            For Y = ViewMinY To ViewMaxY Step SubGridY
                If Y = ViewMinY Then
                    dy = Abs(uy0 - Y)
                    uy = Y
                ElseIf Abs(uy0 - Y) < dy Then
                    dy = Abs(uy0 - Y)
                    uy = Y
                End If
            Next
            If dy > SubGridY / 2 Then
                uy = uy0
            End If
        End If
    End If
End Sub

Sub ConvertPathToUser(ByVal X As Integer, ByVal Y As Integer, ux As Double, uy As Double)
    Dim ux0 As Double, uy0 As Double, x1 As Double, y1 As Double
    
    x1 = (X - ShiftX) / ZoomFactor
    y1 = (Y - ShiftY) / ZoomFactor
    
    If Device_CoordinateMode = 0 Then
        ux0 = x1 / (1# * PathMaxX) * (UserMaxX - UserMinX) + UserMinX
        uy0 = y1 / (1# * PathMaxY) * (UserMaxY - UserMinY) + UserMinY
    Else
        ux0 = y1 / (1# * PathMaxX) * (UserMaxX - UserMinX) + UserMinX
        uy0 = x1 / (1# * PathMaxY) * (UserMaxY - UserMinY) + UserMinY
    End If
    
    'SnapUXUY ux0, uy0, ux, uy
    ux = ux0
    uy = uy0
End Sub

Sub ConvertUserToPath(ByVal ux As Double, ByVal uy As Double, X As Single, Y As Single)
    If Device_CoordinateMode = 0 Then
        X = (ux - UserMinX) / (UserMaxX - UserMinX) * PathMaxX
        Y = (uy - UserMinY) / (UserMaxY - UserMinY) * PathMaxY
    Else
        Y = (ux - UserMinX) / (UserMaxX - UserMinX) * PathMaxX
        X = (uy - UserMinY) / (UserMaxY - UserMinY) * PathMaxY
    End If
    
    X = X * ZoomFactor + ShiftX
    Y = Y * ZoomFactor + ShiftY
End Sub


Sub SetUserScale(ByVal x_width As Double, ByVal y_height As Double, ByVal Margin As Double)
    Dim d As Double
    
    If x_width > Device_UserSize(1) Then
        x_width = Device_UserSize(1)
    End If
    
    If y_height > Device_UserSize(2) Then
        y_height = Device_UserSize(2)
    End If
    
    ViewMinX = 0
    ViewMaxX = x_width
    ViewMinY = 0
    ViewMaxY = y_height
    ViewMargin = Margin
    
    d = (PathMaxY + 1) / (PathMaxX + 1)
    If Device_CoordinateMode = 0 Then
        If y_height / x_width < d Then
            UserMinX = -Margin * x_width
            UserMaxX = x_width + Margin * x_width
            UserMinY = -((UserMaxX - UserMinX) * d - ViewMaxY) / 2
            UserMaxY = (UserMaxX - UserMinX) * d + UserMinY
        Else
            UserMinY = -Margin * y_height
            UserMaxY = y_height + Margin * y_height
            UserMinX = -((UserMaxY - UserMinY) / d - ViewMaxX) / 2
            UserMaxX = (UserMaxY - UserMinY) / d + UserMinX
        End If
    Else
        If x_width / y_height > d Then
            UserMinX = -Margin * x_width
            UserMaxX = (x_width * (1 + 2 * Margin)) / d + UserMinX
            UserMinY = -((UserMaxX - UserMinX) * d - ViewMaxY * d) / 2
            UserMaxY = (UserMaxX - UserMinX) * d + UserMinY
        Else
            UserMinY = -Margin * y_height
            UserMaxY = (y_height * (1 + 2 * Margin)) * d + UserMinY
            UserMinX = -((UserMaxY - UserMinY) / d - ViewMaxX / d) / 2
            UserMaxX = (UserMaxY - UserMinY) / d + UserMinX
        End If
    End If
    UserMaxZ = Device_UserSize(3)
        
    If ViewMaxX < 160 Then
        ScrollFactorXY = 200
    ElseIf ViewMaxX < 320 Then
        ScrollFactorXY = 100
    ElseIf ViewMaxX < 640 Then
        ScrollFactorXY = 50
    ElseIf ViewMaxX < 800 Then
        ScrollFactorXY = 40
    ElseIf ViewMaxX < 1600 Then
        ScrollFactorXY = 20
    ElseIf ViewMaxX < 3200 Then
        ScrollFactorXY = 10
    Else
        ScrollFactorXY = 1
    End If
    
    If Device_UserSize(3) < 160 Then
        ScrollFactorZ = 200
    ElseIf Device_UserSize(3) < 320 Then
        ScrollFactorZ = 100
    ElseIf Device_UserSize(3) < 640 Then
        ScrollFactorZ = 50
    ElseIf Device_UserSize(3) < 800 Then
        ScrollFactorZ = 40
    ElseIf Device_UserSize(3) < 1600 Then
        ScrollFactorZ = 20
    ElseIf Device_UserSize(3) < 3200 Then
        ScrollFactorZ = 10
    Else
        ScrollFactorZ = 1
    End If
    
    PixelWidth = (UserMaxX - UserMinX) / (FrmMain.PicPath.Width - 2)
End Sub

Sub LLine(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal clr As Long, ByVal dw As Integer)
    Dim dX As Single, dy As Single, m As Integer, i As Integer, xm As Single, ym As Single
    
    FrmMain.PicPath.DrawWidth = dw
    dX = x1 - x0
    dy = y1 - y0
    If Abs(dX) > 1500 Or Abs(dy) > 1500 Then 'VB has an error if segment too long
        If Abs(dX) > Abs(dy) Then
            m = Int(Abs(dX) / 1500) + 1
        Else
            m = Int(Abs(dy) / 1500) + 1
        End If
            
        dX = dX / m
        dy = dy / m
        
        For i = 1 To m - 1
            xm = x0 + dX
            ym = y0 + dy
            FrmMain.PicPath.Line (x0, y0)-(xm, ym), clr
            x0 = xm
            y0 = ym
        Next
        FrmMain.PicPath.Line (x0, y0)-(x1, y1), clr
    Else
        FrmMain.PicPath.Line (x0, y0)-(x1, y1), clr
    End If
End Sub

Sub LineOut(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal clr As Long, Optional dw As Integer = 1)
    If XORStack.Enabled = True Then
        FrmMain.PicPath.DrawMode = 7
    Else
        FrmMain.PicPath.DrawMode = 13
    End If
    
    LLine x0, y0, x1, y1, clr, dw
'    FrmMain.PicPath.PSet (x1, y1), RGB(255, 255, 255)
            
    FrmMain.PicPath.DrawWidth = 1
    
    If XORStack.Enabled = True Then
        PushXORStack x0, y0, x1, y1, clr, dw
'        PushXORStack -99999, -99999, x1, y1, RGB(255, 255, 255), dw
    End If
End Sub

Sub DrawGridLines()
    Dim ux As Double, uy As Double, x1 As Single, y1 As Single, x2 As Single, y2 As Single, d As Single, i As Integer
    Dim clr1 As Long, clr2 As Long, clr3 As Long, clr4 As Long
    
    On Error Resume Next
    
    If ColorMode = 0 Then
        clr1 = RGB(60, 180, 60)
        clr2 = RGB(240, 240, 240)
        clr3 = RGB(200, 200, 200)
        clr4 = RGB(0, 180, 180)
    Else
        clr1 = RGB(50, 100, 50)
        clr2 = RGB(20, 20, 20)
        clr3 = RGB(50, 50, 50)
        clr4 = RGB(0, 100, 100)
    End If
        
'    ConvertUserToPath -0.023 * UserMaxX, -0.023 * UserMaxX, x1, y1
'    ConvertUserToPath -0.01 * UserMaxX, -0.01 * UserMaxX, x2, y2
'    d = Abs(y2 - y1)
'
'    LineOut x1, y1, x1, y1 + 4 * d, clr1
'    LineOut x2, y1, x2, y1 + 4 * d, clr1
'    LineOut (x1 + x2) / 2, y1 + 5 * d, x1, y1 + 4 * d, clr1
'    LineOut (x1 + x2) / 2, y1 + 5 * d, x2, y1 + 4 * d, clr1
'
 '   LineOut x1, y1, x1 + 4 * d, y1, clr1
'    LineOut x1, y2, x1 + 4 * d, y2, clr1
'    LineOut x1 + 4 * d, y1, x1 + 5 * d, (y1 + y2) / 2, clr1
'    LineOut x1 + 4 * d, y2, x1 + 5 * d, (y1 + y2) / 2, clr1
'
'    If Device_CoordinateMode = 0 Then
'        FrmMain.PicPath.ForeColor = clr1
 '       FrmMain.PicPath.CurrentX = (x1 + x2) / 2 - 3
 '       FrmMain.PicPath.CurrentY = y1 + 4 * d
'        FrmMain.PicPath.Print "Y"
'        FrmMain.PicPath.CurrentX = x1 + 4 * d - 5
'        FrmMain.PicPath.CurrentY = (y1 + y2) / 2 + 5
'        FrmMain.PicPath.Print "X"
'    Else
'        FrmMain.PicPath.ForeColor = clr1
'        FrmMain.PicPath.CurrentX = (x1 + x2) / 2 - 3
'        FrmMain.PicPath.CurrentY = y1 + 4 * d - 10
'        FrmMain.PicPath.Print "X"
'        FrmMain.PicPath.CurrentX = x1 + 4 * d - 5
'        FrmMain.PicPath.CurrentY = (y1 + y2) / 2 - 5
'        FrmMain.PicPath.Print "Y"
'    End If
    
    'If FrmMain.ChkShowGridLines.Value = 1 Then
    If FrmMain.Toolbar1.Buttons(16).value = tbrPressed Then
        'If SubGridX > 0 And SubGridY > 0 Then
        '    For ux = ViewMinX To ViewMaxX Step SubGridX
        '        For uy = ViewMinY To ViewMaxY Step SubGridY
        '            ConvertUserToPath ux, uy, x1, y1
        '            frmmain.PicPath.PSet (x1, y1), RGB(220, 220, 220)
        '        Next
        '    Next
        'End If
        
        If SubGridX > 0 Then
            For ux = ViewMinX To ViewMaxX Step SubGridX
                ConvertUserToPath ux, ViewMinY, x1, y1
                ConvertUserToPath ux, ViewMaxY, x2, y2
                LineOut x1, y1, x2, y2, clr2
            Next
        End If
        
        If SubGridY > 0 Then
            For uy = ViewMinY To ViewMaxY Step SubGridY
                ConvertUserToPath ViewMinX, uy, x1, y1
                ConvertUserToPath ViewMaxX, uy, x2, y2
                LineOut x1, y1, x2, y2, clr2
            Next
        End If
        
        If MainGridX > 0 Then
            For ux = ViewMinX To ViewMaxX Step MainGridX
                ConvertUserToPath ux, ViewMinY, x1, y1
                ConvertUserToPath ux, ViewMaxY, x2, y2
                LineOut x1, y1, x2, y2, clr3
            Next
        End If
        
        If MainGridY > 0 Then
            For uy = ViewMinY To ViewMaxY Step MainGridY
                ConvertUserToPath ViewMinX, uy, x1, y1
                ConvertUserToPath ViewMaxX, uy, x2, y2
                LineOut x1, y1, x2, y2, clr3
            Next
        End If
    
        ConvertUserToPath ViewMinX, ViewMinY, x1, y1
        ConvertUserToPath ViewMaxX, ViewMaxY, x2, y2
        LineOut x1, y1, x1, y2, clr3
        LineOut x1, y2, x2, y2, clr3
        LineOut x2, y2, x2, y1, clr3
        LineOut x2, y1, x1, y1, clr3
        
    End If
    
    If AuxLineVisible = True Then
        For i = 0 To AuxXLineCount - 1
            ConvertUserToPath AuxXLine(i), ViewMinY, x1, y1
            ConvertUserToPath AuxXLine(i), ViewMaxY, x2, y2
            LineOut x1, y1, x2, y2, clr4
        Next
        
        For i = 0 To AuxYLineCount - 1
            ConvertUserToPath ViewMinX, AuxYLine(i), x1, y1
            ConvertUserToPath ViewMaxX, AuxYLine(i), x2, y2
            LineOut x1, y1, x2, y2, clr4
        Next
    End If
    
    'ConvertUserToPath UserOrgX, UserOrgY, x1, y1
    'PointOut x1, y1, RGB(0, 255, 255), -1
End Sub

Sub PointOut(ByVal X As Single, ByVal Y As Single, ByVal clr As Long, Optional PointMode As ActionType = ActionType.No_Action)
    Dim d As Integer
    d = PointSize
    
    If PointMode = -1 Then 'PointMode = MotionOrg
        LineOut X, Y - d, X, Y - 2, clr
        LineOut X, Y + d, X, Y + 2, clr
        LineOut X - d, Y, X - 2, Y, clr
        LineOut X + d, Y, X + 2, Y, clr
        
        Exit Sub
    End If
    
    LineOut X, Y - d, X, Y + d + 1, clr
    LineOut X - d, Y, X + d + 1, Y, clr
    
    If PointMode = StartDropping Then
        d = d + 2
    
        LineOut X - d, Y, X, Y - d, clr
        LineOut X, Y - d, X + d, Y, clr
        LineOut X + d, Y, X, Y + d, clr
        LineOut X, Y + d, X - d, Y, clr
        
    ElseIf PointMode = StopDropping Then
        d = d

        LineOut X - d, Y - d, X - d, Y + d, clr
        LineOut X + d, Y - d, X + d, Y + d, clr
        LineOut X - d, Y - d, X + d, Y - d, clr
        LineOut X - d, Y + d, X + d + 1, Y + d, clr
        
    End If
    
    LineOut X, Y - 1, X, Y, clr 'set the point
End Sub

Sub PointOut_ForHole(ByVal X As Single, ByVal Y As Single, ByVal clr As Long, Optional Hole As HoleType)
    Dim d As Integer
    d = PointSize
    
    LineOut X, Y - d, X, Y + d + 1, clr
    LineOut X - d, Y, X + d + 1, Y, clr
    
    If Hole = HoleType.HoleType1 Then
        d = 4 * d - 1
    
        LineOut X - d / 2, Y - d, X - d / 2, Y + d, clr
        LineOut X + d / 2, Y - d, X + d / 2, Y + d, clr
        'LineOut X - d / 2, Y - d, X + d, Y - d, clr
        LineOut X - d / 2, Y + d, X + d / 2, Y + d, clr
                
    ElseIf Hole = HoleType.HoleType2 Then
        d = 4 * d - 1

        LineOut X - d / 2, Y - d, X - d / 2, Y + d, clr
        LineOut X + d / 2, Y - d, X + d / 2, Y + d, clr
        LineOut X - d / 2, Y - d, X + d / 2, Y - d, clr
        'LineOut X - d / 2, Y + d, X + d / 2, Y + d, clr
                
        
    ElseIf Hole = HoleType.HoleType3 Then
        d = 4 * d - 1

        LineOut X - d / 2, Y + d / 2, X - d / 3, Y + d, clr
        LineOut X + d / 2, Y + d / 2, X + d / 3, Y + d, clr
        
        LineOut X - d / 2, Y - d, X - d / 2, Y + d / 2, clr
        LineOut X + d / 2, Y - d, X + d / 2, Y + d / 2, clr
        LineOut X - d / 2, Y - d, X + d / 2, Y - d, clr
    End If
    
    LineOut X, Y - 1, X, Y, clr 'set the point
End Sub

Sub DrawPoint(Point As Path_Point, Optional color As Long)
    Dim ux As Double, uy As Double, clr As Long
    Dim X As Single, Y As Single, i As Long
    
    On Error Resume Next
    
    If ShowPoints = False And Point.HoleType = 0 Then
        For i = 1 To SegmentCount
            If SegmentList(i).point0_id = Point.id Or SegmentList(i).point1_id = Point.id Then '两线段的中点
                If Not (Point.action = ActionType.StartDropping Or Point.action = ActionType.StopDropping) Then
                    Exit Sub
                End If
            End If
        Next
    
        For i = 1 To ArcCount '圆弧相关的点
            If ArcList(i).point0_id = Point.id Or _
               ArcList(i).point1_id = Point.id Or _
               ArcList(i).pointm_id = Point.id Then
                Exit Sub
            End If
        Next
        
        If Point.Type = PointType.SPLinePoint Then '样条曲线上的点
            Exit Sub
        End If
    End If
    
    ux = Point.X
    uy = Point.Y
    If color = 0 Then
        If Point.color = 0 Then
            clr = LayerColor(Point.Layer)
            If clr = 0 Then
                clr = RGB(255, 0, 0)
            End If
        Else
            clr = Point.color
        End If
    Else
        clr = color
    End If
    
    If Point.action = ActionType.StartDropping Or _
       Point.action = ActionType.Dropping Or _
       Point.action = ActionType.StopDropping Then
        clr = RGB(255, 255, 127)
    End If
    
    ConvertUserToPath ux, uy, X, Y
    If Point.HoleType = 0 Then
        PointOut X, Y, clr, Point.action
    Else
        PointOut_ForHole X, Y, clr, Point.HoleType
    End If
End Sub

Sub DrawAllPoints()
    Dim id As Long
    
    For id = 1 To PointCount
        DrawPoint PointList(id)
    Next
End Sub

Sub DrawSegment(Segment As Path_Segment, Optional color As Long, Optional ShowCornerArc As Boolean = True)
    Dim ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double, clr As Long
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single, dw As Integer
    
    'Dim dx As Single, dy As Single, m As Integer, i As Integer, xm As Single, ym As Single
    
    If Segment.Type = SegmentType.ReplacedByArc Then
        DrawArc ArcList(Segment.arc_id)
        Exit Sub
    End If
    
    If PointList(Segment.point0_id).method = PointMethod.RoundedCorner Then
        ux0 = PointList(ArcList(PointList(Segment.point0_id).arc_id).point1_id).X
        uy0 = PointList(ArcList(PointList(Segment.point0_id).arc_id).point1_id).Y
    
        If ShowCornerArc Then DrawArc ArcList(PointList(Segment.point0_id).arc_id)
    Else
        ux0 = PointList(Segment.point0_id).X
        uy0 = PointList(Segment.point0_id).Y
    End If
    
    If PointList(Segment.point1_id).method = PointMethod.RoundedCorner Then
        ux1 = PointList(ArcList(PointList(Segment.point1_id).arc_id).point0_id).X
        uy1 = PointList(ArcList(PointList(Segment.point1_id).arc_id).point0_id).Y
    Else
        ux1 = PointList(Segment.point1_id).X
        uy1 = PointList(Segment.point1_id).Y
    End If
    
    If color = 0 Then
        If Segment.color = 0 Then
            clr = LayerColor(Segment.Layer)
            If clr = 0 Then
                clr = RGB(255, 0, 0)
'Debug.Print ">>>>>>>>"
            End If
        Else
            clr = Segment.color
        End If
    Else
        clr = color
    End If
    
    If PointList(Segment.point0_id).action = ActionType.StartDropping Or _
       PointList(Segment.point0_id).action = ActionType.Dropping Then
        'dw = 2
        dw = 1
        clr = RGB(255, 255, 127)
    Else
        dw = 1
    End If
    
    ConvertUserToPath ux0, uy0, x0, y0
    ConvertUserToPath ux1, uy1, x1, y1

    LineOut x0, y0, x1, y1, clr, dw
    
    '--------------------------------------------
    Dim d As Single, dX As Single, dy As Single, dx1 As Single, dy1 As Single, dx2 As Single, dy2 As Single
    Dim xm As Double, ym As Double
    
    If ShowDirection = True And ((x1 - x0) <> 0 Or (y1 - y0) <> 0) Then
        d = 8
        
        dX = d * (x1 - x0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        dy = d * (y1 - y0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        
        Rotate_Z dX, dy, 30 * PI_180, dx1, dy1
        Rotate_Z dX, dy, -30 * PI_180, dx2, dy2
        
        xm = x0 + (x1 - x0) / 2
        ym = y0 + (y1 - y0) / 2
        
        LineOut xm, ym, xm - dx1, ym - dy1, clr
        LineOut xm, ym, xm - dx2, ym - dy2, clr
    End If
End Sub

Function IsSegmentsClockwise(ByVal start_id As Long) As Long
    Dim id As Long
    
    Dim lx() As Double, ly() As Double, cur_list() As Long, n As Long, i As Long, j As Long, clr As Long
    ReDim lx(SegmentCount), ly(SegmentCount), cur_list(SegmentCount)
    
    For id = 1 To SegmentCount
        SegmentList(id).selected = False
    Next
    
    id = start_id
    i = 0
    Do
        SegmentList(id).selected = True
        lx(i + 1) = PointList(SegmentList(id).point0_id).X
        ly(i + 1) = PointList(SegmentList(id).point0_id).Y
        cur_list(i + 1) = id
        i = i + 1
            
        For j = 1 To SegmentCount
            If SegmentList(j).selected = False And SegmentList(j).point0_id = SegmentList(id).point1_id Then
                id = j
                Exit For
            End If
        Next
        If j > SegmentCount Or id = start_id Then
            Exit Do
        End If
    Loop
    
    n = i
    If n > 2 Then
        If IsPathClockwise(n, lx, ly) = True Then
            IsSegmentsClockwise = 1
        Else
            IsSegmentsClockwise = -1
        End If
    Else
        IsSegmentsClockwise = 0
    End If
    
    ReDim lx(0), ly(0), cur_list(0)
End Function

Sub DrawAllSegments()
    Dim id As Long
        
    Dim start_id As Long
    Dim lx() As Double, ly() As Double, cur_list() As Long, n As Long, i As Long, j As Long, clr As Long
    ReDim lx(SegmentCount), ly(SegmentCount), cur_list(SegmentCount)
    
    If DirectionChanged = True Then
        
        For id = 1 To SegmentCount
            SegmentList(id).selected = False
        Next
        
        start_id = 1
        
        Do While start_id <= SegmentCount
            id = start_id
            i = 0
            Do
                SegmentList(id).selected = True
                lx(i + 1) = PointList(SegmentList(id).point0_id).X
                ly(i + 1) = PointList(SegmentList(id).point0_id).Y
                cur_list(i + 1) = id
                i = i + 1
                    
                For j = 1 To SegmentCount
                    If SegmentList(j).selected = False And SegmentList(j).point0_id = SegmentList(id).point1_id Then
                        id = j
                        Exit For
                    End If
                Next
                If j > SegmentCount Or id = start_id Then
                    Exit Do
                End If
            Loop
            
            n = i
            If n > 2 Then
                If IsPathClockwise(n, lx, ly) = True Then
                    clr = RGB(255, 0, 255)
                Else
                    clr = RGB(255, 0, 0)
                End If
            Else
                clr = RGB(0, 255, 0)
            End If
            For id = 1 To n
                SegmentList(cur_list(id)).color = clr
                'DrawSegment SegmentList(cur_list(id)), clr
            Next
            
            For id = 1 To SegmentCount
                If SegmentList(id).selected = False Then
                    Exit For
                End If
            Next
            start_id = id
        Loop
    End If
    
    DirectionChanged = False
    
    For id = 1 To SegmentCount
        DrawSegment SegmentList(id), SegmentList(id).color
    Next
End Sub

Sub DrawArc(Arc As Path_Arc, Optional color As Long = 0, Optional ShowPoint As Boolean = True)
    Dim cx As Double, cy As Double, ux As Double, uy As Double, clr As Long, cclr As Long
    Dim X As Single, Y As Single, x0 As Single, y0 As Single, d As Integer
    Dim angle As Double, Angle0 As Double, Angle1 As Double, angle_step As Double, dw As Integer
    Dim n As Integer, t As Double
    Dim mode As Long, TempSegment As Path_Segment
    
    Dim CS As Double, SN As Double, ux0 As Double, uy0 As Double
    
    cx = Arc.X
    cy = Arc.Y
    
    mode = 0
    If Arc.color = -99999 Then
        clr = LayerColor(Arc.Layer)
        cclr = clr
        mode = 1
        
    ElseIf color = 0 Then
        If Arc.color = 0 Then
            clr = LayerColor(Arc.Layer)
        Else
            clr = Arc.color
        End If
        cclr = RGB(200, 200, 200)
    Else
        clr = color
        cclr = color
    End If
    
    If ShowPoints = True Then
        ConvertUserToPath cx, cy, X, Y
        d = PointSize
        
        LineOut X, Y - d, X, Y + d + 1, cclr
        LineOut X - d, Y, X + d + 1, Y, cclr
        
        LineOut X - (d - 2), Y, X, Y + (d - 2), cclr
        LineOut X, Y + (d - 2), X + (d - 2), Y, cclr
        LineOut X + (d - 2), Y, X, Y - (d - 2), cclr
        LineOut X, Y - (d - 2), X - (d - 2), Y, cclr
    End If
    
    If Arc.point0_id > 0 Then
        If PointList(Arc.point0_id).action = ActionType.StartDropping Or _
           PointList(Arc.point0_id).action = ActionType.Dropping Then
            'dw = 2
            dw = 1
            clr = RGB(255, 255, 127)
        Else
            dw = 1
        End If
    Else
        dw = 1
    End If
    
    If mode = 1 Then
        TempSegment.point0_id = Arc.point0_id
        TempSegment.point1_id = Arc.point1_id
        
        DrawSegment TempSegment, clr
        Exit Sub
    End If
    
    If Arc.a > 0 Then
        'angle0 = Arc.start_angle
        'angle1 = Arc.end_angle
        'If angle0 < angle1 Then
        '    angle_step = 6 * PI_180
        'Else
        '    angle_step = -6 * PI_180
        'End If
        '
        'For angle = angle0 To angle1 Step angle_step
        '    ux = Cos(angle) * Arc.a + cx
        '    uy = Sin(angle) * Arc.b + cy
        '    ConvertUserToPath ux, uy, x, y
        '    If angle = angle0 Then
        '        If ShowPoint Then DrawPoint PointList(Arc.point0_id), clr
        '    Else
        '        LineOut x0, y0, x, y, clr, dw
        '    End If
        '    x0 = x
        '    y0 = y
        'Next
        '
        'If Arc.point1_id > 0 Then
        '    ConvertUserToPath PointList(Arc.point1_id).x, PointList(Arc.point1_id).y, x, y
        '    LineOut x0, y0, x, y, clr, dw
        'Else
        '    ux = Cos(angle1) * Arc.a + cx
        '    uy = Sin(angle1) * Arc.b + cy
        '    ConvertUserToPath ux, uy, x, y
        '    LineOut x0, y0, x, y, clr, dw
        'End If
        
        CS = Cos(Arc.ax_angle)
        SN = Sin(Arc.ax_angle)
        
        Angle0 = Arc.start_angle
        Angle1 = Arc.end_angle
        
        If Angle0 = Angle1 Then Exit Sub
        
        t = Sqr(Arc.a / UserMaxX)
        If t < 0.2 Then t = 0.2
        
        'n = Int((ArcStepFactor * Sqr(Arc.a) / UserMaxX) * (Abs(Angle1 - Angle0) / (6 * PI_180)) + 1)
        n = Int(ArcStepFactor * t * (Abs(Angle1 - Angle0) / PI_2))
        If n < 3 And Abs(Angle1 - Angle0) > Pi Then
            n = 3
        ElseIf n < 2 Then
            n = 2
        End If
        
        angle_step = (Angle1 - Angle0) / n
        Do While Abs(2 * Max(Arc.a, Arc.b) * Sin(angle_step / 2)) < MinPathStep And n > 1
            n = n - 1
            angle_step = (Angle1 - Angle0) / n
        Loop
        
        If ShowPoint Then DrawPoint PointList(Arc.point0_id), clr
        
        If Arc.point0_id > 0 Then
            ux = PointList(Arc.point0_id).X
            uy = PointList(Arc.point0_id).Y
        Else
            'ux = Cos(Angle0) * Arc.a + cx
            'uy = Sin(Angle0) * Arc.B + cy
            
            ux0 = Cos(Angle0) * Arc.a
            uy0 = Sin(Angle0) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
        End If
        ConvertUserToPath ux, uy, x0, y0
        
        For angle = Angle0 + angle_step To Angle1 - 0.999 * angle_step Step angle_step
            'ux = Cos(angle) * Arc.a + cx
            'uy = Sin(angle) * Arc.B + cy
            
            ux0 = Cos(angle) * Arc.a
            uy0 = Sin(angle) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
            
            ConvertUserToPath ux, uy, X, Y
            LineOut x0, y0, X, Y, clr, dw
            
            x0 = X
            y0 = Y
        Next
        If Arc.point1_id > 0 Then
            ux = PointList(Arc.point1_id).X
            uy = PointList(Arc.point1_id).Y
        Else
            'ux = Cos(Angle1) * Arc.a + cx
            'uy = Sin(Angle1) * Arc.B + cy
            
            ux0 = Cos(Angle1) * Arc.a
            uy0 = Sin(Angle1) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
        End If
        ConvertUserToPath ux, uy, X, Y
        LineOut x0, y0, X, Y, clr, dw

        If Angle1 <> Angle0 + PI2 Then
            If ShowPoint Then DrawPoint PointList(Arc.point1_id), clr
        End If
        
        '--------------------------------------------
        Dim dX As Single, dy As Single, dx1 As Single, dy1 As Single, dx2 As Single, dy2 As Single
        Dim am As Double, xm As Single, ym As Single, x1 As Single, y1 As Single
        
        If ShowDirection = True Then
            d = 8
            
            am = Angle0 + (Angle1 - Angle0) / 2
            
            'ux = Cos(am) * Arc.a + cx
            'uy = Sin(am) * Arc.B + cy
            
            ux0 = Cos(am) * Arc.a
            uy0 = Sin(am) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
            
            ConvertUserToPath ux, uy, xm, ym
            
            'ux = Cos(am - Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.a
            'uy = Sin(am - Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.B
            
            ux0 = Cos(am - Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.a
            uy0 = Sin(am - Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.b
            
            ux = CS * ux0 - SN * uy0
            uy = SN * ux0 + CS * uy0
            
            ConvertUserToPath ux, uy, x0, y0
            
            'ux = Cos(am + Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.a
            'uy = Sin(am + Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.B
            
            ux0 = Cos(am + Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.a
            uy0 = Sin(am + Sgn(Angle1 - Angle0) * 10 * PI_180) * Arc.b
            
            ux = CS * ux0 - SN * uy0
            uy = SN * ux0 + CS * uy0
            
            ConvertUserToPath ux, uy, x1, y1
            
            If x1 <> x0 Or y1 <> y0 Then
                dX = d * (x1 - x0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
                dy = d * (y1 - y0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
                
                Rotate_Z dX, dy, 30 * PI_180, dx1, dy1
                Rotate_Z dX, dy, -30 * PI_180, dx2, dy2
                  
                LineOut xm, ym, xm - dx1, ym - dy1, clr
                LineOut xm, ym, xm - dx2, ym - dy2, clr
            End If
        End If
    End If
End Sub

Sub ArcPoints(Arc As Path_Arc, points() As PolygonPoint)
    Dim cx As Double, cy As Double, ux As Double, uy As Double
    'Dim X As Single, Y As Single, x0 As Single, y0 As Single, d As Integer
    Dim angle As Double, Angle0 As Double, Angle1 As Double, angle_step As Double ', dw As Integer
    Dim n As Integer, t As Double, k As Long
    
    Dim CS As Double, SN As Double, ux0 As Double, uy0 As Double
    
    cx = Arc.X
    cy = Arc.Y
        
    If Arc.a > 0 Then
        CS = Cos(Arc.ax_angle)
        SN = Sin(Arc.ax_angle)
        
        Angle0 = Arc.start_angle
        Angle1 = Arc.end_angle
        
        If Angle0 = Angle1 Then Exit Sub
        
        t = Sqr(Arc.a / UserMaxX)
        If t < 0.2 Then t = 0.2
        
        n = Int(ArcStepFactor * t * (Abs(Angle1 - Angle0) / PI_2))
        If n < 3 And Abs(Angle1 - Angle0) > Pi Then
            n = 3
        ElseIf n < 2 Then
            n = 2
        End If
        
        angle_step = (Angle1 - Angle0) / n
        
        ReDim points(n)
        
        If Arc.point0_id > 0 Then
            ux = PointList(Arc.point0_id).X
            uy = PointList(Arc.point0_id).Y
        Else
            ux0 = Cos(Angle0) * Arc.a
            uy0 = Sin(Angle0) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
        End If
        points(0).X = ux
        points(0).Y = uy
        'ConvertUserToPath ux, uy, x0, y0
        k = 0
        For angle = Angle0 + angle_step To Angle1 - 0.999 * angle_step Step angle_step
            ux0 = Cos(angle) * Arc.a
            uy0 = Sin(angle) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
            
            k = k + 1
            points(k).X = ux
            points(k).Y = uy
            'ConvertUserToPath ux, uy, X, Y
            'LineOut x0, y0, X, Y, clr, dw
            'x0 = X
            'y0 = Y
        Next
        If Arc.point1_id > 0 Then
            ux = PointList(Arc.point1_id).X
            uy = PointList(Arc.point1_id).Y
        Else
            ux0 = Cos(Angle1) * Arc.a
            uy0 = Sin(Angle1) * Arc.b
            
            ux = CS * ux0 - SN * uy0 + cx
            uy = SN * ux0 + CS * uy0 + cy
        End If
        k = k + 1
        points(k).X = ux
        points(k).Y = uy
        'ConvertUserToPath ux, uy, X, Y
        'LineOut x0, y0, X, Y, clr, dw
    End If
End Sub

Sub DrawArcStartLine(Arc As Path_Arc, Optional color As Long = 0)
    Dim cx As Double, cy As Double, ux As Double, uy As Double, clr As Long
    Dim X As Single, Y As Single, x0 As Single, y0 As Single
    Dim angle As Double
    
    cx = Arc.X
    cy = Arc.Y
    If color = 0 Then
        clr = Arc.color
    Else
        clr = color
    End If
    
    'angle = Arc.start_angle
    'ux = Cos(angle) * Arc.a + cx
    'uy = Sin(angle) * Arc.b + cy
    
    ConvertUserToPath cx, cy, x0, y0
    'ConvertUserToPath ux, uy, x, y
    
    ConvertUserToPath PointList(Arc.point0_id).X, PointList(Arc.point0_id).Y, X, Y
    LineOut x0, y0, X, Y, clr
End Sub

Sub DrawArcEndLine(Arc As Path_Arc, Optional color As Long = 0)
    Dim cx As Double, cy As Double, ux As Double, uy As Double, clr As Long
    Dim X As Single, Y As Single, x0 As Single, y0 As Single
    Dim angle As Double
    
    cx = Arc.X
    cy = Arc.Y
    If color = 0 Then
        clr = Arc.color
    Else
        clr = color
    End If
    
    'angle = Arc.end_angle
    'ux = Cos(angle) * Arc.a + cx
    'uy = Sin(angle) * Arc.b + cy
    
    ConvertUserToPath cx, cy, x0, y0
    'ConvertUserToPath ux, uy, x, y
    
    ConvertUserToPath PointList(Arc.point1_id).X, PointList(Arc.point1_id).Y, X, Y
    LineOut x0, y0, X, Y, clr
End Sub

Sub DrawAllArcs()
    Dim id As Long
    
    For id = 1 To ArcCount
        DrawArc ArcList(id)
    Next
End Sub

Sub DrawSPLine(CurSPline As Path_SPLine, Optional color As Long = 0, Optional ShowSPLinePoint As Boolean = False)
    Dim Pts() As PolygonPoint
    Dim pts1 As PolygonPoint, pts2 As PolygonPoint
    Dim x0 As Single, y0 As Single, X As Single, Y As Single
    Dim i As Long, dw As Integer, clr As Long, n As Long
    Dim ds As Double
    
    On Error Resume Next
    
    If color = 0 Then
        If CurSPline.color = 0 Then
            clr = LayerColor(CurSPline.Layer)
        Else
            clr = CurSPline.color
        End If
    Else
        clr = color
    End If
    
    If PointList(CurSPline.point0_id).action = ActionType.StartDropping Or _
       PointList(CurSPline.point0_id).action = ActionType.Dropping Then
        'dw = 2
        dw = 1
        clr = RGB(255, 255, 127)
    Else
        dw = 1
    End If
    
    n = SPLine_SegmentBetweenPoints
    Do
        SplinePoints CurSPline, Pts(), n
        
        For i = 2 To UBound(Pts)
            ds = Sqr((Pts(i).X - Pts(i - 1).X) ^ 2 + (Pts(i).Y - Pts(i - 1).Y) ^ 2)
            If ds > 0.0000001 And ds < MinPathStep Then
                If n > 1 Then
                    n = n - 1
                    Exit For
                End If
            End If
        Next
        If i > UBound(Pts) Then Exit Do
        ReDim Pts(0)
    Loop
    
    ConvertUserToPath Pts(0).X, Pts(0).Y, x0, y0
    
    For i = 1 To UBound(Pts)
        ConvertUserToPath Pts(i).X, Pts(i).Y, X, Y
        LineOut x0, y0, X, Y, clr, dw

        'PointOut X, Y, RGB(0, 255, 0), StartDropping 'Show all points

        x0 = X
        y0 = Y
    Next i
    
    If ShowSPLinePoint = True Then
        For i = 0 To CurSPline.vertex_count - 1
            DrawPoint PointList(CurSPline.vertex_id(i)), clr
        Next
    End If
    
    '--------------------------------------------
    Dim dX As Single, dy As Single, dx1 As Single, dy1 As Single, dx2 As Single, dy2 As Single
    Dim xm As Single, ym As Single, x1 As Single, y1 As Single, d As Integer
    
    On Error Resume Next
    
    If ShowDirection = True Then
        d = 8
                                
        For i = 1 To CurSPline.vertex_count - 1
            pts1 = Pts((i - 1) * (CurSPline.segment_between_points + 1) + CurSPline.segment_between_points / 2)
            pts2 = Pts((i - 1) * (CurSPline.segment_between_points + 1) + CurSPline.segment_between_points / 2 + 1)
            
            ConvertUserToPath pts1.X, pts1.Y, x0, y0
            ConvertUserToPath pts2.X, pts2.Y, x1, y1
                    
            xm = x1
            ym = y1
            
            If x1 <> x0 Or y1 <> y0 Then
                dX = d * (x1 - x0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
                dy = d * (y1 - y0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
                
                Rotate_Z dX, dy, 30 * PI_180, dx1, dy1
                Rotate_Z dX, dy, -30 * PI_180, dx2, dy2
                  
                LineOut xm, ym, xm - dx1, ym - dy1, clr
                LineOut xm, ym, xm - dx2, ym - dy2, clr
            End If
        Next
    End If
End Sub

Sub DrawAllSPlines()
    Dim id As Long
    
    For id = 1 To SPLineCount
        DrawSPLine SPLineList(id)
    Next
End Sub

Sub DrawAll(Optional show_calculation_text As Boolean = True)
    'ShowLayerIcons
    
    DrawGridLines
    If ShowPoints = True Then
        DrawAllPoints
    End If
    DrawAllSegments
    DrawAllArcs
    DrawAllSPlines
    
    'DrawLeadingLines
         
    DrawVertPoints
    FrmMain.ShowCalculation show_calculation_text
End Sub

Function CatchPoint(ByVal ux As Double, ByVal uy As Double, ByVal ud As Double, ByVal CatchMode As CatchPointMode) As Long
    Dim i As Long, j As Long, k As Long, dX As Double, dy As Double, m As Double, l As Double, id As Long, p0_id As Long, p1_id As Long, pm_id As Long
    Dim CatchedID() As Long, CatchedIDCount As Long
    
    CatchedIDCount = 0
    
    Do '当被抓取的点不合条件时，继续寻找是否有被压住的点
        id = 0
        For i = 1 To PointCount
            For j = 1 To CatchedIDCount
                If PointList(i).id = CatchedID(j) Then
                    Exit For
                End If
            Next
            If j > CatchedIDCount Then
                'If FrmMain.chkOnlyCurLayer.Value = 0 Or PointList(I).Layer = CurLayer Then
                If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or PointList(i).Layer = CurLayer Then
                    If PointList(i).status = PointStatus.Normal Then
                        dX = ux - PointList(i).X
                        dy = uy - PointList(i).Y
                        If Abs(dX) <= ud And Abs(dy) <= ud Then
                            m = dX * dX + dy * dy
                            If id = 0 Or m < l Then 'm <= l: 若两点重叠则取后点, m < l:若两点重叠则取前点
                                id = i
                                l = m
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        k = 0
        If id > 0 Then
            CatchedIDCount = CatchedIDCount + 1
            ReDim Preserve CatchedID(CatchedIDCount)
            CatchedID(CatchedIDCount) = id
            k = 1
        End If
        
        If id > 0 And CatchMode <> CatchPointMode.Normal Then
            If CatchMode <> OutputStartPoint And CatchMode <> OutputEndPoint Then
                For i = 1 To SegmentCount
                    p0_id = SegmentList(i).point0_id
                    p1_id = SegmentList(i).point1_id
            
                    If p0_id = id And (CatchMode = Alone Or CatchMode = NotStartPoint) Then
                        id = 0
                        'GoTo Exit_Sub
                        GoTo Next_Loop
                    End If
            
                    If p1_id = id And (CatchMode = Alone Or CatchMode = NotEndPoint) Then
                        id = 0
                        'GoTo Exit_Sub
                        GoTo Next_Loop
                    End If
                Next
            
                For i = 1 To ArcCount
                    p0_id = ArcList(i).point0_id
                    p1_id = ArcList(i).point1_id
                    pm_id = ArcList(i).pointm_id
            
                    If p0_id = id And (CatchMode = Alone Or CatchMode = NotStartPoint Or CatchMode = NotArcPoint) Then
                        id = 0
                        'GoTo Exit_Sub
                        GoTo Next_Loop
                    End If
            
                    If p1_id = id And (CatchMode = Alone Or CatchMode = NotEndPoint Or CatchMode = NotArcPoint) Then
                        id = 0
                        'GoTo Exit_Sub
                        GoTo Next_Loop
                    End If
                    
                    If pm_id = id And (CatchMode = Alone Or CatchMode = NotArcCenter Or CatchMode = NotArcPoint) Then
                        id = 0
                        'GoTo Exit_Sub
                        GoTo Next_Loop
                    End If
                Next
            Else
                For i = 1 To OutputStartPointList.Count
                    p0_id = OutputStartPointList.leading_point0(i).id
                    p1_id = OutputStartPointList.leading_point1(i).id
                    
                    If id = p0_id And CatchMode = OutputStartPoint Then
                        CurOutputPointIndex = i
                        GoTo Exit_Sub
                    ElseIf id = p1_id And CatchMode = OutputEndPoint Then
                        CurOutputPointIndex = i
                        GoTo Exit_Sub
                    ElseIf id = p0_id And _
                        PointList(p0_id).X = PointList(p1_id).X And _
                        PointList(p0_id).Y = PointList(p1_id).Y And _
                        CatchMode = OutputEndPoint Then
                        
                        id = p1_id
                        CurOutputPointIndex = i
                        GoTo Exit_Sub
                    End If
                Next
                id = 0
            End If
        End If
Next_Loop:
    Loop Until k = 0 Or (id > 0 And k = 1)
    
Exit_Sub:
    CatchPoint = id
End Function

Function CatchSegment(ByVal ux As Double, ByVal uy As Double, ByVal ud As Double) As Long
    Dim i As Long, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim k As Double, c As Double, uc As Double, m As Double, l As Double, id As Long
    Dim lvl0 As Integer, lvl1 As Integer
    
    id = 0
    For i = 1 To SegmentCount
        lvl0 = PointList(SegmentList(i).point0_id).Layer
        lvl1 = PointList(SegmentList(i).point1_id).Layer
        
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or (lvl0 = CurLayer And lvl1 = CurLayer) Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or (lvl0 = CurLayer And lvl1 = CurLayer) Then
            x0 = PointList(SegmentList(i).point0_id).X
            y0 = PointList(SegmentList(i).point0_id).Y
            x1 = PointList(SegmentList(i).point1_id).X
            y1 = PointList(SegmentList(i).point1_id).Y
            
            If Abs(x1 - x0) > Abs(y1 - y0) Then
                If (ux - x0) * (ux - x1) <= 0 Then
                    k = (y1 - y0) / (x1 - x0)
                    c = -x0 * k + y0
                    uc = -ux * k + uy
                    m = Abs(uc - c)
                    
                    If m <= ud Then
                        If id = 0 Or m < l Then
                            id = i
                            l = m
                        End If
                    End If
                End If
            Else
                If (uy - y0) * (uy - y1) <= 0 Then
                    k = (x1 - x0) / (y1 - y0)
                    c = -y0 * k + x0
                    uc = -uy * k + ux
                    m = Abs(uc - c)
                    
                    If m <= ud Then
                        If id = 0 Or m < l Then
                            id = i
                            l = m
                        End If
                    End If
                End If
            End If
        End If
    Next
    CatchSegment = id
End Function

Function CatchArcCenter(ByVal ux As Double, ByVal uy As Double, ByVal ud As Double) As Long
    Dim i As Long, dX As Double, dy As Double, m As Double, l As Double, id As Long
    
    id = 0
    For i = 1 To ArcCount
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or ArcList(I).Layer = CurLayer Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or ArcList(i).Layer = CurLayer Then
            dX = ux - ArcList(i).X
            dy = uy - ArcList(i).Y
            
            If Abs(dX) <= ud And Abs(dy) <= ud Then
                m = dX * dX + dy * dy
                If id = 0 Or m < l Then
                    id = i
                    l = m
                End If
            End If
        End If
    Next
    CatchArcCenter = id
End Function

Function CatchArc(ByVal ux As Double, ByVal uy As Double, ByVal ud As Double) As Long
    Dim i As Long, dX As Double, dy As Double, id As Long
    Dim a As Double, b As Double, angle As Double, Angle0 As Double, Angle1 As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim CS As Double, SN As Double, x0 As Double, y0 As Double
    Dim ux0 As Double, uy0 As Double
    
    ux0 = ux
    uy0 = uy
    
    id = 0
    For i = 1 To ArcCount
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or ArcList(I).Layer = CurLayer Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or ArcList(i).Layer = CurLayer Then
        
            'ux, uy 围绕圆心 (ArcList(i).X, ArcList(i).Y) 旋转 -ArcList(i).ax_angle 角度
            '---------------------------------------------------------------------------
            ux = ux0 - ArcList(i).X
            uy = uy0 - ArcList(i).Y
            
            CS = Cos(-ArcList(i).ax_angle)
            SN = Sin(-ArcList(i).ax_angle)
            
            x0 = ux
            y0 = uy
            ux = (CS * x0) - (SN * y0)
            uy = (SN * x0) + (CS * y0)
            
            ux = ux + ArcList(i).X
            uy = uy + ArcList(i).Y
            '---------------------------------------------------------------------------
            
            dX = ux - ArcList(i).X
            dy = uy - ArcList(i).Y
            
            a = ArcList(i).a
            b = ArcList(i).b
            
            If dX < a + ud And dX > -a - ud And dy < b + ud And dy > -b - ud Then
                If a = b Then
                    If Abs(Sqr(dX * dX + dy * dy) - a) < ud Then
                        Angle0 = ArcList(i).start_angle
                        Angle1 = ArcList(i).end_angle
                        angle = GetArcAngle(0, 0, dX, dy)
                        If (Angle1 - angle) * (angle - Angle0) > 0 Or _
                           (Angle1 - (angle + PI2)) * ((angle + PI2) - Angle0) > 0 Or _
                           (Angle1 - (angle - PI2)) * ((angle - PI2) - Angle0) > 0 Then
                            id = i
                            Exit For
                        End If
                    End If
                Else
                    If Abs(dX) > Abs(dy) Then
                        x1 = a * b * dX / Sqr(b * dX * b * dX + a * dy * a * dy)
                        x2 = -x1
                        y1 = x1 * dy / dX
                        y2 = -y1
                    Else
                        y1 = a * b * dy / Sqr(b * dX * b * dX + a * dy * a * dy)
                        y2 = -y1
                        x1 = y1 * dX / dy
                        x2 = -x1
                    End If
                    
                    If Sqr((dX - x1) * (dX - x1) + (dy - y1) * (dy - y1)) < ud Or _
                       Sqr((dX - x2) * (dX - x2) + (dy - y2) * (dy - y2)) < ud Then
                        Angle0 = ArcList(i).start_angle
                        Angle1 = ArcList(i).end_angle

                        angle = GetArcAngle(0, 0, dX, dy)
                        
                        If (Angle1 - angle) * (angle - Angle0) > 0 Or _
                           (Angle1 - (angle + PI2)) * ((angle + PI2) - Angle0) > 0 Or _
                           (Angle1 - (angle - PI2)) * ((angle - PI2) - Angle0) > 0 Then
                            id = i
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    CatchArc = id
End Function

Function CatchSPline(ByVal ux As Double, ByVal uy As Double, ByVal ud As Double) As Long
    Dim i As Long, j As Long, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim k As Double, c As Double, uc As Double, m As Double, l As Double, id As Long
    
    id = 0
    For i = 1 To SPLineCount
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or SPLineList(I).Layer = CurLayer Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or SPLineList(i).Layer = CurLayer Then
            For j = 0 To SPLineList(i).segment_point_count - 2
                x0 = SPLineList(i).segment_point(j).X
                y0 = SPLineList(i).segment_point(j).Y
                x1 = SPLineList(i).segment_point(j + 1).X
                y1 = SPLineList(i).segment_point(j + 1).Y
                
                If x1 - x0 <> 0 Or y1 - y0 <> 0 Then
                    If Abs(x1 - x0) > Abs(y1 - y0) Then
                        If (ux - x0) * (ux - x1) <= 0 Then
                            k = (y1 - y0) / (x1 - x0)
                            c = -x0 * k + y0
                            uc = -ux * k + uy
                            m = Abs(uc - c)
                            
                            If m <= ud Then
                                If id = 0 Or m < l Then
                                    id = i
                                    l = m
                                End If
                            End If
                        End If
                    Else
                        If (uy - y0) * (uy - y1) <= 0 Then
                            k = (x1 - x0) / (y1 - y0)
                            c = -y0 * k + x0
                            uc = -uy * k + ux
                            m = Abs(uc - c)
                            
                            If m <= ud Then
                                If id = 0 Or m < l Then
                                    id = i
                                    l = m
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    CatchSPline = id
End Function

Sub GetSPlinePoint(ByVal id As Long, ByVal ux As Double, ByVal uy As Double, ByRef X As Double, ByRef Y As Double, ByRef sid As Long)
    Dim j As Long, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim k As Double, c As Double, uc As Double, m As Double, l As Double
    
    l = -1
    For j = 0 To SPLineList(id).segment_point_count - 2
        x0 = SPLineList(id).segment_point(j).X
        y0 = SPLineList(id).segment_point(j).Y
        x1 = SPLineList(id).segment_point(j + 1).X
        y1 = SPLineList(id).segment_point(j + 1).Y
        
        If x1 - x0 <> 0 Or y1 - y0 <> 0 Then
            If Abs(x1 - x0) > Abs(y1 - y0) Then
                If (ux - x0) * (ux - x1) <= 0 Then
                    k = (y1 - y0) / (x1 - x0)
                    c = -x0 * k + y0
                    uc = -ux * k + uy
                    m = Abs(uc - c)
                    
                    If l = -1 Or m < l Then
                        If k = 0 Then
                            X = ux
                            Y = y0
                        Else
                            X = (uy + ux / k - c) / (k + 1 / k)
                            Y = k * X + c
                        End If
                        l = m
                        sid = j
                    End If
                End If
            Else
                If (uy - y0) * (uy - y1) <= 0 Then
                    k = (x1 - x0) / (y1 - y0)
                    c = -y0 * k + x0
                    uc = -uy * k + ux
                    m = Abs(uc - c)
                    
                    If l = -1 Or m < l Then
                        If k = 0 Then
                            Y = uy
                            X = x0
                        Else
                            Y = (ux + uy / k - c) / (k + 1 / k)
                            X = k * Y + c
                        End If
                        l = m
                        sid = j
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CatchElement(ux As Double, uy As Double, TrapWidth As Double, id As Long, ElementType As CatchedElementType, Param As Long, Param2 As Long) As CatchedElementType
    Dim i As Long, dX As Double, dy As Double, m As Double, l As Double, j As Long
    Dim seg0_id As Long, seg1_id As Long
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim k As Double, c As Double, uc As Double
    Dim lyr0 As Long, lyr1 As Long
    Dim a As Double, b As Double, angle As Double, Angle0 As Double, Angle1 As Double
    Dim x2 As Double, y2 As Double
    
    CatchElement = ElementNone
    
    '------------------------------------------------------------------------------------
    id = 0
    For i = 1 To PointCount
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or PointList(I).Layer = CurLayer Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or PointList(i).Layer = CurLayer Then
            dX = ux - PointList(i).X
            dy = uy - PointList(i).Y
            
            If Abs(dX) <= TrapWidth And Abs(dy) <= TrapWidth Then
                m = dX * dX + dy * dy
                If id = 0 Or m < l Then
                    id = i
                    l = m
                End If
            End If
        End If
    Next
    
    If id > 0 Then
        CatchElement = ElementPoint
        
        For i = 1 To SegmentCount
            If SegmentList(i).point0_id = id Then
                seg1_id = i
            ElseIf SegmentList(i).point1_id = id Then
                seg0_id = i
            End If
        Next
        
        If seg0_id > 0 And seg1_id > 0 Then
            ElementType = PointWithBothSegments
            Param = seg0_id
            Param2 = seg1_id
            Exit Function
        ElseIf seg0_id > 0 Then
            ElementType = PointWithSegment0
            Param = seg0_id
            Exit Function
        ElseIf seg1_id > 0 Then
            ElementType = PointWithSegment1
            Param = seg1_id
            Exit Function
        End If
        
        For i = 1 To ArcCount
            If ArcList(i).point0_id = id Then
                ElementType = PointOfArcPoint0
                Param = i
                Exit Function
            ElseIf ArcList(i).point1_id = id Then
                ElementType = PointOfArcPoint1
                Param = i
                Exit Function
            ElseIf ArcList(i).pointm_id = id Then
                ElementType = PointOfArcPointm
                Param = i
                Exit Function
            End If
        Next
                
        For i = 1 To SPLineCount
            If SPLineList(i).point0_id = id Then
                ElementType = PointOfSPLinePoint0
                Param = i
                Exit Function
            ElseIf SPLineList(i).point1_id = id Then
                ElementType = PointOfSPLinePoint1
                Param = i
                Exit Function
            Else
                For j = 1 To SPLineList(i).vertex_count - 2
                    If SPLineList(i).vertex_id(j) = id Then
                        ElementType = PointOfSPLinePointm
                        Param = i
                        Param2 = j
                        Exit Function
                    End If
                Next
            End If
        Next
    End If
    
    '------------------------------------------------------------------------------------
    For i = 1 To SegmentCount
        lyr0 = PointList(SegmentList(i).point0_id).Layer
        lyr1 = PointList(SegmentList(i).point1_id).Layer
        
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or (lyr0 = CurLayer And lyr1 = CurLayer) Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or (lyr0 = CurLayer And lyr1 = CurLayer) Then
            x0 = PointList(SegmentList(i).point0_id).X
            y0 = PointList(SegmentList(i).point0_id).Y
            x1 = PointList(SegmentList(i).point1_id).X
            y1 = PointList(SegmentList(i).point1_id).Y
            
            If Abs(x1 - x0) > Abs(y1 - y0) Then
                If (ux - x0) * (ux - x1) <= 0 Then
                    k = (y1 - y0) / (x1 - x0)
                    c = -x0 * k + y0
                    uc = -ux * k + uy
                    m = Abs(uc - c)
                    
                    If m <= TrapWidth Then
                        If id = 0 Or m < l Then
                            id = i
                            l = m
                        End If
                    End If
                End If
            Else
                If (uy - y0) * (uy - y1) <= 0 Then
                    k = (x1 - x0) / (y1 - y0)
                    c = -y0 * k + x0
                    uc = -uy * k + ux
                    m = Abs(uc - c)
                    
                    If m <= TrapWidth Then
                        If id = 0 Or m < l Then
                            id = i
                            l = m
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If id > 0 Then
        CatchElement = ElementSegment
        ElementType = ElementSegment
        Param = SegmentList(id).point0_id
        Param2 = SegmentList(id).point1_id
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------
    For i = 1 To ArcCount
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or ArcList(I).Layer = CurLayer Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or ArcList(i).Layer = CurLayer Then
            dX = ux - ArcList(i).X
            dy = uy - ArcList(i).Y
            
            If Abs(dX) <= TrapWidth And Abs(dy) <= TrapWidth Then
                m = dX * dX + dy * dy
                If id = 0 Or m < l Then
                    id = i
                    l = m
                End If
            End If
        End If
    Next
    
    If id > 0 Then
        CatchElement = ElementArc
        ElementType = ArcCenterPoint
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------
    For i = 1 To ArcCount
        'If FrmMain.chkOnlyCurLayer.Value = 0 Or ArcList(I).Layer = CurLayer Then
        If FrmMain.Toolbar1.Buttons(22).value = tbrUnpressed Or ArcList(i).Layer = CurLayer Then
            dX = ux - ArcList(i).X
            dy = uy - ArcList(i).Y
            
            a = ArcList(i).a
            b = ArcList(i).b
            
            If dX < a + TrapWidth And dX > -a - TrapWidth And dy < b + TrapWidth And dy > -b - TrapWidth Then
                If a = b Then
                    If Abs(Sqr(dX * dX + dy * dy) - a) < TrapWidth Then
                        Angle0 = ArcList(i).start_angle + ArcList(i).ax_angle
                        Angle1 = ArcList(i).end_angle + ArcList(i).ax_angle
                        angle = GetArcAngle(0, 0, dX, dy)
                        If (Angle1 - angle) * (angle - Angle0) > 0 Or _
                           (Angle1 - (angle + PI2)) * ((angle + PI2) - Angle0) > 0 Or _
                           (Angle1 - (angle - PI2)) * ((angle - PI2) - Angle0) > 0 Then
                            id = i
                            Exit For
                        End If
                    End If
                Else
                    If Abs(dX) > Abs(dy) Then
                        x1 = a * b * dX / Sqr(b * dX * b * dX + a * dy * a * dy)
                        x2 = -x1
                        y1 = x1 * dy / dX
                        y2 = -y1
                    Else
                        y1 = a * b * dy / Sqr(b * dX * b * dX + a * dy * a * dy)
                        y2 = -y1
                        x1 = y1 * dX / dy
                        x2 = -x1
                    End If
                    
                    If Sqr((dX - x1) * (dX - x1) + (dy - y1) * (dy - y1)) < TrapWidth Or _
                       Sqr((dX - x2) * (dX - x2) + (dy - y2) * (dy - y2)) < TrapWidth Then
                        Angle0 = GetArcAngle(ArcList(i).X, ArcList(i).Y, PointList(ArcList(i).point0_id).X, PointList(ArcList(i).point0_id).Y)
                        Angle1 = GetArcAngle(ArcList(i).X, ArcList(i).Y, PointList(ArcList(i).point1_id).X, PointList(ArcList(i).point1_id).Y)
                        
                        angle = GetArcAngle(0, 0, dX, dy)
                        If (Angle1 - angle) * (angle - Angle0) > 0 Or _
                           (Angle1 - (angle + PI2)) * ((angle + PI2) - Angle0) > 0 Or _
                           (Angle1 - (angle - PI2)) * ((angle - PI2) - Angle0) > 0 Then
                            id = i
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If id > 0 Then
        CatchElement = ElementArc
        ElementType = ArcEdge
        Exit Function
    End If
End Function

Function GetArcAngle(ByVal cx As Double, ByVal cy As Double, ByVal X As Double, ByVal Y As Double) As Double
    Dim xl As Double, yl As Double
    Dim angle As Double
    
    xl = X - cx
    yl = Y - cy
    
    If xl = 0 And yl = 0 Then
        GetArcAngle = -999999
        Exit Function
    End If
    
    If Abs(xl) > Abs(yl) Then
        angle = Atn(Abs(yl) / Abs(xl))
    Else
        angle = PI_2 - Atn(Abs(xl) / Abs(yl))
    End If
    
    If xl >= 0 And yl < 0 Then
        angle = -angle
    ElseIf xl < 0 And yl >= 0 Then
        angle = Pi - angle
    ElseIf xl < 0 And yl < 0 Then
        angle = angle - Pi
    End If
    
    GetArcAngle = angle
End Function

Function RoundCorner(ByVal id As Long, ByVal r As Double) As Boolean
    Dim i As Long, k As Long
    Dim seg0_id As Long, seg1_id As Long, seg_p0_id As Long, seg_p1_id As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim cx As Double, cy As Double, gx0 As Double, gy0 As Double, gx1 As Double, gy1 As Double
    Dim dx0 As Double, dy0 As Double, dx1 As Double, dy1 As Double
    Dim a0 As Double, b0 As Double, a1 As Double, b1 As Double
    Dim fa0 As Double, fb0 As Double, fa1 As Double, fb1 As Double, d As Double
    Dim ga0 As Double, gb0 As Double, ga1 As Double, gb1 As Double
    Dim vx0 As Double, vy0 As Double, vx1 As Double, vy1 As Double
    Dim Angle0 As Double, Angle1 As Double
    
    Dim px1 As Single, py1 As Single
    Dim px2 As Single, py2 As Single
    
    k = 0
    For i = 1 To SegmentCount
        If SegmentList(i).point0_id = id Then
            seg_p1_id = SegmentList(i).point1_id
            seg1_id = i
            k = k + 1
        End If
        
        If SegmentList(i).point1_id = id Then
            seg_p0_id = SegmentList(i).point0_id
            seg0_id = i
            k = k + 1
        End If
    Next
    
    If k <> 2 Then '不是两线段的首尾相交点
        MsgBox " 不符合倒角条件，不能对该顶点进行倒角处理 ! ", vbExclamation + vbOKOnly
        RoundCorner = False
        Exit Function
    End If
    
    If seg0_id > 0 And seg1_id > 0 Then 'Middle Point
        x0 = PointList(seg_p0_id).X
        y0 = PointList(seg_p0_id).Y
        
        x1 = PointList(id).X
        y1 = PointList(id).Y
        
        x2 = PointList(seg_p1_id).X
        y2 = PointList(seg_p1_id).Y
        
        dx0 = x1 - x0
        dy0 = y1 - y0
        dx1 = x2 - x1
        dy1 = y2 - y1
        
        If Abs(dx0 * dy1 - dx1 * dy0) < 0.5 Then '两线段平行,连接成直线
            MsgBox " 不符合倒角条件，不能对该顶点进行倒角处理 ! ", vbExclamation + vbOKOnly
            RoundCorner = False
            Exit Function
        End If
        
        '转成两个矢量
        vx0 = dx0
        vy0 = dy0
        vx1 = x2 - x0
        vy1 = y2 - y0
        
        If vx0 * vy1 - vx1 * vy0 > 0 Then '叉积,得出面法线的方向
            d = 1
        Else
            d = -1
        End If
        
        If dx0 <> 0 And dx1 <> 0 And dy0 <> 0 And dy1 <> 0 Then
            'L0: y = a0 * x + b0 '线段方程
            a0 = dy0 / dx0
            b0 = y0 - dy0 / dx0 * x0
            
            'L1: y = a1 * x + b1
            a1 = dy1 / dx1
            b1 = y1 - dy1 / dx1 * x1
            
            'F0: y = fa0 * x + fb0 '平行线方程, d 确定左右侧
            fa0 = a0
            fb0 = b0 + d * r * Sqr(dx0 * dx0 + dy0 * dy0) / dx0
            
            
            'F1: y = fa1 * x + fb1
            fa1 = a1
            fb1 = b1 + d * r * Sqr(dx1 * dx1 + dy1 * dy1) / dx1
                    
            'F0 & F1 '平行线相交求圆心
            cx = (fb1 - fb0) / (fa0 - fa1)
            cy = fa0 * cx + fb0
            
            '------------------------------------------------
            'G0: y = ga0 * x + gb0 '过圆心的垂线方程
            ga0 = -dx0 / dy0
            gb0 = dx0 / dy0 * cx + cy
            
            'G1: y = ga1 * x + gb1
            ga1 = -dx1 / dy1
            gb1 = dx1 / dy1 * cx + cy
            
            '------------------------------------------------
            'L0 & G0 '垂线与线段相交求切点
            gx0 = (gb0 - b0) / (a0 - ga0)
            gy0 = a0 * gx0 + b0
            
            'L1 & G1
            gx1 = (gb1 - b1) / (a1 - ga1)
            gy1 = a1 * gx1 + b1
            
        ElseIf dx0 = 0 And dy1 = 0 Then
            'L0: x = x0 '线段方程
            'L1: y = y1
            
            'F0: x = x0 + d * r '平行线方程, d 确定左右侧
            'F1: y = y1 + d * r
                    
            'F0 & F1 '平行线相交求圆心
            cx = x0 - d * r * Sgn(dy0)
            cy = y1 + d * r * Sgn(dx1)
            
            '------------------------------------------------
            'G0: y = cy '过圆心的垂线方程
            'G1: x=cx
            
            
            '------------------------------------------------
            'L0 & G0 '垂线与线段相交求切点
            gx0 = x0
            gy0 = cy
            
            'L1 & G1
            gx1 = cx
            gy1 = y1
            
        ElseIf dy0 = 0 And dx1 = 0 Then
            'L0: y = y0 '线段方程
            'L1: x = x1
            
            'F0: y = y0 + d * r '平行线方程, d 确定左右侧
            'F1: x = x1 + d * r
                    
            'F0 & F1 '平行线相交求圆心
            cx = x1 - d * r * Sgn(dy1)
            cy = y0 + d * r * Sgn(dx0)
            
            '------------------------------------------------
            'G0: y = cy '过圆心的垂线方程
            'G1: x = cx
    
            '------------------------------------------------
            'L0 & G0 '垂线与线段相交求切点
            gx0 = cx
            gy0 = y0
            
            'L1 & G1
            gx1 = x1
            gy1 = cy
            
        ElseIf dx0 = 0 Then
            'L0: x = x0 '线段方程
            
            'L1: y = a1 * x + b1
            a1 = dy1 / dx1
            b1 = y1 - dy1 / dx1 * x1
            
            'F0: x = x0 + d * r '平行线方程, d 确定左右侧
                        
            'F1: y = fa1 * x + fb1
            fa1 = a1
            fb1 = b1 + d * r * Sqr(dx1 * dx1 + dy1 * dy1) / dx1
                    
            'F0 & F1 '平行线相交求圆心
            cx = x0 - d * r * Sgn(dy0)
            cy = fa1 * cx + fb1
            
            '------------------------------------------------
            'G0: y = cy '过圆心的垂线方程
            
            'G1: y = ga1 * x + gb1
            ga1 = -dx1 / dy1
            gb1 = dx1 / dy1 * cx + cy
            
            '------------------------------------------------
            'L0 & G0 '垂线与线段相交求切点
            gx0 = x0
            gy0 = cy
            
            'L1 & G1
            gx1 = (gb1 - b1) / (a1 - ga1)
            gy1 = a1 * gx1 + b1
            
        ElseIf dx1 = 0 Then
            'L0: y = a0 * x + b0
            a0 = dy0 / dx0
            b0 = y0 - dy0 / dx0 * x0
            
            'L1: x = x1 '线段方程
            
            'F0: y = fa0 * x + fb0
            fa0 = a0
            fb0 = b0 + d * r * Sqr(dx0 * dx0 + dy0 * dy0) / dx0
                    
            'F1: x = x1 + d * r '平行线方程, d 确定左右侧
                        
            'F0 & F1 '平行线相交求圆心
            cx = x1 - d * r * Sgn(dy1)
            cy = fa0 * cx + fb0
            
            '------------------------------------------------
            'G0: y = ga0 * x + gb0
            ga0 = -dx0 / dy0
            gb0 = dx0 / dy0 * cx + cy
            
            'G1: y = cy '过圆心的垂线方程
            
            '------------------------------------------------
            'L0 & G0
            gx0 = (gb0 - b0) / (a0 - ga0)
            gy0 = a0 * gx0 + b0
                        
            'L1 & G1 '垂线与线段相交求切点
            gx1 = x1
            gy1 = cy
            
        ElseIf dy0 = 0 Then
            'L0: y = y0 '线段方程
            
            'L1: y = a1 * x + b1
            a1 = dy1 / dx1
            b1 = y1 - dy1 / dx1 * x1
            
            'F0: y = y0 + d * r '平行线方程, d 确定左右侧
            
            'F1: y = fa1 * x + fb1
            fa1 = a1
            fb1 = b1 + d * r * Sqr(dx1 * dx1 + dy1 * dy1) / dx1
                    
            'F0 & F1 '平行线相交求圆心
            cy = y0 + d * r * Sgn(dx0)
            cx = (cy - fb1) / fa1
            
            '------------------------------------------------
            'G0: x = cx '过圆心的垂线方程
            
            'G1: y = ga1 * x + gb1
            ga1 = -dx1 / dy1
            gb1 = dx1 / dy1 * cx + cy
            
            '------------------------------------------------
            'L0 & G0 '垂线与线段相交求切点
            gx0 = cx
            gy0 = y0
            
            'L1 & G1
            gx1 = (gb1 - b1) / (a1 - ga1)
            gy1 = a1 * gx1 + b1
            
        ElseIf dy1 = 0 Then
            'L1: y = y1 '线段方程
            
            'L0: y = a0 * x + b0
            a0 = dy0 / dx0
            b0 = y0 - dy0 / dx0 * x0
            
            'F1: y = y1 + d * r '平行线方程, d 确定左右侧
            
            'F0: y = fa0 * x + fb0
            fa0 = a0
            fb0 = b0 + d * r * Sqr(dx0 * dx0 + dy0 * dy0) / dx0
                    
            'F0 & F1 '平行线相交求圆心
            cy = y1 + d * r * Sgn(dx1)
            cx = (cy - fb0) / fa0
            
            '------------------------------------------------
            'G1: x = cx '过圆心的垂线方程
            
            'G0: y = ga0 * x + gb0
            ga0 = -dx0 / dy0
            gb0 = dx0 / dy0 * cx + cy
            
            '------------------------------------------------
            'L1 & G1 '垂线与线段相交求切点
            gx1 = cx
            gy1 = y1
            
            'L0 & G0
            gx0 = (gb0 - b0) / (a0 - ga0)
            gy0 = a0 * gx0 + b0
        End If
        
        'CurArc.x = cx
        'CurArc.y = cy
        'CurArc.a = r
        'CurArc.b = r
            
        Angle0 = GetArcAngle(cx, cy, gx0, gy0)
        Angle1 = GetArcAngle(cx, cy, gx1, gy1)

        If Angle0 - Angle1 > Pi And Angle1 < 0 Then
            Angle1 = Angle1 + PI2
        ElseIf Angle0 - Angle1 < -Pi And Angle0 < 0 Then
            Angle0 = Angle0 + PI2
        End If
        
        'CurArc.start_angle = angle0
        'CurArc.end_angle = angle1
        
        If PointList(id).method = PointMethod.RoundedCorner Then 'Redo
            ArcList(PointList(id).arc_id).X = cx
            ArcList(PointList(id).arc_id).Y = cy
            ArcList(PointList(id).arc_id).a = r
            ArcList(PointList(id).arc_id).b = r
            ArcList(PointList(id).arc_id).ax_angle = 0
            ArcList(PointList(id).arc_id).start_angle = Angle0
            ArcList(PointList(id).arc_id).end_angle = Angle1
            
            PointList(ArcList(PointList(id).arc_id).point0_id).X = gx0
            PointList(ArcList(PointList(id).arc_id).point0_id).Y = gy0
            PointList(ArcList(PointList(id).arc_id).pointm_id).X = cx
            PointList(ArcList(PointList(id).arc_id).pointm_id).Y = cy
            PointList(ArcList(PointList(id).arc_id).point1_id).X = gx1
            PointList(ArcList(PointList(id).arc_id).point1_id).Y = gy1
            
        Else
            AddArc cx, cy, PointList(id).z, r, r, Angle0, Angle1, 0, 0, 0, PointList(id).Layer, ArcType.RoundedCorner
            ArcList(ArcCount).point_id = id
            ArcList(ArcCount).body_id = PointList(id).body_id
            ArcList(ArcCount).group_id = PointList(id).group_id
                        
            AddPoint gx0, gy0, LayerZValue(CurLayer), CurLayer, ArcPoint
            ArcList(ArcCount).point0_id = PointList(PointCount).id
            PointList(PointCount).body_id = PointList(id).body_id
            PointList(PointCount).group_id = PointList(id).group_id
            
            AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, ArcPoint
            ArcList(ArcCount).pointm_id = PointList(PointCount).id
            PointList(PointCount).body_id = PointList(id).body_id
            PointList(PointCount).group_id = PointList(id).group_id
            
            AddPoint gx1, gy1, LayerZValue(CurLayer), CurLayer, ArcPoint
            ArcList(ArcCount).point1_id = PointList(PointCount).id
            PointList(PointCount).body_id = PointList(id).body_id
            PointList(PointCount).group_id = PointList(id).group_id
            
            PointList(id).method = PointMethod.RoundedCorner
            PointList(id).arc_id = ArcCount
        End If
    End If
    
    RoundCorner = True
End Function

Function GetCircleBy3Points(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, cx As Double, cy As Double, r As Double, sa As Double, ea As Double) As Boolean
    Dim dx0 As Double, dy0 As Double, k0 As Integer, dx1 As Double, dy1 As Double, k1 As Integer
    Dim xm0 As Double, ym0 As Double, xm1 As Double, ym1 As Double
    Dim a0 As Double, b0 As Double, a1 As Double, b1 As Double
    
    dx0 = x1 - x0
    dy0 = y1 - y0
    
    xm0 = x0 + (x1 - x0) / 2
    ym0 = y0 + (y1 - y0) / 2
    
    dx1 = x2 - x1
    dy1 = y2 - y1
    
    xm1 = x1 + (x2 - x1) / 2
    ym1 = y1 + (y2 - y1) / 2
    
    If Abs(dx0 * dy1 - dx1 * dy0) < 0.00001 Then '三点连接成直线
'        MsgBox " 点的数据错误，无法得出对应的圆形 ! ", vbExclamation + vbOKOnly
        GetCircleBy3Points = False
        Exit Function
    End If
        
    If Abs(dy0) > Abs(dx0) Then
        a0 = -dx0 / dy0
        b0 = -xm0 * a0 + ym0
        k0 = 0
    Else
        a0 = -dy0 / dx0
        b0 = -ym0 * a0 + xm0
        k0 = 1
    End If
    
    If Abs(dy1) > Abs(dx1) Then
        a1 = -dx1 / dy1
        b1 = -xm1 * a1 + ym1
        k1 = 0
    Else
        a1 = -dy1 / dx1
        b1 = -ym1 * a1 + xm1
        k1 = 1
    End If
    
    If k0 = 0 And k1 = 0 Then
        cx = (b1 - b0) / (a0 - a1)
        cy = a0 * cx + b0
    ElseIf k0 = 0 And k1 = 1 Then
        cx = (a1 * b0 + b1) / (1 - a0 * a1)
        cy = a0 * cx + b0
    ElseIf k0 = 1 And k1 = 0 Then
        cx = (a0 * b1 + b0) / (1 - a0 * a1)
        cy = a1 * cx + b1
    ElseIf k0 = 1 And k1 = 1 Then
        cy = (b1 - b0) / (a0 - a1)
        cx = a0 * cy + b0
    End If
    
    r = Sqr((cx - x0) * (cx - x0) + (cy - y0) * (cy - y0))
    
    sa = GetArcAngle(cx, cy, x0, y0)
    ea = GetArcAngle(cx, cy, x2, y2)
    
    If dx0 * dy1 - dx1 * dy0 > 0 And sa > ea Then
        ea = ea + PI2
    ElseIf dx0 * dy1 - dx1 * dy0 < 0 And sa < ea Then
        sa = sa + PI2
    End If
    
    GetCircleBy3Points = True
End Function

Sub DrawCursorReferenceLines(ByVal X As Single, ByVal Y As Single, ByVal mode As Integer)
    Dim dm As Integer
    
    Static x0 As Single, y0 As Single, NotFirst As Boolean
    
    If X = -9999 Then
        NotFirst = False
        x0 = -1
        y0 = -1
        Exit Sub
    End If
    
    If CanShowCursorReferenceLines = False Or FrmMain.Toolbar1.Buttons(15).value = tbrUnpressed Then
        Exit Sub
    End If
    
    dm = FrmMain.PicPath.DrawMode
    FrmMain.PicPath.DrawMode = 7
    
    If mode = 0 Then
        If NotFirst Then
            LLine 0, y0, PathMaxX, y0, RGB(128, 128, 0), 1
            LLine x0, 0, x0, PathMaxY, RGB(128, 128, 0), 1
        End If
    Else
        LLine 0, Y, PathMaxX, Y, RGB(128, 128, 0), 1
        LLine X, 0, X, PathMaxY, RGB(128, 128, 0), 1
    
        x0 = X
        y0 = Y
    End If
    NotFirst = True
    
    FrmMain.PicPath.DrawMode = dm
End Sub

Sub SetStartDroppingOnChain(ByVal Start_pid As Long, ByVal Stop_pid As Long, ByRef end_pid As Long)
    Dim i As Long, j As Long, q As Integer
    Static k As Long, pid0 As Long
    
    'If PointList(Start_pid).action = ActionType.StartDropping Then
    '    Exit Sub
    'End If
    
    k = k + 1
    If k = 1 Then
        PointList(Start_pid).action = ActionType.StartDropping
        pid0 = Start_pid
        
        If PointList(Start_pid).Type = PointType.SPLinePoint Then
            For i = 1 To SPLineCount
                For j = 0 To SPLineList(i).vertex_count - 1
                    If SPLineList(i).vertex_id(j) = Start_pid And SPLineList(i).point1_id <> Start_pid Then
                        Exit For
                    End If
                Next
                If j < SPLineList(i).vertex_count Then
                    Exit For
                End If
            Next
            If i <= SPLineCount Then
                For j = 0 To SPLineList(i).vertex_count - 1
                    If SPLineList(i).vertex_id(j) = SPLineList(i).point0_id Then
                        PointList(SPLineList(i).vertex_id(j)).action = ActionType.StartDropping
                    ElseIf SPLineList(i).vertex_id(j) <> SPLineList(i).point1_id Then
                        PointList(SPLineList(i).vertex_id(j)).action = ActionType.Dropping
                    End If
                Next
                SetStartDroppingOnChain SPLineList(i).point1_id, Stop_pid, end_pid
                GoTo Exit_Sub
            End If
        End If
    Else
        q = 0
        For i = 1 To SegmentCount
            If SegmentList(i).point0_id = Start_pid Then
                If PointList(Start_pid).action = ActionType.StartDropping Then
                    PointList(Start_pid).action = ActionType.Dropping
                    For j = 1 To OutputStartPointList.Count
                        If OutputStartPointList.point_id(j) = Start_pid Then
                            OutputStartPointList.point_id(j) = pid0
                            Exit For
                        End If
                    Next
                    GoTo Exit_Sub
                ElseIf PointList(Start_pid).action <> ActionType.StopDropping Then
                    q = 1
                End If
                Exit For
            End If
        Next
        If q = 0 Then
            For i = 1 To ArcCount
                If ArcList(i).point0_id = Start_pid Then
                    If PointList(Start_pid).action = ActionType.StartDropping Then
                        PointList(Start_pid).action = ActionType.Dropping
                        For j = 1 To OutputStartPointList.Count
                            If OutputStartPointList.point_id(j) = Start_pid Then
                                OutputStartPointList.point_id(j) = pid0
                                Exit For
                            End If
                        Next
                        GoTo Exit_Sub
                    ElseIf PointList(Start_pid).action <> ActionType.StopDropping Then
                        q = 1
                    End If
                    Exit For
                End If
            Next
        End If
        If q = 0 Then
            For i = 1 To SPLineCount
                If SPLineList(i).point0_id = Start_pid Then
                    If PointList(Start_pid).action = ActionType.StartDropping Then
                        PointList(Start_pid).action = ActionType.Dropping
                        For j = 1 To OutputStartPointList.Count
                            If OutputStartPointList.point_id(j) = Start_pid Then
                                OutputStartPointList.point_id(j) = pid0
                                Exit For
                            End If
                        Next
                        GoTo Exit_Sub
                    ElseIf PointList(Start_pid).action <> ActionType.StopDropping Then
                        q = 1
                    End If
                    Exit For
                End If
            Next
        End If
        If q = 1 Then
            PointList(Start_pid).action = ActionType.Dropping
        Else
            PointList(Start_pid).action = ActionType.StopDropping
            end_pid = Start_pid
            
            GoTo Exit_Sub
        End If
    End If
    
    If PointList(Start_pid).method = PointMethod.RoundedCorner Then
        PointList(ArcList(PointList(Start_pid).arc_id).point0_id).action = ActionType.Dropping
        PointList(ArcList(PointList(Start_pid).arc_id).point1_id).action = ActionType.Dropping
    End If
    
    For i = 1 To SegmentCount
        If SegmentList(i).point0_id = Start_pid Then
            If SegmentList(i).point1_id <> Stop_pid And SegmentList(i).point1_id <> pid0 Then
                SetStartDroppingOnChain SegmentList(i).point1_id, Stop_pid, end_pid
            ElseIf SegmentList(i).point1_id = Stop_pid Then
                end_pid = Stop_pid
            Else
                end_pid = pid0
            End If
            GoTo Exit_Sub
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).point0_id = Start_pid Then
            If ArcList(i).point1_id <> Stop_pid And ArcList(i).point1_id <> pid0 Then
                SetStartDroppingOnChain ArcList(i).point1_id, Stop_pid, end_pid
            ElseIf ArcList(i).point1_id = Stop_pid Then
                end_pid = Stop_pid
            Else
                end_pid = pid0
            End If
            GoTo Exit_Sub
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).point0_id = Start_pid Then
            If SPLineList(i).point1_id <> Stop_pid And SPLineList(i).point1_id <> pid0 Then
                SetStartDroppingOnChain SPLineList(i).point1_id, Stop_pid, end_pid
            ElseIf SPLineList(i).point1_id = Stop_pid Then
                end_pid = Stop_pid
            Else
                end_pid = pid0
            End If
            GoTo Exit_Sub
        End If
    Next
    
Exit_Sub:
    If k = 1 Then
        pid0 = 0
    End If
    k = k - 1
End Sub

Sub SetStopDroppingOnChain(ByVal Start_pid As Long, ByVal Stop_pid As Long)
    Dim i As Long, j As Long
    Static k As Long, pid0 As Long
    
    k = k + 1
    If k = 1 Then
        PointList(Start_pid).action = ActionType.StopDropping
        pid0 = Start_pid
        
        If PointList(Start_pid).Type = PointType.SPLinePoint Then
            For i = 1 To SPLineCount
                For j = 0 To SPLineList(i).vertex_count - 1
                    If SPLineList(i).vertex_id(i) = Start_pid And SPLineList(i).point1_id <> Start_pid Then
                        Exit For
                    End If
                Next
                If j < SPLineList(i).vertex_count Then
                    Exit For
                End If
            Next
            If i <= SPLineCount Then
                For j = 0 To SPLineList(i).vertex_count - 1
                    If SPLineList(i).vertex_id(j) = SPLineList(i).point0_id Then
                        PointList(SPLineList(i).vertex_id(j)).action = ActionType.StopDropping
                    ElseIf SPLineList(i).vertex_id(j) <> SPLineList(i).point1_id Then
                        PointList(SPLineList(i).vertex_id(j)).action = ActionType.No_Action
                    End If
                Next
                SetStopDroppingOnChain SPLineList(i).point1_id, Stop_pid
                GoTo Exit_Sub
            End If
        End If
        
    Else
        If PointList(Start_pid).action = ActionType.StartDropping Then
            GoTo Exit_Sub
        End If
    
        PointList(Start_pid).action = ActionType.No_Action
    End If
    
    If PointList(Start_pid).method = PointMethod.RoundedCorner Then
        PointList(ArcList(PointList(Start_pid).arc_id).point0_id).action = ActionType.No_Action
        PointList(ArcList(PointList(Start_pid).arc_id).point1_id).action = ActionType.No_Action
    End If
    
    For i = 1 To SegmentCount
        If SegmentList(i).point0_id = Start_pid Then
            If SegmentList(i).point1_id <> Stop_pid And _
               SegmentList(i).point1_id <> pid0 And _
               PointList(SegmentList(i).point1_id).action <> ActionType.StartDropping Then
                
                SetStopDroppingOnChain SegmentList(i).point1_id, Stop_pid
            End If
            GoTo Exit_Sub
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).point0_id = Start_pid Then
            If ArcList(i).point1_id <> Stop_pid And _
               ArcList(i).point1_id <> pid0 And _
               PointList(ArcList(i).point1_id).action <> ActionType.StartDropping Then
                SetStopDroppingOnChain ArcList(i).point1_id, Stop_pid
            End If
            GoTo Exit_Sub
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).point0_id = Start_pid Then
            If SPLineList(i).point1_id <> Stop_pid And _
               SPLineList(i).point1_id <> pid0 And _
               PointList(SPLineList(i).point1_id).action <> ActionType.StartDropping Then
                SetStopDroppingOnChain SPLineList(i).point1_id, Stop_pid
            End If
            GoTo Exit_Sub
        End If
    Next
    
Exit_Sub:
    If k = 1 Then
        pid0 = 0
    End If
    k = k - 1
End Sub

Function GetUserDistance(ByVal x0 As Integer, ByVal y0 As Integer, ByVal x1 As Integer, ByVal y1 As Integer) As Double
    Dim ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double
    
    ConvertPathToUser x0, y0, ux0, uy0
    ConvertPathToUser x1, y1, ux1, uy1
    
'Debug.Print ux0; uy0, ux1; uy1, Sqr((ux1 - ux0) * (ux1 - ux0) + (uy1 - uy0) * (uy1 - uy0))
    GetUserDistance = Sqr((ux1 - ux0) * (ux1 - ux0) + (uy1 - uy0) * (uy1 - uy0))
End Function

Public Function DistanceOfPtToLine(x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
     DistanceOfPtToLine = (Abs((y1 - y2) * x0 + (x2 - x1) * y0 + x1 * y2 - x2 * y1)) / Sqr((y1 - y2) * (y1 - y2) + (x2 - x1) * (x2 - x1))
End Function

Public Function PointOnSegment(x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, SnapDis As Double) As Boolean
    Dim dX As Double, dy As Double
    
    dX = x2 - x1
    dy = y2 - y1
    
    If dX = 0 And dy = 0 Then
        If Sqr((x0 - x1) * (x0 - x1) + (y0 - y1) * (y0 - y1)) < SnapDis Then
            PointOnSegment = True
        Else
            PointOnSegment = False
        End If
        Exit Function
    End If
    
    If DistanceOfPtToLine(x0, y0, x1, y1, x2, y2) < SnapDis Then
        If Abs(dX) <> Abs(dy) Then
            If (x1 - x0) * (x0 - x2) >= 0 Then
                PointOnSegment = True
            Else
                PointOnSegment = False
            End If
        Else
            If (y1 - y0) * (y0 - y2) >= 0 Then
                PointOnSegment = True
            Else
                PointOnSegment = False
            End If
        End If
    Else
        PointOnSegment = False
    End If
End Function

Public Sub DrawBody(ByVal BodyID As Long, Optional clr As Long = 0, Optional ShowPoint As Boolean = False)
    Dim i As Long
    
    For i = 1 To SegmentCount
        If SegmentList(i).body_id = BodyID Then
            DrawSegment SegmentList(i), clr, False
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).body_id = BodyID Then
            DrawArc ArcList(i), clr, ShowPoint
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).body_id = BodyID Then
            DrawSPLine SPLineList(i), clr, ShowPoint
        End If
    Next
End Sub

Public Sub DrawGroup(ByVal GroupID As Long, Optional clr As Long = 0, Optional ShowPoint As Boolean = False)
    Dim i As Long
    
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = GroupID Then
            DrawBody BodyList(i).body_id, clr, ShowPoint
        End If
    Next
End Sub

Public Sub DrawBodyExcept(ByVal BodyID As Long, Optional clr As Long = 0, Optional ShowPoint As Boolean = False)
    Dim i As Long
    
    For i = 1 To SegmentCount
        If SegmentList(i).body_id <> BodyID Then
            DrawSegment SegmentList(i), clr, ShowPoint
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).body_id <> BodyID Then
            DrawArc ArcList(i), clr, ShowPoint
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).body_id <> BodyID Then
            DrawSPLine SPLineList(i), clr, ShowPoint
        End If
    Next
End Sub

Public Sub DrawGroupExcept(ByVal GroupID As Long, Optional clr As Long = 0, Optional ShowPoint As Boolean = False)
    Dim i As Long
    
    'MaxBodyID = GetBodyList(BodyList)
    
    For i = 1 To MaxBodyID
        If BodyList(i).group_id <> GroupID Then
            DrawBody BodyList(i).body_id, clr, ShowPoint
        End If
    Next
End Sub

Public Sub DrawLine2P(ByVal ux0 As Double, ByVal uy0 As Double, ByVal ux1 As Double, ByVal uy1 As Double, ByVal clr As Long)
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single, dw As Integer
    
    ConvertUserToPath ux0, uy0, x0, y0
    ConvertUserToPath ux1, uy1, x1, y1

    LineOut x0, y0, x1, y1, clr
    
    '--------------------------------------------
    Dim d As Single, dX As Single, dy As Single, dx1 As Single, dy1 As Single, dx2 As Single, dy2 As Single
    Dim xm As Double, ym As Double
    
    If ((x1 - x0) <> 0 Or (y1 - y0) <> 0) Then
        d = 8
    
        dX = d * (x1 - x0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        dy = d * (y1 - y0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        
        Rotate_Z dX, dy, 30 * PI_180, dx1, dy1
        Rotate_Z dX, dy, -30 * PI_180, dx2, dy2
        
        xm = x0 + (x1 - x0) / 2
        ym = y0 + (y1 - y0) / 2
        
        LineOut xm, ym, xm - dx1, ym - dy1, clr
        LineOut xm, ym, xm - dx2, ym - dy2, clr
    End If
End Sub

Public Sub DrawLeadingLines()
    Dim i As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim X As Single, Y As Single
    
    For i = 1 To OutputStartPointList.Count
        x0 = OutputStartPointList.leading_point0(i).X
        y0 = OutputStartPointList.leading_point0(i).Y
        
        x1 = PointList(OutputStartPointList.leading_point0(i).id).X
        y1 = PointList(OutputStartPointList.leading_point0(i).id).Y
        
        If x0 <> x1 Or y0 <> y1 Then
            DrawLine2P x0, y0, x1, y1, RGB(127, 127, 255)
            ConvertUserToPath x0, y0, X, Y
            PointOut X, Y, RGB(127, 127, 255)
        End If
        
        x0 = PointList(OutputStartPointList.leading_point1(i).id).X
        y0 = PointList(OutputStartPointList.leading_point1(i).id).Y
        
        x1 = OutputStartPointList.leading_point1(i).X
        y1 = OutputStartPointList.leading_point1(i).Y
        
        If x0 <> x1 Or y0 <> y1 Then
            DrawLine2P x0, y0, x1, y1, RGB(127, 127, 255)
            ConvertUserToPath x1, y1, X, Y
            PointOut X, Y, RGB(127, 127, 255)
        End If
    Next
End Sub


Sub DrawVertPoints()
    Dim i As Long, ux As Double, uy As Double, X As Single, Y As Single, k As Long, t As Long
    
    t = 0
    For i = 1 To PathOutputPointCount
        If PathOutputPoint(i).VertType <= 0 Then
            ux = PathOutputPoint(i).ux
            uy = PathOutputPoint(i).uy
            k = PathOutputPoint(i).Type
            
            ConvertUserToPath ux, uy, X, Y

            If k = 88888 Then
                FrmMain.PicPath.PSet (X, Y)
                t = 1
            ElseIf k = 99999 Then
                FrmMain.PicPath.Line -(X, Y), RGB(255, 255, 0)
                t = 0
            ElseIf t = 1 Then
                FrmMain.PicPath.Line -(X, Y), RGB(255, 255, 0)
            End If
        End If
    Next
    
    k = 0
    For i = 1 To PathOutputPointCount
        If PathOutputPoint(i).VertType = -1 Then
            ux = PathOutputPoint(i).ux
            uy = PathOutputPoint(i).uy
            
            ConvertUserToPath ux, uy, X, Y
            
            FrmMain.PicPath.Circle (X, Y), 8, RGB(0, 255, 0)
            k = k + 1
        End If
    Next
'Debug.Print "vert count="; k
End Sub

