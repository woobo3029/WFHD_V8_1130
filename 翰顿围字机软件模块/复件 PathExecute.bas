Attribute VB_Name = "PathExecute"
Option Explicit

Public Type PathOutputPointType
    ux As Double
    uy As Double
    Type As Long
    LengthFromStart As Double
    AngleToNext As Double
    Radius3P As Double
    VertType As Long 'HoleType
End Type

Public Type OutputStartPoint
    Count As Long
    point_id() As Long
    leading_point0() As Path_Point
    leading_point1() As Path_Point
End Type

Enum OutputMode
    Calculate
    OutputSegments
    
    CalculateStartPoint
    CalculateEndPoint
'    TextOutput
'    ScreenTestOutput
'    DeviceTestOutput
'    DeviceOutput
'    PointListOutput
End Enum

Public OutputStartPointList As OutputStartPoint

Public DemoStep As Double

Public StopOutput As Boolean
Private StillOutput As Boolean

Private Device_ux0 As Double
Private Device_uy0 As Double
Private Device_uz0 As Double

Private Device_xpuls0 As Double
Private Device_ypuls0 As Double
Private Device_zplus0 As Double

Private Device_HeadDown As Double

Public OutputPoint() As PolygonPoint
Public OutputPointCount As Long

Public PathOutLength As Double
Public PathOutAngle As Double
Public PathOutputPointStart As Long
Public TotalPathOutLength As Double '输出路径总长度

Public SumCount As Long
Public SumTotalPathOutLength As Double
Public SumTotalPathOutLength0 As Double

Public MaxPathOutAngle As Double
Public MinPathOutAngle As Double

Public TotalWorkLength As Double
Public TotalWorkBendCount As Long
Public TotalWorkCount As Long
Public TotalWorkTime As Double

Public PathHoleCount As Long
Public PathHolePointID(200) As Long
Public PathHolePos(200) As Double
Public PathHoleType(200) As Long
Public PathHoleWidth(200) As Double

Public PathOutputPointCount As Long     '输出节总点数
Public PathOutputPoint() As PathOutputPointType

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public VT1 As Long, VT2 As Long, VTDir As Long
Dim AdjustMM_0() As Double, AdjustMM_1() As Double
Public PathSmooth As Boolean

Sub OutputPath(ByVal mode As OutputMode, ByVal start_id As Long)
    Dim I As Long, cur_id As Long, p As Path_Point, X As Single, Y As Single
    Dim p_count As Long
    Dim ux As Double, uy As Double, uz As Double
    
    StillOutput = True
        
'    For I = 1 To Min(OutputStartPointList.Count, 1)
'        start_id = OutputStartPointList.point_id(I)
'
'If OutputStartPointList.Count > 0 Then
'    start_id = OutputStartPointList.point_id(1)
'Else
'    start_id = 1
'End If

        p_count = 0
                
        'PathOutLength = 0
        'PathOutAngle = 0
        MaxPathOutAngle = 0
        MinPathOutAngle = 180
        
        cur_id = start_id
        OutputPathPointChain start_id, cur_id, 0, p_count, mode
'    Next

    'Debug.Print ">1--------------------------------------------------"
    'For i = 1 To PathOutputPointCount
    '    Debug.Print i; PathOutputPoint(i).LengthFromStart, PathOutputPoint(i).VertType, PathOutputPoint(i).AngleToNext, PathOutputPoint(i).Radius3P
    'Next
    
    
    StillOutput = False
    StopOutput = False
End Sub

Sub OutputAllPath(ByVal mode As OutputMode)
    Dim I As Long, start_id As Long, cur_id As Long, p As Path_Point, X As Single, Y As Single
    Dim p_count As Long
    Dim ux As Double, uy As Double, uz As Double
    
    SetDroppingByDrawingOrder
    
    StillOutput = True
    p_count = 0
            
    PathOutLength = 0
    PathOutAngle = 0
    MaxPathOutAngle = 0
    MinPathOutAngle = 180
        
    For I = 1 To Max(OutputStartPointList.Count, 1)
        start_id = OutputStartPointList.point_id(I)
        
        cur_id = start_id
        PathOutputPointStart = 0
        OutputPathPointChain start_id, cur_id, 0, p_count, mode
        'PathOutputPoint(PathOutputPointCount).Type = 1 'end point
    Next
    
    StillOutput = False
    StopOutput = False
    
    EraseDroppingSetting
End Sub

Sub OutputPathPointChain(ByVal Start_pid As Long, ByRef Cur_pid As Long, ByVal Stop_pid As Long, ByRef pcount As Long, ByVal mode As OutputMode)
    Dim I As Long, j As Long
    
    If StopOutput Then Exit Sub
    
    pcount = pcount + 1
    
'    If PointList(Cur_pid).stay_time > 0 Then
'        If mode > ScreenTestOutput Then
'            Device_Wait PointList(Cur_pid).stay_time
'            'Debug.Print Cur_pid; "StayTime="; PointList(Cur_pid).stay_time
'        End If
'    End If
    
If PointList(Cur_pid).action = StopDropping Then
    If mode = Calculate Then
        For j = 1 To SegmentCount
            If SegmentList(j).point0_id = Cur_pid Then
                OutputSegment SegmentList(j), CalculateEndPoint
                Exit For
            End If
        Next
    End If
    Exit Sub
End If

    If PointList(Cur_pid).action = ActionType.StartDropping Or _
       PointList(Cur_pid).action = ActionType.Dropping Or _
       mode = OutputSegments Then
    
        'If PointList(Start_pid).type = PointType.RoundedCorner Then
        '    OutputArc ArcList(PointList(Start_pid).arc_id)
        'End If
        
        For I = 1 To SegmentCount
            If SegmentList(I).point0_id = Cur_pid Then
            
            
If mode = Calculate Then
    If Cur_pid = Start_pid Then
        For j = 1 To SegmentCount
            If SegmentList(j).point1_id = Cur_pid Then
                OutputSegment SegmentList(j), CalculateStartPoint
                Exit For
            End If
        Next
    End If
End If

                OutputSegment SegmentList(I), mode
                Cur_pid = SegmentList(I).point1_id
                pcount = pcount + 1
                If SegmentList(I).point1_id <> Stop_pid And SegmentList(I).point1_id <> Start_pid Then
                    OutputPathPointChain Start_pid, Cur_pid, Stop_pid, pcount, mode
                Else
If mode = Calculate Then
    If Cur_pid = Start_pid Then
        For j = 1 To SegmentCount
            If SegmentList(j).point0_id = Cur_pid Then
                OutputSegment SegmentList(j), CalculateEndPoint
                Exit For
            End If
        Next
    End If
End If
                End If
                Exit Sub
            End If
        Next
        
        For I = 1 To ArcCount
            If ArcList(I).point0_id = Cur_pid Then
                OutputArc ArcList(I), mode
                Cur_pid = ArcList(I).point1_id
                pcount = pcount + 1
                If ArcList(I).point1_id <> Stop_pid And ArcList(I).point1_id <> Start_pid Then
                    OutputPathPointChain Start_pid, Cur_pid, Stop_pid, pcount, mode
                End If
                Exit Sub
            End If
        Next
        
        For I = 1 To SPLineCount
            If SPLineList(I).point0_id = Cur_pid Then
                OutputSPline SPLineList(I), mode
                Cur_pid = SPLineList(I).point1_id
                pcount = pcount + 1
                If SPLineList(I).point1_id <> Stop_pid And SPLineList(I).point1_id <> Start_pid Then
                    OutputPathPointChain Start_pid, Cur_pid, Stop_pid, pcount, mode
                End If
                Exit Sub
            End If
        Next
    End If
End Sub

Sub StopOutputPath()
    If StillOutput Then
        StopOutput = True
    End If
End Sub

Sub MarkPoint(ByVal X As Single, ByVal Y As Single, Optional clr As Long = 0)
    Dim d As Integer
    d = 8
    
    FrmMain.PicPath.DrawWidth = 3
    FrmMain.PicPath.Line (X, Y - d)-(X, Y + d + 1), IIf(clr = 0, RGB(0, 255, 255), clr)
    FrmMain.PicPath.Line (X - d - 1, Y)-(X + d + 1, Y), IIf(clr = 0, RGB(0, 255, 255), clr)
    
    FrmMain.PicPath.CurrentX = X
    FrmMain.PicPath.CurrentY = Y
    FrmMain.PicPath.DrawWidth = 1
End Sub

Sub MarkPathinnercorner(ByVal X As Single, ByVal Y As Single, Optional clr As Long = 0)
    Dim d As Integer
    d = 5
    
    FrmMain.PicPath.DrawWidth = 3
    FrmMain.PicPath.Line (X, Y - d)-(X, Y + d + 1), IIf(clr = 0, RGB(0, 255, 0), clr)
    FrmMain.PicPath.Line (X - d - 1, Y)-(X + d + 1, Y), IIf(clr = 0, RGB(0, 255, 0), clr)
    
    FrmMain.PicPath.CurrentX = X
    FrmMain.PicPath.CurrentY = Y
    FrmMain.PicPath.DrawWidth = 1
End Sub

Sub ScreenLine(ByVal ux0 As Double, ByVal uy0 As Double, uz0 As Double, ByVal ux1 As Double, ByVal uy1 As Double, uz1 As Double)
    Dim s As Double, ds As Double, l As Double, dux As Double, duy As Double, duz As Double, ux As Double, uy As Double, uz As Double
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    
    If StopOutput Then Exit Sub
    
    s = Sqr((ux1 - ux0) * (ux1 - ux0) + (uy1 - uy0) * (uy1 - uy0) + (uz1 - uz0) * (uz1 - uz0))
    ds = DemoStep
Debug.Print "ds="; ds
    If s = 0 Then Exit Sub

    dux = ds / s * (ux1 - ux0)
    duy = ds / s * (uy1 - uy0)
    duz = ds / s * (uz1 - uz0)
    
    ux = ux0 + dux
    uy = uy0 + duy
    uz = uz0 + duz
    
    l = ds
    
    'ConvertUserToPath ux0, uy0, x0, y0
    ConvertUserToPath ux0, uy0, x0, y0
    'ShowPosition ux0, uy0, uz0, OnlyStautsBar
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint x0, y0
    
    Do Until l >= s
        MarkPoint x0, y0
        FrmMain.PicPath.DrawMode = 13
        
        If StopOutput Then Exit Sub

        'ConvertUserToPath ux, uy, x1, y1
        ConvertUserToPath ux, uy, x1, y1
        'ShowPosition ux, uy, uz, OnlyStautsBar
        LLine x0, y0, x1, y1, IIf(ColorMode = 0, 0, RGB(255, 255, 255)), 2
        
        FrmMain.PicPath.DrawMode = 7
        MarkPoint x1, y1
                
        ux = ux + dux
        uy = uy + duy
        uz = uz + duz
        
        l = l + ds
        
        x0 = x1
        y0 = y1
        
        Wait 0.02
    Loop
    MarkPoint x0, y0
    FrmMain.PicPath.DrawMode = 13
    
    If StopOutput Then Exit Sub

    'ConvertUserToPath ux1, uy1, x1, y1
    ConvertUserToPath ux1, uy1, x1, y1
    'ShowPosition ux1, uy1, uz1, OnlyStautsBar
    LLine x0, y0, x1, y1, IIf(ColorMode = 0, 0, RGB(255, 255, 255)), 2
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint x1, y1
    
    Wait 0.02
    
    MarkPoint x1, y1
    FrmMain.PicPath.DrawMode = 13
End Sub

Sub ScreenLine2(ByVal ux0 As Double, ByVal uy0 As Double, uz0 As Double, ByVal ux1 As Double, ByVal uy1 As Double, uz1 As Double)
    Dim s As Double
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    
    If StopOutput Then Exit Sub
    
    s = Sqr((ux1 - ux0) * (ux1 - ux0) + (uy1 - uy0) * (uy1 - uy0) + (uz1 - uz0) * (uz1 - uz0))
    
    If s = 0 Then Exit Sub
    
    ConvertUserToPath ux0, uy0, x0, y0
    'ShowPosition ux0, uy0, uz0, OnlyStautsBar
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint x0, y0
    
    FrmMain.PicPath.DrawMode = 13
    
    If StopOutput Then Exit Sub

    ConvertUserToPath ux1, uy1, x1, y1
    'ShowPosition ux1, uy1, uz1, OnlyStautsBar
    LLine x0, y0, x1, y1, IIf(ColorMode = 0, 0, RGB(255, 255, 255)), 2
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint x1, y1
    
    MarkPoint x1, y1
    FrmMain.PicPath.DrawMode = 13
End Sub

Sub OutputLine(ByVal ux0 As Double, ByVal uy0 As Double, uz0 As Double, ByVal ux1 As Double, ByVal uy1 As Double, uz1 As Double, Optional mode As OutputMode)
    Dim Dis As Double, Ang As Double, ds As Double, id As Long, hole_type As HoleType, hole_mm As Double, hole_dmm As Double
    Dim radius As Double, cx As Double, cy As Double, sa As Double, ea As Double
    Static ux00 As Double, uy00 As Double
    
    If StopOutput Then Exit Sub
    
    If mode = CalculateStartPoint Then
        ux00 = ux0
        uy00 = uy0
        Exit Sub
    End If
    
    If mode = CalculateEndPoint Then
        Ang = GetAngle(ux00, uy00, ux0, uy0, ux1, uy1)
        If GetCircleBy3Points(ux00, uy00, ux0, uy0, ux1, uy1, cx, cy, radius, sa, ea) = False Then
            radius = 0
        End If

        PathOutputPoint(PathOutputPointCount).AngleToNext = Ang
        PathOutputPoint(PathOutputPointCount).Radius3P = IIf(Ang > 0, 1, -1) * radius
        PathOutputPoint(PathOutputPointCount).Type = 99999 'end point
        
        If Device_AmericanMaterial = True Then
            PathOutputPointCount = PathOutputPointCount + 1
            ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
            PathOutLength = PathOutLength + Device_ExtendMM
            PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutLength
            PathOutputPoint(PathOutputPointCount).ux = ux0
            PathOutputPoint(PathOutputPointCount).uy = uy0
            PathOutputPoint(PathOutputPointCount).AngleToNext = 0
            PathOutputPoint(PathOutputPointCount).Radius3P = 0
            PathOutputPoint(PathOutputPointCount).Type = -99999 'extented point
        End If
        Exit Sub
    End If
    
    If mode = OutputSegments Then
        If PathOutputPointCount = 0 Or PathOutputPointStart = 0 Then
            PathOutputPointCount = PathOutputPointCount + 1
            ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
            PathOutputPoint(PathOutputPointCount).ux = ux0
            PathOutputPoint(PathOutputPointCount).uy = uy0
            PathOutputPointStart = 1
        End If
        
        PathOutputPointCount = PathOutputPointCount + 1
        ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
        PathOutputPoint(PathOutputPointCount).ux = ux1
        PathOutputPoint(PathOutputPointCount).uy = uy1
        Exit Sub
    End If
    
    If mode = Calculate Then
        If PathOutputPointCount = 0 Or PathOutputPointStart = 0 Then
            PathOutputPointCount = PathOutputPointCount + 1
            ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
            If PathOutputPointCount > 1 Then
                PathOutLength = PathOutLength + Device_DoneDistance
            End If
            
            PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutLength
            PathOutputPoint(PathOutputPointCount).ux = ux0
            PathOutputPoint(PathOutputPointCount).uy = uy0
            PathOutputPoint(PathOutputPointCount).Type = 88888 'start point
            
            'PathOutLength = 0
            PathOutAngle = 0
            PathOutputPointStart = 1
        End If
    End If
    
    'If PathOutLength > 0 Then
        Ang = GetAngle(ux00, uy00, ux0, uy0, ux1, uy1)
        
        If GetCircleBy3Points(ux00, uy00, ux0, uy0, ux1, uy1, cx, cy, radius, sa, ea) = False Then
            radius = 0
        End If

        
        If Abs(Ang) > Abs(MaxPathOutAngle) Then
            MaxPathOutAngle = Ang
        End If
        If Abs(Ang) < Abs(MinPathOutAngle) Then
            MinPathOutAngle = Ang
        End If
    'Else
    '    Ang = 0
    '    radius = 0
    'End If
    PathOutAngle = PathOutAngle + Ang
    
    Dis = Sqr((ux1 - ux0) * (ux1 - ux0) + (uy1 - uy0) * (uy1 - uy0))
    PathOutLength = PathOutLength + Dis
    
    If mode = Calculate Then
        PathOutputPointCount = PathOutputPointCount + 1
        ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
        PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutLength
        PathOutputPoint(PathOutputPointCount).ux = ux1
        PathOutputPoint(PathOutputPointCount).uy = uy1
        
        PathOutputPoint(PathOutputPointCount - 1).AngleToNext = Ang
        PathOutputPoint(PathOutputPointCount - 1).Radius3P = IIf(Ang > 0, 1, -1) * radius
    End If
    
    ux00 = ux0
    uy00 = uy0
End Sub

Sub OutputSegment(Segment As Path_Segment, ByVal mode As OutputMode)
    Dim ux0 As Double, uy0 As Double, uz0 As Double, ux1 As Double, uy1 As Double, uz1 As Double, id As Long
    
    If StopOutput Then Exit Sub
    
    If PointList(Segment.point0_id).method = PointMethod.RoundedCorner Then
        id = ArcList(PointList(Segment.point0_id).arc_id).point1_id
        ux0 = PointList(id).X + PointList(id).xp / 100
        uy0 = PointList(id).Y + PointList(id).yp / 100
        
        OutputArc ArcList(PointList(Segment.point0_id).arc_id), mode
    Else
        id = Segment.point0_id
        ux0 = PointList(id).X + PointList(id).xp / 100
        uy0 = PointList(id).Y + PointList(id).yp / 100
    End If
    
    If PointList(Segment.point1_id).method = PointMethod.RoundedCorner Then
        id = ArcList(PointList(Segment.point1_id).arc_id).point0_id
        ux1 = PointList(id).X + PointList(id).xp / 100
        uy1 = PointList(id).Y + PointList(id).yp / 100
    Else
        id = Segment.point1_id
        ux1 = PointList(id).X + PointList(id).xp / 100
        uy1 = PointList(id).Y + PointList(id).yp / 100
    End If
    
    OutputLine ux0, uy0, uz0, ux1, uy1, uz1, mode
End Sub

Sub OutputArc(Arc As Path_Arc, ByVal mode As OutputMode)
    Dim cx As Double, cy As Double, ux0 As Double, uy0 As Double, uz0 As Double, ux As Double, uy As Double, uz As Double
    Dim n As Integer, angle As Double, Angle0 As Double, Angle1 As Double, angle_step As Double, t As Double
    Dim TempSegment As Path_Segment
    
    
    Dim CS As Double, SN As Double, ux00 As Double, uy00 As Double
    
    cx = Arc.X
    cy = Arc.Y
    
    '如果 Arc.a=Arc.b，可用圆弧插补实现
    
    'Debug.Print ">>> Arc"; Arc.id
    
    If StopOutput Then Exit Sub
    
    If Arc.color = -99999 Then
        TempSegment.point0_id = Arc.point0_id
        TempSegment.point1_id = Arc.point1_id
        
        OutputSegment TempSegment, mode
        Exit Sub
    End If
    
    If Arc.a > 0 Then
        CS = Cos(Arc.ax_angle)
        SN = Sin(Arc.ax_angle)
        
        Angle0 = Arc.start_angle '+ Arc.ax_angle
        Angle1 = Arc.end_angle '+ Arc.ax_angle
        
        t = Sqr(Arc.a / UserMaxX)
        If t < 0.2 Then t = 0.2
        
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
        
        
        ux0 = PointList(Arc.point0_id).X + PointList(Arc.point0_id).xp / 100
        uy0 = PointList(Arc.point0_id).Y + PointList(Arc.point0_id).yp / 100
        
        For angle = Angle0 + angle_step To Angle1 - 0.999 * angle_step Step angle_step
        
            If StopOutput Then Exit Sub
    
            'ux = Cos(angle) * Arc.a + cx
            'uy = Sin(angle) * Arc.B + cy
            uz = uz0 '可利用起点和终点的Z值进行插值计算
            
            ux00 = Cos(angle) * Arc.a
            uy00 = Sin(angle) * Arc.b
            
            ux = CS * ux00 - SN * uy00 + cx
            uy = SN * ux00 + CS * uy00 + cy
            
            OutputLine ux0, uy0, uz0, ux, uy, uz, mode
            ux0 = ux
            uy0 = uy
        Next
        
        If StopOutput Then Exit Sub
    
        ux = PointList(Arc.point1_id).X + PointList(Arc.point1_id).xp / 100
        uy = PointList(Arc.point1_id).Y + PointList(Arc.point1_id).yp / 100
        OutputLine ux0, uy0, uz0, ux, uy, uz, mode
    End If
    
    'Debug.Print "<<< Arc"
End Sub

Sub OutputSPline(CurSPline As Path_SPLine, ByVal mode As OutputMode)
    Dim Pts() As PolygonPoint
    Dim ux0 As Double, uy0 As Double, uz0 As Double, ux As Double, uy As Double, uz As Double
    Dim I As Long, n As Long, d As Double
    
    'Debug.Print ">>> SPLine"; CurSPline.id
    
    If StopOutput Then Exit Sub
    
    'SplinePoints CurSPline, Pts(), SPLine_SegmentBetweenPoints
    n = SPLine_SegmentBetweenPoints
    Do
        SplinePoints CurSPline, Pts(), n
        
        For I = 2 To UBound(Pts)
            d = Sqr((Pts(I).X - Pts(I - 1).X) ^ 2 + (Pts(I).Y - Pts(I - 1).Y) ^ 2)
            If d > 0.000001 And d < MinPathStep Then
                If n > 1 Then
                    n = n - 1
                    Exit For
                End If
            End If
        Next
        If I > UBound(Pts) Then Exit Do
        ReDim Pts(0)
    Loop
    
    ux0 = Pts(0).X
    uy0 = Pts(0).Y
    uz0 = LayerZValue(CurSPline.Layer) + IIf(CurSPline.Layer > 0, LayerZValue(0), 0)
    
    For I = 1 To UBound(Pts)
        ux = Pts(I).X
        uy = Pts(I).Y
        uz = uz0
        
        OutputLine ux0, uy0, uz0, ux, uy, uz0, mode
        
        If StopOutput Then
            Exit Sub
        End If
        
        ux0 = ux
        uy0 = uy
    Next I
    
    'Debug.Print "<<< SPLine"
End Sub

Sub CalculatePath(ByVal start_id As Long, Optional USA_Module As Boolean = True)
'计算路径， 将文字路径的基本信息输出到 PointList.txt
    Dim I As Long, j As Long, t As PathOutputPointType
    Dim PathOutputPointCount0 As Long
    Dim d1 As Double, d2 As Double
    Dim ff As Integer
    
    Dim ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, ux As Double, uy As Double, ML As Double
    Dim Sum_Of_AngleAdjustMM As Double
    
    PathOutputPointCount = 0
    
    StopOutput = False
    PathOutLength = 0
    OutputPath Calculate, start_id
        
    If PathOutputPointCount <= 1 Then
        Exit Sub
    End If
    
    ff = FreeFile
    Open "PointList.txt" For Output As #ff
    Print #ff, "序号", Tab(18); "总长度", Tab(35); "点距", Tab(51); "夹角", Tab(67); "半径"
    For I = 1 To PathOutputPointCount
        Print #ff, Mid(str(I) & "    ", 1, 4); Tab(8); Round(PathOutputPoint(I).LengthFromStart, 4); Tab(36); IIf(I = 1, 0, Round(PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart, 4)); Tab(56); Round(PathOutputPoint(I).AngleToNext, 4); Tab(76); Round(PathOutputPoint(I).Radius3P, 4); Tab(94); PathOutputPoint(I).Type
    Next
    Close #ff
    
'    '最后扩充一段长度为 ML 的线段
'    If USA_Module = True And PathOutputPoint(1).ux = PathOutputPoint(PathOutputPointCount).ux And PathOutputPoint(1).uy = PathOutputPoint(PathOutputPointCount).uy Then
'        ux0 = PathOutputPoint(1).ux
'        uy0 = PathOutputPoint(1).uy
'
'        ux1 = PathOutputPoint(2).ux
'        uy1 = PathOutputPoint(2).uy
'
'        ML = Device_ExtendMM
'
'        If Abs(ux1 - ux0) > Abs(uy1 - uy0) Then
'            x1 = ML / Sqr((1 + (uy1 - uy0) ^ 2 / (ux1 - ux0) ^ 2)) + ux0
'            y1 = (uy1 - uy0) / (ux1 - ux0) * (x1 - ux0) + uy0
'
'            x2 = -ML / Sqr((1 + (uy1 - uy0) ^ 2 / (ux1 - ux0) ^ 2)) + ux0
'            y2 = (uy1 - uy0) / (ux1 - ux0) * (x2 - ux0) + uy0
'        Else
'            y1 = ML / Sqr((1 + (ux1 - ux0) ^ 2 / (uy1 - uy0) ^ 2)) + uy0
'            x1 = (ux1 - ux0) / (uy1 - uy0) * (y1 - uy0) + ux0
'
'            y2 = -ML / Sqr((1 + (ux1 - ux0) ^ 2 / (uy1 - uy0) ^ 2)) + uy0
'            x2 = (ux1 - ux0) / (uy1 - uy0) * (y2 - uy0) + ux0
'       End If
'
'        If Sqr((x1 - ux1) ^ 2 + (y1 - uy1) ^ 2) < Sqr((x2 - ux1) ^ 2 + (y2 - uy1) ^ 2) Then
'            ux = x1
'            uy = y1
'        Else
'            ux = x2
'            uy = y2
'        End If
'
'Debug.Print "s="; Sqr((uy - uy0) ^ 2 + (ux - ux0) ^ 2)
'
'        OutputLine ux0, uy0, 0, ux, uy, 0, Calculate
'    End If

    RoundPathInnerCorner
    SetAngleAdjustMM
    TotalPathOutLength = PathOutLength
    
    'Debug.Print ">2--------------------------------------------------"
    For I = 1 To PathOutputPointCount
        PathOutputPoint(I).LengthFromStart = PathOutputPoint(I).LengthFromStart + Device_HeadDistance
        'Debug.Print I; PathOutputPoint(I).LengthFromStart, PathOutputPoint(I).VertType, PathOutputPoint(I).AngleToNext, PathOutputPoint(I).Radius3P
    Next
    
    PathOutputPointCount0 = PathOutputPointCount
    For I = 1 To PathOutputPointCount0
        PathOutputPoint(I).Type = 0
        PathOutputPoint(I).VertType = 0
        If I = 1 Or I = PathOutputPointCount0 Or Abs(PathOutputPoint(I).AngleToNext) >= Device_VertMinAngle Then
            PathOutputPointCount = PathOutputPointCount + 1
            ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
            PathOutputPoint(PathOutputPointCount).ux = PathOutputPoint(I).ux
            PathOutputPoint(PathOutputPointCount).uy = PathOutputPoint(I).uy
            
            ''If USA_Module = False Or (I > 1 And I < PathOutputPointCount0) Then
            'If I > 1 And I < PathOutputPointCount0 Then
            '    If PathOutputPoint(I).AngleToNext < 0 Then
            '        PathOutputPoint(PathOutputPointCount).VertType = 1
            '    Else
            '        PathOutputPoint(PathOutputPointCount).VertType = 2
            '    End If
            'ElseIf I = 1 Then
            '    PathOutputPoint(PathOutputPointCount).VertType = 3 '线段起始端
            'Else
            '    PathOutputPoint(PathOutputPointCount).VertType = 4 '(最后扩充的)线段末端
            'End If
            
            If I = 1 Then
                PathOutputPoint(PathOutputPointCount).Type = 3 '线段起始端
            ElseIf I = PathOutputPointCount0 Then
                PathOutputPoint(PathOutputPointCount).Type = 4 '(最后扩充的)线段末端
            Else
                PathOutputPoint(PathOutputPointCount).Type = 0
            End If
            
            If PathOutputPoint(I).AngleToNext < 0 Then
                PathOutputPoint(PathOutputPointCount).VertType = 1
            Else
                PathOutputPoint(PathOutputPointCount).VertType = 2
            End If
            
            PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutputPoint(I).LengthFromStart - Device_HeadDistance
            
            PathOutputPoint(PathOutputPointCount).AngleToNext = PathOutputPoint(I).AngleToNext
            
            PathOutputPoint(I).VertType = -1
            PathOutputPoint(I).Radius3P = 0
            
            'PathOutputPoint(I).LengthFromStart = PathOutputPoint(I).LengthFromStart + Device_TurnPointOffsetMM '折角点偏移补偿
            
            'PathOutputPoint(I).AngleToNext = 0 '不折角时设为0， 折角时取消此语句
       End If
       
' ?????????? 造成Test图形漏点
'       If I > 1 And I < PathOutputPointCount0 Then
'            d1 = PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart
'            d2 = PathOutputPoint(I + 1).LengthFromStart - PathOutputPoint(I).LengthFromStart
'
'            If d2 / d1 > 3 Then '将PathOutputPoint(i)，PathOutputPoint(i+1)之间作为直线处理
'                PathOutputPoint(I).Radius3P = 0
'                PathOutputPoint(I).AngleToNext = 0
'            End If
'       End If
    Next
    
    '排序
    For I = 1 To PathOutputPointCount
        For j = I + 1 To PathOutputPointCount
            If PathOutputPoint(I).LengthFromStart > PathOutputPoint(j).LengthFromStart Then
                t = PathOutputPoint(I)
                PathOutputPoint(I) = PathOutputPoint(j)
                PathOutputPoint(j) = t
            End If
        Next
    Next
       
    'Debug.Print ">3--------------------------------------------------"
    'For i = 1 To PathOutputPointCount
    '    Debug.Print i; Round(PathOutputPoint(i).LengthFromStart, 4), PathOutputPoint(i).VertType, Round(PathOutputPoint(i).AngleToNext, 4), Round(PathOutputPoint(i).Radius3P, 4)
    'Next
End Sub

Sub CalculateAllPath(Optional repeat_count As Long = 1)
'从PointList.txt 文件中提取数据，计算处理后的路径点数据输出到 outputpointlist.txt 文本中
'repeat_count: 重复次数
    Dim I As Long, j As Long, t As PathOutputPointType
    Dim PathOutputPointCount0 As Long
    Dim ff As Integer
    Dim start_id As Long
    
    Dim ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double
    Dim ux As Double, uy As Double, l As Double, L0 As Double, DL As Double
    Dim ds(5) As Double, dr(5) As Double, cdr As Double, Kp As Long, KN As Long
    
    Dim Sum_Of_AngleAdjustMM As Double
    
    PathOutputPointCount = 0
    
    StopOutput = False
    PathOutLength = 0
    
    For I = 1 To repeat_count
        For j = 1 To OutputStartPointList.Count
            start_id = OutputStartPointList.point_id(j)
            PathOutputPointStart = 0
            OutputPath Calculate, start_id
        Next
    Next
    
    ff = FreeFile
    Open "PointList.txt" For Output As #ff
    Print #ff, "序号"; Tab(10); "总长度", Tab(30); "点距", Tab(51); "夹角", Tab(67); "半径"
    For I = 1 To PathOutputPointCount
        If PathOutputPoint(I).Type < 88888 Then
            Print #ff, Mid(str(I) & "      ", 1, 7); Tab(12); Round(PathOutputPoint(I).LengthFromStart, 4); Tab(36); IIf(I = 1, 0, Round(PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart, 4)); Tab(58); Round(PathOutputPoint(I).AngleToNext, 4); Tab(76); Round(PathOutputPoint(I).Radius3P, 4) '; Tab(94); PathOutputPoint(i).Type
        Else
            Print #ff, Mid(str(I) & "******", 1, 7); Tab(12); Round(PathOutputPoint(I).LengthFromStart, 4); Tab(36); IIf(I = 1, 0, Round(PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart, 4)); Tab(58); Round(PathOutputPoint(I).AngleToNext, 4); Tab(76); Round(PathOutputPoint(I).Radius3P, 4) '; Tab(94); PathOutputPoint(i).Type
        End If
    Next
    Close #ff
    
    If PathOutputPointCount <= 1 Then
        Exit Sub
    End If
    
    If PathSmooth = True Then
    '拉直孤立偏移点，过滤曲线上的凹坑-------------------------------------------------------过滤突兀点
        For I = 3 To PathOutputPointCount - 2
            Kp = 0
            KN = 0
            For j = 1 To 5 '考察包括前2点及后2点在内的5个点, j=3时为点i
                ds(j) = PathOutputPoint(I - 3 + j).LengthFromStart - IIf(I = 3, 0, PathOutputPoint(I - 4 + j).LengthFromStart) '线段长度
                If ds(j) > 10 Then '线段长度大于10则不考虑
                    Exit For
                End If
                dr(j) = Sgn(PathOutputPoint(I - 3 + j).AngleToNext) '前进方向夹角
                If j <> 3 Then
                    If dr(j) > 0 Then Kp = Kp + 1
                    If dr(j) < 0 Then KN = KN + 1
                End If
            Next
            
            If j > 5 Then '所有ds(i)<10
                cdr = PathOutputPoint(I).AngleToNext
                If (Kp > 0 And KN = 0 And cdr < 0 And cdr > -30) Or (Kp = 0 And KN > 0 And cdr > 0 And cdr < 30) Then '前进方向夹角与前后均不一致
                    ux0 = PathOutputPoint(I - 1).ux
                    uy0 = PathOutputPoint(I - 1).uy
                    ux1 = PathOutputPoint(I + 1).ux
                    uy1 = PathOutputPoint(I + 1).uy
                    
                    ux = (ux0 + ux1) / 2
                    uy = (uy0 + uy1) / 2
                    
                    PathOutputPoint(I).ux = ux '调整为前后点的中点，三点成一直线
                    PathOutputPoint(I).uy = uy
                    PathOutputPoint(I).AngleToNext = 0
                    PathOutputPoint(I).Radius3P = 0
                End If
            End If
        Next
    End If
    '在长直线的两侧加点，以控制直线走向
    PathOutputPointCount0 = PathOutputPointCount
    For I = 2 To PathOutputPointCount0
        l = PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart
        If I > 2 And l > 20 And l > 2 * L0 Then '长于20mm且比上一线段长2倍以上的线段要加点
            ux0 = PathOutputPoint(I - 1).ux
            uy0 = PathOutputPoint(I - 1).uy
            ux1 = PathOutputPoint(I).ux
            uy1 = PathOutputPoint(I).uy
            DL = 10 '前端加点的位置
            
            PathOutputPointCount = PathOutputPointCount + 1
            ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
            PathOutputPoint(PathOutputPointCount).ux = ux0 + (ux1 - ux0) * DL / l
            PathOutputPoint(PathOutputPointCount).uy = uy0 + (uy1 - uy0) * DL / l
            PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutputPoint(I - 1).LengthFromStart + DL
            PathOutputPoint(PathOutputPointCount).AngleToNext = 0
            PathOutputPoint(PathOutputPointCount).Radius3P = 0
            
            If l > 30 Then '后端加点的位置为(l - DL)
                PathOutputPointCount = PathOutputPointCount + 1
                ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
                PathOutputPoint(PathOutputPointCount).ux = ux0 + (ux1 - ux0) * (l - DL) / l
                PathOutputPoint(PathOutputPointCount).uy = uy0 + (uy1 - uy0) * (l - DL) / l
                PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutputPoint(I - 1).LengthFromStart + l - DL
                PathOutputPoint(PathOutputPointCount).AngleToNext = 0
                PathOutputPoint(PathOutputPointCount).Radius3P = 0
            End If
       End If
       L0 = l
    Next
    '应该在此计算 Radius3P (而不是在OutputPath中计算)
    
    '排序
    For I = 1 To PathOutputPointCount
        For j = I + 1 To PathOutputPointCount
            If PathOutputPoint(I).LengthFromStart > PathOutputPoint(j).LengthFromStart Then
                t = PathOutputPoint(I)
                PathOutputPoint(I) = PathOutputPoint(j)
                PathOutputPoint(j) = t
            End If
        Next
    Next
    
    'Debug.Print ">1--------------------------------------------------"
    'For i = 1 To PathOutputPointCount
    '    Debug.Print i; PathOutputPoint(i).Type, PathOutputPoint(i).VertType, Round(PathOutputPoint(i).LengthFromStart, 4), Round(PathOutputPoint(i).AngleToNext, 4), Round(PathOutputPoint(i).Radius3P, 4)
    'Next
    
    RoundPathInnerCorner
    'Debug.Print ">1a--------------------------------------------------"
    'For i = 1 To PathOutputPointCount
    '    Debug.Print i; PathOutputPoint(i).Type, PathOutputPoint(i).VertType, Round(PathOutputPoint(i).LengthFromStart, 4), Round(PathOutputPoint(i).AngleToNext, 4), Round(PathOutputPoint(i).Radius3P, 4)
    'Next
    
    SetAngleAdjustMM
    TotalPathOutLength = PathOutLength
    
    'Debug.Print ">2--------------------------------------------------"
    For I = 1 To PathOutputPointCount
        PathOutputPoint(I).LengthFromStart = PathOutputPoint(I).LengthFromStart + Device_HeadDistance
    '    Debug.Print i; PathOutputPoint(i).Type, PathOutputPoint(i).VertType, Round(PathOutputPoint(i).LengthFromStart, 4), Round(PathOutputPoint(i).AngleToNext, 4), Round(PathOutputPoint(i).Radius3P, 4)
    Next
    
    PathOutputPointCount0 = PathOutputPointCount
    For I = 1 To PathOutputPointCount0
        'PathOutputPoint(I).Type = 0
        PathOutputPoint(I).VertType = 0
        'If I = 1 Or I = PathOutputPointCount0 Or Abs(PathOutputPoint(I).AngleToNext) >= Device_VertMinAngle Or PathOutputPoint(I).Type >= 88888 Then
        If PathOutputPoint(I).Type = 88888 Or PathOutputPoint(I).Type = 99999 Or PathOutputPoint(I).Type = -99999 Or Abs(PathOutputPoint(I).AngleToNext) >= Device_VertMinAngle Then
            PathOutputPointCount = PathOutputPointCount + 1
            ReDim Preserve PathOutputPoint(PathOutputPointCount)
            
            PathOutputPoint(PathOutputPointCount).ux = PathOutputPoint(I).ux
            PathOutputPoint(PathOutputPointCount).uy = PathOutputPoint(I).uy
                        
            'If I = 1 Then
            If PathOutputPoint(I).Type = 88888 Then
                PathOutputPoint(PathOutputPointCount).Type = 3 '线段起始端
            'ElseIf I = PathOutputPointCount0 Then
            ElseIf PathOutputPoint(I).Type = 99999 Then
                PathOutputPoint(PathOutputPointCount).Type = 4 '线段末端
            ElseIf PathOutputPoint(I).Type = -99999 Then
                PathOutputPoint(PathOutputPointCount).Type = 5 '(美国型材最后扩充的)线段末端
            Else
                PathOutputPoint(PathOutputPointCount).Type = 0
            End If
            
            If PathOutputPoint(I).AngleToNext < 0 Then
                PathOutputPoint(PathOutputPointCount).VertType = 1
            Else
                PathOutputPoint(PathOutputPointCount).VertType = 2
            End If
            
            PathOutputPoint(PathOutputPointCount).LengthFromStart = PathOutputPoint(I).LengthFromStart - Device_HeadDistance
            
            PathOutputPoint(PathOutputPointCount).AngleToNext = PathOutputPoint(I).AngleToNext
            
            PathOutputPoint(I).VertType = -1
            PathOutputPoint(I).Radius3P = 0
            
            'PathOutputPoint(I).LengthFromStart = PathOutputPoint(I).LengthFromStart + Device_TurnPointOffsetMM '折角点偏移补偿
            
            'PathOutputPoint(I).AngleToNext = 0 '不折角时设为0， 折角时取消此语句
       End If
       
    Next
    
    '排序
    For I = 1 To PathOutputPointCount
        For j = I + 1 To PathOutputPointCount
            If PathOutputPoint(I).LengthFromStart > PathOutputPoint(j).LengthFromStart Then
                t = PathOutputPoint(I)
                PathOutputPoint(I) = PathOutputPoint(j)
                PathOutputPoint(j) = t
            End If
        Next
    Next
       
    'Debug.Print ">3--------------------------------------------------"
    ff = FreeFile
    Open "OutputPointList.txt" For Output As #ff
    Print #ff, "序号"; Tab(7); "类型"; Tab(12); "子类"; Tab(18); "总长度"; Tab(35); "点距"; Tab(54); "夹角"; Tab(72); "半径"
    For I = 1 To PathOutputPointCount
        Print #ff, I; Tab(8); PathOutputPoint(I).Type; Tab(16); PathOutputPoint(I).VertType; Tab(24); Round(PathOutputPoint(I).LengthFromStart, 4); Tab(44); Round(IIf(I = 1, 0, PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart), 4); Tab(64); Round(PathOutputPoint(I).AngleToNext, 4); Tab(84); Round(PathOutputPoint(I).Radius3P, 4)
    Next
    Close #ff
End Sub

Sub ConvertAllToSegments()
    Dim I As Long, start_id As Long, group_id As Long
    
    PathOutputPointCount = 0
    StopOutput = False
    OutputAllPath OutputSegments
    
    PointCount = PathOutputPointCount
    ReDim PointList(PointCount)
    
    SegmentCount = PointCount - 1
    ReDim SegmentList(SegmentCount)
    
    ArcCount = 0
    SPLineCount = 0
    
    start_id = 1
    group_id = 1
    
    For I = 1 To PointCount
        PointList(I).X = PathOutputPoint(I).ux
        PointList(I).Y = PathOutputPoint(I).uy
        PointList(I).id = I
        PointList(I).Layer = 1
        PointList(I).body_id = group_id
        PointList(I).group_id = group_id
        
        If I > 1 Then
            If PathOutputPoint(I - 1).Type = 1 Then 'end point
                If PointList(start_id).X = PointList(I - 1).X And PointList(start_id).Y = PointList(I - 1).Y Then
                    SegmentCount = SegmentCount + 1
                    ReDim Preserve SegmentList(SegmentCount)
                    
                    SegmentList(I - 1).point0_id = I - 1
                    SegmentList(I - 1).point1_id = start_id
                    SegmentList(I - 1).Layer = 1
                    SegmentList(I - 1).body_id = group_id
                    SegmentList(I - 1).group_id = group_id
                End If
                
                start_id = I
                group_id = group_id + 1
            Else
                SegmentList(I - 1).point0_id = I - 1
                SegmentList(I - 1).point1_id = I
                SegmentList(I - 1).Layer = 1
                SegmentList(I - 1).body_id = group_id
                SegmentList(I - 1).group_id = group_id
            End If
        End If
    Next
    
    If PointList(start_id).X = PointList(PointCount).X And PointList(start_id).Y = PointList(PointCount).Y Then
        SegmentCount = SegmentCount + 1
        ReDim Preserve SegmentList(SegmentCount)
        
        SegmentList(SegmentCount).point0_id = PointCount
        SegmentList(SegmentCount).point1_id = start_id
        SegmentList(SegmentCount).Layer = 1
        SegmentList(SegmentCount).body_id = group_id
        SegmentList(SegmentCount).group_id = group_id
    End If
End Sub

Sub ReversePath()
    Dim I As Long, sid As Long, cid As Long, spid As Long, t As Integer
    
    FrmMain.ProgressBar.value = 0
    FrmMain.ProgressBar.Max = Max(SegmentCount + ArcCount + SPLineCount, 1)
    
    For I = 1 To SegmentCount
        FrmMain.ProgressBar.value = I
        
        If SegmentList(I).selected = False Then
            ReverseDirection I, 0, True, True
        End If
    Next
    
    For I = 1 To ArcCount
        FrmMain.ProgressBar.value = SegmentCount + I
        
        ReverseDirection I, 1, True, True
    Next
    
    For I = 1 To SPLineCount
        FrmMain.ProgressBar.value = SegmentCount + ArcCount + I
        ReverseDirection I, 2, True, True
    Next
End Sub

Sub FindNewPointByCutRadius(ByVal id As Long, ByVal last_id As Long, ByVal next_id As Long, ByVal r As Double, ByRef X As Double, ByRef Y As Double)
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim cx As Double, cy As Double, gx0 As Double, gy0 As Double, gx1 As Double, gy1 As Double
    Dim dx0 As Double, dy0 As Double, dx1 As Double, dy1 As Double
    Dim a0 As Double, b0 As Double, a1 As Double, b1 As Double
    Dim fa0 As Double, fb0 As Double, fa1 As Double, fb1 As Double, d As Double
    'Dim ga0 As Double, gb0 As Double, ga1 As Double, gb1 As Double
    Dim vx0 As Double, vy0 As Double, vx1 As Double, vy1 As Double
    
    'Dim px1 As Single, py1 As Single
    'Dim px2 As Single, py2 As Single
    
    x0 = PathOutputPoint(last_id).ux
    y0 = PathOutputPoint(last_id).uy
    
    x1 = PathOutputPoint(id).ux
    y1 = PathOutputPoint(id).uy
    
    x2 = PathOutputPoint(next_id).ux
    y2 = PathOutputPoint(next_id).uy
    
    dx0 = Round(x1 - x0, 8)
    dy0 = Round(y1 - y0, 8)
    dx1 = Round(x2 - x1, 8)
    dy1 = Round(y2 - y1, 8)
    
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
        
    ElseIf dx0 = 0 And dy1 = 0 Then
        'L0: x = x0 '线段方程
        'L1: y = y1
        
        'F0: x = x0 + d * r '平行线方程, d 确定左右侧
        'F1: y = y1 + d * r
                
        'F0 & F1 '平行线相交求圆心
        cx = x0 - d * r * Sgn(dy0)
        cy = y1 + d * r * Sgn(dx1)
        
    ElseIf dy0 = 0 And dx1 = 0 Then
        'L0: y = y0 '线段方程
        'L1: x = x1
        
        'F0: y = y0 + d * r '平行线方程, d 确定左右侧
        'F1: x = x1 + d * r
                
        'F0 & F1 '平行线相交求圆心
        cx = x1 - d * r * Sgn(dy1)
        cy = y0 + d * r * Sgn(dx0)
        
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
    
    End If
    
    d = Sqr((x1 - cx) ^ 2 + (y1 - cy) ^ 2)
    X = r * (x1 - cx) / d + cx
    Y = r * (y1 - cy) / d + cy
End Sub

Sub RoundPathInnerCorner()
    Dim I As Long, I1 As Long, j As Long, k As Long, ux As Double, uy As Double, X As Single, Y As Single, lfs As Double
    Dim ux1 As Double, uy1 As Double, ux2 As Double, uy2 As Double, Ang As Double, n As Long, dX As Double, dy As Double
    Dim begin_next_piece As Boolean
    
    If Device_CutRadiusMM < 0.001 Then
        Exit Sub
    End If
    
    For I = 1 To PathOutputPointCount - 1
        If PathOutputPoint(I).AngleToNext < -90.1 And PathOutputPoint(I).Type = 88888 Then
            I1 = 0
            For j = I + 1 To PathOutputPointCount
                If PathOutputPoint(j).Type = 99999 Then
                    If PathOutputPoint(I).ux = PathOutputPoint(j).ux And PathOutputPoint(I).uy = PathOutputPoint(j).uy Then 'closed
                        I1 = j
                        Exit For
                    End If
                End If
            Next
            
            If I1 > 0 Then
                FindNewPointByCutRadius I, I1 - 1, I + 1, Device_CutRadiusMM, ux, uy
                k = j
                
Debug.Print ">>>x,y="; ux; uy
ConvertUserToPath ux, uy, X, Y
MarkPathinnercorner X, Y

                PathOutputPoint(I).ux = ux
                PathOutputPoint(I).uy = uy
                
                PathOutputPoint(I1).ux = ux
                PathOutputPoint(I1).uy = uy
                
                '往前
                For j = I + 1 To PathOutputPointCount
                    If j = PathOutputPointCount Or Abs(PathOutputPoint(j).AngleToNext) >= Device_VertMinAngle Or PathOutputPoint(j).Type = 99999 Then
                        ux1 = PathOutputPoint(j).ux
                        uy1 = PathOutputPoint(j).uy
                        Exit For
                    Else
                        ux1 = PathOutputPoint(j).ux
                        uy1 = PathOutputPoint(j).uy
                        ux2 = PathOutputPoint(j + 1).ux
                        uy2 = PathOutputPoint(j + 1).uy
                        
                        Ang = GetAngle(ux, uy, ux1, uy1, ux2, uy2)
                        If Abs(Ang) < Device_VertMinAngle Then
                            Exit For
                        End If
                    End If
                Next
                
                n = j - (I + 1) + 1
                dX = (ux1 - ux) / n
                dy = (uy1 - uy) / n
                For k = 1 To n - 1
                    PathOutputPoint(I + k).ux = PathOutputPoint(I).ux + k * dX
                    PathOutputPoint(I + k).uy = PathOutputPoint(I).uy + k * dy
                Next
                
                '往后
                For j = I1 - 1 To 1 Step -1
                    If j = 1 Or Abs(PathOutputPoint(j).AngleToNext) >= Device_VertMinAngle Or PathOutputPoint(j).Type = 99999 Then
                        ux1 = PathOutputPoint(j).ux
                        uy1 = PathOutputPoint(j).uy
                        Exit For
                    Else
                        ux1 = PathOutputPoint(j).ux
                        uy1 = PathOutputPoint(j).uy
                        ux2 = PathOutputPoint(j - 1).ux
                        uy2 = PathOutputPoint(j - 1).uy
                        
                        Ang = GetAngle(ux2, uy2, ux1, uy1, ux, uy)
                        If Abs(Ang) < Device_VertMinAngle Then
                            Exit For
                        End If
                    End If
                Next
                
                n = (I1 - 1) - j + 1
                dX = (ux1 - ux) / n
                dy = (uy1 - uy) / n
                For k = 1 To n - 1
                    PathOutputPoint(I1 - k).ux = PathOutputPoint(I).ux + k * dX
                    PathOutputPoint(I1 - k).uy = PathOutputPoint(I).uy + k * dy
                Next
            End If
            
        ElseIf PathOutputPoint(I).AngleToNext < -90.1 And Abs(PathOutputPoint(I).Type) <> 99999 Then
            FindNewPointByCutRadius I, I - 1, I + 1, Device_CutRadiusMM, ux, uy
            
Debug.Print "i,x,y="; I, Round(ux, 4); Round(uy, 4)
ConvertUserToPath ux, uy, X, Y
MarkPathinnercorner X, Y

            PathOutputPoint(I).ux = ux
            PathOutputPoint(I).uy = uy
            
            '往前
            For j = I + 1 To PathOutputPointCount
                If j = PathOutputPointCount Or Abs(PathOutputPoint(j).AngleToNext) >= Device_VertMinAngle Or PathOutputPoint(j).Type = 99999 Then
                    ux1 = PathOutputPoint(j).ux
                    uy1 = PathOutputPoint(j).uy
                    Exit For
                Else
                    ux1 = PathOutputPoint(j).ux
                    uy1 = PathOutputPoint(j).uy
                    ux2 = PathOutputPoint(j + 1).ux
                    uy2 = PathOutputPoint(j + 1).uy
                    
                    Ang = GetAngle(ux, uy, ux1, uy1, ux2, uy2)
                    If Abs(Ang) < Device_VertMinAngle Then
                        Exit For
                    End If
                End If
            Next
            
            n = j - (I + 1) + 1
            dX = (ux1 - ux) / n
            dy = (uy1 - uy) / n
            For k = 1 To n - 1
                PathOutputPoint(I + k).ux = PathOutputPoint(I).ux + k * dX
                PathOutputPoint(I + k).uy = PathOutputPoint(I).uy + k * dy
            Next
            
            '往后
            For j = I - 1 To 1 Step -1
                If j = 1 Or Abs(PathOutputPoint(j).AngleToNext) >= Device_VertMinAngle Or PathOutputPoint(j).Type = 99999 Then
                    ux1 = PathOutputPoint(j).ux
                    uy1 = PathOutputPoint(j).uy
                    Exit For
                Else
                    ux1 = PathOutputPoint(j).ux
                    uy1 = PathOutputPoint(j).uy
                    ux2 = PathOutputPoint(j - 1).ux
                    uy2 = PathOutputPoint(j - 1).uy
                    
                    Ang = GetAngle(ux2, uy2, ux1, uy1, ux, uy)
                    If Abs(Ang) < Device_VertMinAngle Then
                        Exit For
                    End If
                End If
            Next
            
            n = (I - 1) - j + 1
            dX = (ux1 - ux) / n
            dy = (uy1 - uy) / n
            For k = 1 To n - 1
                PathOutputPoint(I - k).ux = PathOutputPoint(I).ux + k * dX
                PathOutputPoint(I - k).uy = PathOutputPoint(I).uy + k * dy
            Next
        End If
    Next
    
'    'set LengthFormStart again
'    begin_next_piece = False
'    For I = 1 To PathOutputPointCount
'        If I = 1 Then
'            lfs = 0
'        ElseIf PathOutputPoint(I).Type = 88888 And begin_next_piece = True Then
'            lfs = lfs + Device_DoneDistance
'        Else
'            If PathOutputPoint(I).Type = 99999 Then
'                begin_next_piece = True
'            End If
'            lfs = lfs + Sqr((PathOutputPoint(I).ux - PathOutputPoint(I - 1).ux) ^ 2 + (PathOutputPoint(I).uy - PathOutputPoint(I - 1).uy) ^ 2)
'        End If
'        PathOutputPoint(I).LengthFromStart = lfs
'    Next
End Sub

Sub SetAngleAdjustMM()
    Dim I As Long, lfs As Double
    Dim begin_next_piece As Boolean
    
    Dim lx() As Double, ly() As Double, n As Long, i0 As Long
    ReDim lx(PathOutputPointCount) As Double, ly(PathOutputPointCount) As Double
    
    ReDim AdjustMM_0(PathOutputPointCount) As Double, AdjustMM_1(PathOutputPointCount) As Double
    
    For I = 1 To PathOutputPointCount
        AdjustMM_0(I) = 0
        AdjustMM_1(I) = 0
    Next
    
    'For I = 2 To PathOutputPointCount - 1'原来补偿算法，首尾端点补偿减半
    For I = 1 To PathOutputPointCount
        If PathOutputPoint(I).AngleToNext <= -Device_VertMinAngle Then
            If VT1 = 1 Then '内角为<= -Device_VertMinAngle
                AdjustMM_0(I) = Device_InnerAngleAdjustMM / 2
                AdjustMM_1(I) = Device_InnerAngleAdjustMM / 2
                
            Else 'VT1=2,内角为>=Device_VertMinAngle
                AdjustMM_0(I) = Device_OuterAngleAdjustMM / 2
                AdjustMM_1(I) = Device_OuterAngleAdjustMM / 2
            End If
                
        ElseIf PathOutputPoint(I).AngleToNext >= Device_VertMinAngle Then
            If VT1 = 1 Then '内角为<= -Device_VertMinAngle
                AdjustMM_0(I) = Device_OuterAngleAdjustMM / 2
                AdjustMM_1(I) = Device_OuterAngleAdjustMM / 2
                
            Else 'VT1=2,内角为>=Device_VertMinAngle
                AdjustMM_0(I) = Device_InnerAngleAdjustMM / 2
                AdjustMM_1(I) = Device_InnerAngleAdjustMM / 2
            End If
            
        Else
            AdjustMM_0(I) = Device_MaterialThickMM * PathOutputPoint(I).AngleToNext * PI_180 / 2
            AdjustMM_1(I) = Device_MaterialThickMM * PathOutputPoint(I).AngleToNext * PI_180 / 2
        
        End If
    Next
    
    i0 = 0
    For I = 1 To PathOutputPointCount
        If PathOutputPoint(I).Type = 88888 Then
            i0 = I
            lx(1) = PathOutputPoint(I).ux
            ly(1) = PathOutputPoint(I).uy
            
            AdjustMM_1(I) = AdjustMM_1(I) + Device_StartPointAdjustMM
        ElseIf PathOutputPoint(I).Type = 99999 Then
            If i0 > 0 Then
                n = I - i0 + 1
                lx(n) = PathOutputPoint(I).ux
                ly(n) = PathOutputPoint(I).uy
                
                If lx(1) = lx(n) And ly(1) = ly(n) Then 'closed
                    If IsPathClockwise(n, lx, ly) = True Then
                        AdjustMM_0(i0 + 1) = AdjustMM_0(i0 + 1) + Device_OuterLineTerminalAdjustMM / 2
                        AdjustMM_1(n - 1) = AdjustMM_1(n - 1) + Device_OuterLineTerminalAdjustMM / 2
                    Else
                        AdjustMM_0(i0 + 1) = AdjustMM_0(i0 + 1) + Device_InnerLineTerminalAdjustMM / 2
                        AdjustMM_1(n - 1) = AdjustMM_1(n - 1) + Device_InnerLineTerminalAdjustMM / 2
                    End If
                End If
                i0 = 0
                
                AdjustMM_0(I) = AdjustMM_0(I) + Device_EndPointAdjustMM
            End If
                    
        ElseIf i0 > 0 Then
            lx(I - i0 + 1) = PathOutputPoint(I).ux
            ly(I - i0 + 1) = PathOutputPoint(I).uy
        End If
    Next
    
    'set LengthFormStart again
    begin_next_piece = False
    For I = 1 To PathOutputPointCount
        If I = 1 Then
            lfs = 0
        ElseIf PathOutputPoint(I).Type = 88888 And begin_next_piece = True Then
            lfs = lfs + Device_DoneDistance
        ElseIf PathOutputPoint(I).Type = -99999 Then
            lfs = lfs + Device_ExtendMM
        Else
            If PathOutputPoint(I).Type = 99999 Then
                begin_next_piece = True
            End If
            lfs = lfs + AdjustMM_1(I - 1) + AdjustMM_0(I) + Sqr((PathOutputPoint(I).ux - PathOutputPoint(I - 1).ux) ^ 2 + (PathOutputPoint(I).uy - PathOutputPoint(I - 1).uy) ^ 2)
        End If
        PathOutputPoint(I).LengthFromStart = lfs
    Next
    PathOutLength = lfs
End Sub

Function GetOutputEndPointAngle(ByVal id As Long) As Double
    Dim I As Long
    Dim x0 As Double, y0 As Double, X As Double, Y As Double, x1 As Double, y1 As Double
    
    X = PathOutputPoint(id).ux
    Y = PathOutputPoint(id).uy
    
    For I = 1 To SegmentCount
        If PointList(SegmentList(I).point1_id).X = X And PointList(SegmentList(I).point1_id).Y = Y Then
            x0 = PointList(SegmentList(I).point0_id).X
            y0 = PointList(SegmentList(I).point0_id).Y
        End If
        If PointList(SegmentList(I).point0_id).X = X And PointList(SegmentList(I).point0_id).Y = Y Then
            x1 = PointList(SegmentList(I).point1_id).X
            y1 = PointList(SegmentList(I).point1_id).Y
        End If
    Next
    
    GetOutputEndPointAngle = GetAngle(x0, y0, X, Y, x1, y1)
End Function

Sub CheckOuterAndInnerLines()
    Dim id As Long, I As Long, j As Long, k As Long, q As Long, u As Long, v As Long
        
    Dim start_id As Long, start_id_list() As Long, start_id_count As Long
    Dim list_n() As Long, Max_X() As Double, Max_Y() As Double, Min_X() As Double, Min_Y() As Double
    
    Dim X As Double, Y As Double, x0 As Double, y0 As Double, x1 As Double, y1 As Double, px As Double
    
    'Dim lx() As Double, ly() As Double, cur_list() As Long, n As Long, clr As Long
    'ReDim lx(SegmentCount), ly(SegmentCount), cur_list(SegmentCount)

    For id = 1 To SegmentCount
        SegmentList(id).selected = False
    Next
    
    start_id = 1
    start_id_count = 0
    ReDim list_n(0) As Long, Max_X(0) As Double, Max_Y(0) As Double, Min_X(0) As Double, Min_Y(0) As Double
    
    Do While start_id <= SegmentCount
        start_id_count = start_id_count + 1
        ReDim Preserve start_id_list(start_id_count), list_n(start_id_count), Max_X(start_id_count), Max_Y(start_id_count), Min_X(start_id_count), Min_Y(start_id_count)
        start_id_list(start_id_count) = start_id
        
        id = start_id
        I = 0
        Do
            SegmentList(id).selected = True
            
            If I = 0 Then
                Min_X(start_id_count) = PointList(SegmentList(id).point0_id).X
                Max_X(start_id_count) = PointList(SegmentList(id).point0_id).X
                Min_Y(start_id_count) = PointList(SegmentList(id).point0_id).Y
                Max_Y(start_id_count) = PointList(SegmentList(id).point0_id).Y
            Else
                Min_X(start_id_count) = IIf(Min_X(start_id_count) > PointList(SegmentList(id).point0_id).X, PointList(SegmentList(id).point0_id).X, Min_X(start_id_count))
                Max_X(start_id_count) = IIf(Max_X(start_id_count) < PointList(SegmentList(id).point0_id).X, PointList(SegmentList(id).point0_id).X, Max_X(start_id_count))
                Min_Y(start_id_count) = IIf(Min_Y(start_id_count) > PointList(SegmentList(id).point0_id).Y, PointList(SegmentList(id).point0_id).Y, Min_Y(start_id_count))
                Max_Y(start_id_count) = IIf(Max_Y(start_id_count) < PointList(SegmentList(id).point0_id).Y, PointList(SegmentList(id).point0_id).Y, Max_Y(start_id_count))
            End If
            
            Min_X(start_id_count) = IIf(Min_X(start_id_count) > PointList(SegmentList(id).point1_id).X, PointList(SegmentList(id).point1_id).X, Min_X(start_id_count))
            Max_X(start_id_count) = IIf(Max_X(start_id_count) < PointList(SegmentList(id).point1_id).X, PointList(SegmentList(id).point1_id).X, Max_X(start_id_count))
            Min_Y(start_id_count) = IIf(Min_Y(start_id_count) > PointList(SegmentList(id).point1_id).Y, PointList(SegmentList(id).point1_id).Y, Min_Y(start_id_count))
            Max_Y(start_id_count) = IIf(Max_Y(start_id_count) < PointList(SegmentList(id).point1_id).Y, PointList(SegmentList(id).point1_id).Y, Max_Y(start_id_count))
            I = I + 1
            
            For j = 1 To SegmentCount
                If SegmentList(j).selected = False And SegmentList(j).point0_id = SegmentList(id).point1_id Then
                    id = j
                    Exit For
                End If
            Next
            If j > SegmentCount Or id = start_id Then
                list_n(start_id_count) = I
                Exit Do
            End If
        Loop
                
        For id = 1 To SegmentCount
            If SegmentList(id).selected = False Then
                Exit For
            End If
        Next
        start_id = id
    Loop
    
    For I = 1 To start_id_count
        'Debug.Print i, start_id_list(i), list_n(i), Round(Min_X(i), 2), Round(Max_X(i), 2), Round(Min_Y(i), 2), Round(Max_Y(i), 2)
        X = PointList(SegmentList(start_id_list(I)).point0_id).X
        Y = PointList(SegmentList(start_id_list(I)).point0_id).Y
        
        v = 0
        For j = 1 To start_id_count
            If I <> j Then
                If X >= Min_X(j) And X <= Max_X(j) And Y >= Min_Y(j) And Y <= Max_Y(j) Then
                    u = 0
                    For k = start_id_list(j) To start_id_list(j) + list_n(j) - 1
                        x0 = PointList(SegmentList(k).point0_id).X
                        y0 = PointList(SegmentList(k).point0_id).Y
                        x1 = PointList(SegmentList(k).point1_id).X
                        y1 = PointList(SegmentList(k).point1_id).Y
                        
                        If y0 <> y1 Then
                            If Y >= Min(y0, y1) Then
                                If Y < Max(y0, y1) Then
                                    px = (Y - y0) * (x1 - x0) / (y1 - y0) + x0
                                    If px > X Then
                                        u = u + 1
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    If u Mod 2 = 1 Then
                        v = v + 1
                    End If
                End If
            End If
        Next
                
        If v Mod 2 = 0 Then
            'Debug.Print "Outer Line"
            v = IsSegmentsClockwise(SegmentList(start_id_list(I)).point0_id)
            If v = -1 Then
                ReverseDirection start_id_list(I), 0, True, True
                
                For j = start_id_list(I) To start_id_list(I) + list_n(I) - 1
                    SegmentList(j).color = RGB(255, 0, 255)
                Next
                DirectionChanged = True
            End If
        Else
            'Debug.Print "Inner Line"
            v = IsSegmentsClockwise(SegmentList(start_id_list(I)).point0_id)
            If v = 1 Then
                ReverseDirection start_id_list(I), 0, True, True
                
                For j = start_id_list(I) To start_id_list(I) + list_n(I) - 1
                    SegmentList(j).color = RGB(255, 0, 0)
                Next
                DirectionChanged = True
            End If
        End If
    Next
End Sub


