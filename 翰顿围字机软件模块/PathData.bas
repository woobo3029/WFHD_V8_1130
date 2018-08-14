Attribute VB_Name = "PathData"
Option Explicit

Enum PointStatus
    Normal
    temp
    Used
    Deleted
End Enum

Public Type Path_Point
    id As Long
    body_id As Long
    group_id As Long
    X As Double
    Y As Double
    z As Double
    xp As Double 'X轴的补偿值
    yp As Double 'Y轴的补偿值
    zp As Double 'Z轴的补偿值
    color As Long
    Layer As Long
    Type As Long
    method As Long
    arc_id As Long
    action As ActionType
    stay_time As Long
    v_down As Double
    v_up As Double
    selected As Boolean
    status As PointStatus
    HoleType As HoleType
End Type

Public Type Path_Segment
    id As Long
    body_id As Long
    group_id As Long
    point0_id As Long
    point1_id As Long
    color As Long
    Layer As Long
    Type As Long
    arc_id As Long
    action As ActionType
    selected As Boolean
    clockwise As Boolean
End Type

Public Type Path_Arc
    id As Long
    body_id As Long
    group_id As Long
    X As Double
    Y As Double
    z As Double
    a As Double
    b As Double
    ax_angle As Double
    start_angle As Double
    end_angle As Double
    point0_id As Long
    pointm_id As Long
    point1_id As Long
    color As Long
    Layer As Long
    Type As Long
    point_id As Long
    action As ActionType
    selected As Boolean
End Type

Public Type PolygonPoint
    X As Double
    Y As Double
End Type

Public Type Path_SPLine
    id As Long
    body_id As Long
    group_id As Long
    vertex_count As Long
    vertex_id() As Long
    point0_id As Long
    point1_id As Long
    color As Long
    Layer As Long
    Type As Long
    action As ActionType
    segment_between_points As Long '应可单独设置
    segment_point_count As Long
    segment_point() As PolygonPoint
    selected As Boolean
End Type

Public Type BodyInGroup
    body_id As Long
    group_id As Long
End Type

Public Enum PointType
    NormalPoint = 0
    RoundCornerPoint
    ArcPoint
    BoxPoint
    SPLinePoint
End Enum

Public Enum PointMethod
    No_Method
    RoundedCorner
End Enum

Public Enum SegmentType
    NormalSegment
    ReplacedByArc
End Enum

Public Enum ArcType
    CircleCR
    'Circle3P
    'Circle2PR
    Ellipse
    ReplacSegment
    RoundedCorner
End Enum

Public Enum ActionType
    No_Action
    StartDropping
    BothStartAndStopDropping
    Dropping
    StopDropping
    'PointDropping
End Enum

Public Enum HoleType
    No_Hole
    HoleType1
    HoleType2
    HoleType3
End Enum

Public catched_pid As Long, catched_sid As Long, catched_cid As Long, catched_aid As Long, catched_spid As Long, catched_bodyid As Long, catched_groupid As Long

Public CurPoint As Path_Point, LastPoint As Path_Point, CurPointIndex As Long
Public CurSegment As Path_Segment, CurSegmentIndex As Long
Public CurArc As Path_Arc, LastArc As Path_Arc, CurArcIndex As Long
Public CurBodyID As Long, CurBodyCenterX As Double, CurBodyCenterY As Double, BodyList() As BodyInGroup, MaxBodyID As Long
Public CurGroupID As Long, CurGroupCenterX As Double, CurGroupCenterY As Double, CurUnittedGroupID As Long, CurArrayedGroupID As Long
Public CurBoxFirstSegmentID As Long

Public PointList() As Path_Point
Public SegmentList() As Path_Segment
Public ArcList() As Path_Arc
Public SPLineList() As Path_SPLine

Public PointCount As Long
Public SegmentCount As Long
Public ArcCount As Long
Public BodyCount As Long
Public SPLineCount As Long
Public GroupCount As Long

Sub AddPoint(ByVal X As Double, ByVal Y As Double, ByVal z As Double, ByVal Layer As Integer, ByVal ptype As PointType, Optional BodyID As Long = 0)
    PointCount = PointCount + 1
    ReDim Preserve PointList(PointCount)
    
    PointList(PointCount).id = PointCount
    PointList(PointCount).X = X
    PointList(PointCount).Y = Y
    PointList(PointCount).z = z
    PointList(PointCount).Layer = Layer
    PointList(PointCount).Type = ptype
    PointList(PointCount).body_id = BodyID
    PointList(PointCount).group_id = BodyID
    
    LastPoint = CurPoint
    
    CurPoint = PointList(PointCount)
End Sub

Sub AddSegment(ByVal p0_id As Long, ByVal p1_id As Long)
    Dim body_id As Long, group_id As Long
    
    SegmentCount = SegmentCount + 1
    ReDim Preserve SegmentList(SegmentCount)
    
    SegmentList(SegmentCount).id = SegmentCount
    SegmentList(SegmentCount).point0_id = p0_id
    SegmentList(SegmentCount).point1_id = p1_id

    If PointList(p0_id).body_id = 0 Then
        BodyCount = BodyCount + 1
        body_id = BodyCount
        group_id = BodyCount
        PointList(p0_id).body_id = body_id
        PointList(p0_id).group_id = group_id
    Else
        body_id = PointList(p0_id).body_id
        group_id = PointList(p0_id).group_id
    End If
    SegmentList(SegmentCount).body_id = body_id
    SegmentList(SegmentCount).group_id = group_id
    
    If PointList(p1_id).body_id > 0 Then
        ReplaceBodyID PointList(p1_id).body_id, body_id
        ReplaceGroupID PointList(p1_id).group_id, group_id
    End If
    PointList(p1_id).body_id = body_id
    PointList(p1_id).group_id = group_id
    
    If PointList(p0_id).Layer = PointList(p1_id).Layer Then
        SegmentList(SegmentCount).Layer = PointList(p0_id).Layer
    Else
        SegmentList(SegmentCount).Layer = 0
    End If
End Sub
    
Sub DeletePoint(ByVal id As Long)
    Dim i As Long, j As Long
    Dim seg0_id As Long, seg1_id As Long, seg_p0_id As Long, seg_p1_id As Long
    
    If id = 0 Then
        Exit Sub
    End If
    
    If PointList(id).arc_id > 0 Then
        If ArcList(PointList(id).arc_id).point0_id <> id And _
           ArcList(PointList(id).arc_id).pointm_id <> id And _
           ArcList(PointList(id).arc_id).point1_id <> id Then
            If PointList(id).method = PointMethod.RoundedCorner Then
                DeletePoint ArcList(PointList(id).arc_id).point0_id
                DeletePoint ArcList(PointList(id).arc_id).point1_id
                DeletePoint ArcList(PointList(id).arc_id).pointm_id
            End If
        
            DeleteArc PointList(id).arc_id
        End If
    End If
    
    For i = id To PointCount - 1
        PointList(i) = PointList(i + 1)
        PointList(i).id = i
    Next
    PointCount = PointCount - 1
    ReDim Preserve PointList(PointCount)
    
    For i = 1 To SegmentCount
        If SegmentList(i).point0_id = id Then
            'seg_p1_id = SegmentList(I).point1_id
            seg1_id = i
        End If
        
        If SegmentList(i).point1_id = id Then
            'seg_p0_id = SegmentList(I).point0_id
            seg0_id = i
        End If
    Next
    
    If seg0_id > 0 And seg1_id > 0 Then 'Middle Point
        'SegmentList(seg0_id).point1_id = seg_p1_id
        
        DeleteSegment seg1_id
    ElseIf seg0_id > 0 Then 'End Point
        DeleteSegment seg0_id
    ElseIf seg1_id > 0 Then 'Start Point
        DeleteSegment seg1_id
    End If
    
    id = Min(id, PointCount)
    
    For i = 1 To SegmentCount
        If SegmentList(i).point0_id > id Then
            SegmentList(i).point0_id = SegmentList(i).point0_id - 1
        End If

        If SegmentList(i).point1_id > id Then
            SegmentList(i).point1_id = SegmentList(i).point1_id - 1
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).point_id > id Then
            ArcList(i).point_id = ArcList(i).point_id - 1
        End If
        
        If ArcList(i).point0_id > id Then
            ArcList(i).point0_id = ArcList(i).point0_id - 1
        End If
        
        If ArcList(i).point1_id > id Then
            ArcList(i).point1_id = ArcList(i).point1_id - 1
        End If
        
        If ArcList(i).pointm_id > id Then
            ArcList(i).pointm_id = ArcList(i).pointm_id - 1
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).point0_id > id Then
            SPLineList(i).point0_id = SPLineList(i).point0_id - 1
        End If

        If SPLineList(i).point1_id > id Then
            SPLineList(i).point1_id = SPLineList(i).point1_id - 1
        End If
        For j = 0 To SPLineList(i).vertex_count - 1
            If SPLineList(i).vertex_id(j) > id Then
                SPLineList(i).vertex_id(j) = SPLineList(i).vertex_id(j) - 1
            End If
        Next
    Next
    
    For i = 1 To OutputStartPointList.Count
        If OutputStartPointList.point_id(i) = id Then
            For j = i + 1 To OutputStartPointList.Count
                OutputStartPointList.point_id(j - 1) = OutputStartPointList.point_id(j)
                OutputStartPointList.leading_point0(j - 1) = OutputStartPointList.leading_point0(j)
                OutputStartPointList.leading_point1(j - 1) = OutputStartPointList.leading_point1(j)
            Next
            OutputStartPointList.Count = OutputStartPointList.Count - 1
            ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.Count)
            ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.Count)
            ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.Count)
            Exit For
        End If
    Next
    For i = 1 To OutputStartPointList.Count
        If OutputStartPointList.point_id(i) > id Then
            OutputStartPointList.point_id(i) = OutputStartPointList.point_id(i) - 1
        End If
        If OutputStartPointList.leading_point0(i).id > id Then
            OutputStartPointList.leading_point0(i).id = OutputStartPointList.leading_point0(i).id - 1
        End If
        If OutputStartPointList.leading_point1(i).id > id Then
            OutputStartPointList.leading_point1(i).id = OutputStartPointList.leading_point1(i).id - 1
        End If
    Next
    
    LastPoint.id = 0
    CurPoint.id = 0
End Sub

Sub DeleteSegment(ByVal id As Long)
    Dim i As Long, id0 As Long, id1 As Long
    
    id0 = SegmentList(id).point0_id
    id1 = SegmentList(id).point1_id
    
    For i = 1 To SegmentCount
        If SegmentList(i).point1_id = id0 Then
            SegmentList(i).point1_id = id1
        End If
    Next
    
    For i = id To SegmentCount - 1
        SegmentList(i) = SegmentList(i + 1)
        SegmentList(i).id = i
    Next
    SegmentCount = SegmentCount - 1
    ReDim Preserve SegmentList(SegmentCount)
End Sub

Sub InsertSegment(ByVal id As Long, ByVal X As Double, ByVal Y As Double)
    Dim p0_id As Long, p1_id As Long
    Dim x0 As Double, y0 As Double, z0 As Double, x1 As Double, y1 As Double, z1 As Double, xm As Double, ym As Double, zm As Double
    Dim lvl0 As Integer, lvl1 As Integer
    
    p0_id = SegmentList(id).point0_id
    p1_id = SegmentList(id).point1_id
        
    lvl0 = PointList(p0_id).Layer
    lvl1 = PointList(p1_id).Layer
    
    x0 = PointList(p0_id).X
    y0 = PointList(p0_id).Y
    z0 = PointList(p0_id).z
    
    x1 = PointList(p1_id).X
    y1 = PointList(p1_id).Y
    z1 = PointList(p1_id).z
    
    If Abs(x1 - x0) > Abs(y1 - y0) Then
        If (x1 - X) * (X - x0) > 0 Then
            xm = X
        Else
            If Abs(X - x1) < Abs(X - x0) Then
                xm = x0 + 0.9 * (x1 - x0)
            Else
                xm = x0 + 0.1 * (x1 - x0)
            End If
        End If
        ym = y0 + (xm - x0) / (x1 - x0) * (y1 - y0)
        zm = z0 + (xm - x0) / (x1 - x0) * (z1 - z0)
        
    Else
        If (y1 - Y) * (Y - y0) > 0 Then
            ym = Y
        Else
            If Abs(Y - y1) < Abs(Y - y0) Then
                ym = y0 + 0.9 * (y1 - y0)
            Else
                ym = y0 + 0.1 * (y1 - y0)
            End If
        End If
        xm = x0 + (ym - y0) / (y1 - y0) * (x1 - x0)
        zm = z0 + (ym - y0) / (y1 - y0) * (z1 - z0)
        
    End If
    
    'xm = x0 + (x1 - x0) / 2
    'ym = y0 + (y1 - y0) / 2
    'zm = z0 + (z1 - z0) / 2
    
    AddPoint xm, ym, zm, IIf(lvl0 = lvl1, lvl0, 0), PointList(p0_id).Type
    PointList(PointCount).body_id = PointList(p0_id).body_id
    PointList(PointCount).group_id = PointList(p0_id).group_id
    If (PointList(p0_id).action = StartDropping Or PointList(p0_id).action = Dropping) And (PointList(p1_id).action = StopDropping Or PointList(p1_id).action = Dropping) Then
        PointList(PointCount).action = Dropping
    Else
        PointList(PointCount).action = No_Action
    End If
    
    SegmentList(id).point1_id = PointCount
    'SegmentList(id).color = IIf(lvl0 = lvl1, PointList(p0_id).color, 0)
    
    AddSegment PointCount, p1_id
    SegmentList(SegmentCount).body_id = SegmentList(id).body_id
    SegmentList(SegmentCount).group_id = SegmentList(id).group_id
    SegmentList(SegmentCount).color = SegmentList(id).color
    
End Sub

Sub AddArc(X As Double, Y As Double, z As Double, a As Double, b As Double, Angle0 As Double, Angle1 As Double, pid0 As Long, pid1 As Long, pidm As Long, Layer As Long, atype As Integer)
    Dim body_id As Long, group_id As Long
    
    ArcCount = ArcCount + 1
    ReDim Preserve ArcList(ArcCount)
    
    ArcList(ArcCount).id = ArcCount
    ArcList(ArcCount).X = X
    ArcList(ArcCount).Y = Y
    ArcList(ArcCount).z = z
    ArcList(ArcCount).a = a
    ArcList(ArcCount).b = b
    ArcList(ArcCount).ax_angle = 0
    ArcList(ArcCount).start_angle = Angle0
    ArcList(ArcCount).end_angle = Angle1
    ArcList(ArcCount).point0_id = pid0
    ArcList(ArcCount).point1_id = pid1
    ArcList(ArcCount).pointm_id = pidm
    ArcList(ArcCount).Layer = Layer
    ArcList(ArcCount).Type = atype
    ArcList(ArcCount).color = LayerColor(Layer)
    
    If pid0 > 0 Then
        If PointList(pid0).body_id = 0 Then
            BodyCount = BodyCount + 1
            body_id = BodyCount
            group_id = BodyCount
        Else
            body_id = PointList(pid0).body_id
            group_id = PointList(pid0).group_id
        End If
        PointList(pid0).body_id = body_id
        PointList(pid1).body_id = body_id
        PointList(pidm).body_id = body_id
    
        PointList(pid0).group_id = group_id
        PointList(pid1).group_id = group_id
        PointList(pidm).group_id = group_id
    Else
        BodyCount = BodyCount + 1
        body_id = BodyCount
        group_id = BodyCount
    End If
    ArcList(ArcCount).body_id = body_id
    ArcList(ArcCount).group_id = group_id
End Sub

Sub DeleteArc(ByVal id As Long)
    Dim i As Long
    
    For i = id To ArcCount - 1
        ArcList(i) = ArcList(i + 1)
        ArcList(i).id = i
    Next
    ArcCount = ArcCount - 1
    ReDim Preserve ArcList(ArcCount)
    
    For i = 1 To PointCount
        If PointList(i).arc_id = id Then
            PointList(i).arc_id = 0
            If PointList(i).method = PointMethod.RoundedCorner Then
                PointList(i).method = PointMethod.No_Method
            End If
        ElseIf PointList(i).arc_id > id Then
            PointList(i).arc_id = PointList(i).arc_id - 1
        End If
    Next

    For i = 1 To SegmentCount
        If SegmentList(i).arc_id = id Then
            SegmentList(i).arc_id = 0
            If SegmentList(i).Type = SegmentType.ReplacedByArc Then
                SegmentList(i).Type = SegmentType.NormalSegment
            End If
        ElseIf SegmentList(i).arc_id > id Then
            SegmentList(i).arc_id = SegmentList(i).arc_id - 1
        End If
    Next
End Sub

Sub AddBox(lp As Path_Point, cp As Path_Point)
    Dim p0 As Path_Point, P1 As Path_Point
    Dim p_id0 As Long, p_id1 As Long, p_id2 As Long, p_id3 As Long
    Dim s_id0 As Long, s_id1 As Long, s_id2 As Long, s_id3 As Long
    
    p0 = lp
    P1 = cp
    
    BodyCount = BodyCount + 1
    
    If p0.id = 0 Then
        AddPoint p0.X, p0.Y, LayerZValue(CurLayer), CurLayer, PointType.BoxPoint
        p_id0 = PointCount
    Else
        p_id0 = p0.id
    End If
    PointList(p_id0).body_id = BodyCount
    PointList(p_id0).group_id = BodyCount
    
    AddPoint p0.X, P1.Y, LayerZValue(CurLayer), CurLayer, PointType.BoxPoint
    p_id1 = PointCount
    PointList(p_id1).body_id = BodyCount
    PointList(p_id1).group_id = BodyCount
    
    If P1.id = 0 Then
        AddPoint P1.X, P1.Y, LayerZValue(CurLayer), CurLayer, PointType.BoxPoint
        p_id2 = PointCount
    Else
        p_id2 = P1.id
    End If
    PointList(p_id2).body_id = BodyCount
    PointList(p_id2).group_id = BodyCount
    
    AddPoint P1.X, p0.Y, LayerZValue(CurLayer), CurLayer, PointType.BoxPoint
    p_id3 = PointCount
    PointList(p_id3).body_id = BodyCount
    PointList(p_id3).group_id = BodyCount

    AddSegment p_id0, p_id1
    s_id0 = SegmentCount
    SegmentList(s_id0).body_id = BodyCount
    SegmentList(s_id0).group_id = BodyCount
    
    AddSegment p_id1, p_id2
    s_id1 = SegmentCount
    SegmentList(s_id1).body_id = BodyCount
    SegmentList(s_id1).group_id = BodyCount
    
    AddSegment p_id2, p_id3
    s_id2 = SegmentCount
    SegmentList(s_id2).body_id = BodyCount
    SegmentList(s_id2).group_id = BodyCount
    
    AddSegment p_id3, p_id0
    s_id3 = SegmentCount
    SegmentList(s_id3).body_id = BodyCount
    SegmentList(s_id3).group_id = BodyCount
    
    CurBoxFirstSegmentID = s_id0
End Sub

Function GetBodyList(ByRef BodyList() As BodyInGroup) As Long
    Dim i As Long, max_body_id As Long
    
    ReDim BodyList(0) As BodyInGroup
    
    max_body_id = 0
    For i = 1 To PointCount
        If PointList(i).body_id > max_body_id Then
            max_body_id = PointList(i).body_id
            
            ReDim Preserve BodyList(max_body_id) As BodyInGroup
            BodyList(max_body_id).body_id = PointList(i).body_id
            BodyList(max_body_id).group_id = PointList(i).group_id
        Else
            If BodyList(PointList(i).body_id).body_id = 0 Then
                BodyList(PointList(i).body_id).body_id = PointList(i).body_id
                BodyList(PointList(i).body_id).group_id = PointList(i).group_id
            End If
        End If
    Next
    
    GetBodyList = max_body_id
End Function

Sub DeleteBody(ByVal id As Long)
    Dim i As Long
    
    For i = SegmentCount To 1 Step -1
        If SegmentList(i).body_id = id Then
            DeleteSegment i
        End If
    Next

    For i = ArcCount To 1 Step -1
        If ArcList(i).body_id = id Then
            DeleteArc i
        End If
    Next
        
    For i = SPLineCount To 1 Step -1
        If SPLineList(i).body_id = id Then
            DeleteSPLine i
        End If
    Next
    
    For i = PointCount To 1 Step -1
        If PointList(i).body_id = id Then
            DeletePoint i
        End If
    Next

End Sub

Sub DeleteGroup(ByVal id As Long)
    Dim i As Long
    
    For i = SegmentCount To 1 Step -1
        If SegmentList(i).group_id = id Then
            DeleteSegment i
        End If
    Next

    For i = ArcCount To 1 Step -1
        If ArcList(i).group_id = id Then
            DeleteArc i
        End If
    Next
        
    For i = SPLineCount To 1 Step -1
        If SPLineList(i).group_id = id Then
            DeleteSPLine i
        End If
    Next
    
    For i = PointCount To 1 Step -1
        If PointList(i).group_id = id Then
            DeletePoint i
        End If
    Next

End Sub

Public Function CopyBody(ByVal id As Long) As Long
    Dim i As Long, j As Long, PCount0 As Long, Count0 As Long
    
    BodyCount = BodyCount + 1
    
    PCount0 = PointCount
    For i = 1 To PCount0
        If PointList(i).body_id = id Then
            PointCount = PointCount + 1
            ReDim Preserve PointList(PointCount)
            PointList(PointCount) = PointList(i)
            PointList(PointCount).id = PointCount
            PointList(PointCount).body_id = BodyCount
            PointList(PointCount).group_id = BodyCount
            PointList(i).status = PointCount
        End If
    Next

    Count0 = SegmentCount
    For i = 1 To Count0
        If SegmentList(i).body_id = id Then
            SegmentCount = SegmentCount + 1
            ReDim Preserve SegmentList(SegmentCount)
            SegmentList(SegmentCount) = SegmentList(i)
            SegmentList(SegmentCount).id = SegmentCount
            SegmentList(SegmentCount).point0_id = PointList(SegmentList(SegmentCount).point0_id).status
            SegmentList(SegmentCount).point1_id = PointList(SegmentList(SegmentCount).point1_id).status
            SegmentList(SegmentCount).body_id = BodyCount
            SegmentList(SegmentCount).group_id = BodyCount
       End If
    Next

    Count0 = ArcCount
    For i = 1 To Count0
        If ArcList(i).body_id = id Then
            ArcCount = ArcCount + 1
            ReDim Preserve ArcList(ArcCount)
            ArcList(ArcCount) = ArcList(i)
            ArcList(ArcCount).id = ArcCount
            ArcList(ArcCount).point_id = PointList(ArcList(ArcCount).point_id).status
            ArcList(ArcCount).point0_id = PointList(ArcList(ArcCount).point0_id).status
            ArcList(ArcCount).point1_id = PointList(ArcList(ArcCount).point1_id).status
            ArcList(ArcCount).pointm_id = PointList(ArcList(ArcCount).pointm_id).status
            ArcList(ArcCount).body_id = BodyCount
            ArcList(ArcCount).group_id = BodyCount
            
            For j = PCount0 + 1 To PointCount
                If ArcList(PointList(j).arc_id).body_id = id Then
                    PointList(j).arc_id = ArcCount
                End If
            Next
        End If
    Next
    
    Count0 = SPLineCount
    For i = 1 To Count0
        If SPLineList(i).body_id = id Then
            SPLineCount = SPLineCount + 1
            ReDim Preserve SPLineList(SPLineCount)
            SPLineList(SPLineCount) = SPLineList(i)
            For j = 0 To SPLineList(SPLineCount).vertex_count - 1
                SPLineList(SPLineCount).vertex_id(j) = PointList(SPLineList(SPLineCount).vertex_id(j)).status
            Next
            SPLineList(SPLineCount).body_id = BodyCount
            SPLineList(SPLineCount).group_id = BodyCount
            SPLineList(SPLineCount).id = SPLineCount
            SPLineList(SPLineCount).point0_id = PointList(SPLineList(i).point0_id).status
            SPLineList(SPLineCount).point1_id = PointList(SPLineList(i).point1_id).status
        End If
    Next
    
    Count0 = OutputStartPointList.Count
    For i = 1 To Count0
        If PointList(OutputStartPointList.point_id(i)).body_id = id Then
            OutputStartPointList.Count = OutputStartPointList.Count + 1
            ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.Count)
            ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.Count)
            ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.Count)
            
            OutputStartPointList.point_id(OutputStartPointList.Count) = PointList(OutputStartPointList.point_id(i)).status
            OutputStartPointList.leading_point0(OutputStartPointList.Count) = OutputStartPointList.leading_point0(i)
            OutputStartPointList.leading_point0(OutputStartPointList.Count).id = PointList(OutputStartPointList.leading_point0(i).id).status
            
            OutputStartPointList.leading_point1(OutputStartPointList.Count) = OutputStartPointList.leading_point1(i)
            OutputStartPointList.leading_point1(OutputStartPointList.Count).id = PointList(OutputStartPointList.leading_point1(i).id).status
        End If
    Next
    
    For i = 1 To PointCount
        PointList(i).status = PointStatus.Normal
    Next
    
    CopyBody = BodyCount
End Function

Public Function CopyGroup(ByVal id As Long) As Long
    Dim i As Long, body_id As Long, group_id As Long
    
    MaxBodyID = GetBodyList(BodyList)
    group_id = MaxBodyID + 1
    
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = id Then
            body_id = CopyBody(BodyList(i).body_id)
            ReplaceGroupID body_id, group_id
        End If
    Next
    
    CopyGroup = group_id
    MaxBodyID = GetBodyList(BodyList)
End Function

Sub AddSPLine()
    SPLineCount = SPLineCount + 1
    ReDim Preserve SPLineList(SPLineCount)
    ReDim SPLineList(SPLineCount).vertex_id(0)
    
    SPLineList(SPLineCount).id = SPLineCount
    SPLineList(SPLineCount).vertex_count = 0
    SPLineList(SPLineCount).segment_between_points = SPLine_SegmentBetweenPoints
    SPLineList(SPLineCount).color = LayerColor(CurLayer)
    
    BodyCount = BodyCount + 1
    SPLineList(SPLineCount).body_id = BodyCount
    SPLineList(SPLineCount).group_id = BodyCount
End Sub

Sub AddSPLineByTempSPline()
    Dim k As Long, i As Long, id As Long
    Dim body_id As Long, group_id As Long

    AddSPLine
    
    k = TempSPline.vertex_count - 2
    ReDim SPLineList(SPLineCount).vertex_id(k)
    
    If PointList(TempSPline.vertex_id(0)).body_id = 0 Then
        body_id = BodyCount
        group_id = BodyCount
    Else
        BodyCount = BodyCount - 1
        body_id = PointList(TempSPline.vertex_id(0)).body_id
        group_id = PointList(TempSPline.vertex_id(0)).group_id
    End If
    
    For i = 0 To k
        id = TempSPline.vertex_id(i)
        SPLineList(SPLineCount).vertex_id(i) = id
        
        PointList(id).Type = PointType.SPLinePoint
        PointList(id).body_id = body_id
        PointList(id).group_id = group_id
    Next
    SPLineList(SPLineCount).point0_id = TempSPline.vertex_id(0)
    SPLineList(SPLineCount).point1_id = TempSPline.vertex_id(k)
    SPLineList(SPLineCount).vertex_count = k + 1
    
    SPLineList(SPLineCount).Layer = TempSPline.Layer
    SPLineList(SPLineCount).color = TempSPline.color
    SPLineList(SPLineCount).body_id = body_id
    SPLineList(SPLineCount).group_id = group_id
End Sub

Sub DeleteSPLine(ByVal id As Long)
    Dim i As Long
    
    If SPLineList(id).point0_id = SPLineList(id).point1_id Then 'closed
        SPLineList(id).vertex_count = SPLineList(id).vertex_count - 1
        SPLineList(id).point1_id = SPLineList(id).vertex_id(SPLineList(id).vertex_count - 1)
    End If
        
    For i = id To SPLineCount - 1
        ReDim Preserve SPLineList(i).vertex_id(SPLineList(i + 1).vertex_count - 1)
        ReDim Preserve SPLineList(i).segment_point(SPLineList(i + 1).segment_point_count - 1)
        SPLineList(i) = SPLineList(i + 1)
        SPLineList(i).id = i
    Next
    SPLineCount = SPLineCount - 1
    ReDim Preserve SPLineList(SPLineCount)
End Sub

Sub DeleteAll()
    ReDim PointList(0)
    ReDim SegmentList(0)
    ReDim ArcList(0)
    ReDim SPLineList(0)
    ReDim BodyList(0)
    ReDim GroupList(0)
    
    PointCount = 0
    SegmentCount = 0
    ArcCount = 0
    SPLineCount = 0
    BodyCount = 0
    GroupCount = 0
    
    OutputStartPointList.Count = 0
    ReDim OutputStartPointList.point_id(0)
    ReDim OutputStartPointList.leading_point0(0)
    ReDim OutputStartPointList.leading_point1(0)
    
    AuxXLineCount = 0
    AuxYLineCount = 0
    ReDim AuxXLine(0)
    ReDim AuxYLine(0)
    
    catched_pid = 0
    catched_sid = 0
    catched_aid = 0
    catched_cid = 0
    catched_spid = 0
End Sub

Sub ReverseSegmentDirection(sid As Long, do_s0 As Boolean, do_s1 As Boolean)
    Dim pid0 As Long, pid1 As Long, sid0 As Long, sid1 As Long
    Dim i As Long
    
    pid0 = SegmentList(sid).point0_id
    pid1 = SegmentList(sid).point1_id
    
    SegmentList(sid).point0_id = pid1
    SegmentList(sid).point1_id = pid0
    
    If do_s0 = True And pid0 > 0 Then
        For i = 1 To SegmentCount
            If i <> sid And SegmentList(i).point1_id = pid0 Then
                ReverseSegmentDirection i, True, False
            End If
        Next
    End If
    
    If do_s1 = True And pid1 > 0 Then
        For i = 1 To SegmentCount
            If i <> sid And SegmentList(i).point0_id = pid1 Then
                ReverseSegmentDirection i, False, True
            End If
        Next
    End If
End Sub

Sub ReverseDirection(id As Long, id_type As Long, do_0 As Boolean, do_1 As Boolean)
    Dim pid0 As Long, pid1 As Long, sid0 As Long, sid1 As Long, sa As Double, ea As Double, atype As ArcType
    Dim i As Long, j As Long, n As Long
    
    Static last_pid0 As Long, last_pid1 As Long, last_rc_pid0 As Long, k As Long

    If id_type = 0 Then 'segment
        pid0 = SegmentList(id).point0_id
        pid1 = SegmentList(id).point1_id
        
        SegmentList(id).point0_id = pid1
        SegmentList(id).point1_id = pid0
        
        If do_0 And pid0 <> last_pid1 Then
            If PointList(pid0).action = ActionType.StartDropping Then
                PointList(pid0).action = ActionType.StopDropping
            ElseIf PointList(pid0).action = ActionType.StopDropping Then
                PointList(pid0).action = ActionType.StartDropping
            End If
            last_pid0 = pid0
        End If
        
        If do_1 And pid1 <> last_pid0 Then
            If PointList(pid1).action = ActionType.StartDropping Then
                PointList(pid1).action = ActionType.StopDropping
            ElseIf PointList(pid1).action = ActionType.StopDropping Then
                PointList(pid1).action = ActionType.StartDropping
            End If
            last_pid1 = pid1
        End If
    ElseIf id_type = 1 Then ' arc
        pid0 = ArcList(id).point0_id
        pid1 = ArcList(id).point1_id
        
        ArcList(id).point0_id = pid1
        ArcList(id).point1_id = pid0
        
        sa = ArcList(id).start_angle
        ea = ArcList(id).end_angle
        
        ArcList(id).start_angle = ea
        ArcList(id).end_angle = sa
        
        atype = ArcList(id).Type
        
        If do_0 Then
            If PointList(pid0).action = ActionType.StartDropping Then
                PointList(pid0).action = ActionType.StopDropping
            ElseIf PointList(pid0).action = ActionType.StopDropping Then
                PointList(pid0).action = ActionType.StartDropping
            End If
        End If
        
        If do_1 Then
            If PointList(pid1).action = ActionType.StartDropping Then
                PointList(pid1).action = ActionType.StopDropping
            ElseIf PointList(pid1).action = ActionType.StopDropping Then
                PointList(pid1).action = ActionType.StartDropping
            End If
        End If
    ElseIf id_type = 2 Then 'SPline
        pid0 = SPLineList(id).point0_id
        pid1 = SPLineList(id).point1_id
        
        SPLineList(id).point0_id = pid1
        SPLineList(id).point1_id = pid0
        
        n = SPLineList(id).vertex_count
        For i = 0 To Int((n - 1) / 2)
            j = SPLineList(id).vertex_id(i)
            SPLineList(id).vertex_id(i) = SPLineList(id).vertex_id(n - 1 - i)
            SPLineList(id).vertex_id(n - 1 - i) = j
        Next
        
        If do_0 And pid0 <> last_pid1 Then
            If PointList(pid0).action = ActionType.StartDropping Then
                PointList(pid0).action = ActionType.StopDropping
            ElseIf PointList(pid0).action = ActionType.StopDropping Then
                PointList(pid0).action = ActionType.StartDropping
            End If
            last_pid0 = pid0
        End If
        
        If do_1 And pid1 <> last_pid0 Then
            If PointList(pid1).action = ActionType.StartDropping Then
                PointList(pid1).action = ActionType.StopDropping
            ElseIf PointList(pid1).action = ActionType.StopDropping Then
                PointList(pid1).action = ActionType.StartDropping
            End If
            last_pid1 = pid1
        End If
    End If
    
    k = k + 1
    
    If do_0 = True Then
        For i = 1 To SegmentCount
            If id_type = 0 Then
                If i <> id And SegmentList(i).point1_id = pid0 Then
                    ReverseDirection i, 0, True, False
                    Exit For
                End If
            ElseIf id_type = 1 Then
                If atype = ArcType.RoundedCorner And SegmentList(i).point1_id = id Then
                    ReverseDirection i, 0, True, False
                    Exit For
                ElseIf SegmentList(i).point1_id = pid0 Then
                    ReverseDirection i, 0, True, False
                    Exit For
                End If
            ElseIf id_type = 2 Then
                If SegmentList(i).point1_id = pid0 Then
                    ReverseDirection i, 0, True, False
                    Exit For
                End If
            End If
        Next
        
        For i = 1 To ArcCount
            If id_type = 0 Or id_type = 2 Then
                If ArcList(i).point_id = pid0 And ArcList(i).Type = ArcType.RoundedCorner Then
                    ReverseDirection i, 1, False, False
                    If last_rc_pid0 = 0 Then last_rc_pid0 = pid0
                    Exit For
                ElseIf ArcList(i).point1_id = pid0 Then
                    ReverseDirection i, 1, True, False
                    Exit For
                End If
            ElseIf id_type = 1 Then
                If i <> id And ArcList(i).point1_id = pid0 Then
                    ReverseDirection i, 1, True, False
                    Exit For
                End If
            End If
        Next
    
        For i = 1 To SPLineCount
            If id_type = 0 Then
                If SPLineList(i).point1_id = pid0 Then
                    ReverseDirection i, 2, True, False
                    Exit For
                End If
            ElseIf id_type = 1 Then
                If atype = ArcType.RoundedCorner And SPLineList(i).point1_id = id Then
                    ReverseDirection i, 2, True, False
                    Exit For
                ElseIf SPLineList(i).point1_id = pid0 Then
                    ReverseDirection i, 2, True, False
                    Exit For
                End If
            ElseIf id_type = 2 Then
                If i <> id And SPLineList(i).point1_id = pid0 Then
                    ReverseDirection i, 2, True, False
                    Exit For
                End If
            End If
        Next
        
    End If
    
    If do_1 = True Then
        For i = 1 To SegmentCount
            If id_type = 0 Then
                If i <> id And SegmentList(i).point0_id = pid1 And pid1 <> last_rc_pid0 Then
                    ReverseDirection i, 0, False, True
                    Exit For
                End If
            ElseIf id_type = 1 Then
                If atype = ArcType.RoundedCorner And SegmentList(i).point0_id = id Then
                    ReverseDirection i, 0, False, True
                    Exit For
                ElseIf SegmentList(i).point0_id = pid1 Then
                    ReverseDirection i, 0, False, True
                    Exit For
                End If
            ElseIf id_type = 2 Then
                If SegmentList(i).point0_id = pid1 Then
                    ReverseDirection i, 0, False, True
                    Exit For
                End If
            End If
        Next
        
        For i = 1 To ArcCount
            If id_type = 0 Or id_type = 2 Then
                If pid1 = last_rc_pid0 Then '避免循环时重做首尾
                    last_rc_pid0 = 0
                    Exit For
                ElseIf ArcList(i).point_id = pid1 And ArcList(i).Type = ArcType.RoundedCorner Then
                    ReverseDirection i, 1, False, False
                    Exit For
                ElseIf ArcList(i).point0_id = pid1 Then
                    ReverseDirection i, 1, False, True
                    Exit For
                End If
            ElseIf id_type = 1 Then
                If i <> id And ArcList(i).point0_id = pid1 Then
                    ReverseDirection i, 1, False, True
                    Exit For
                End If
            End If
        Next
        
        For i = 1 To SPLineCount
            If id_type = 0 Then
                If SPLineList(i).point0_id = pid1 And pid1 <> last_rc_pid0 Then
                    ReverseDirection i, 2, False, True
                    Exit For
                End If
            ElseIf id_type = 1 Then
                If atype = ArcType.RoundedCorner And SPLineList(i).point0_id = id Then
                    ReverseDirection i, 2, False, True
                    Exit For
                ElseIf SPLineList(i).point0_id = pid1 Then
                    ReverseDirection i, 2, False, True
                    Exit For
                End If
            ElseIf id_type = 2 Then
                If i <> id And SPLineList(i).point0_id = pid1 And pid1 <> last_rc_pid0 Then
                    ReverseDirection i, 2, False, True
                    Exit For
                End If
            End If
        Next
        
    End If
    
    k = k - 1
    If k = 0 Then
        last_pid0 = 0
        last_pid1 = 0
        last_rc_pid0 = 0
    End If
End Sub

Sub SetDroppingByDrawingOrder()
    Dim i As Long, j As Long, k As Long, id As Long, min_id As Long, end_pid As Long
    
    
    EraseDroppingSetting

    '-----------------------
    OutputStartPointList.Count = 0
    '----------------------------
    
    For i = 1 To PointCount
        If PointList(i).action <> ActionType.Dropping Then
            k = 0
            If PointList(i).Type = PointType.SPLinePoint Then
                For j = 1 To SPLineCount
                    If SPLineList(j).point0_id = PointList(i).id Then '初始点
                        k = 1
                        Exit For
                    End If
                Next
            End If
            
            If PointList(i).Type = PointType.ArcPoint Then
                For j = 1 To ArcCount
                    If PointList(i).id = ArcList(j).point0_id And ArcList(j).Type <> ArcType.RoundedCorner Then '初始点
                        k = 1
                        Exit For
                    End If
                Next
            End If
                    
            'If PointList(I).Type = PointType.NormalPoint Then
                For j = 1 To SegmentCount
                    If PointList(i).id = SegmentList(j).point1_id Then
                        Exit For
                    End If
                Next
                If j > SegmentCount Then '孤点或起点
                    For j = 1 To SegmentCount
                        If PointList(i).id = SegmentList(j).point0_id Then
                            Exit For
                        End If
                    Next
                    If j <= SegmentCount Then '不是孤点
                        k = 1
                    End If
                Else
                    id = PointList(i).id
                    min_id = id
                    Do
                        For j = 1 To SegmentCount
                            If id = SegmentList(j).point0_id Then
                                id = SegmentList(j).point1_id
                                If id < min_id Then
                                    min_id = id
                                End If
                                Exit For
                            End If
                        Next
                        If j > SegmentCount Then '不封闭
                            Exit Do
                        Else
                            If id = PointList(i).id Then '循环一圈
                                If min_id = id Then '最小编号
                                    k = 1
                                End If
                                Exit Do
                            End If
                        End If
                    Loop
                End If
            'End If
            
            If k = 1 Then
                SetStartDroppingOnChain PointList(i).id, 0, end_pid
                    
                For j = 1 To OutputStartPointList.Count
                    If OutputStartPointList.point_id(j) = PointList(i).id Then
                        Exit For
                    End If
                Next
                If j > OutputStartPointList.Count Then
                    OutputStartPointList.Count = OutputStartPointList.Count + 1
                    ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.Count)
                    ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.Count)
                    ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.Count)
                    j = OutputStartPointList.Count
                    
                    OutputStartPointList.leading_point1(j) = PointList(end_pid)
                End If
                OutputStartPointList.point_id(j) = PointList(i).id
                OutputStartPointList.leading_point0(j) = PointList(i)
            End If
        End If
    Next
End Sub

Sub EraseDroppingSetting()
    Dim i As Long
    
    For i = 1 To PointCount
        PointList(i).action = No_Action
    Next
    
    OutputStartPointList.Count = 0
    ReDim OutputStartPointList.point_id(0)
    ReDim OutputStartPointList.leading_point0(0)
    ReDim OutputStartPointList.leading_point1(0)
    
    PathOutputPointCount = 0
    ReDim PathOutputPoint(0)
    
    'SaveUndo
    
    'FrmMain.PicPathCls
    'DrawAll
End Sub

Function SeekSPlineByPointID(ByVal pid As Long) As Long
    Dim i As Long, j As Long
    
    SeekSPlineByPointID = 0
    For i = 1 To SPLineCount
        For j = 0 To SPLineList(i).vertex_count - 1
            If SPLineList(i).vertex_id(j) = pid Then
                SeekSPlineByPointID = i
                Exit Function
            End If
        Next
    Next
End Function

Function SeekAnySPlineByPointID(ByVal pid As Long, ByVal num As Long) As Long
    Dim i As Long, j As Long, k As Long
    
    k = 0
    SeekAnySPlineByPointID = 0
    For i = 1 To SPLineCount
        For j = 0 To SPLineList(i).vertex_count - 1
            If SPLineList(i).vertex_id(j) = pid Then
                k = k + 1
                If k = num Then
                    SeekAnySPlineByPointID = i
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Public Sub ReplaceBodyID(ByVal BodyID0 As Long, ByVal BodyID1 As Long)
    Dim i As Long
    
    For i = 1 To SegmentCount
        If SegmentList(i).body_id = BodyID0 Then
            SegmentList(i).body_id = BodyID1
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).body_id = BodyID0 Then
            ArcList(i).body_id = BodyID1
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).body_id = BodyID0 Then
            SPLineList(i).body_id = BodyID1
        End If
    Next
    
    For i = 1 To PointCount
        If PointList(i).body_id = BodyID0 Then
            PointList(i).body_id = BodyID1
        End If
    Next
End Sub

Public Sub ReplaceGroupID(ByVal GroupID0 As Long, ByVal GroupID1 As Long)
    Dim i As Long
    
    For i = 1 To SegmentCount
        If SegmentList(i).group_id = GroupID0 Then
            SegmentList(i).group_id = GroupID1
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).group_id = GroupID0 Then
            ArcList(i).group_id = GroupID1
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).group_id = GroupID0 Then
            SPLineList(i).group_id = GroupID1
        End If
    Next
    
    For i = 1 To PointCount
        If PointList(i).group_id = GroupID0 Then
            PointList(i).group_id = GroupID1
        End If
    Next
End Sub

Public Sub SeperateBodyFromGroup(ByVal GroupID As Long)
    Dim i As Long
    
    For i = 1 To SegmentCount
        If SegmentList(i).group_id = GroupID Then
            SegmentList(i).group_id = SegmentList(i).body_id
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).group_id = GroupID Then
            ArcList(i).group_id = ArcList(i).body_id
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).group_id = GroupID Then
            SPLineList(i).group_id = SPLineList(i).body_id
        End If
    Next
    
    For i = 1 To PointCount
        If PointList(i).group_id = GroupID Then
            PointList(i).group_id = PointList(i).body_id
        End If
    Next
End Sub

Public Function SegmentOnBox(ByVal sid As Long, X As Double, Y As Double, w As Double, h As Double, start_sid As Long, body_id As Long) As Boolean
    Dim i As Long, j As Long, k As Long, k0 As Long, id0 As Long, id1 As Long
    Dim sid0 As Long, sid_min As Long
    
    SegmentOnBox = False
    sid0 = sid
    sid_min = sid
    k0 = -1
    For i = 1 To 4
        id0 = SegmentList(sid).point0_id
        id1 = SegmentList(sid).point1_id
        If PointList(id0).X = PointList(id1).X Then
            k = 1
        ElseIf PointList(id0).Y = PointList(id1).Y Then
            k = 2
        Else
            k = 0
        End If
        If k = 0 Or k0 = k Then
            Exit Function
        End If
        k0 = k
        For j = 1 To SegmentCount
            If SegmentList(j).point0_id = id1 Then
                sid = j
                Exit For
            End If
        Next
        If j > SegmentCount Then
            Exit Function
        End If
        If sid_min > sid Then
            sid_min = sid
        End If
    Next
    If sid = sid0 Then
        SegmentOnBox = True
        
        X = PointList(SegmentList(sid_min).point0_id).X
        Y = PointList(SegmentList(sid_min).point0_id).Y
        
        For i = 1 To SegmentCount
            If SegmentList(i).point0_id = SegmentList(sid_min).point1_id Then
                w = PointList(SegmentList(i).point1_id).X - X
                h = PointList(SegmentList(i).point1_id).Y - Y
                Exit For
            End If
        Next
        
        start_sid = sid_min
        body_id = PointList(SegmentList(sid_min).point0_id).body_id
    End If
End Function

Public Sub GetBodyScale(ByVal body_id As Long, ByRef Min_X As Double, ByRef Min_Y As Double, ByRef Max_X As Double, ByRef Max_Y As Double)
    Dim i As Long, j As Long, k As Integer
    Dim MinX As Double, MaxX As Double
    Dim MinY As Double, MaxY As Double
    
    Dim AllPoints() As PolygonPoint
    
    k = 0
    For i = 1 To PointCount
        If PointList(i).body_id = body_id And PointList(i).arc_id = 0 Then
            If k = 0 Then
                MinX = PointList(i).X
                MaxX = MinX
                
                MinY = PointList(i).Y
                MaxY = MinY
                k = 1
            Else
                If MinX > PointList(i).X Then
                    MinX = PointList(i).X
                ElseIf MaxX < PointList(i).X Then
                    MaxX = PointList(i).X
                End If
                
                If MinY > PointList(i).Y Then
                    MinY = PointList(i).Y
                ElseIf MaxY < PointList(i).Y Then
                    MaxY = PointList(i).Y
                End If
            End If
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).body_id = body_id Then
            ReDim AllPoints(0)
            ArcPoints ArcList(i), AllPoints
            For j = 0 To UBound(AllPoints)
                If k = 0 Then
                    MinX = AllPoints(j).X
                    MaxX = MinX
                    
                    MinY = AllPoints(j).Y
                    MaxY = MinY
                    k = 1
                Else
                    If MinX > AllPoints(j).X Then
                        MinX = AllPoints(j).X
                    ElseIf MaxX < AllPoints(j).X Then
                        MaxX = AllPoints(j).X
                    End If
                    
                    If MinY > AllPoints(j).Y Then
                        MinY = AllPoints(j).Y
                    ElseIf MaxY < AllPoints(j).Y Then
                        MaxY = AllPoints(j).Y
                    End If
                End If
            Next
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).body_id = body_id Then
            ReDim AllPoints(0)
            SplinePoints SPLineList(i), AllPoints, SPLine_SegmentBetweenPoints
            For j = 0 To UBound(AllPoints)
                If k = 0 Then
                    MinX = AllPoints(j).X
                    MaxX = MinX
                    
                    MinY = AllPoints(j).Y
                    MaxY = MinY
                    k = 1
                Else
                    If MinX > AllPoints(j).X Then
                        MinX = AllPoints(j).X
                    ElseIf MaxX < AllPoints(j).X Then
                        MaxX = AllPoints(j).X
                    End If
                    
                    If MinY > AllPoints(j).Y Then
                        MinY = AllPoints(j).Y
                    ElseIf MaxY < AllPoints(j).Y Then
                        MaxY = AllPoints(j).Y
                    End If
                End If
            Next
        End If
    Next
    
    Min_X = MinX
    Min_Y = MinY
    Max_X = MaxX
    Max_Y = MaxY
End Sub

Public Sub GetGroupScale(ByVal group_id As Long, ByRef Min_X As Double, ByRef Min_Y As Double, ByRef Max_X As Double, ByRef Max_Y As Double)
    Dim i As Long, k As Long
    Dim Min_X0 As Double, Max_X0 As Double, MinX As Double, MaxX As Double
    Dim Min_Y0 As Double, Max_Y0 As Double, MinY As Double, MaxY As Double
    
    k = 0
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = group_id Then
            GetBodyScale BodyList(i).body_id, MinX, MinY, MaxX, MaxY
            k = k + 1
            If k = 1 Then
                Min_X0 = MinX
                Min_Y0 = MinY
                Max_X0 = MaxX
                Max_Y0 = MaxY
            Else
                If MinX < Min_X0 Then Min_X0 = MinX
                If MinY < Min_Y0 Then Min_Y0 = MinY
                If MaxX > Max_X0 Then Max_X0 = MaxX
                If MaxY > Max_Y0 Then Max_Y0 = MaxY
            End If
        End If
    Next
    
    Min_X = Min_X0
    Min_Y = Min_Y0
    Max_X = Max_X0
    Max_Y = Max_Y0
End Sub

Public Sub GetBodyCenter(ByVal body_id As Long, ByRef cx As Double, ByRef cy As Double)
    Dim MinX As Double, MaxX As Double
    Dim MinY As Double, MaxY As Double
    
    GetBodyScale body_id, MinX, MinY, MaxX, MaxY
    
    cx = (MinX + MaxX) / 2
    cy = (MinY + MaxY) / 2
End Sub

Public Sub GetGroupCenter(ByVal group_id As Long, ByRef cx As Double, ByRef cy As Double)
    Dim i As Long, k As Long
    Dim Min_X As Double, Max_X As Double, MinX As Double, MaxX As Double
    Dim Min_Y As Double, Max_Y As Double, MinY As Double, MaxY As Double
    
    k = 0
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = group_id Then
            GetBodyScale BodyList(i).body_id, MinX, MinY, MaxX, MaxY
            k = k + 1
            If k = 1 Then
                Min_X = MinX
                Min_Y = MinY
                Max_X = MaxX
                Max_Y = MaxY
            Else
                If MinX < Min_X Then Min_X = MinX
                If MinY < Min_Y Then Min_Y = MinY
                If MaxX > Max_X Then Max_X = MaxX
                If MaxY > Max_Y Then Max_Y = MaxY
            End If
        End If
    Next
    
    cx = (Min_X + Max_X) / 2
    cy = (Min_Y + Max_Y) / 2
End Sub

Public Sub RotateBody(ByVal body_id As Long, ByVal cx As Double, ByVal cy As Double, ByVal angle As Double)
    Dim i As Long, CS As Double, SN As Double, x0 As Double, y0 As Double
    
    CS = Cos(angle)
    SN = Sin(angle)
    
    'shift to org
    '-------------------------------------------------
    For i = 1 To PointCount
        If PointList(i).body_id = body_id Then
            PointList(i).X = PointList(i).X - cx
            PointList(i).Y = PointList(i).Y - cy
        End If
    Next
    For i = 1 To ArcCount
        If ArcList(i).body_id = body_id Then
            ArcList(i).X = ArcList(i).X - cx
            ArcList(i).Y = ArcList(i).Y - cy
        End If
    Next
    For i = 1 To OutputStartPointList.Count
        If PointList(OutputStartPointList.point_id(i)).body_id = body_id Then
            OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X - cx
            OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y - cy
            OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X - cx
            OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y - cy
        End If
    Next
    '-------------------------------------------------
    
    'rotate
    '-------------------------------------------------
    For i = 1 To PointCount
        If PointList(i).body_id = body_id Then
            x0 = PointList(i).X
            y0 = PointList(i).Y
            PointList(i).X = (CS * x0) - (SN * y0)
            PointList(i).Y = (SN * x0) + (CS * y0)
        End If
    Next
    For i = 1 To ArcCount
        If ArcList(i).body_id = body_id Then
            x0 = ArcList(i).X
            y0 = ArcList(i).Y
            ArcList(i).X = (CS * x0) - (SN * y0)
            ArcList(i).Y = (SN * x0) + (CS * y0)
            
            ArcList(i).ax_angle = ArcList(i).ax_angle + angle
        End If
    Next
    For i = 1 To OutputStartPointList.Count
        If PointList(OutputStartPointList.point_id(i)).body_id = body_id Then
            x0 = OutputStartPointList.leading_point0(i).X
            y0 = OutputStartPointList.leading_point0(i).Y
            OutputStartPointList.leading_point0(i).X = (CS * x0) - (SN * y0)
            OutputStartPointList.leading_point0(i).Y = (SN * x0) + (CS * y0)
            
            
            x0 = OutputStartPointList.leading_point1(i).X
            y0 = OutputStartPointList.leading_point1(i).Y
            OutputStartPointList.leading_point1(i).X = (CS * x0) - (SN * y0)
            OutputStartPointList.leading_point1(i).Y = (SN * x0) + (CS * y0)
        End If
    Next
    '-------------------------------------------------
    
    'shift back
    '-------------------------------------------------
    For i = 1 To PointCount
        If PointList(i).body_id = body_id Then
            PointList(i).X = PointList(i).X + cx
            PointList(i).Y = PointList(i).Y + cy
        End If
    Next
    For i = 1 To ArcCount
        If ArcList(i).body_id = body_id Then
            ArcList(i).X = ArcList(i).X + cx
            ArcList(i).Y = ArcList(i).Y + cy
        End If
    Next
    For i = 1 To OutputStartPointList.Count
        If PointList(OutputStartPointList.point_id(i)).body_id = body_id Then
            OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X + cx
            OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y + cy
            OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X + cx
            OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y + cy
        End If
    Next
End Sub

Public Sub RotateGroup(ByVal group_id As Long, ByVal cx As Double, ByVal cy As Double, ByVal angle As Double)
    Dim i As Long
    
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = group_id Then
            RotateBody BodyList(i).body_id, cx, cy, angle
        End If
    Next
End Sub

Public Sub MoveBody(ByVal BodyID As Long, ByVal dX As Double, ByVal dy As Double)
    Dim i As Long, j As Long
    
    For i = 1 To PointCount
        PointList(i).status = PointStatus.Normal
    Next
    
    For i = 1 To SegmentCount
        If SegmentList(i).body_id = BodyID Then
            If PointList(SegmentList(i).point0_id).status = PointStatus.Normal Then
                PointList(SegmentList(i).point0_id).X = PointList(SegmentList(i).point0_id).X + dX
                PointList(SegmentList(i).point0_id).Y = PointList(SegmentList(i).point0_id).Y + dy
                PointList(SegmentList(i).point0_id).status = PointStatus.Used
            End If
            
            If PointList(SegmentList(i).point1_id).status = PointStatus.Normal Then
                PointList(SegmentList(i).point1_id).X = PointList(SegmentList(i).point1_id).X + dX
                PointList(SegmentList(i).point1_id).Y = PointList(SegmentList(i).point1_id).Y + dy
                PointList(SegmentList(i).point1_id).status = PointStatus.Used
            End If
        End If
    Next
    
    For i = 1 To ArcCount
        If ArcList(i).body_id = BodyID Then
            ArcList(i).X = ArcList(i).X + dX
            ArcList(i).Y = ArcList(i).Y + dy
            
            If PointList(ArcList(i).point0_id).status = PointStatus.Normal Then
                PointList(ArcList(i).point0_id).X = PointList(ArcList(i).point0_id).X + dX
                PointList(ArcList(i).point0_id).Y = PointList(ArcList(i).point0_id).Y + dy
                PointList(ArcList(i).point0_id).status = PointStatus.Used
            End If
            
            If PointList(ArcList(i).point1_id).status = PointStatus.Normal Then
                PointList(ArcList(i).point1_id).X = PointList(ArcList(i).point1_id).X + dX
                PointList(ArcList(i).point1_id).Y = PointList(ArcList(i).point1_id).Y + dy
                PointList(ArcList(i).point1_id).status = PointStatus.Used
            End If
            
            If PointList(ArcList(i).pointm_id).status = PointStatus.Normal Then
                PointList(ArcList(i).pointm_id).X = PointList(ArcList(i).pointm_id).X + dX
                PointList(ArcList(i).pointm_id).Y = PointList(ArcList(i).pointm_id).Y + dy
                PointList(ArcList(i).pointm_id).status = PointStatus.Used
            End If
        End If
    Next
    
    For i = 1 To SPLineCount
        If SPLineList(i).body_id = BodyID Then
            For j = 0 To SPLineList(i).vertex_count - 1
                If PointList(SPLineList(i).vertex_id(j)).status = PointStatus.Normal Then
                    PointList(SPLineList(i).vertex_id(j)).X = PointList(SPLineList(i).vertex_id(j)).X + dX
                    PointList(SPLineList(i).vertex_id(j)).Y = PointList(SPLineList(i).vertex_id(j)).Y + dy
                    PointList(SPLineList(i).vertex_id(j)).status = PointStatus.Used
                End If
            Next
        End If
    Next
    
    For i = 1 To OutputStartPointList.Count
        If PointList(OutputStartPointList.point_id(i)).body_id = BodyID Then
            OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X + dX
            OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y + dy
            OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X + dX
            OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y + dy
        End If
    Next
    
    For i = 1 To PointCount
        PointList(i).status = PointStatus.Normal
    Next
End Sub

Public Sub MoveAllBody(ByVal dX As Double, ByVal dy As Double)
    Dim i As Long, j As Long
    
    For i = 1 To PointCount
        PointList(i).status = PointStatus.Normal
    Next
    
    For i = 1 To SegmentCount
        If PointList(SegmentList(i).point0_id).status = PointStatus.Normal Then
            PointList(SegmentList(i).point0_id).X = PointList(SegmentList(i).point0_id).X + dX
            PointList(SegmentList(i).point0_id).Y = PointList(SegmentList(i).point0_id).Y + dy
            PointList(SegmentList(i).point0_id).status = PointStatus.Used
        End If
        
        If PointList(SegmentList(i).point1_id).status = PointStatus.Normal Then
            PointList(SegmentList(i).point1_id).X = PointList(SegmentList(i).point1_id).X + dX
            PointList(SegmentList(i).point1_id).Y = PointList(SegmentList(i).point1_id).Y + dy
            PointList(SegmentList(i).point1_id).status = PointStatus.Used
        End If
    Next
    
    For i = 1 To ArcCount
        ArcList(i).X = ArcList(i).X + dX
        ArcList(i).Y = ArcList(i).Y + dy
        
        If PointList(ArcList(i).point0_id).status = PointStatus.Normal Then
            PointList(ArcList(i).point0_id).X = PointList(ArcList(i).point0_id).X + dX
            PointList(ArcList(i).point0_id).Y = PointList(ArcList(i).point0_id).Y + dy
            PointList(ArcList(i).point0_id).status = PointStatus.Used
        End If
        
        If PointList(ArcList(i).point1_id).status = PointStatus.Normal Then
            PointList(ArcList(i).point1_id).X = PointList(ArcList(i).point1_id).X + dX
            PointList(ArcList(i).point1_id).Y = PointList(ArcList(i).point1_id).Y + dy
            PointList(ArcList(i).point1_id).status = PointStatus.Used
        End If
        
        If PointList(ArcList(i).pointm_id).status = PointStatus.Normal Then
            PointList(ArcList(i).pointm_id).X = PointList(ArcList(i).pointm_id).X + dX
            PointList(ArcList(i).pointm_id).Y = PointList(ArcList(i).pointm_id).Y + dy
            PointList(ArcList(i).pointm_id).status = PointStatus.Used
        End If
    Next
    
    For i = 1 To SPLineCount
        For j = 0 To SPLineList(i).vertex_count - 1
            If PointList(SPLineList(i).vertex_id(j)).status = PointStatus.Normal Then
                PointList(SPLineList(i).vertex_id(j)).X = PointList(SPLineList(i).vertex_id(j)).X + dX
                PointList(SPLineList(i).vertex_id(j)).Y = PointList(SPLineList(i).vertex_id(j)).Y + dy
                PointList(SPLineList(i).vertex_id(j)).status = PointStatus.Used
            End If
        Next
    Next
    
    For i = 1 To OutputStartPointList.Count
        OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X + dX
        OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y + dy
        OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X + dX
        OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y + dy
    Next
    
    For i = 1 To PointCount
        PointList(i).status = PointStatus.Normal
    Next
End Sub

Public Sub MoveGroup(ByVal GroupID As Long, ByVal dX As Double, ByVal dy As Double)
    Dim i As Long
    
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = GroupID Then
            MoveBody BodyList(i).body_id, dX, dy
        End If
    Next
 End Sub

Public Sub MirrorBody(ByVal body_id As Long, ByVal cx As Double)
    Dim i As Long, dX As Double, da As Double
    
    For i = 1 To PointCount
        If PointList(i).body_id = body_id Then
            dX = PointList(i).X - cx
            PointList(i).X = PointList(i).X - 2 * dX
        End If
    Next
    For i = 1 To ArcCount
        If ArcList(i).body_id = body_id Then
            dX = ArcList(i).X - cx
            ArcList(i).X = ArcList(i).X - 2 * dX
            
            ArcList(i).ax_angle = Pi - ArcList(i).ax_angle
            
            da = ArcList(i).end_angle - ArcList(i).start_angle
            ArcList(i).start_angle = -ArcList(i).start_angle
            ArcList(i).end_angle = ArcList(i).start_angle - da
        End If
    Next
    For i = 1 To OutputStartPointList.Count
        If PointList(OutputStartPointList.point_id(i)).body_id = body_id Then
            dX = OutputStartPointList.leading_point0(i).X - cx
            OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X - 2 * dX
            
            dX = OutputStartPointList.leading_point1(i).X - cx
            OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X - 2 * dX
        End If
    Next
End Sub

Public Sub MirrorGroup(ByVal GroupID As Long, ByVal cx As Double)
    Dim i As Long
    
    For i = 1 To MaxBodyID
        If BodyList(i).group_id = GroupID Then
            MirrorBody BodyList(i).body_id, cx
        End If
    Next
End Sub

Function IsPathClockwise(ByVal n As Long, ByRef X() As Double, ByRef Y() As Double) As Boolean
    Dim s As Double
    Dim i As Integer, k As Integer
    
    s = 0
    For i = 1 To n
        If i < n Then
            k = i + 1
        Else
            k = 1
        End If
        s = s + X(i) * Y(k) - X(k) * Y(i)
    Next

    IsPathClockwise = IIf(s > 0, True, False)
End Function

