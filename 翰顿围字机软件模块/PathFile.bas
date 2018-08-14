Attribute VB_Name = "PathFile"
Option Explicit

Public Const RecentFileCount = 5
Public RecentFile() As String

Private CurUndoCount As Long
Private CurUndoIndex As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Sub WriteFile(ByVal fn As String)
    Dim f As Long, I As Long, j As Long
    
    f = FreeFile
    Open fn For Output As f
    
    Print #f, "BEGIN POINT"
    Print #f, PointCount
    For I = 1 To PointCount
        Print #f, PointList(I).id; ",";
        Print #f, PointList(I).body_id; ",";
        Print #f, PointList(I).group_id; ",";
        Print #f, PointList(I).X; ",";
        Print #f, PointList(I).Y; ",";
        Print #f, PointList(I).z; ",";
        Print #f, PointList(I).xp; ",";
        Print #f, PointList(I).yp; ",";
        Print #f, PointList(I).zp; ",";
        Print #f, PointList(I).color; ",";
        Print #f, PointList(I).Layer; ",";
        Print #f, PointList(I).Type; ",";
        Print #f, PointList(I).method; ",";
        Print #f, PointList(I).arc_id; ",";
        Print #f, PointList(I).action; ",";
        Print #f, PointList(I).stay_time; ",";
        Print #f, PointList(I).v_down; ",";
        Print #f, PointList(I).v_up; ",";
        Print #f, PointList(I).HoleType
        
    Next
    Print #f, "END POINT"
    Print #f,
    
    Print #f, "BEGIN SEGMENT"
    Print #f, SegmentCount
    For I = 1 To SegmentCount
        Print #f, SegmentList(I).id; ",";
        Print #f, SegmentList(I).body_id; ",";
        Print #f, SegmentList(I).group_id; ",";
        Print #f, SegmentList(I).point0_id; ",";
        Print #f, SegmentList(I).point1_id; ",";
        Print #f, SegmentList(I).color; ",";
        Print #f, SegmentList(I).Layer; ",";
        Print #f, SegmentList(I).Type; ",";
        Print #f, SegmentList(I).arc_id; ",";
        Print #f, SegmentList(I).action
    Next
    Print #f, "END SEGMENT"
    Print #f,
    
    Print #f, "BEGIN ARC"
    Print #f, ArcCount
    For I = 1 To ArcCount
        Print #f, ArcList(I).id; ",";
        Print #f, ArcList(I).body_id; ",";
        Print #f, ArcList(I).group_id; ",";
        Print #f, ArcList(I).X; ",";
        Print #f, ArcList(I).Y; ",";
        Print #f, ArcList(I).z; ",";
        Print #f, ArcList(I).a; ",";
        Print #f, ArcList(I).b; ",";
        Print #f, ArcList(I).ax_angle; ",";
        Print #f, ArcList(I).start_angle; ",";
        Print #f, ArcList(I).end_angle; ",";
        Print #f, ArcList(I).point0_id; ",";
        Print #f, ArcList(I).point1_id; ",";
        Print #f, ArcList(I).pointm_id; ",";
        Print #f, ArcList(I).color; ",";
        Print #f, ArcList(I).Layer; ",";
        Print #f, ArcList(I).Type; ",";
        Print #f, ArcList(I).point_id; ",";
        Print #f, ArcList(I).action
    Next
    Print #f, "END ARC"
    Print #f,
    
    Print #f, "BEGIN SPLINE"
    Print #f, SPLineCount
    For I = 1 To SPLineCount
        Print #f, SPLineList(I).id; ",";
        Print #f, SPLineList(I).body_id; ",";
        Print #f, SPLineList(I).group_id; ",";
        Print #f, SPLineList(I).vertex_count; ",";
        For j = 0 To SPLineList(I).vertex_count - 1
            Print #f, SPLineList(I).vertex_id(j); ",";
        Next
        Print #f, SPLineList(I).point0_id; ",";
        Print #f, SPLineList(I).point1_id; ",";
        Print #f, SPLineList(I).color; ",";
        Print #f, SPLineList(I).Layer; ",";
        Print #f, SPLineList(I).Type; ",";
        Print #f, SPLineList(I).action; ",";
        Print #f, SPLineList(I).segment_between_points
    Next
    Print #f, "END SPLINE"
    Print #f,
        
    Print #f, "BEGIN BODY"
    Print #f, BodyCount
    Print #f, "END BODY"
    Print #f,
    
    Print #f, "BEGIN START_POINT"
    Print #f, OutputStartPointList.Count
    If OutputStartPointList.Count > 0 Then
        For I = 1 To OutputStartPointList.Count - 1
            Print #f, OutputStartPointList.point_id(I); ",";
            Print #f, OutputStartPointList.leading_point0(I).id; ",";
            Print #f, OutputStartPointList.leading_point0(I).X; ",";
            Print #f, OutputStartPointList.leading_point0(I).Y; ",";
            Print #f, OutputStartPointList.leading_point1(I).id; ",";
            Print #f, OutputStartPointList.leading_point1(I).X; ",";
            Print #f, OutputStartPointList.leading_point1(I).Y; ",";
        Next
        Print #f, OutputStartPointList.point_id(I); ",";
        Print #f, OutputStartPointList.leading_point0(I).id; ",";
        Print #f, OutputStartPointList.leading_point0(I).X; ",";
        Print #f, OutputStartPointList.leading_point0(I).Y; ",";
        Print #f, OutputStartPointList.leading_point1(I).id; ",";
        Print #f, OutputStartPointList.leading_point1(I).X; ",";
        Print #f, OutputStartPointList.leading_point1(I).Y
    End If
    Print #f, "END START_POINT"
    Print #f,
    
    Print #f, "BEGIN ENVIRONMENT"
    Print #f, MainGridX; ","; MainGridY; ","; SubGridX; ","; SubGridY
    Print #f, UserOrgX; ","; UserOrgY; ","; UserOrgZ
    
    Print #f, LayerMax; ",";
    For I = 0 To LayerMax - 1
        Print #f, LayerZValue(I); ","; LayerColor(I); ",";
    Next
    Print #f, LayerZValue(I); ","; LayerColor(I)
    
    Print #f, AuxXLineCount; ",";
    For I = 0 To AuxXLineCount - 1
        Print #f, AuxXLine(I); ",";
    Next
    Print #f,
    
    Print #f, AuxYLineCount; ",";
    For I = 0 To AuxYLineCount - 1
        Print #f, AuxYLine(I); ",";
    Next
    Print #f,
    
    Print #f, PointSize; ","; TrapWidth; ","; HVTrapWidth
    Print #f, CornerR; ","; ArcStepFactor; ","; SPLine_SegmentBetweenPoints
    Print #f, "END ENVIRONMENT"
        
    Close f
End Sub

Sub ReadFile(ByVal fn As String)
    Dim f As Long, s As String, v As Variant, stp As Integer, I As Long, j As Long
    
    If fn = "" Then Exit Sub
    If dir(fn) = "" Then Exit Sub
    
    On Error Resume Next
    
    f = FreeFile
    Open fn For Input As f
    Do While Not EOF(f)
        Line Input #f, s
        If Trim(s) <> "" Then
            Select Case Trim(UCase(s))
                Case "BEGIN POINT"
                    stp = 10
                Case "END POINT"
                    stp = 0
                Case "BEGIN SEGMENT"
                    stp = 20
                Case "END SEGMENT"
                    stp = 0
                Case "BEGIN ARC"
                    stp = 30
                Case "END ARC"
                    stp = 0
                Case "BEGIN SPLINE"
                    stp = 40
                Case "END SPLINE"
                    stp = 0
                Case "BEGIN BODY"
                    stp = 50
                Case "END BODY"
                    stp = 0
                    
                Case "BEGIN START_POINT"
                    stp = 990
                Case "END START_POINT"
                    stp = 0
                Case "BEGIN ENVIRONMENT"
                    stp = 1000
                Case "END ENVIRONMENT"
                    stp = 0
            End Select
            
            Select Case stp
                Case 10
                    stp = 11
                Case 11
                    PointCount = Val(s)
                    ReDim PointList(PointCount)
                    I = 0
                    stp = 12
                Case 12
                    I = I + 1
                    v = Split(s, ",")
                    PointList(I).id = Val(v(0))
                    PointList(I).body_id = Val(v(1))
                    PointList(I).group_id = Val(v(2))
                    PointList(I).X = Val(v(3))
                    PointList(I).Y = Val(v(4))
                    PointList(I).z = Val(v(5))
                    PointList(I).xp = Val(v(6))
                    PointList(I).yp = Val(v(7))
                    PointList(I).zp = Val(v(8))
                    PointList(I).color = Val(v(9))
                    PointList(I).Layer = Val(v(10))
                    PointList(I).Type = Val(v(11))
                    PointList(I).method = Val(v(12))
                    PointList(I).arc_id = Val(v(13))
                    PointList(I).action = Val(v(14))
                    PointList(I).stay_time = Val(v(15))
                    PointList(I).v_down = Val(v(16))
                    PointList(I).v_up = Val(v(17))
                    PointList(I).HoleType = Val(v(18))
                    
                Case 20
                    stp = 21
                Case 21
                    SegmentCount = Val(s)
                    ReDim SegmentList(SegmentCount)
                    I = 0
                    stp = 22
                Case 22
                    I = I + 1
                    v = Split(s, ",")
                    SegmentList(I).id = Val(v(0))
                    SegmentList(I).body_id = Val(v(1))
                    SegmentList(I).group_id = Val(v(2))
                    SegmentList(I).point0_id = Val(v(3))
                    SegmentList(I).point1_id = Val(v(4))
                    SegmentList(I).color = Val(v(5))
                    SegmentList(I).Layer = Val(v(6))
                    SegmentList(I).Type = Val(v(7))
                    SegmentList(I).arc_id = Val(v(8))
                    SegmentList(I).action = Val(v(9))
                    
                Case 30
                    stp = 31
                Case 31
                    ArcCount = Val(s)
                    ReDim ArcList(ArcCount)
                    I = 0
                    stp = 32
                Case 32
                    I = I + 1
                    v = Split(s, ",")
                    ArcList(I).id = Val(v(0))
                    ArcList(I).body_id = Val(v(1))
                    ArcList(I).group_id = Val(v(2))
                    ArcList(I).X = Val(v(3))
                    ArcList(I).Y = Val(v(4))
                    ArcList(I).z = Val(v(5))
                    ArcList(I).a = Val(v(6))
                    ArcList(I).b = Val(v(7))
                    ArcList(I).ax_angle = Val(v(8))
                    ArcList(I).start_angle = Val(v(9))
                    ArcList(I).end_angle = Val(v(10))
                    ArcList(I).point0_id = Val(v(11))
                    ArcList(I).point1_id = Val(v(12))
                    ArcList(I).pointm_id = Val(v(13))
                    ArcList(I).color = Val(v(14))
                    ArcList(I).Layer = Val(v(15))
                    ArcList(I).Type = Val(v(16))
                    ArcList(I).point_id = Val(v(17))
                    ArcList(I).action = Val(v(18))
            
                Case 40
                    stp = 41
                Case 41
                    SPLineCount = Val(s)
                    ReDim SPLineList(SPLineCount)
                    I = 0
                    stp = 42
                Case 42
                    I = I + 1
                    v = Split(s, ",")
                    SPLineList(I).id = Val(v(0))
                    SPLineList(I).body_id = Val(v(1))
                    SPLineList(I).group_id = Val(v(2))
                    SPLineList(I).vertex_count = Val(v(3))
                    ReDim SPLineList(I).vertex_id(SPLineList(I).vertex_count - 1)
                    For j = 0 To SPLineList(I).vertex_count - 1
                        SPLineList(I).vertex_id(j) = Val(v(4 + j))
                    Next
                    SPLineList(I).point0_id = Val(v(4 + j))
                    SPLineList(I).point1_id = Val(v(5 + j))
                    SPLineList(I).color = Val(v(6 + j))
                    SPLineList(I).Layer = Val(v(7 + j))
                    SPLineList(I).Type = Val(v(8 + j))
                    SPLineList(I).action = Val(v(9 + j))
                    SPLineList(I).segment_between_points = Val(v(10 + j))
                                        
                Case 50
                    stp = 51
                Case 51
                    BodyCount = Val(s)
                    I = 0
                    stp = 52
                Case 52
                    
                Case 990
                    stp = 991
                Case 991
                    OutputStartPointList.Count = Val(s)
                    ReDim OutputStartPointList.point_id(OutputStartPointList.Count)
                    ReDim OutputStartPointList.leading_point0(OutputStartPointList.Count)
                    ReDim OutputStartPointList.leading_point1(OutputStartPointList.Count)
                    stp = 992
                Case 992
                    v = Split(s, ",")
                    For I = 1 To OutputStartPointList.Count
                        OutputStartPointList.point_id(I) = Val(v((I - 1) * 7))
                        OutputStartPointList.leading_point0(I).id = Val(v((I - 1) * 7 + 1))
                        OutputStartPointList.leading_point0(I).X = Val(v((I - 1) * 7 + 2))
                        OutputStartPointList.leading_point0(I).Y = Val(v((I - 1) * 7 + 3))
                        OutputStartPointList.leading_point1(I).id = Val(v((I - 1) * 7 + 4))
                        OutputStartPointList.leading_point1(I).X = Val(v((I - 1) * 7 + 5))
                        OutputStartPointList.leading_point1(I).Y = Val(v((I - 1) * 7 + 6))
                    Next
                    
                Case 1000
                    I = 0
                    stp = 1001
                Case 1001
                    I = I + 1
                    v = Split(s, ",")
                    Select Case I
                        Case 1
                            MainGridX = Val(v(0))
                            MainGridY = Val(v(1))
                            SubGridX = Val(v(2))
                            SubGridY = Val(v(3))
                        
                        Case 2
                            UserOrgX = Val(v(0))
                            UserOrgY = Val(v(1))
                            UserOrgZ = Val(v(2))
                            
                        Case 3
                            LayerMax = Val(v(0))
                            For j = 0 To LayerMax
                                LayerZValue(j) = Val(v(2 * j + 1))
                                LayerColor(j) = Val(v(2 * j + 2))
                            Next
                            
                        Case 4
                            AuxXLineCount = Val(v(0))
                            If AuxXLineCount > 0 Then
                                ReDim Preserve AuxXLine(AuxXLineCount - 1)
                                For j = 0 To AuxXLineCount - 1
                                    AuxXLine(j) = Val(v(j + 1))
                                Next
                            End If
                        Case 5
                            AuxYLineCount = Val(v(0))
                            If AuxYLineCount > 0 Then
                                ReDim Preserve AuxYLine(AuxYLineCount - 1)
                                For j = 0 To AuxYLineCount - 1
                                    AuxYLine(j) = Val(v(j + 1))
                                Next
                            End If
                        Case 6
                            'PointSize = Val(v(0))
                            TrapWidth = Val(v(1))
                            HVTrapWidth = Val(v(2))
                            
                        Case 7
                            CornerR = Val(v(0))
                            ArcStepFactor = Val(v(1))
                            SPLine_SegmentBetweenPoints = Val(v(2))
    
                     End Select
            End Select
        End If
    Loop
    Close f
End Sub

Sub SaveUndo()
    Dim UndoFileName As String, UndoFilePath As String
    
    On Error Resume Next
    
    CurUndoIndex = CurUndoIndex + 1
    CurUndoCount = CurUndoIndex
    
    UndoFilePath = App.Path
    If Right(UndoFilePath, 1) <> "\" Then UndoFilePath = UndoFilePath & "\Temp\"
    
    MkDir UndoFilePath
    UndoFileName = UndoFilePath & "UndoFile" & Format(CurUndoIndex, "0000") & ".TMP"
    
    WriteFile UndoFileName
    
    FrmMain.MnuUndo.Enabled = True
    FrmMain.MnuRedo.Enabled = False
    
    FrmMain.Toolbar1.Buttons(9).Enabled = True
    FrmMain.Toolbar1.Buttons(10).Enabled = False
    
    DataChanged = True
End Sub

Sub Undo(Optional KeepStatus As Boolean = False)
    Dim UndoFileName As String, UndoFilePath As String
    
    If CurUndoIndex > 0 Then
        CurUndoIndex = CurUndoIndex - 1
        
        If CurUndoIndex > 0 Then
            UndoFilePath = App.Path
            If Right(UndoFilePath, 1) <> "\" Then UndoFilePath = UndoFilePath & "\Temp\"
            UndoFileName = UndoFilePath & "UndoFile" & Format(CurUndoIndex, "0000") & ".TMP"
            
            ReadFile UndoFileName
            
            CloseXORStack
            FrmMain.PicPathCls
            DrawAll
        Else
            CloseXORStack
            DeleteAll
            FrmMain.PicPathCls
            DrawAll
        End If
        
        If KeepStatus = False Then
            CurTool = ToolType.None
            FrmMain.TxtCurTool = ""
            CurToolStep = 0
            FrmMain.FraEdit.Visible = False
        End If
    End If
    
    If CurUndoIndex = 0 Then
        FrmMain.MnuUndo.Enabled = False
        FrmMain.Toolbar1.Buttons(9).Enabled = False
    End If
    FrmMain.MnuRedo.Enabled = True
    FrmMain.Toolbar1.Buttons(10).Enabled = True
End Sub

Sub Redo()
    Dim UndoFileName As String, UndoFilePath As String
    
    If CurUndoIndex < CurUndoCount Then
        CurUndoIndex = CurUndoIndex + 1
        
        UndoFilePath = App.Path
        If Right(UndoFilePath, 1) <> "\" Then UndoFilePath = UndoFilePath & "\Temp\"
        UndoFileName = UndoFilePath & "UndoFile" & Format(CurUndoIndex, "0000") & ".TMP"

        ReadFile UndoFileName
        CloseXORStack
        FrmMain.PicPathCls
        DrawAll
        
        CurTool = ToolType.None
        FrmMain.TxtCurTool = ""
        CurToolStep = 0
    End If

    If CurUndoIndex = CurUndoCount Then
        FrmMain.MnuRedo.Enabled = False
        FrmMain.Toolbar1.Buttons(10).Enabled = False
    
    End If
    FrmMain.MnuUndo.Enabled = True
    FrmMain.Toolbar1.Buttons(9).Enabled = True
End Sub

Sub DeleteAllUndoFiles()
    Dim UndoFileName As String, UndoFilePath As String
    
    On Error Resume Next
    
    UndoFilePath = App.Path
    If Right(UndoFilePath, 1) <> "\" Then UndoFilePath = UndoFilePath & "\Temp\"
    
    UndoFileName = dir(UndoFilePath & "*.*")
    Do While UndoFileName <> ""
        If UndoFileName <> "." And UndoFileName <> ".." Then
            DeleteFile UndoFilePath & UndoFileName
        End If
        UndoFileName = dir
    Loop
    RmDir UndoFilePath
    
    CurUndoIndex = 0
    CurUndoCount = 0
    
    FrmMain.MnuUndo.Enabled = False
    FrmMain.MnuRedo.Enabled = False
    
    FrmMain.Toolbar1.Buttons(9).Enabled = False
    FrmMain.Toolbar1.Buttons(10).Enabled = False
End Sub

Function FileName(ByVal File As String) As String
    Dim p As Long, s As String
    
    If File = "" Then
        FileName = ""
        Exit Function
    End If
    
    p = InStrRev(File, "\")
    s = Mid(File, p + 1)
    p = InStrRev(s, ".")
    FileName = UCase(Mid(s, 1, p - 1))
End Function

Sub WriteUserParameter()
    Dim f As Long, I As Long, fn As String
    
    fn = App.Path & "\" & App.EXEName & "_Data"
    
    f = FreeFile
    Open fn For Output As f
    
    Print #f, "BEGIN ENVIRONMENT"
    Print #f, MainGridX; ","; MainGridY; ","; SubGridX; ","; SubGridY
    Print #f, UserOrgX; ","; UserOrgY; ","; UserOrgZ
    
    Print #f, LayerMax; ",";
    For I = 0 To LayerMax - 1
        Print #f, LayerZValue(I); ","; LayerColor(I); ",";
    Next
    Print #f, LayerZValue(I); ","; LayerColor(I)
    
    Print #f, AuxXLineCount; ",";
    For I = 0 To AuxXLineCount - 1
        Print #f, AuxXLine(I); ",";
    Next
    Print #f,
    
    Print #f, AuxYLineCount; ",";
    For I = 0 To AuxYLineCount - 1
        Print #f, AuxYLine(I); ",";
    Next
    Print #f,
    
    Print #f, PointSize; ","; TrapWidth; ","; HVTrapWidth
    Print #f, CornerR; ","; ArcStepFactor; ","; MinPathStep; ","; SPLine_SegmentBetweenPoints
    Print #f, "END ENVIRONMENT"
        
    Close f
End Sub


Sub ReadUserParameter()
    Dim f As Long, s As String, v As Variant, stp As Integer, I As Long, j As Long
    Dim fn As String
                            
    PointSize = 5
    
    fn = App.Path & "\" & App.EXEName & "_Data"
    If dir(fn) = "" Then Exit Sub
    
    On Error Resume Next
    
    f = FreeFile
    Open fn For Input As f
    Do While Not EOF(f)
        Line Input #f, s
        If Trim(s) <> "" Then
            Select Case Trim(UCase(s))
                Case "BEGIN ENVIRONMENT"
                    stp = 1000
                Case "END ENVIRONMENT"
                    stp = 0
            End Select
            
            Select Case stp
                Case 1000
                    I = 0
                    stp = 1001
                Case 1001
                    I = I + 1
                    v = Split(s, ",")
                    Select Case I
                        Case 1
                            MainGridX = Val(v(0))
                            MainGridY = Val(v(1))
                            SubGridX = Val(v(2))
                            SubGridY = Val(v(3))
                        
                        Case 2
                            UserOrgX = Val(v(0))
                            UserOrgY = Val(v(1))
                            UserOrgZ = Val(v(2))
                            
                        Case 3
                            LayerMax = Val(v(0))
                            For j = 0 To LayerMax
                                LayerZValue(j) = Val(v(2 * j + 1))
                                LayerColor(j) = Val(v(2 * j + 2))
                            Next
                            
                        Case 4
                            AuxXLineCount = Val(v(0))
                            If AuxXLineCount > 0 Then
                                ReDim Preserve AuxXLine(AuxXLineCount - 1)
                                For j = 0 To AuxXLineCount - 1
                                    AuxXLine(j) = Val(v(j + 1))
                                Next
                            End If
                        Case 5
                            AuxYLineCount = Val(v(0))
                            If AuxYLineCount > 0 Then
                                ReDim Preserve AuxYLine(AuxYLineCount - 1)
                                For j = 0 To AuxYLineCount - 1
                                    AuxYLine(j) = Val(v(j + 1))
                                Next
                            End If
                        Case 6
                            'PointSize = Val(v(0))
                            TrapWidth = Val(v(1))
                            HVTrapWidth = Val(v(2))
                            
                        Case 7
                            CornerR = Val(v(0))
                            ArcStepFactor = Val(v(1))
                            MinPathStep = Val(v(2))
                            SPLine_SegmentBetweenPoints = Val(v(3))
                     End Select
            End Select
        End If
    Loop
    Close f
End Sub


