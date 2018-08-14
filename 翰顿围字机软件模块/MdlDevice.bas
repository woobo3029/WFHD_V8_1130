Attribute VB_Name = "MdlDevice"
Option Explicit

'====================================================================
Public IsRunning As Boolean
Public StopRunning As Boolean
Public PauseRunning As Boolean
Public IsFeeding As Boolean

Public VertResetOK As Boolean

''Public Const CtrlCardType = 1       '1 代表 9030， 0 代表adt8940a
'
''====================================================================
''If CtrlCardType = 0 Then
''    Public Const FeedAxis = 1
''    Public Const BendAxis = 2
''   Public Const VertAxis = 3
''    Public Const VertUpDownAxis = 4
''Else
'    Public Const FeedAxis = 0
'    Public Const BendAxis = 1
'    Public Const VertAxis = 2
'    Public Const VertUpDownAxis = 3
'
''End If

Public Const TimToChgSta = 20

Public Const VertOrgSensor = 22
Public Const VertLowSensor = 17
Public Const VertHighSensor = 23
Public Const KniftOrgSensor = 16
Public Const FeedOrgSensor = 5

Public Const VertMotorPort = 1

Public Const VertMoveUpDownPort = 1

Public Const VertMoveUpPort = 1
Public Const VertMoveDownPort = 2

Public Const VertClosePort = 3
Public Const FeedFWPort = 4
Public Const FeedBWPort = 5
Public Const FeedFWPort2 = 6
Public Const FeedBWPort2 = 7
Public Const MagnetClampPort = 5


Public PrePauseSwitchVal As Integer
Public CurPauseSwitchVal As Integer
Public Const PauseSwitch = 11   '暂停开关I11(转接板I7)

Public Const PauseSwitch_GALIL = 7   '暂停开关IN7
'Public Const UseMagnetDO = True
Public Const TopSwitchIn = 7
Public Const BottomSwitchIn = 9

Public Const ElevatorUpSensor = 10
Public Const ElevatorDownSensor = 11

'Public Const ElevatorUpPort = 6
'Public Const ElevatorDownPort = 7

Public LastBendDir As Long '0-right, 1-left

Public PortBit(6) As Long
Public FeedPulsPerMM As Double

Public pospre0, pospre1, pospre2, pospre3 As Long
Public poscur0, poscur1, poscur2, poscur3 As Long

'====================================================================

Public Device_PulsPerMM As Double
Public Device_EncoderPulsPerMM As Double
Public Device_UseEncoder As Boolean
Public Device_BenderHome As Boolean
Public Device_VertUpdownHome As Boolean

Public Device_PulsPerDegree As Double

Public Device_AdjustmentDegree As Double
Public Device_SearchDegree As Double

Public Device_EmptyDegree As Double

'Public Device_AdjustmentDegree2 As Double
Public Device_EmptyDegree2 As Double

'Public Device_WaitUpTime As Double
'Public Device_WaitDownTime As Double

Public Device_VertMotorDrive As Boolean
Public Device_VertAllHigh As Boolean
Public Device_VertNoTurn As Boolean
    
Public Device_VertPulsPerDegree As Double
Public Device_VertAdjustmentDegree As Double
    
Public Device_VertUpDownPulsPerMM As Double
Public Device_VertUpDownAdjustmentMM As Double

Public Device_MinBendDisMM As Double
    
Public Device_HeadDistance As Double
Public Device_DoneDistance As Double
Public Device_BackSet As Double
Public Device_DoneWaitingTime As Double
Public Device_ExtendMM As Double

Public Device_CurMaterial As String
Public Device_MaterialName(10) As String

Public Device_FeedStartV As Double
Public Device_FeedSpeed As Double
Public Device_FeedAccel As Double
Public Device_FeedOffset As Double

'Public Device_ManualFeedStartV As Double
'Public Device_ManualFeedSpeed As Double
'Public Device_ManualFeedAccel As Double
'Public Device_ManualFeedOffset As Double

Public Device_BeatAngModify As Double
Public Device_BeatPtOffset As Double

Public Device_BendStartV As Double
Public Device_BendSpeed As Double
Public Device_BendAccel As Double

Public Device_ManualBendStartV As Double
Public Device_ManualBendSpeed As Double
Public Device_ManualBendAccel As Double

Public Device_ResetBendStartV As Double
Public Device_ResetBendSpeed As Double
Public Device_ResetBendAccel As Double

Public Device_TurnFeedStartV As Double
Public Device_TurnFeedSpeed As Double
Public Device_TurnFeedAccel As Double

Public Device_VertStartV As Double
Public Device_VertSpeed As Double
Public Device_VertAccel As Double

Public Device_ResetVertStartV As Double
Public Device_ResetVertSpeed As Double
Public Device_ResetVertAccel As Double

Public Device_VertUpDownStartV As Double
Public Device_VertUpDownSpeed As Double
Public Device_VertUpDownAccel As Double

Public Device_VertUpDownMM As Double

Public Device_VertMinAngle As Double
Public Device_VertMinDistance As Double
Public Device_BeatMaxRadius As Double

Public Device_TurnFeedMM As Double
Public Device_CutRadiusMM As Double

Public Device_CutDepth As Double
Public Device_CutDepth2 As Double
Public Device_Linearization As Double

Public Device_CutoffHeight As Double

Public Device_TurnPointOffsetMM As Double
Public Device_MinContinuousMM As Double

Public Device_VertKnifeDegree As Double

Public Device_VertMaxInnerAngle As Double
Public Device_VertMaxOuterAngle As Double

Public Device_InnerAngleAdjustMM As Double
Public Device_OuterAngleAdjustMM As Double

Public Device_InnerLineTerminalAdjustMM As Double
Public Device_OuterLineTerminalAdjustMM As Double

Public Device_BenderBacklash As Double
Public Device_BenderSpringback As Double

Public Device_TurnAngleDeg As Double

Public Device_FastSpeedMinLenMM As Double

Public Device_AmericanMaterial As Boolean
Public Device_TailVertAngle As Double
Public Device_VertUpDownMM_A As Double

Public Device_KareanMaterial As Boolean

Public Device_MaterialThickMM As Double
Public Device_InnerCompRatio As Double

Public Device_ArcTailModify As Integer

Public Device_StartComp As Double
Public Device_StartComp2 As Double
Public Device_EndComp As Double
Public Device_EndComp2 As Double


Public Device_StartPointAdjustMM As Double
Public Device_EndPointAdjustMM As Double

Public VertThreadStep As Long
Public VertThreadAngle As Double
Public VertThreadTime As Double

Public Device_VertMotorZoneMM As Double
Public FeedIntoVertMotorZone As Boolean
Public VertUpToMiddleWay As Boolean

Public IsCutoff As Boolean

Public BendAngle_dir0 As Long
Public Device_TotalAddDoneDistance As Double
    
Sub GetDeviceParameters()
    Dim v As Double, I As Long
    
    Device_PulsPerMM = GetValueFromINI("Device", "PulsPerMM", "100", App.Path & "\Parameters.ini")
    Device_EncoderPulsPerMM = GetValueFromINI("Device", "EncoderPulsPerMM", "100", App.Path & "\Parameters.ini")
    
    If FeedByDCMotor = True Then
        Device_UseEncoder = True
    Else
        v = GetValueFromINI("Device", "UseEncoder", "0", App.Path & "\Parameters.ini")
        Device_UseEncoder = IIf(v = 0, False, True)
    End If
    
    If FeedByDCMotor = True Then
        Device_BenderHome = True
    Else
        v = GetValueFromINI("Device", "BenderHome", "0", App.Path & "\Parameters.ini")
        Device_BenderHome = IIf(v = 0, False, True)
    End If
    
    v = GetValueFromINI("Device", "VertUpdownHome", "0", App.Path & "\Parameters.ini")
    Device_VertUpdownHome = IIf(v = 0, False, True)
    
    Device_PulsPerDegree = GetValueFromINI("Device", "PulsPerDegree", "100", App.Path & "\Parameters.ini")
    
    Device_AdjustmentDegree = GetValueFromINI("Device", "AdjustmentDegree", "0", App.Path & "\Parameters.ini")
    Device_SearchDegree = GetValueFromINI("Device", "SearchDegree", "0", App.Path & "\Parameters.ini")
    
    Device_EmptyDegree = GetValueFromINI("Device", "EmptyDegree", "0", App.Path & "\Parameters.ini")
    
    'Device_AdjustmentDegree2 = GetValueFromINI("Device", "AdjustmentDegree2", "0", App.Path & "\Parameters.ini")
    Device_EmptyDegree2 = GetValueFromINI("Device", "EmptyDegree2", "0", App.Path & "\Parameters.ini")
    
    'Device_WaitUpTime = GetValueFromINI("Device", "WaitUpTime", "0", App.Path & "\Parameters.ini")
    'Device_WaitDownTime = GetValueFromINI("Device", "WaitDownTime", "0", App.Path & "\Parameters.ini")
    
    v = GetValueFromINI("Device", "VertMotorDrive", "0", App.Path & "\Parameters.ini")
    Device_VertMotorDrive = IIf(v = 0, False, True)
    v = GetValueFromINI("Device", "VertAllHigh", "0", App.Path & "\Parameters.ini")
    Device_VertAllHigh = IIf(v = 0, False, True)
    
    If bNoBendUIKorea = True Then
        'v = GetValueFromINI("Device", "VertNoTurn", "0", App.Path & "\Parameters.ini")
        Device_VertNoTurn = True 'IIf(v = 0, False, True)
    Else
        v = GetValueFromINI("Device", "VertNoTurn", "0", App.Path & "\Parameters.ini")
        Device_VertNoTurn = IIf(v = 0, False, True)
    End If
    Device_VertPulsPerDegree = GetValueFromINI("Device", "VertPulsPerDegree", "100", App.Path & "\Parameters.ini")
    Device_VertAdjustmentDegree = GetValueFromINI("Device", "VertAdjustmentDegree", "0", App.Path & "\Parameters.ini")
        
    Device_VertUpDownPulsPerMM = GetValueFromINI("Device", "VertUpDownPulsPerMM", "100", App.Path & "\Parameters.ini")
    Device_VertUpDownAdjustmentMM = GetValueFromINI("Device", "VertUpDownAdjustmentMM", "0", App.Path & "\Parameters.ini")
    Device_MinBendDisMM = GetValueFromINI("Device", "MinBendDisMM", "0", App.Path & "\Parameters.ini")
    
    Device_VertUpDownMM = GetValueFromINI("Device", "VertUpDownMM", "100", App.Path & "\Parameters.ini")
    Device_InnerCompRatio = GetValueFromINI("Device", "InnerCompRatio", "1", App.Path & "\Parameters.ini")
    
    Device_ArcTailModify = GetValueFromINI("Device", "ArcTailModify", "1", App.Path & "\Parameters.ini")
    
    Device_StartComp = GetValueFromINI("Device", "StartComp", "1", App.Path & "\Parameters.ini")
    Device_EndComp = GetValueFromINI("Device", "EndComp", "1", App.Path & "\Parameters.ini")
    Device_StartComp2 = GetValueFromINI("Device", "StartComp2", "1", App.Path & "\Parameters.ini")
    Device_EndComp2 = GetValueFromINI("Device", "EndComp2", "1", App.Path & "\Parameters.ini")
    
    Device_StartPointAdjustMM = GetValueFromINI("Device", "StartPointAdjustMM", "1", App.Path & "\Parameters.ini")
    Device_EndPointAdjustMM = GetValueFromINI("Device", "EndPointAdjustMM", "1", App.Path & "\Parameters.ini")
        
    Device_HeadDistance = GetValueFromINI("Device", "HeadDistance", "0", App.Path & "\Parameters.ini")
    Device_DoneDistance = GetValueFromINI("Device", "DoneDistance", "0", App.Path & "\Parameters.ini")
    
    If Device_DoneDistance = 0 Then
        Device_DoneDistance = 0.1
    End If
    Device_BackSet = GetValueFromINI("Device", "BackSet", "0", App.Path & "\Parameters.ini")
    
    Device_DoneWaitingTime = GetValueFromINI("Device", "DoneWaitingTime", "0", App.Path & "\Parameters.ini")
    Device_ExtendMM = GetValueFromINI("Device", "ExtendMM", "0", App.Path & "\Parameters.ini")
    
    Device_CurMaterial = GetStringFromINI("Device", "CurMaterial", "Material00", App.Path & "\Parameters.ini")
    For I = 1 To 10
        Device_MaterialName(I) = GetStringFromINI("MaterialName", str(I), "#" & Trim(str(I)), App.Path & "\Parameters.ini")
    Next
    
    Device_FeedStartV = GetValueFromINI("Device", "FeedStartV", "1000", App.Path & "\Parameters.ini")
    Device_FeedSpeed = GetValueFromINI("Device", "FeedSpeed", "2000", App.Path & "\Parameters.ini")
    Device_FeedAccel = GetValueFromINI("Device", "FeedAccel", "1000", App.Path & "\Parameters.ini")
    Device_FeedOffset = GetValueFromINI("Device", "FeedOffset", "0", App.Path & "\Parameters.ini")
    
    'Device_ManualFeedStartV = GetValueFromINI("Device", "ManualFeedStartV", "1000", App.Path & "\Parameters.ini")
    'Device_ManualFeedSpeed = GetValueFromINI("Device", "ManualFeedSpeed", "2000", App.Path & "\Parameters.ini")
    'Device_ManualFeedAccel = GetValueFromINI("Device", "ManualFeedAccel", "1000", App.Path & "\Parameters.ini")
    'Device_ManualFeedOffset = GetValueFromINI("Device", "ManualFeedOffset", "10", App.Path & "\Parameters.ini")
    
    Device_BendStartV = GetValueFromINI("Device", "BendStartV", "1000", App.Path & "\Parameters.ini")
    Device_BendSpeed = GetValueFromINI("Device", "BendSpeed", "2000", App.Path & "\Parameters.ini")
    Device_BendAccel = GetValueFromINI("Device", "BendAccel", "1000", App.Path & "\Parameters.ini")
    
    Device_ManualBendStartV = GetValueFromINI("Device", "ManualBendStartV", "1000", App.Path & "\Parameters.ini")
    Device_ManualBendSpeed = GetValueFromINI("Device", "ManualBendSpeed", "2000", App.Path & "\Parameters.ini")
    Device_ManualBendAccel = GetValueFromINI("Device", "ManualBendAccel", "1000", App.Path & "\Parameters.ini")
    
    Device_ResetBendStartV = GetValueFromINI("Device", "ResetBendStartV", "1000", App.Path & "\Parameters.ini")
    Device_ResetBendSpeed = GetValueFromINI("Device", "ResetBendSpeed", "2000", App.Path & "\Parameters.ini")
    Device_ResetBendAccel = GetValueFromINI("Device", "ResetBendAccel", "1000", App.Path & "\Parameters.ini")
    
    'Device_TurnFeedStartV = GetValueFromINI("Device", "TurnFeedStartV", "200", App.Path & "\Parameters.ini")
    'Device_TurnFeedSpeed = GetValueFromINI("Device", "TurnFeedSpeed", "400", App.Path & "\Parameters.ini")
    'Device_TurnFeedAccel = GetValueFromINI("Device", "TurnFeedAccel", "400", App.Path & "\Parameters.ini")
    
    Device_VertUpDownStartV = GetValueFromINI("Device", "VertUpDownStartV", "100", App.Path & "\Parameters.ini")
    Device_VertUpDownSpeed = GetValueFromINI("Device", "VertUpDownSpeed", "1000", App.Path & "\Parameters.ini")
    Device_VertUpDownAccel = GetValueFromINI("Device", "VertUpDownAccel", "1000", App.Path & "\Parameters.ini")
    
    Device_VertStartV = GetValueFromINI("Device", "VertStartV", "1000", App.Path & "\Parameters.ini")
    Device_VertSpeed = GetValueFromINI("Device", "VertSpeed", "2000", App.Path & "\Parameters.ini")
    Device_VertAccel = GetValueFromINI("Device", "VertAccel", "1000", App.Path & "\Parameters.ini")
    
    Device_ResetVertStartV = GetValueFromINI("Device", "ResetVertStartV", "100", App.Path & "\Parameters.ini")
    Device_ResetVertSpeed = GetValueFromINI("Device", "ResetVertSpeed", "1000", App.Path & "\Parameters.ini")
    Device_ResetVertAccel = GetValueFromINI("Device", "ResetVertAccel", "1000", App.Path & "\Parameters.ini")
    
    Device_VertMinAngle = GetValueFromINI("Device", "VertMinAngle", "15", App.Path & "\Parameters.ini")
    Device_VertMinDistance = GetValueFromINI("Device", "VertMinDistance", "5", App.Path & "\Parameters.ini")
    Device_BeatMaxRadius = GetValueFromINI("Device", "BeatMaxRadius", "300", App.Path & "\Parameters.ini")
    Device_TurnFeedMM = GetValueFromINI("Device", "TurnFeedMM", "3", App.Path & "\Parameters.ini")
    Device_CutRadiusMM = GetValueFromINI("Device", "CutRadiusMM", "3", App.Path & "\Parameters.ini")
    Device_CutDepth = GetValueFromINI("Device", "CutDepth", "3", App.Path & "\Parameters.ini")
    Device_CutDepth2 = GetValueFromINI("Device", "CutDepth2", "3", App.Path & "\Parameters.ini")
    Device_Linearization = GetValueFromINI("Device", "Linearization", "3", App.Path & "\Parameters.ini")
    
    Device_CutoffHeight = GetValueFromINI("Device", "CutoffHeight", "3", App.Path & "\Parameters.ini")
    
    Device_TurnPointOffsetMM = GetValueFromINI("Device", "TurnPointOffsetMM", "3", App.Path & "\Parameters.ini")
    Device_MinContinuousMM = GetValueFromINI("Device", "MinContinuousMM", "3", App.Path & "\Parameters.ini")
    
    Device_VertKnifeDegree = GetValueFromINI("Device", "VertKnifeDegree", "45", App.Path & "\Parameters.ini")
    
    Device_VertMaxInnerAngle = GetValueFromINI("Device", "VertMaxInnerAngle", "120", App.Path & "\Parameters.ini")
    Device_VertMaxOuterAngle = GetValueFromINI("Device", "VertMaxOuterAngle", "80", App.Path & "\Parameters.ini")
    
    Device_InnerAngleAdjustMM = GetValueFromINI("Device", "InnerAngleAdjustMM", "0", App.Path & "\Parameters.ini")
    Device_OuterAngleAdjustMM = GetValueFromINI("Device", "OuterAngleAdjustMM", "0", App.Path & "\Parameters.ini")
    
    Device_InnerLineTerminalAdjustMM = GetValueFromINI("Device", "InnerLineTerminalAdjustMM", "0", App.Path & "\Parameters.ini")
    Device_OuterLineTerminalAdjustMM = GetValueFromINI("Device", "OuterLineTerminalAdjustMM", "0", App.Path & "\Parameters.ini")
    
    Device_BenderBacklash = GetValueFromINI("Device", "BenderBacklash", "0", App.Path & "\Parameters.ini")
    Device_BenderSpringback = GetValueFromINI("Device", "BenderSpringback", "0.5", App.Path & "\Parameters.ini")
    Device_TurnAngleDeg = GetValueFromINI("Device", "TurnAngleDeg", "35", App.Path & "\Parameters.ini")
    
    Device_FastSpeedMinLenMM = GetValueFromINI("Device", "FastSpeedMinLenMM", "20", App.Path & "\Parameters.ini")
    Device_VertMotorZoneMM = GetValueFromINI("Device", "VertMotorZoneMM", "50", App.Path & "\Parameters.ini")
    
    v = GetValueFromINI("Device", "AmericanMaterial", "0", App.Path & "\Parameters.ini")
    Device_AmericanMaterial = IIf(v = 0, False, True)
    
    Device_TailVertAngle = GetValueFromINI("Device", "TailVertAngle", "130", App.Path & "\Parameters.ini")
    Device_VertUpDownMM_A = GetValueFromINI("Device", "VertUpDownMM_A", "100", App.Path & "\Parameters.ini")
    
    v = GetValueFromINI("Device", "KareanMaterial", "0", App.Path & "\Parameters.ini")
    Device_KareanMaterial = IIf(v = 0, False, True)
    
    Device_MaterialThickMM = GetValueFromINI("MaterialThickMM", Device_CurMaterial, "0.8", App.Path & "\Parameters.ini")
    
    '----------------------------------------------------------------------
    'FrmMain.ChkStartPointVert90.Visible = Not Device_AmericanMaterial
    'FrmMain.ChkEndPointVert90.Visible = Not Device_AmericanMaterial
End Sub

Sub SetDeviceParameters()
    Dim I As Long
    
    WritePrivateProfileString "Device", "PulsPerMM", str(Device_PulsPerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EncoderPulsPerMM", str(Device_EncoderPulsPerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "UseEncoder", IIf(Device_UseEncoder = True, "1", "0"), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "BenderHome", IIf(Device_BenderHome = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpdownHome", IIf(Device_VertUpdownHome = True, "1", "0"), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "PulsPerDegree", str(Device_PulsPerDegree), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "AdjustmentDegree", str(Device_AdjustmentDegree), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "SearchDegree", str(Device_SearchDegree), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "EmptyDegree", str(Device_EmptyDegree), App.Path & "\Parameters.ini"
    
    'WritePrivateProfileString "Device", "AdjustmentDegree2", Str(Device_AdjustmentDegree2), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EmptyDegree2", str(Device_EmptyDegree2), App.Path & "\Parameters.ini"
    
    'WritePrivateProfileString "Device", "WaitUpTime", Str(Device_WaitUpTime), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "WaitDownTime", Str(Device_WaitDownTime), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertMotorDrive", IIf(Device_VertMotorDrive = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertAllHigh", IIf(Device_VertAllHigh = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertNoTurn", IIf(Device_VertNoTurn = True, "1", "0"), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertUpDownPulsPerMM", str(Device_VertUpDownPulsPerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownAdjustmentMM", str(Device_VertUpDownAdjustmentMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "MinBendDisMM", str(Device_MinBendDisMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownMM", str(Device_VertUpDownMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "InnerCompRatio", str(Device_InnerCompRatio), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "ArcTailModify", str(Device_ArcTailModify), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "StartComp", str(Device_StartComp), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EndComp", str(Device_EndComp), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "StartComp2", str(Device_StartComp2), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EndComp2", str(Device_EndComp2), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "StartPointAdjustMM", str(Device_StartPointAdjustMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EndPointAdjustMM", str(Device_EndPointAdjustMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertPulsPerDegree", str(Device_VertPulsPerDegree), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertAdjustmentDegree", str(Device_VertAdjustmentDegree), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "HeadDistance", str(Device_HeadDistance), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "DoneDistance", str(Device_DoneDistance), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "BackSet", str(Device_BackSet), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "DoneWaitingTime", str(Device_DoneWaitingTime), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ExtendMM", str(Device_ExtendMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "CurMaterial", Device_CurMaterial, App.Path & "\Parameters.ini"
    For I = 1 To 10
        WritePrivateProfileString "MaterialName", str(I), Device_MaterialName(I), App.Path & "\Parameters.ini"
    Next
    
    WritePrivateProfileString "Device", "FeedStartV", str(Device_FeedStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "FeedSpeed", str(Device_FeedSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "FeedAccel", str(Device_FeedAccel), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "FeedOffset", str(Device_FeedOffset), App.Path & "\Parameters.ini"
    
    'WritePrivateProfileString "Device", "ManualFeedStartV", str(Device_ManualFeedStartV), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "ManualFeedSpeed", str(Device_ManualFeedSpeed), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "ManualFeedAccel", str(Device_ManualFeedAccel), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "ManualFeedOffset", str(Device_ManualFeedOffset), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "BendStartV", str(Device_BendStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "BendSpeed", str(Device_BendSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "BendAccel", str(Device_BendAccel), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "ManualBendStartV", str(Device_ManualBendStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ManualBendSpeed", str(Device_ManualBendSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ManualBendAccel", str(Device_ManualBendAccel), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "ResetBendStartV", str(Device_ResetBendStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ResetBendSpeed", str(Device_ResetBendSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ResetBendAccel", str(Device_ResetBendAccel), App.Path & "\Parameters.ini"
    
    'WritePrivateProfileString "Device", "TurnFeedStartV", Str(Device_TurnFeedStartV), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "TurnFeedSpeed", Str(Device_TurnFeedSpeed), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "TurnFeedAccel", Str(Device_TurnFeedAccel), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownStartV", str(Device_VertUpDownStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownSpeed", str(Device_VertUpDownSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownAccel", str(Device_VertUpDownAccel), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertStartV", str(Device_VertStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertSpeed", str(Device_VertSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertAccel", str(Device_VertAccel), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "ResetVertStartV", str(Device_ResetVertStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ResetVertSpeed", str(Device_ResetVertSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ResetVertAccel", str(Device_ResetVertAccel), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertMinAngle", str(Device_VertMinAngle), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertMinDistance", str(Device_VertMinDistance), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "BeatMaxRadius", str(Device_BeatMaxRadius), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "TurnFeedMM", str(Device_TurnFeedMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "CutRadiusMM", str(Device_CutRadiusMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "CutDepth", str(Device_CutDepth), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "CutDepth2", str(Device_CutDepth2), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "Linearization", str(Device_Linearization), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "CutoffHeight", str(Device_CutoffHeight), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "TurnPointOffsetMM", str(Device_TurnPointOffsetMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "MinContinuousMM", str(Device_MinContinuousMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertKnifeDegree", str(Device_VertKnifeDegree), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertMaxInnerAngle", str(Device_VertMaxInnerAngle), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertMaxOuterAngle", str(Device_VertMaxOuterAngle), App.Path & "\Parameters.ini"

    WritePrivateProfileString "Device", "InnerAngleAdjustMM", str(Device_InnerAngleAdjustMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "OuterAngleAdjustMM", str(Device_OuterAngleAdjustMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "InnerLineTerminalAdjustMM", str(Device_InnerLineTerminalAdjustMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "OuterLineTerminalAdjustMM", str(Device_OuterLineTerminalAdjustMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "BenderBacklash", str(Device_BenderBacklash), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "BenderSpringback", str(Device_BenderSpringback), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "TurnAngleDeg", str(Device_TurnAngleDeg), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "FastSpeedMinLenMM", str(Device_FastSpeedMinLenMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertMotorZoneMM", str(Device_VertMotorZoneMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "AmericanMaterial", IIf(Device_AmericanMaterial = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "TailVertAngle", str(Device_TailVertAngle), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownMM_A", str(Device_VertUpDownMM_A), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "KareanMaterial", IIf(Device_KareanMaterial = True, "1", "0"), App.Path & "\Parameters.ini"
End Sub

Public Sub DCMotorFeedFWOn()
'Debug.Print "DCMotorFeedFWOn"
    write_bit 0, FeedFWPort, 1
    write_bit 0, FeedFWPort, 1
    write_bit 0, FeedFWPort, 1
End Sub

Public Sub DCMotorFeedFWOff()
'Debug.Print "DCMotorFeedFWOff"
    write_bit 0, FeedFWPort, 0
    write_bit 0, FeedFWPort, 0
    write_bit 0, FeedFWPort, 0
End Sub

Public Sub DCMotorFeedFWOn2()
'Debug.Print "DCMotorFeedFWOn"
    write_bit 0, FeedFWPort2, 1
    write_bit 0, FeedFWPort2, 1
    write_bit 0, FeedFWPort2, 1
End Sub

Public Sub DCMotorFeedFWOff2()
'Debug.Print "DCMotorFeedFWOff"
    write_bit 0, FeedFWPort2, 0
    write_bit 0, FeedFWPort2, 0
    write_bit 0, FeedFWPort2, 0
End Sub

Public Sub DCMotorFeedBWOn()
'Debug.Print "DCMotorFeedBWOn"
    write_bit 0, FeedBWPort, 1
    write_bit 0, FeedBWPort, 1
    write_bit 0, FeedBWPort, 1
End Sub

Public Sub DCMotorFeedBWOff()
'Debug.Print "DCMotorFeedBWOff"
    write_bit 0, FeedBWPort, 0
    write_bit 0, FeedBWPort, 0
    write_bit 0, FeedBWPort, 0
End Sub

Public Sub DCMotorFeedBWOn2()
'Debug.Print "DCMotorFeedBWOn"
    write_bit 0, FeedBWPort2, 1
    write_bit 0, FeedBWPort2, 1
    write_bit 0, FeedBWPort2, 1
End Sub

Public Sub DCMotorFeedBWOff2()
'Debug.Print "DCMotorFeedBWOff"
    write_bit 0, FeedBWPort2, 0
    write_bit 0, FeedBWPort2, 0
    write_bit 0, FeedBWPort2, 0
End Sub

Sub BeatAngle(ByVal deg As Double, Optional called_by_turn As Boolean = False)
    Dim Ret As Long, cur_pos As Long, Puls As Long, Status As Long
    Dim FeedPuls As Long
    
    Dim GapPos As Long
    
    IsRunning = True
    
    If deg > 0 Then
        GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsPerDegree
    Else
        GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsPerDegree
    End If
    
    Ret = set_startv(0, BendAxis, Device_BendStartV)
    Ret = set_speed(0, BendAxis, Device_BendSpeed)
    Ret = set_acc(0, BendAxis, Device_BendAccel)
        
    FrmMain.TmrBend.Enabled = True
    
    Ret = get_command_pos(0, BendAxis, cur_pos)
    Puls = deg * Device_PulsPerDegree - cur_pos
    Ret = pmove(0, BendAxis, Puls)
    Do
        get_status 0, BendAxis, Status
        If Status = 0 Then
            Wait 0.1
            Exit Do
        End If
        DoEvents
    Loop
    
    Puls = deg * Device_PulsPerDegree - GapPos

    If called_by_turn = False Then
        Ret = pmove(0, BendAxis, -Puls)
    Else
        Ret = set_startv(0, FeedAxis, Device_TurnFeedStartV)
        Ret = set_speed(0, FeedAxis, Device_TurnFeedSpeed)
        Ret = set_acc(0, FeedAxis, Device_TurnFeedAccel)
        
        FeedPuls = Device_TurnFeedMM * Device_PulsPerMM
        'ret = inp_move2(0, BendAxis, -Puls, FeedAxis, FeedPuls)
        Ret = pmove(0, BendAxis, -Puls)
        Ret = pmove(0, FeedAxis, FeedPuls)
    End If
    Do
        get_status 0, BendAxis, Status
        If Status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    IsRunning = False
End Sub

Sub BeatRealAngle(ByVal real_deg As Double, ByVal Dis As Double)
    Dim Ret As Long, deg As Double, p As Long, q As Long, Status As Long

    Dim Curpos As Long
    Dim GapPos As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    If real_deg = 0 Then        '弯弧角为0则不必后续处理
        Exit Sub
    End If
    
    Do While StopRunning = False
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status = 0 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status = 0 Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
    
    If real_deg > 0 Then
        'GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsPerDegree
        GapPos = Device_BenderSpringback * Device_EmptyDegree2 * Device_PulsPerDegree
    Else
        'GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsPerDegree
        GapPos = -Device_BenderSpringback * Device_EmptyDegree * Device_PulsPerDegree
    End If
    
    If CtrlCardType = 0 Then
        get_command_pos 0, BendAxis, Curpos
    ElseIf CtrlCardType = 4 Then
        Curpos = GetPos(hDmc, BendAxis)
    Else
        Curpos = ReadAxisPos_9030(0, BendAxis)
    End If
            
    '------拍弧-----------------------------------------------
    IsRunning = True
    If CtrlCardType = 0 Then
        Ret = set_startv(0, BendAxis, Device_BendStartV)
        Ret = set_speed(0, BendAxis, Device_BendSpeed)
        Ret = set_acc(0, BendAxis, Device_BendAccel)
    ElseIf CtrlCardType = 4 Then
        
        Ret = SetVel(hDmc, BendAxis, Device_BendSpeed)
        Ret = SetAcc(hDmc, BendAxis, Device_BendAccel * 25)
        Ret = SetDec(hDmc, BendAxis, Device_BendAccel * 25)
    Else
        Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
        Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel * 25)
        Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel * 25)
    End If
        
    'TmrBend.Enabled = True

    'Debug.Print "Real angle,Beat Angle="; real_deg, GetBeatAngleByRealAngle(real_deg, Dis)
    If real_deg > 0 Then
        deg = GetBeatAngleByRealAngle(real_deg, Dis) + Device_EmptyDegree2
    Else
        deg = -GetBeatAngleByRealAngle(real_deg, Dis) - Device_EmptyDegree
    End If

    If StopRunning = True Then
        Exit Sub
    End If
    
    FrmMain.TmrBend_Timer
    
    'p = deg * Device_PulsPerDegree - CurPos'原来算法
    p = deg * Device_PulsPerDegree '由于9030使用绝对值编程，不用减去CurPos
    
    If CtrlCardType = 0 Then
        pmove 0, BendAxis, p
    ElseIf CtrlCardType = 4 Then
        PosMoveAbs hDmc, BendAxis, p
    Else
        SetAxisPos_9030 0, BendAxis, p
        StartAxis_9030 0, BendAxis
    End If
    
    Wait 0.01
    Do While StopRunning = False
        'get_status 0, BendAxis, status
        'If status = 0 Then
        '    Exit Do
        'End If
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        
        FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
    
    'Wait 2  '调试用以延时看拍弧位置
    
    '------复位--------------------------------------------
    If StopRunning = True Then
        Exit Sub
    End If
    
    If CtrlCardType = 0 Then
        q = -deg * Device_PulsPerDegree + GapPos
        pmove 0, BendAxis, q
    ElseIf CtrlCardType = 4 Then
        q = deg * Device_PulsPerDegree - GapPos
        PosMoveAbs hDmc, BendAxis, q
    Else
        q = deg * Device_PulsPerDegree - GapPos
        SetAxisPos_9030 0, BendAxis, q
        StartAxis_9030 0, BendAxis
    End If
    
    
    Wait 0.01
    Do While StopRunning = False
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        
        FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
End Sub
Sub TurnRealAngle_9030(ByVal real_deg As Double)
'此函数用来在将铣刀点移动到弯弧器位置时，折边。
'过程是：获取当前弯弧器位置P0，由实际角判断方向，得出一个相对于当前弯弧器角度的相对折边角P1（固定可设置，如30°）
'弯弧器移动到P0+P1-->等待运动完成-->弯弧器返回，即移动到P0-->等待运动完成。
'折边过程结束
    Dim Ret As Long, deg As Double, p As Long, q As Long, Status As Long

    Dim Curpos As Long
    'Dim GapPos As Long
    
    Dim FeedPuls As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    If real_deg = 0 Then
        Exit Sub
    End If
    
    Do While StopRunning = False
        
        Status = ReadAxisState_9030(0, BendAxis)
        If Status <> 1 Then
            Exit Do '弯弧器非运动中则进行下一步
        End If
        
    Loop
    
    '计算左右偏转空程
'    If real_deg > 0 Then
'        GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsPerDegree
'    Else
'        GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsPerDegree
'    End If
    
    IsRunning = True
    
    '获取当前弯弧器位置
        Curpos = ReadAxisPos_9030(0, BendAxis)
        
        '--- 折角 --------------------------------
        Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
        Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel * 25)
        Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel * 25)

    
'    If real_deg > 0 Then
'        deg = Device_TurnAngleDeg
'    Else
'        deg = -1 * Device_TurnAngleDeg
'    End If

    If real_deg > 0 Then
        If real_deg > 90 Then real_deg = 90
        'deg = GetTurnAngleByRealAngle(real_deg) + Device_EmptyDegree2 / 2
        deg = GetTurnAngleByRealAngle(real_deg) + Device_TurnAngleDeg
    Else
        If real_deg < -90 Then real_deg = -90
        'deg = -GetTurnAngleByRealAngle(real_deg) - Device_EmptyDegree / 2
        deg = -GetTurnAngleByRealAngle(real_deg) - Device_TurnAngleDeg
    End If
    
    'ElevatorUp
        
    If StopRunning = True Then
        Exit Sub
    End If
    
    'FrmMain.TmrBend_Timer 显示位置
    
    'p = Curpos + deg * Device_PulsPerDegree '从当前位置偏移deg
    
    p = deg * Device_PulsPerDegree
    
    Ret = SetAxisPos_9030(0, BendAxis, p)
    Ret = StartAxis_9030(0, BendAxis)
    
    'Sleep (1)
    Status = ReadAxisState_9030(0, BendAxis)
    Do While Status = 0
        Status = ReadAxisState_9030(0, BendAxis)
        DoEvents
    Loop
'Debug.Print "p="; p, "ret="; ret
    Do While StopRunning = False
       
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then '等待折边动作完成
                Exit Do
            End If
        
        'FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
'Debug.Print "<"

    '------折边复位--------------------------------------------
    

'    Ret = SetAxisStartVel_9030(0, FeedAxis, Device_TurnFeedStartV)
'    Ret = SetAxisVel_9030(0, FeedAxis, Device_TurnFeedSpeed)
'    Ret = SetAxisAcc_9030(0, FeedAxis, Device_TurnFeedAccel)
'    Ret = SetAxisDec_9030(0, FeedAxis, Device_TurnFeedAccel)
    
'    Ret = SetAxisStartVel_9030(0, FeedAxis, 1000)
'    Ret = SetAxisVel_9030(0, FeedAxis, 5000)
'    Ret = SetAxisAcc_9030(0, FeedAxis, 10000)
'    Ret = SetAxisDec_9030(0, FeedAxis, 10000)

        Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
        Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel * 25)
        Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel * 25)
        
    If real_deg > 0 Then
        If real_deg > 90 Then real_deg = 90
        deg = GetTurnAngleByRealAngle(real_deg) ' + Device_TurnAngleDeg
    Else
        If real_deg < -90 Then real_deg = -90
        deg = -GetTurnAngleByRealAngle(real_deg) ' - Device_TurnAngleDeg
    End If
    
    q = deg * Device_PulsPerDegree
    
    'q = Curpos
    
    Ret = SetAxisPos_9030(0, BendAxis, q)
    StartAxis_9030 0, BendAxis
   
'    If StopRunning = True Then
'        Exit Sub
'    End If
    
    '按原来程序设计FeedPulse为进料电机移动一段偏移量，应该去掉
    'FeedPuls = Device_TurnFeedMM * Device_PulsPerMM + ReadAxisPos_9030(0, FeedAxis)
    
    'Ret = SetAxisPos_9030(0, FeedAxis, FeedPuls)
    'StartAxis_9030 0, FeedAxis
    
    
    Do While StopRunning = False
        
        Status = ReadAxisState_9030(0, BendAxis)
        If Status <> 1 Then
            Exit Do
        End If
        
        
        'FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop

End Sub



Sub TurnRealAngle_GALIL(ByVal real_deg As Double)
'此函数用来在将铣刀点移动到弯弧器位置时，折边。
'过程是：获取当前弯弧器位置P0，由实际角判断方向，得出一个相对于当前弯弧器角度的相对折边角P1（固定可设置，如30°）
'弯弧器移动到P0+P1-->等待运动完成-->弯弧器返回，即移动到P0-->等待运动完成。
'折边过程结束
    Dim Ret As Long, deg As Double, p As Long, q As Long, Status As Long

    Dim Curpos As Long
    'Dim GapPos As Long
    
    Dim FeedPuls As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    If real_deg = 0 Then
        Exit Sub
    End If
    
    Do While StopRunning = False
        
        Status = GetStatus(hDmc, BendAxis)
        If Status <> 1 Then
            Exit Do '弯弧器非运动中则进行下一步
        End If
        
    Loop
    
   
    
    IsRunning = True
    
    '获取当前弯弧器位置
        Curpos = GetPos(0, BendAxis)
        
        '--- 折角 --------------------------------
       
        Ret = SetVel(hDmc, BendAxis, Device_BendSpeed)
        Ret = SetAcc(hDmc, BendAxis, Device_BendAccel * 25)
        Ret = SetDec(hDmc, BendAxis, Device_BendAccel * 25)

    
'    If real_deg > 0 Then
'        deg = Device_TurnAngleDeg
'    Else
'        deg = -1 * Device_TurnAngleDeg
'    End If

    If real_deg > 0 Then
        If real_deg > 90 Then real_deg = 90
        'deg = GetTurnAngleByRealAngle(real_deg) + Device_EmptyDegree2 / 2
        deg = GetTurnAngleByRealAngle(real_deg) + Device_TurnAngleDeg
    Else
        If real_deg < -90 Then real_deg = -90
        'deg = -GetTurnAngleByRealAngle(real_deg) - Device_EmptyDegree / 2
        deg = -GetTurnAngleByRealAngle(real_deg) - Device_TurnAngleDeg
    End If
    
    'ElevatorUp
        
    If StopRunning = True Then
        Exit Sub
    End If
    
    'FrmMain.TmrBend_Timer 显示位置
    
    'p = Curpos + deg * Device_PulsPerDegree '从当前位置偏移deg
    
    p = deg * Device_PulsPerDegree
    
    Ret = PosMoveAbs(hDmc, BendAxis, p)
    
    Status = GetStatus(hDmc, BendAxis)
    Sleep (1)
    
    Do While Status = 0
        Status = GetStatus(hDmc, BendAxis)
        DoEvents
    Loop
'Debug.Print "p="; p, "ret="; ret
    Do While StopRunning = False
       
           Status = GetStatus(hDmc, BendAxis)
            If Status <> 1 Then '等待折边动作完成
                Exit Do
            End If
        
        'FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
'Debug.Print "<"

    '------折边复位--------------------------------------------
        
    If real_deg > 0 Then
        If real_deg > 90 Then real_deg = 90
        deg = GetTurnAngleByRealAngle(real_deg) ' + Device_TurnAngleDeg
    Else
        If real_deg < -90 Then real_deg = -90
        deg = -GetTurnAngleByRealAngle(real_deg) ' - Device_TurnAngleDeg
    End If
    
    q = deg * Device_PulsPerDegree
    
    Ret = PosMoveAbs(hDmc, BendAxis, q)
    
    
    Do While StopRunning = False
        
       Status = GetStatus(hDmc, BendAxis)
        If Status <> 1 Then
            Exit Do
        End If
        
        
        'FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop

End Sub


Sub TurnRealAngle(ByVal real_deg As Double)
'此函数用来在将铣刀点移动到弯弧器位置时，折边。
'过程是：获取当前弯弧器位置P0，由实际角判断方向，得出一个相对于当前弯弧器角度的相对折边角P1（固定可设置，如30°）
'弯弧器移动到P0+P1-->等待运动完成-->弯弧器返回，即移动到P0-->等待运动完成。
'折边过程结束
    Dim Ret As Long, deg As Double, p As Long, q As Long, Status As Long

    Dim Curpos As Long
    Dim GapPos As Long
    
    Dim FeedPuls As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    If real_deg = 0 Then
        Exit Sub
    End If
    
    Do While StopRunning = False
        
        'get_status 0, BendAxis, status
        'If status = 0 Then
        '    Exit Do
        'End If
        'DoEvents
        
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
    Loop
    
    If real_deg > 0 Then
        GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsPerDegree
    Else
        GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsPerDegree
    End If
    
    IsRunning = True
    
    If CtrlCardType = 0 Then
        get_command_pos 0, BendAxis, Curpos
        
        '--- 折角 --------------------------------
        Ret = set_startv(0, BendAxis, Device_BendStartV)
        Ret = set_speed(0, BendAxis, Device_BendSpeed)
        Ret = set_acc(0, BendAxis, Device_BendAccel)
    Else
        Curpos = ReadAxisPos_9030(0, BendAxis)
        
        '--- 折角 --------------------------------
        Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
        Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel * 25)
        Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel * 25)

    End If
    If real_deg > 0 Then
        If real_deg > 90 Then real_deg = 90
        deg = GetTurnAngleByRealAngle(real_deg) + Device_EmptyDegree2
    Else
        If real_deg < -90 Then real_deg = -90
        deg = -GetTurnAngleByRealAngle(real_deg) - Device_EmptyDegree
    End If

    'ElevatorUp
        
    If StopRunning = True Then
        Exit Sub
    End If
    
    FrmMain.TmrBend_Timer
    
    p = deg * Device_PulsPerDegree - Curpos
    If CtrlCardType = 0 Then
        Ret = pmove(0, BendAxis, p)
    Else
        Ret = SetAxisPos_9030(0, BendAxis, p)
        Ret = StartAxis_9030(0, BendAxis)
    End If
'Debug.Print "p="; p, "ret="; ret
    Do While StopRunning = False
        'get_status 0, BendAxis, status
        'If status = 0 Then
        '    Exit Do
        'End If
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
'Debug.Print "<"

    '------复位--------------------------------------------
    If StopRunning = True Then
        Exit Sub
    End If
    
    If CtrlCardType = 0 Then
        Ret = set_startv(0, FeedAxis, Device_TurnFeedStartV)
        Ret = set_speed(0, FeedAxis, Device_TurnFeedSpeed)
        Ret = set_acc(0, FeedAxis, Device_TurnFeedAccel)
        
        q = -deg * Device_PulsPerDegree + GapPos
        Ret = pmove(0, BendAxis, q)
    Else
        Ret = SetAxisStartVel_9030(0, FeedAxis, Device_TurnFeedStartV)
        Ret = SetAxisVel_9030(0, FeedAxis, Device_TurnFeedSpeed)
        Ret = SetAxisAcc_9030(0, FeedAxis, Device_TurnFeedAccel)
        Ret = SetAxisDec_9030(0, FeedAxis, Device_TurnFeedAccel)
    
    
        
        q = -deg * Device_PulsPerDegree + GapPos
        Ret = SetAxisPos_9030(0, BendAxis, q)
        StartAxis_9030 0, BendAxis
    End If
    
    
    
    FeedPuls = Device_TurnFeedMM * Device_PulsPerMM
    If CtrlCardType = 0 Then
        Ret = pmove(0, FeedAxis, FeedPuls)
    Else
        Ret = SetAxisPos_9030(0, FeedAxis, FeedPuls)
        StartAxis_9030 0, FeedAxis
    End If
    
    Do While StopRunning = False
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        
        FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
'Debug.Print "<<"
    
    Ret = set_startv(0, FeedAxis, Device_FeedStartV)
    Ret = set_speed(0, FeedAxis, Device_FeedSpeed)
    Ret = set_acc(0, FeedAxis, Device_FeedAccel)
        
    'ElevatorDown
End Sub

Sub BendAngle(ByVal deg As Double, Optional ByManual As Boolean = False)
    Dim Ret As Long, cur_pos As Long, Puls As Long, Status As Long, backlash_puls As Long, dir As Long
    Dim dir0 As Long
        
    dir0 = BendAngle_dir0
    IsRunning = True
    
    If CtrlCardType = 0 Then
        If ByManual = True Then
            Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
            Ret = set_speed(0, BendAxis, Device_ManualBendSpeed)
            Ret = set_acc(0, BendAxis, Device_ManualBendAccel)
        Else
            Ret = set_startv(0, BendAxis, Device_BendStartV)
            Ret = set_speed(0, BendAxis, Device_BendSpeed)
            Ret = set_acc(0, BendAxis, Device_BendAccel)
        End If
    ElseIf CtrlCardType = 4 Then
        If ByManual = True Then
            'Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
            Ret = SetVel(hDmc, BendAxis, Device_ManualBendSpeed)
            Ret = SetAcc(hDmc, BendAxis, Device_ManualBendAccel)
        Else
            'Ret = set_startv(0, BendAxis, Device_BendStartV)
            Ret = SetVel(hDmc, BendAxis, Device_BendSpeed)
            Ret = SetAcc(hDmc, BendAxis, Device_BendAccel)
        End If
    Else
        If ByManual = True Then
            Ret = SetAxisStartVel_9030(0, BendAxis, Device_ManualBendStartV)
            Ret = SetAxisVel_9030(0, BendAxis, Device_ManualBendSpeed)
            Ret = SetAxisAcc_9030(0, BendAxis, Device_ManualBendAccel)
            Ret = SetAxisDec_9030(0, BendAxis, Device_ManualBendAccel)
        Else
            Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
            Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
            Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel)
            Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel)
        End If
    End If
    
    FrmMain.TmrBend.Enabled = True
    
    dir = Sgn(deg)
    If dir0 = 0 Then
        backlash_puls = dir * Device_BenderBacklash / 2
    ElseIf dir <> dir0 Then
        backlash_puls = dir * Device_BenderBacklash
    Else
        backlash_puls = 0
    End If
    BendAngle_dir0 = dir0
            
    If CtrlCardType = 0 Then
        Ret = get_command_pos(0, BendAxis, cur_pos)
        Puls = deg * Device_PulsPerDegree - cur_pos + backlash_puls
        Ret = pmove(0, BendAxis, Puls)
    ElseIf CtrlCardType = 4 Then
        
        Puls = deg * Device_PulsPerDegree + backlash_puls
        Ret = PosMoveRel(hDmc, BendAxis, Puls)
    Else
        cur_pos = ReadAxisPos_9030(0, BendAxis)
        Puls = deg * Device_PulsPerDegree + cur_pos + backlash_puls
        SetAxisPos_9030 0, BendAxis, Puls
        Ret = StartAxis_9030(0, BendAxis)
    End If
    
    
    Wait 0.02
    Do
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status = 0 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
    IsRunning = False
    
    dir0 = dir
End Sub
Sub BendAngleAbs(ByVal deg As Double, Optional ByManual As Boolean = False)
    Dim Ret As Long, cur_pos As Long, Puls As Long, Status As Long, backlash_puls As Long, dir As Long
    Dim dir0 As Long
        
    dir0 = BendAngle_dir0
    IsRunning = True
    
    If CtrlCardType = 0 Then
    ElseIf CtrlCardType = 4 Then
        If ByManual = True Then
            'Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
            Ret = SetVel(hDmc, BendAxis, Device_ManualBendSpeed)
            Ret = SetAcc(hDmc, BendAxis, Device_ManualBendAccel)
        Else
            'Ret = set_startv(0, BendAxis, Device_BendStartV)
            Ret = SetVel(hDmc, BendAxis, Device_BendSpeed)
            Ret = SetAcc(hDmc, BendAxis, Device_BendAccel)
        End If
    Else
        If ByManual = True Then
            Ret = SetAxisStartVel_9030(0, BendAxis, Device_ManualBendStartV)
            Ret = SetAxisVel_9030(0, BendAxis, Device_ManualBendSpeed)
            Ret = SetAxisAcc_9030(0, BendAxis, Device_ManualBendAccel)
            Ret = SetAxisDec_9030(0, BendAxis, Device_ManualBendAccel)
        Else
            Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
            Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
            Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel)
            Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel)
        End If
    End If
    
    FrmMain.TmrBend.Enabled = True
    
    dir = Sgn(deg)
    If dir0 = 0 Then
        backlash_puls = dir * Device_BenderBacklash / 2
    ElseIf dir <> dir0 Then
        backlash_puls = dir * Device_BenderBacklash
    Else
        backlash_puls = 0
    End If
    BendAngle_dir0 = dir0
            
    If CtrlCardType = 0 Then
    ElseIf CtrlCardType = 4 Then
        
        Puls = deg * Device_PulsPerDegree + backlash_puls
        Ret = PosMoveAbs(hDmc, BendAxis, Puls)
    Else
        'cur_pos = ReadAxisPos_9030(0, BendAxis)
        Puls = deg * Device_PulsPerDegree + backlash_puls
        SetAxisPos_9030 0, BendAxis, Puls
        Ret = StartAxis_9030(0, BendAxis)
    End If
    
    
    Wait 0.02
    Do
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status = 0 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
    IsRunning = False
    
    dir0 = dir
End Sub

Sub BendAngleByRadius(ByVal radius As Double, Optional check_done As Boolean = True)
    Dim Ret As Long, deg0 As Double, deg As Double, cur_pos As Long, Puls As Long, Status As Long
    Dim dir0 As Long, backlash_puls As Long, dir As Long
        
    dir0 = BendAngle_dir0
    
    IsRunning = True
    
    If CtrlCardType = 0 Then
        sudden_stop 0, BendAxis
    ElseIf CtrlCardType = 4 Then
        StopAxis hDmc, BendAxis
    Else
        CeaseAxis_9030 0, BendAxis
    End If
    Do
        'get_status 0, BendAxis, status
        'If status = 0 Then
        '    Exit Do
        'End If
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
       
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
'Debug.Print "BendAngleByRadius="; radius,

    If CtrlCardType = 0 Then
        Ret = set_startv(0, BendAxis, Device_BendStartV)
        Ret = set_speed(0, BendAxis, Device_BendSpeed)
        Ret = set_acc(0, BendAxis, Device_BendAccel)
            
        Puls = -Sgn(deg) * Device_PulsPerDegree
        Ret = pmove(0, BendAxis, Puls)
    ElseIf CtrlCardType = 4 Then
        'Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
        Ret = SetVel(hDmc, BendAxis, Device_BendSpeed)
        Ret = SetAcc(hDmc, BendAxis, Device_BendAccel * 25)
        Ret = SetDec(hDmc, BendAxis, Device_BendAccel * 25)
            
        Puls = -Sgn(deg) * Device_PulsPerDegree
    Else
        Ret = SetAxisStartVel_9030(0, BendAxis, Device_BendStartV)
        Ret = SetAxisVel_9030(0, BendAxis, Device_BendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_BendAccel * 25)
        Ret = SetAxisDec_9030(0, BendAxis, Device_BendAccel * 25)
            
        Puls = -Sgn(deg) * Device_PulsPerDegree
        'Puls = 0
        'Ret = SetAxisPos_9030(0, BendAxis, Puls)
        'StartAxis_9030 0, BendAxis
    End If
    
    'Sleep 3
    Do
        If CtrlCardType = 0 Then
            get_status 0, BendAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        ElseIf CtrlCardType = 4 Then
            Status = GetStatus(hDmc, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        Else
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
    
    'FrmMain.TmrBend_Timer
    
    If CtrlCardType = 0 Then
        Ret = get_command_pos(0, BendAxis, cur_pos)
        
    ElseIf CtrlCardType = 4 Then
        cur_pos = GetPos(hDmc, BendAxis)
    Else
        cur_pos = ReadAxisPos_9030(0, BendAxis)
    End If
    
   
        If radius = 0 Then
            deg0 = 0
            deg = 0
        Else
            If radius < 0 Then
                deg0 = -GetBendAngleByRadius(radius)
                deg = deg0 - Device_EmptyDegree
            Else
                deg0 = GetBendAngleByRadius(radius)
                deg = deg0 + Device_EmptyDegree2
            End If
        End If
    
    
    BendAngle_dir0 = dir0
    
    
'    dir = Sgn(deg)
'    If dir0 = 0 Then
'        backlash_puls = dir * Device_BenderBacklash / 2
'    ElseIf dir <> dir0 Then
'        backlash_puls = dir * Device_BenderBacklash
'    Else
'        backlash_puls = 0
'    End If
            
    'Puls = deg * Device_PulsPerDegree - cur_pos ' + backlash_puls
    Puls = deg * Device_PulsPerDegree + backlash_puls ' 原来的算法如上，由于9030采用绝对位置编程，不用在减去当前位置
    
    If CtrlCardType = 0 Then
        Ret = pmove(0, BendAxis, Puls)
    ElseIf CtrlCardType = 4 Then
        Ret = PosMoveAbs(hDmc, BendAxis, Puls)
    Else
        Ret = SetAxisPos_9030(0, BendAxis, Puls)
        StartAxis_9030 0, BendAxis
    End If
'Debug.Print "degree="; deg, "pulse="; Puls
    
    Wait 0.005
    'Status = ReadAxisState_9030(0, BendAxis)
    
    If check_done = True Then
        Sleep 3
        Do
            If CtrlCardType = 0 Then
                get_status 0, BendAxis, Status
                If Status = 0 Then
                    Exit Do
                End If
            ElseIf CtrlCardType = 4 Then
                Status = GetStatus(hDmc, BendAxis)
                If Status <> 1 Then
                    Exit Do
                End If
            Else
                Status = ReadAxisState_9030(0, BendAxis)
                If Status <> 1 Then
                    Exit Do
                End If
            End If
        
        '    TmrBend_Timer
        '    ShowFeedMarkPoint
        '    ShowVertMarkPoint
            DoEvents
        Loop
    End If
    IsRunning = False
End Sub
Sub BendReset_9030()
    Dim Ret As Long, Puls As Long, Status As Long
  
    SetAxisAcc_9030 0, BendAxis, Device_ResetBendAccel * 25
    SetAxisDec_9030 0, BendAxis, Device_ResetBendAccel * 25
    SetAxisStopDec_9030 0, BendAxis, Device_ResetBendAccel * 2500
    Puls = Device_AdjustmentDegree * Device_PulsPerDegree
    
    Ret = GoHome_9030(0, BendAxis, Device_ResetBendSpeed, -Device_ResetBendSpeed, Puls, 0, 0, 0)
    
    Do
        Sleep (1)
        Status = ReadAxisState_9030(0, BendAxis)
        If Status = -1 Then     '等待回零完成
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    
    
'    SetAxisAcc_9030 0, BendAxis, Device_ResetBendAccel * 25
'    SetAxisDec_9030 0, BendAxis, Device_ResetBendAccel * 25
'    SetAxisStopDec_9030 0, BendAxis, Device_ResetBendAccel * 2500
'    Puls = Device_AdjustmentDegree * Device_PulsPerDegree
'
'    Ret = GoHome_9030(0, BendAxis, Device_ResetBendSpeed / 10, -Device_ResetBendSpeed / 20, 0, 0, 0, 0)
'
'    Do
'        Sleep (1)
'        status = ReadAxisState_9030(0, BendAxis)
'        If status = -1 Then     '等待回零完成
'            Exit Do
'        End If
'
'        If StopRunning = True Then
'            Exit Sub
'        End If
'        DoEvents
'    Loop
'
'    SetAxisStartVel_9030 0, BendAxis, Device_ResetBendStartV
'    SetAxisVel_9030 0, BendAxis, Device_ResetBendSpeed
'    SetAxisPos_9030 0, BendAxis, Puls
'    StartAxis_9030 0, BendAxis
'
'    Do
'        Sleep (1)
'        status = ReadAxisState_9030(0, BendAxis)
'        If status = 0 Then     '等待回零完成
'            Exit Do
'        End If
'
'        If StopRunning = True Then
'            Exit Sub
'        End If
'        DoEvents
'    Loop
    
    Home_9030 0, BendAxis
    'BendAngle 10, True
    If StopRunning = True Then
        Exit Sub
    End If
End Sub
Sub BendReset_9030_V8() '自制机械式复位开关
    Dim Ret As Long, Puls As Long, Status As Long
  
    SetAxisIO_9030 0, BendAxis, 2, 3, 1, 5  '将bendaxis原点恢复设置成原点限位
    
    SetAxisAcc_9030 0, BendAxis, Device_ResetBendAccel * 25
    SetAxisDec_9030 0, BendAxis, Device_ResetBendAccel * 25
    SetAxisStopDec_9030 0, BendAxis, Device_ResetBendAccel * 2500
    
    '先按回零反方向运行左空程角度
    SetAxisVel_9030 0, BendAxis, Device_ResetBendSpeed
    SetAxisPos_9030 0, BendAxis, -1 * Device_SearchDegree * Device_PulsPerDegree
    StartAxis_9030 0, BendAxis
    '等待完成
    Wait 0.1
    Do
        Sleep (1)
        Status = ReadAxisState_9030(0, BendAxis)
        If Status = 0 Then    '等待回零完成
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    '回零
    
    Puls = Device_AdjustmentDegree * Device_PulsPerDegree
    
    Ret = GoHome_9030(0, BendAxis, Device_ResetBendSpeed, -Device_ResetBendSpeed, Puls, 0, 0, 0) '第一次回零回退5mm
    
    Do
        Sleep (1)
        Status = ReadAxisState_9030(0, BendAxis)
        If Status = -1 Then     '等待回零完成
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    '第二次要压在弯弧原点上停止，先将弯弧原点设为正限位
    
'    SetAxisAcc_9030 0, BendAxis, Device_ResetBendAccel * 25
'    SetAxisDec_9030 0, BendAxis, Device_ResetBendAccel * 25
'    SetAxisStopDec_9030 0, BendAxis, Device_ResetBendAccel * 2500
'    Puls = Device_AdjustmentDegree * Device_PulsPerDegree
'
'    SetAxisIO_9030 0, BendAxis, 1, 3, 1, 5  '将bendaxis原点暂时设置成正限位
'    SetAxisPos_9030 0, BendAxis, 20 * Device_PulsPerDegree
'    StartAxis_9030 0, BendAxis
'    status = ReadAxisState_9030(0, BendAxis)
'    Do
'        Sleep (1)
'        status = ReadAxisState_9030(0, BendAxis)
'        If status = -7 Then     '等待碰到正限位
'            Exit Do
'        End If
'
'        If StopRunning = True Then
'            Exit Sub
'        End If
'        DoEvents
'    Loop
''    MsgBox "Bendaxis home ok"
'
'    CeaseAxis_9030 0, BendAxis
'    SetAxisIO_9030 0, BendAxis, 1, 3, 1, 0  '将bendaxis原点恢复设置成原点限位
'
'    SetAxisAcc_9030 0, BendAxis, Device_ResetBendAccel * 25
'    SetAxisDec_9030 0, BendAxis, Device_ResetBendAccel * 25
'    SetAxisStopDec_9030 0, BendAxis, Device_ResetBendAccel * 2500
'
'    '先按回零反方向运行左空程角度
'    SetAxisVel_9030 0, BendAxis, Device_ResetBendSpeed
'    SetAxisPos_9030 0, BendAxis, Device_AdjustmentDegree * Device_PulsPerDegree
'    StartAxis_9030 0, BendAxis
'    '等待完成
'    Wait 0.1
'    Do
'        Sleep (1)
'        status = ReadAxisState_9030(0, BendAxis)
'        If status = 0 Then    '等待回零完成
'            Exit Do
'        End If
'
'        If StopRunning = True Then
'            Exit Sub
'        End If
'        DoEvents
'    Loop
    
    Home_9030 0, BendAxis
    Sleep 100
    Home_9030 0, BendAxis
    Sleep 100
    Home_9030 0, BendAxis
    'BendAngle 10, True
    If StopRunning = True Then
        Exit Sub
    End If
End Sub
Sub BendReset_GALIL_V8() '自制机械式复位开关
    Dim Ret As Long, Puls As Long, Status As Long
    
    
    '先按回零反方向运行搜索角度
    
    SetAcc hDmc, BendAxis, Device_ResetBendSpeed * 10
    SetDec hDmc, BendAxis, Device_ResetBendSpeed * 10
    SetVel hDmc, BendAxis, Device_ResetBendSpeed
    PosMoveRel hDmc, BendAxis, Device_SearchDegree * Device_PulsPerDegree
    
    
    '等待运动完成
    Sleep (5)
    Wait 0.1
    Do
        
        Status = GetStatus(hDmc, BendAxis)
        If Status = 0 Then
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    '回零启动
    
    GoHome hDmc, BendAxis, Device_ResetBendSpeed
    
    
    
    
    '等待回零完成
    Sleep (1)
    Do
        
        Status = GetStatus(hDmc, BendAxis)
        If Status = 0 Then     '等待回零完成
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    '回退
    Puls = -Device_AdjustmentDegree * Device_PulsPerDegree
    PosMoveAbs hDmc, BendAxis, Puls
    Sleep (5)
    Do
        Sleep (1)
        Status = GetStatus(hDmc, BendAxis)
        If Status = 0 Then     '等待回零完成
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
   
   DefinePos hDmc, BendAxis, 0
   DefineEnc hDmc, BendAxis, 0
    If StopRunning = True Then
        Exit Sub
    End If
End Sub
Sub BendReset()
    Dim Ret As Long, Puls As Long, Status As Long
    '反向转20度
    Ret = set_startv(0, BendAxis, Device_ResetBendStartV)
    Ret = set_speed(0, BendAxis, Device_ResetBendSpeed)
    Ret = set_acc(0, BendAxis, Device_ResetBendAccel)
    Puls = -20 * Device_PulsPerDegree
    Ret = pmove(0, BendAxis, Puls)
    Do
        get_status 0, BendAxis, Status
        If Status = 0 Then
            Exit Do
        End If
        '如被停止，则退出函数
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    If StopRunning = True Then
        Exit Sub
    End If
    '反转过程中执行找零
    Ret = home1(0, BendAxis, 0, 0, -1, Device_ResetBendStartV, Device_ResetBendSpeed, Device_ResetBendAccel, 3 * Device_PulsPerDegree, Device_ResetBendSpeed / 8, 0, 360 * Device_PulsPerDegree)
    Do
        get_status 0, BendAxis, Status
        If Status = 0 Then
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    '回零完成后，回到调整位置
    Ret = set_startv(0, BendAxis, Device_ResetBendStartV)
    Ret = set_speed(0, BendAxis, Device_ResetBendSpeed)
    Ret = set_acc(0, BendAxis, Device_ResetBendAccel)
    Puls = Device_AdjustmentDegree * Device_PulsPerDegree
    Ret = pmove(0, BendAxis, Puls)
    '等待调整结束
    Do
        get_status 0, BendAxis, Status
        If Status = 0 Then
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    set_command_pos 0, BendAxis, 0
    set_actual_pos 0, BendAxis, 0
    
    BendAngle 0, True
End Sub

Sub FeedV3_9030(Optional dr As Long = 1)
'   进料轴以位置模式 从当前位置增量运动 dr*1000000个脉冲
'   参数 dr 表示方向
    Dim Ret As Long
    Dim pos_cur As Long
    Dim outval_vertmotor As Integer
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
    If StopRunning = True Then
        IsRunning = False
        Exit Sub
    End If
    
    FrmMain.LblSpeedMode.BackColor = RGB(0, 255, 0)
        
    If VertUpDownByDCMotor = True Then
        outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
        Do
            outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
            If outval_vertmotor = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
    End If
    
    Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
    Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedSpeed)   '大于减速距离部分的进料，速度以Device_FeedSpeed
    'Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
    'Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
    pos_cur = ReadAxisPos_9030(0, FeedAxis)
    Ret = SetAxisPos_9030(0, FeedAxis, dr * 10000000 + pos_cur)
    Ret = StartAxis_9030(0, FeedAxis)
        
End Sub


Sub FeedV3_GALIL(Optional dr As Long = 1)
'   进料轴以位置模式 从当前位置增量运动 dr*1000000个脉冲
'   参数 dr 表示方向
    Dim Ret As Long
    Dim pos_cur As Long
    Dim outval_vertmotor As Integer
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
    If StopRunning = True Then
        IsRunning = False
        Exit Sub
    End If
    
    FrmMain.LblSpeedMode.BackColor = RGB(0, 255, 0)
    
    SetVel hDmc, FeedAxis, Device_FeedSpeed
    SetAcc hDmc, FeedAxis, Device_FeedAccel
    SetDec hDmc, FeedAxis, Device_FeedAccel * 25
    PosMoveRel hDmc, FeedAxis, dr * 10000000
    
        
End Sub

Sub FeedV(Optional dr As Long = 1)
'   进料轴以位置模式 从当前位置增量运动 dr*1000000个脉冲
'   参数 dr 表示方向
    Dim Ret As Long
    Dim pos_cur As Long
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
    
    FrmMain.LblSpeedMode.BackColor = RGB(0, 255, 0)
        
    If FeedByDCMotor = True Then
        'Debug.Print "FeedV"
        DCMotorFeedFWOff2
        DCMotorFeedFWOn
    Else
        If CtrlCardType = 0 Then
            Ret = set_startv(0, FeedAxis, Device_FeedStartV / 5)
            Ret = set_speed(0, FeedAxis, Device_FeedStartV)
            Ret = set_acc(0, FeedAxis, Device_FeedAccel / 5)
                
            Ret = pmove(0, FeedAxis, dr * 10000000)
        ElseIf CtrlCardType = 4 Then
            SetAcc hDmc, FeedAxis, Device_FeedAccel / 5
            SetDec hDmc, FeedAxis, Device_FeedAccel
            SetVel hDmc, FeedAxis, Device_FeedStartV
            PosMoveRel hDmc, FeedAxis, dr * 10000000
        Else
            Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
            Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedStartV)   '小于减速距离的进料，速度以StartV
            'Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
            'Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
            pos_cur = ReadAxisPos_9030(0, FeedAxis)
            Ret = SetAxisPos_9030(0, FeedAxis, dr * 10000000 + pos_cur)
            Ret = StartAxis_9030(0, FeedAxis)
        End If
        
    End If
End Sub
Sub FeedV_GALIL(Optional dr As Long = 1)
'   进料轴以位置模式 从当前位置增量运动 dr*1000000个脉冲
'   参数 dr 表示方向
    Dim Ret As Long
    Dim pos_cur As Long
    Dim outval_vertmotor As Integer
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
    
    If StopRunning = True Then
        IsRunning = False
        Exit Sub
    End If
    
    FrmMain.LblSpeedMode.BackColor = RGB(0, 255, 0)
    
    If VertUpDownByDCMotor = True Then
        outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
        Do
            outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
            If outval_vertmotor = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
    End If
        
   
    SetAcc hDmc, FeedAxis, Device_FeedAccel / 5
    SetDec hDmc, FeedAxis, Device_FeedAccel * 5
    Ret = SetVel(hDmc, FeedAxis, Device_FeedStartV)   '小于减速距离的进料，速度以StartV
    
    Ret = PosMoveRel(hDmc, FeedAxis, dr * 10000000)
    
        
End Sub
Sub FeedV_9030(Optional dr As Long = 1)
'   进料轴以位置模式 从当前位置增量运动 dr*1000000个脉冲
'   参数 dr 表示方向
    Dim Ret As Long
    Dim pos_cur As Long
    Dim outval_vertmotor As Integer
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
    
    If StopRunning = True Then
        IsRunning = False
        Exit Sub
    End If
    
    FrmMain.LblSpeedMode.BackColor = RGB(0, 255, 0)
    
    If VertUpDownByDCMotor = True Then
        outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
        Do
            outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
            If outval_vertmotor = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
    End If
        
   
    Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV / 2)
    Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedStartV)   '小于减速距离的进料，速度以StartV
    'Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
    'Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
    pos_cur = ReadAxisPos_9030(0, FeedAxis)
    Ret = SetAxisPos_9030(0, FeedAxis, dr * 10000000 + pos_cur)
    Ret = StartAxis_9030(0, FeedAxis)
        
End Sub

Sub FeedV2(Optional dr As Long = 1)
    Dim Ret As Long
    Dim pos_cur As Long
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
        
    FrmMain.LblSpeedMode.BackColor = RGB(255, 0, 0)

    If FeedByDCMotor = True Then
        'Debug.Print "FeedV2"
        DCMotorFeedFWOff
        DCMotorFeedFWOn2
    Else
        If CtrlCardType = 0 Then
            Ret = set_startv(0, FeedAxis, Device_FeedStartV)
            Ret = set_speed(0, FeedAxis, Device_FeedSpeed)
            Ret = set_acc(0, FeedAxis, Device_FeedAccel)
                
            Ret = pmove(0, FeedAxis, dr * 10000000)
        ElseIf CtrlCardType = 4 Then
            SetAcc hDmc, FeedAxis, Device_FeedAccel
            SetDec hDmc, FeedAxis, Device_FeedAccel * 25
            SetVel hDmc, FeedAxis, Device_FeedSpeed
            PosMoveRel hDmc, FeedAxis, dr * 10000000
        Else
            Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
            Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedSpeed)
            Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
            Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
            pos_cur = ReadAxisPos_9030(0, FeedAxis)
            Ret = SetAxisPos_9030(0, FeedAxis, dr * 10000000 + pos_cur)
            Ret = StartAxis_9030(0, FeedAxis)
            
        End If
    End If
End Sub

Sub FeedV3(ByVal high_speed_encoder_feed_pulse As Long)
    Dim Ret As Long
    Dim pulse As Long
    Dim pos_cur As Long
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
        
    FrmMain.LblSpeedMode.BackColor = RGB(255, 0, 0)

    If FeedByDCMotor = True Then
        'Debug.Print "FeedV2"
        DCMotorFeedFWOff
        DCMotorFeedFWOn2
    Else
        If CtrlCardType = 0 Then
            Ret = set_startv(0, FeedAxis, Device_FeedStartV)
            Ret = set_speed(0, FeedAxis, Device_FeedSpeed)
            Ret = set_acc(0, FeedAxis, Device_FeedAccel)
               
            If Device_EncoderPulsPerMM > 0 Then
                pulse = high_speed_encoder_feed_pulse * Device_PulsPerMM / Device_EncoderPulsPerMM
                Ret = pmove(0, FeedAxis, pulse)
            End If
        ElseIf CtrlCardType = 4 Then
            
            Ret = SetVel(hDmc, FeedAxis, Device_FeedSpeed)    '大于减速距离部分的进料，速度以Device_FeedSpeed
            Ret = SetAcc(hDmc, FeedAxis, Device_FeedAccel)
            Ret = SetDec(hDmc, FeedAxis, Device_FeedAccel)
               
            If Device_EncoderPulsPerMM > 0 Then
                pulse = high_speed_encoder_feed_pulse * Device_PulsPerMM / Device_EncoderPulsPerMM
                PosMoveRel hDmc, FeedAxis, pulse
                
            End If
        Else
           Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
            Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedSpeed)    '大于减速距离部分的进料，速度以Device_FeedSpeed
            Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
            Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
               
            If Device_EncoderPulsPerMM > 0 Then
                pulse = high_speed_encoder_feed_pulse * Device_PulsPerMM / Device_EncoderPulsPerMM
                pos_cur = ReadAxisPos_9030(0, FeedAxis)
                Ret = SetAxisPos_9030(0, FeedAxis, pulse + pos_cur)
                Ret = StartAxis_9030(0, FeedAxis)
            End If
        End If
        FrmMain.TmrFeedV3Thread.Enabled = True
    End If
End Sub

Sub StopFeedV(Optional id As Long = 0)
    Dim cur_encoder_pos As Long, target_encoder_pos As Long, dep As Long, DP As Long, Ret As Long, Status As Long
        
    FrmMain.LblSpeedMode.BackColor = RGB(240, 240, 240)
    IsFeeding = False

    If FeedByDCMotor = True Then
        'Debug.Print "StopFeedV"
        DCMotorFeedFWOff
        DCMotorFeedBWOff
        DCMotorFeedFWOff2
        DCMotorFeedBWOff2
    Else
        'reset_fifo 0
        If CtrlCardType = 0 Then
            sudden_stop 0, FeedAxis
        ElseIf CtrlCardType = 4 Then
            SetDec hDmc, FeedAxis, 20000000 '模拟急停使得加速度足够大，否则程序出错，可以设置5000来观察错误情况
            StopAxis hDmc, FeedAxis
            Do
                DoEvents
                If GetStatus(hDmc, FeedAxis) = 0 Then
                    Exit Do
                End If
            Loop
            SetDec hDmc, FeedAxis, Device_FeedAccel
        Else
            CeaseAxis_9030 0, FeedAxis
        End If
        
'        If Device_UseEncoder = True And id > 0 Then
'            get_actual_pos 0, FeedAxis, cur_encoder_pos
'            target_encoder_pos = PathOutputPoint(id).LengthFromStart * Device_EncoderPulsPerMM
'            dep = cur_encoder_pos - target_encoder_pos
'            If Device_EncoderPulsPerMM > 0 Then
'                DP = Int(-dep * Device_PulsPerMM / Device_EncoderPulsPerMM)
'            End If
'
'            Ret = set_startv(0, FeedAxis, Device_FeedStartV / 5)
'            Ret = set_speed(0, FeedAxis, Device_FeedStartV)
'            Ret = set_acc(0, FeedAxis, Device_FeedAccel / 5)
'
'            Ret = pmove(0, FeedAxis, DP)
'            Do
'                get_status 0, FeedAxis, status
'                If status = 0 Then
'                    Exit Do
'                End If
'
'                If StopRunning = True Then
'                    Exit Sub
'                End If
'                DoEvents
'            Loop
'        End If
    End If
End Sub

Sub FeedMM(ByVal MM As Double, ByVal use_encoder As Boolean, ByVal wait_sec As Double, Optional ShowText As Boolean = True)
    Dim Ret As Long, Status As Long, cur_pos As Long, feed_puls As Double, cur_feed_puls As Double, t0 As Double, t As Double
    Dim fastMM As Double, slowMM As Double
    
    Dim fast_startv As Long
    Dim fast_speed As Long
    Dim fast_accel As Long
    
    Dim slow_startv As Long
    Dim slow_speed As Long
    Dim slow_accel As Long
    
    Dim nLogPos As Long                   '逻辑位置
    Dim nActPos As Long                   '实际位置
    Dim nSpeed As Long                    '运行速度
    
    Dim target_encoder_pos As Long, dep As Long, DP As Long
    
    Dim npos As Long
    Dim wucha As Double
    
    'On Error Resume Next
    
    IsRunning = True
    StopRunning = False
        
    fast_startv = Device_FeedStartV
    fast_speed = Device_FeedSpeed
    fast_accel = Device_FeedAccel
    
    slow_startv = Device_FeedStartV
    slow_speed = Device_FeedStartV
    slow_accel = Device_FeedAccel
    
    If use_encoder = False Then '不使用编码器
        If CtrlCardType = 0 Then
            Ret = set_startv(0, FeedAxis, fast_startv)
            Ret = set_speed(0, FeedAxis, fast_speed)
            Ret = set_acc(0, FeedAxis, fast_accel)
            
            feed_puls = MM * Device_PulsPerMM
            Ret = pmove(0, FeedAxis, feed_puls)
        ElseIf CtrlCardType = 4 Then
            'Ret = set_startv(0, FeedAxis, fast_startv)
            Ret = SetVel(hDmc, FeedAxis, fast_speed)
            Ret = SetAcc(hDmc, FeedAxis, fast_accel)
            Ret = SetDec(hDmc, FeedAxis, fast_accel)
            
            feed_puls = MM * Device_PulsPerMM
            Ret = PosMoveRel(hDmc, FeedAxis, feed_puls)
        Else
            Ret = SetAxisStartVel_9030(0, FeedAxis, fast_startv)
            Ret = SetAxisVel_9030(0, FeedAxis, fast_speed)
            Ret = SetAxisAcc_9030(0, FeedAxis, fast_accel)
            Ret = SetAxisDec_9030(0, FeedAxis, fast_accel)
            
            feed_puls = MM * Device_PulsPerMM + ReadAxisPos_9030(0, FeedAxis)
            Ret = SetAxisPos_9030(0, FeedAxis, feed_puls)
            Ret = StartAxis_9030(0, FeedAxis)
            
        End If
        
    Else    '使用编码器手动进退料
        If CtrlCardType = 0 Then    '8941A卡
            get_actual_pos 0, FeedAxis, cur_pos
            If Abs(MM) - Device_FastSpeedMinLenMM > 0 Then
                fastMM = Abs(MM) - Device_FastSpeedMinLenMM
                slowMM = Device_FastSpeedMinLenMM
            Else
                fastMM = 0
                slowMM = MM
            End If
        
            If fastMM > 0 Then
                
                Ret = set_startv(0, FeedAxis, fast_startv)
                Ret = set_speed(0, FeedAxis, fast_speed)
                Ret = set_acc(0, FeedAxis, fast_accel)
            
                feed_puls = Sgn(MM) * fastMM * Device_PulsPerMM
                Ret = pmove(0, FeedAxis, feed_puls)
                
                Do
                    
                    get_status 0, FeedAxis, Status
                    If Status = 0 Then
                        Exit Do
                    End If
                    
                    If StopRunning = True Then
                        Exit Sub
                    End If
                    DoEvents
                Loop
            
           End If
            
            
            Ret = set_startv(0, FeedAxis, slow_startv)
            Ret = set_speed(0, FeedAxis, slow_speed)
            Ret = set_acc(0, FeedAxis, slow_accel)
            
            feed_puls = Abs(MM) * Device_EncoderPulsPerMM
            Ret = pmove(0, FeedAxis, IIf(MM > 0, 10000000, -10000000))
            
        
            Do
                Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
            
                If nSpeed = 0 Then
                    Exit Do
                End If
                
                cur_feed_puls = nActPos - cur_pos
            
                If Abs(cur_feed_puls) >= feed_puls - Device_FeedOffset Then
                    sudden_stop 0, FeedAxis
                    Exit Do
                End If
               
                DoEvents
            Loop
            
            '========================================================================
            target_encoder_pos = MM * Device_EncoderPulsPerMM + cur_pos
            dep = nActPos - target_encoder_pos
            If Device_EncoderPulsPerMM > 0 Then
                DP = Int(-dep * Device_PulsPerMM / Device_EncoderPulsPerMM)
            End If
                
            
            Ret = pmove(0, FeedAxis, DP)
            
            Do
                
                get_status 0, FeedAxis, Status
                
                If Status = 0 Then
                    Exit Do
                End If
                
                If StopRunning = True Then
                    Exit Sub
                End If
                DoEvents
            Loop
        ElseIf CtrlCardType = 4 Then        '使用编码器进料，galil卡
            cur_pos = GetPosEnc(hDmc, FeedAxis)
            If Abs(MM) - Device_FastSpeedMinLenMM > 0 Then
                fastMM = Abs(MM) - Device_FastSpeedMinLenMM
                slowMM = Device_FastSpeedMinLenMM
            Else
                fastMM = 0
                slowMM = MM
            End If
            If fastMM > 0 Then
                'Ret = SetAxisStartVel_9030(0, FeedAxis, Sgn(MM) * fast_startv)
                Ret = SetAcc(hDmc, FeedAxis, fast_accel)
                Ret = SetDec(hDmc, FeedAxis, fast_accel)
                ContinousMove hDmc, FeedAxis, Sgn(MM) * fast_speed
                Do
                    nActPos = GetPosEnc(hDmc, FeedAxis)
                    If Abs(nActPos - cur_pos) > fastMM * Device_EncoderPulsPerMM Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                
            Else
                
                Ret = SetAcc(hDmc, FeedAxis, slow_accel)
                Ret = SetDec(hDmc, FeedAxis, slow_accel)
                ContinousMove hDmc, FeedAxis, Sgn(MM) * slow_speed
                
            End If
            
            Do
                            
                nActPos = GetPosEnc(hDmc, FeedAxis)
                If Abs(nActPos - cur_pos) > fastMM * Device_EncoderPulsPerMM Then
                    SetVel hDmc, FeedAxis, slow_speed
                    Exit Do
                End If
               
                DoEvents
            Loop
            
            Do
                            
                nActPos = GetPosEnc(hDmc, FeedAxis)
                If Abs(nActPos - cur_pos) > Abs(MM) * Device_EncoderPulsPerMM - Device_FeedOffset Then
                    SetDec hDmc, FeedAxis, 90000000
                    StopAxis hDmc, FeedAxis
                    Exit Do
                End If
               
                DoEvents
            Loop
            
            'nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
            'npos = ReadAxisPos_9030(0, FeedAxis)
            'wucha = nActPos - Abs(MM) * Device_EncoderPulsPerMM
            'Ret = SetAxisPos_9030(0, FeedAxis, npos - wucha * Device_EncoderPulsPerMM)
            'Ret = StartAxis_9030(0, FeedAxis)
            
       
        Else        '使用编码器进料，9030卡
            cur_pos = ReadAxisEncodePos_9030(0, FeedAxis)
            If Abs(MM) - Device_FastSpeedMinLenMM > 0 Then
                fastMM = Abs(MM) - Device_FastSpeedMinLenMM
                slowMM = Device_FastSpeedMinLenMM
            Else
                fastMM = 0
                slowMM = MM
            End If
            If fastMM > 0 Then
                Ret = SetAxisStartVel_9030(0, FeedAxis, Sgn(MM) * fast_startv)
                Ret = SetAxisAcc_9030(0, FeedAxis, fast_accel)
                Ret = SetAxisDec_9030(0, FeedAxis, fast_accel)
                Ret = StartAxisVel_9030(0, FeedAxis, Sgn(MM) * fast_speed)
                Do
                    nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                    If Abs(nActPos - cur_pos) > fastMM * Device_EncoderPulsPerMM Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                'Ret = SetAxisStartVel_9030(0, FeedAxis, Sgn(MM) * slow_startv)
                'Ret = SetAxisAcc_9030(0, FeedAxis, slow_accel)
                'Ret = SetAxisDec_9030(0, FeedAxis, slow_accel)
                Ret = StartAxisVel_9030(0, FeedAxis, Sgn(MM) * slow_speed)
            Else
                Ret = SetAxisStartVel_9030(0, FeedAxis, Sgn(MM) * slow_startv)
                Ret = SetAxisAcc_9030(0, FeedAxis, slow_accel)
                Ret = SetAxisDec_9030(0, FeedAxis, slow_accel)
                Ret = StartAxisVel_9030(0, FeedAxis, Sgn(MM) * slow_speed)
            End If
            
            Do
                            
                nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                If Abs(nActPos - cur_pos) > Abs(MM) * Device_EncoderPulsPerMM - Device_FeedOffset Then
                    CeaseAxis_9030 0, FeedAxis
                    Exit Do
                End If
               
                DoEvents
            Loop
            
            'nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
            'npos = ReadAxisPos_9030(0, FeedAxis)
            'wucha = nActPos - Abs(MM) * Device_EncoderPulsPerMM
            'Ret = SetAxisPos_9030(0, FeedAxis, npos - wucha * Device_EncoderPulsPerMM)
            'Ret = StartAxis_9030(0, FeedAxis)
            
        End If
    End If
    
    
        
    t0 = Timer
    Do
        t = Timer
        If TimeDiff(t, t0) > wait_sec Then
            Exit Do
        End If
    
        DoEvents
    Loop
        
    IsRunning = False
End Sub

Sub FeedMMByDCMotor(ByVal MM As Double, ByVal wait_sec As Double, Optional ShowText As Boolean = True)
    Dim Ret As Long, Status As Long, cur_pos As Long, cur_pos0 As Long, feed_puls As Double, cur_feed_puls As Double, t0 As Double, t As Double
    Dim obj As Object, stk As Long
    
    Dim feed_startv As Long
    Dim feed_speed As Long
    Dim feed_accel As Long
    Dim feed_offset As Long
    Dim feed_puls_before_change_speed As Long
    
    Dim nLogPos As Long                   '逻辑位置
    Dim nActPos As Long                   '实际位置
    Dim nSpeed As Long                    '运行速度
    
    Dim nLogPos0 As Long
    
    'On Error Resume Next
    
    IsRunning = True
    StopRunning = False
        
    feed_puls = Abs(MM) * Device_EncoderPulsPerMM
    feed_puls_before_change_speed = feed_puls - Device_FastSpeedMinLenMM * Device_EncoderPulsPerMM
    get_actual_pos 0, FeedAxis, cur_pos
    cur_pos0 = cur_pos
    
    'FrmMain.ShowFeedPos
    If MM > Device_FastSpeedMinLenMM Then
        DCMotorFeedBWOff2
        DCMotorFeedFWOn2
    ElseIf MM > 0 Then
        DCMotorFeedBWOff
        DCMotorFeedFWOn
    ElseIf MM < -Device_FastSpeedMinLenMM Then
        DCMotorFeedFWOff2
        DCMotorFeedBWOn2
    Else
        DCMotorFeedFWOff
        DCMotorFeedBWOn
    End If
    
    feed_offset = Device_FeedOffset
    
    FrmMain.LblFeedStatus.BackColor = RGB(255, 0, 0)
    stk = 0
    t0 = Timer
    Do
        If StopRunning = True Then
            Exit Do
        End If
        
        Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
        cur_feed_puls = nActPos - cur_pos
    
        If feed_puls_before_change_speed > 0 Then
            If Abs(cur_feed_puls) >= feed_puls_before_change_speed Then
                If MM > 0 Then
                    DCMotorFeedFWOff2
                    DCMotorFeedFWOn
                Else
                    DCMotorFeedBWOff2
                    DCMotorFeedBWOn
                End If
            End If
        End If
        
        If Abs(cur_feed_puls) >= feed_puls - feed_offset Then
            If MM > 0 Then
                DCMotorFeedFWOff
            Else
                DCMotorFeedBWOff
            End If
            
            'Debug.Print "采样次数："; stk; " 编码器理论脉冲："; feed_puls; " 脉冲提前量："; feed_offset; " 当前脉冲："; cur_feed_puls; " 误差脉冲："; Abs(cur_feed_puls) - (feed_puls - feed_offset)
            'If FrmTestVisible = True And ShowText Then
            '    FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "采样次数：" + str(stk) + " 编码器理论脉冲：" + str(feed_puls) + " 脉冲提前量：" + str(feed_offset) + " 当前脉冲：" + str(cur_feed_puls) + " 误差脉冲：" + str(Round(Abs(cur_feed_puls) - (feed_puls - feed_offset), 3)) + " 当前电机脉冲：" + str(nLogPos) + "(" + str(nLogPos - nLogPos0) + ")" + vbCrLf
            'End If
            cur_pos = nActPos
            nLogPos0 = nLogPos
            Exit Do
        End If
        
        DoEvents
    Loop
    FrmMain.LblFeedStatus.BackColor = RGB(220, 220, 220)
    
    t0 = Timer
    Do
        If StopRunning = True Then
            Exit Do
        End If
        
        t = Timer
        If TimeDiff(t, t0) > wait_sec Then
            Exit Do
        End If
        
        'FrmMain.ShowFeedPos
        DoEvents
    Loop
    
    Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
    FrmMain.TxtStatistics.Text = "停止所需脉冲：" + vbCrLf + str(nActPos - cur_pos) + vbCrLf + vbCrLf + "刹车距离：" + vbCrLf + str(Round((nActPos - cur_pos) / Device_EncoderPulsPerMM, 2)) + " mm" + vbCrLf + vbCrLf + "运行总脉冲：" + vbCrLf + str(nActPos - cur_pos0) + vbCrLf + vbCrLf + "进料总距离：" + vbCrLf + str(Round((nActPos - cur_pos0) / Device_EncoderPulsPerMM, 2)) + " mm" + vbCrLf + vbCrLf
    'If FrmTestVisible = True And ShowText Then
    '    FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "停止所用脉冲：" + str(nActPos - cur_pos) + " 刹车距离：" + str(str(Round((nActPos - cur_pos) / Device_EncoderPulsPerMM, 4))) + " mm" + " 当前电机脉冲：" + str(nLogPos) + "(" + str(nLogPos - nLogPos0) + ")" + vbCrLf
    'End If
    
    IsRunning = False
End Sub

Public Sub Vert(ByVal low_or_high As Integer, ByVal motor As Integer)
    Dim Sensor As Long
    Dim Ret As Long, cur_puls As Long, Puls As Long, Status As Long
    
    If motor = 1 Then
        PortBit(1) = 1
        PortBit(2) = 1
        If CtrlCardType = 0 Then
            write_bit 0, VertMotorPort, 1      '铣刀旋转马达开
            'write_bit 0, VertClosePort, 1          '铣刀靠紧
        Else
            WriteIoBit_9030 0, 1, VertMotorPort + 1
        End If
    End If
    
    If Device_VertMotorDrive = False Then
        'Sensor = IIf(low_or_high = 0, VertLowSensor, VertHighSensor)
        Sensor = VertHighSensor
        
        PortBit(3) = 1
        PortBit(4) = 0
        write_bit 0, VertMoveUpPort, 1         '铣刀向上运动
        write_bit 0, VertMoveDownPort, 0
        
        Do
            If read_bit(0, Sensor) = 0 Then   '如果遇到低位传感器
                PortBit(2) = 0
                'write_bit 0, VertClosePort, 0        '铣刀离开
                PortBit(3) = 0
                write_bit 0, VertMoveUpPort, 0       '铣刀向下运动
                Wait 0.1
                PortBit(4) = 1
                write_bit 0, VertMoveDownPort, 1
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Do
            End If
            
            Do While PauseRunning = True
                DoEvents
            Loop
            
            DoEvents
        Loop
        
        Do
            'If read_bit(0, VertOrgSensor) = 0 Then '如果遇到原点传感器
            If read_bit(0, VertLowSensor) = 0 Then '如果遇到原点传感器
                PortBit(3) = 0
                PortBit(4) = 0
                write_bit 0, VertMoveUpPort, 0      '铣刀停止运动
                write_bit 0, VertMoveDownPort, 0
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Do
            End If
            
            Do While PauseRunning = True
                DoEvents
            Loop
            
            DoEvents
        Loop
    Else
        Ret = get_command_pos(0, VertAxis, cur_puls)
        
        Ret = set_startv(0, VertAxis, Device_VertStartV)
        Ret = set_speed(0, VertAxis, Device_VertSpeed)
        Ret = set_acc(0, VertAxis, Device_VertAccel)
        
        Puls = Device_VertAdjustmentDegree * Device_VertPulsPerDegree - cur_puls
            
        Ret = pmove(0, VertAxis, Puls)
        Do
            get_status 0, VertAxis, Status
            If Status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            
            Do While PauseRunning = True
                DoEvents
            Loop
            
            DoEvents
        Loop
        
        Ret = get_command_pos(0, VertAxis, cur_puls)
        Ret = pmove(0, VertAxis, -cur_puls)
        Do
            get_status 0, VertAxis, Status
            If Status = 0 Then
                Exit Do
            End If
        
            If StopRunning = True Then
                Exit Sub
            End If
            
            Do While PauseRunning = True
                DoEvents
            Loop
            
            DoEvents
        Loop
    End If
    
    If motor = 1 Then
        PortBit(1) = 0
        If CtrlCardType = 0 Then
            write_bit 0, VertMotorPort, 0      '铣刀旋转马达关
        Else
            WriteIoBit_9030 0, 0, VertMotorPort + 1
        End If
    End If
    
    Wait 0.1
    
End Sub
Sub VertUp(ByVal low_or_high As Integer, ByVal motor As Integer)
    Dim Sensor As Long
    Dim Ret As Long, cur_puls As Long, Puls As Long, Status As Long
    
    If motor = 1 Then
        If CtrlCardType = 0 Then
            write_bit 0, VertMotorPort, 1      '铣刀旋转马达开
            write_bit 0, VertClosePort, 1          '铣刀靠紧
        Else
            WriteIoBit_9030 0, 1, VertMotorPort + 1    '铣刀旋转马达开
            WriteIoBit_9030 0, 1, VertClosePort + 1        '铣刀靠紧
        End If
    End If
    
    If Device_VertMotorDrive = False Then
        Sensor = IIf(low_or_high = 0, VertLowSensor, VertHighSensor)
        
        write_bit 0, VertMoveUpPort, 1         '铣刀向上运动
        write_bit 0, VertMoveDownPort, 0
        Do
            If read_bit(0, Sensor) = 0 Or read_bit(0, VertHighSensor) = 0 Then   '如果遇到低位传感器
                write_bit 0, VertClosePort, 0        '铣刀离开
                write_bit 0, VertMoveUpPort, 0       '铣刀停止向上运动
                'Wait 0.1
                'write_bit 0, VertMoveDownPort, 1
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Do
            End If
            
            DoEvents
        Loop
    Else
        Ret = get_command_pos(0, VertAxis, cur_puls)
        
        Ret = set_startv(0, VertAxis, Device_VertStartV)
        Ret = set_speed(0, VertAxis, Device_VertSpeed)
        Ret = set_acc(0, VertAxis, Device_VertAccel)
        
            Puls = Device_VertAdjustmentDegree * Device_VertPulsPerDegree - cur_puls
            
        Ret = pmove(0, VertAxis, Puls)
        
        
        Do
            get_status 0, VertAxis, Status
            If Status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        
        'If low_or_high = 0 Then
        '    Puls = -Device_VertLowMM * Device_VertPulsPerDegree
        'Else
        '    Puls = -Device_VertAdjustmentDegree * Device_VertPulsPerDegree
        'End If
        '
        'ret = pmove(0, VertAxis, Puls)
        'Do
        '    get_status 0, VertAxis, status
        '    If status = 0 Then
        '        Exit Do
        '    End If
        '
        '    If StopRunning = True Then
        '        Exit Sub
        '    End If
        '    DoEvents
        'Loop
    End If
    
    If motor = 1 Then
        If CtrlCardType = 0 Then
            write_bit 0, VertMotorPort, 0      '铣刀旋转马达关
            write_bit 0, VertClosePort, 0      '铣刀离开
        Else
            WriteIoBit_9030 0, 0, VertMotorPort + 1    '铣刀旋转马达开
            WriteIoBit_9030 0, 0, VertClosePort + 1        '铣刀靠紧
        End If
    End If
End Sub
Sub VertUpDownReset_9030()
    '================================================================================================================================
    '铣刀升降复位
    Dim Sensor As Long
    Dim Ret As Long, Puls As Long, Status As Long
        
    If VertUpDownByDCMotor = True Then      '如果由直流电机驱动上下
        'write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
        'Wait 0.1
        'write_bit 0, VertMoveDownPort, 0
        
   
    Else                                    '非由直流电机驱动上下
        
        
        '回零后向上移动一段距离
        Ret = SetAxisStartVel_9030(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = SetAxisAcc_9030(0, VertUpDownAxis, Device_VertUpDownAccel)
        Ret = SetAxisDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel)
        SetAxisStopDec_9030 0, VertUpDownAxis, Device_ResetBendAccel * 2500
        Puls = Device_VertUpDownAdjustmentMM * Device_VertUpDownPulsPerMM
        Ret = pmove(0, VertUpDownAxis, Puls)
        Ret = GoHome_9030(0, VertUpDownAxis, -1 * Device_VertUpDownSpeed, Device_VertUpDownSpeed, Puls, 0, 0, 0)
        Sleep (5)
        Do
            Sleep (1)
            Status = ReadAxisState_9030(0, VertUpDownAxis)
            If Status = -1 Then     '等待VertUpDownAxis回零完成
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        
    End If
End Sub
Sub VertUpDownReset_V81()
    '================================================================================================================================
    '铣刀升降复位
    Dim t As Double, t0 As Double, b As Long
    Dim Ret As Long, Puls As Long, Status As Long
        
    If VertUpDownByDCMotor = True Then      '如果由直流电机驱动上下
           
        WriteIoBit_9030 0, 1, VertMoveUpDownPort
        
        
            FrmMain.TmrDevicePortChecking.Enabled = True
            t0 = Timer
            Do
                If StopRunning = True Then
                    Exit Do
                End If
            
                Do While PauseRunning = True
                    DoEvents
                Loop
                
                b = ReadIOBit_9030(0, BottomSwitchIn)
                If b = 1 Then
                    FrmMain.LblVertHighSensor.BackColor = RGB(255, 0, 0)
                    WriteIoBit_9030 0, 0, VertMoveUpDownPort
                     Exit Do
                End If
                
                t = Timer
                If TimeDiff(t, t0) > 5 Then
                    Exit Do
                End If
            Loop
       
    End If
End Sub

Sub VertReset_9030()
     
    FrmMain.CmdMotorStop_Click
    Dim Curpos As Long
    'CmdVertUp_Click
    'CmdVertDown_Click
    'Exit Sub
    Dim Sensor As Long
    Dim Ret As Long, Puls As Long, Status As Long
    
    '铣刀转角复位
    Ret = SetAxisStartVel_9030(0, VertAxis, Device_ResetVertStartV)
    'Ret = set_speed(0, VertAxis, Device_ResetVertSpeed)
    Ret = SetAxisAcc_9030(0, VertAxis, Device_ResetVertAccel)
    Ret = SetAxisDec_9030(0, VertAxis, Device_ResetVertAccel)
    SetAxisStopDec_9030 0, VertAxis, Device_ResetBendAccel * 2500
    Puls = Device_VertAdjustmentDegree * Device_VertPulsPerDegree
    Ret = Ret = GoHome_9030(0, VertAxis, Device_ResetVertSpeed, -1 * Device_ResetVertSpeed, Puls, 0, 0, 0)
    'Ret = Ret = GoHome_9030(0, VertAxis, -1 * Device_ResetVertSpeed, Device_ResetVertSpeed, 2000, 0, 0, 0)
    Do
        Sleep (1)
        Status = ReadAxisState_9030(0, VertAxis)
        If Status = -1 Then     '等待VertAxis回零完成
            VertResetOK = True
            Exit Do
        End If
        
        If StopRunning = True Then
            VertResetOK = False
            Exit Sub
        End If
        DoEvents
    Loop
    
    '================================================================================================================================
'    '铣刀升降复位
'    If Device_VertUpdownHome = True Then
'
'
''        Curpos = ReadAxisPos_9030(0, VertUpDownAxis)
''        '向上移动一段距离
'        Ret = SetAxisStartVel_9030(0, VertUpDownAxis, Device_VertUpDownStartV)
'        Ret = SetAxisVel_9030(0, VertUpDownAxis, Device_VertUpDownSpeed)
'        Ret = SetAxisAcc_9030(0, VertUpDownAxis, Device_VertUpDownAccel)
'        Ret = SetAxisDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel)
'        Ret = SetAxisStopDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 2500)
''        Puls = Curpos + 10 * Device_VertUpDownPulsPerMM
''        Ret = SetAxisPos_9030(0, VertUpDownAxis, Puls)
''        Ret = StartAxis_9030(0, VertUpDownAxis)
''
''        Sleep (1)
''        Do
''            status = ReadAxisState_9030(0, VertUpDownAxis)
''            If status <> 1 Then    '等待VertUpDownAxis上拉完成
''                Exit Do
''            End If
''
''            If StopRunning = True Then
''                Exit Sub
''            End If
''            DoEvents
''        Loop
'
'        '铣刀升降复位位置
'        Puls = Device_VertUpDownAdjustmentMM * Device_VertUpDownPulsPerMM
'
'        Ret = GoHome_9030(0, VertUpDownAxis, Device_VertUpDownSpeed, -Device_VertUpDownSpeed, Puls, 0, 0, 0)
'        Sleep (1)
'        Do
'            status = ReadAxisState_9030(0, VertUpDownAxis)
'            If status = -1 Then     '等待VertUpDownAxis回零完成
'                Exit Do
'            End If
'
'            If StopRunning = True Then
'                Exit Sub
'            End If
'            DoEvents
'        Loop
'
'
'    End If
    '================================================================================================================================
    
End Sub
Sub VertReset_GALIL_V8()
     
    FrmMain.CmdMotorStop_Click
    Dim Curpos As Long
    'CmdVertUp_Click
    'CmdVertDown_Click
    'Exit Sub
    Dim Sensor As Long
    Dim Ret As Long, Puls As Long, Status As Long
    
    '铣刀转角复位
    
    Ret = SetAcc(hDmc, VertAxis, Device_ResetVertAccel)
    Ret = SetDec(hDmc, VertAxis, Device_ResetVertAccel * 100)
    
    
    Ret = GoHome(hDmc, VertAxis, Device_ResetVertSpeed)
    
    Do
        Sleep (1)
        Status = GetStatus(hDmc, VertAxis)
        If Status = 0 Then     '等待VertAxis回零完成
            VertResetOK = True
            Exit Do
        End If
        
        If StopRunning = True Then
            VertResetOK = False
            Exit Sub
        End If
        DoEvents
    Loop
    'Wait 2
    
    Puls = Device_VertAdjustmentDegree * Device_VertPulsPerDegree
    PosMoveAbs hDmc, VertAxis, Puls
    
    Do
        Sleep (1)
        Status = GetStatus(hDmc, VertAxis)
        If Status = 0 Then     '等待VertAxis回零完成
            VertResetOK = True
            Exit Do
        End If
        
        If StopRunning = True Then
            VertResetOK = False
            Exit Sub
        End If
        DoEvents
    Loop
    
    Wait 2
    DefinePos hDmc, VertAxis, 0
    DefineEnc hDmc, VertAxis, 0
    
End Sub



Sub VertReset()
     
    FrmMain.CmdMotorStop_Click
    'CmdVertUp_Click
    'CmdVertDown_Click
    'Exit Sub
    
    '================================================================================================================================
    '铣刀升降复位
    Dim Sensor As Long
    Dim Ret As Long, Puls As Long, Status As Long
        
    If VertUpDownByDCMotor = True Then      '如果由直流电机驱动上下
        write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
        Wait 0.01
        write_bit 0, VertMoveDownPort, 0
        
    '   Do
    '        If read_bit(0, VertOrgSensor) = 0 Then '如果遇到原点传感器
    '            write_bit 0, VertMoveUpPort, 0      '铣刀停止运动
    '            write_bit 0, VertMoveDownPort, 0
    '            Exit Do
    '        End If
    '
    '        If StopRunning = True Then
    '            Exit Do
    '        End If
    '
    '        DoEvents
    '    Loop
    Else                                    '非由直流电机驱动上下
        Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed)
        Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
        Puls = -5 * Device_VertUpDownPulsPerMM
        Ret = pmove(0, VertUpDownAxis, Puls)
        Do
            get_status 0, VertUpDownAxis, Status
            If Status = 0 Then
                Exit Do
            End If
    
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        If StopRunning = True Then
            Exit Sub
        End If
        '升降电机执行回零
        Ret = home1(0, VertUpDownAxis, 0, 0, -1, Device_VertUpDownStartV, Device_VertUpDownSpeed, Device_VertUpDownAccel, 1 * Device_VertUpDownPulsPerMM, Device_VertUpDownSpeed / 5, 0, 10 * Device_VertUpDownPulsPerMM)
        '等待回零完成，或被终止
        Do
            get_status 0, VertUpDownAxis, Status
            If Status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        '回零后向上移动一段距离
        Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed)
        Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
        Puls = 2 * Device_VertUpDownPulsPerMM
        Ret = pmove(0, VertUpDownAxis, Puls)
        Do
            get_status 0, VertUpDownAxis, Status
            If Status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        '位置寄存器清零
        set_command_pos 0, VertUpDownAxis, 0
        set_actual_pos 0, VertUpDownAxis, 0
    End If
    
    '================================================================================================================================
    '铣刀偏角复位
    Ret = set_startv(0, VertAxis, Device_ResetVertStartV)
    Ret = set_speed(0, VertAxis, Device_ResetVertSpeed)
    Ret = set_acc(0, VertAxis, Device_ResetVertAccel)
    Puls = -90 * Device_VertPulsPerDegree
    Ret = pmove(0, VertAxis, Puls)
    Do
        get_status 0, VertAxis, Status
        If Status = 0 Then
            Exit Do
        End If

        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    If StopRunning = True Then
        Exit Sub
    End If
    
    'ret = home1(0, VertAxis, 1, 0, -1, Device_VertStartV, Device_VertSpeed, Device_VertAccel, 5 * Device_VertPulsPerDegree, Device_VertSpeed / 5, 0, 100 * Device_VertPulsPerDegree)
    Ret = home1(0, VertAxis, 0, 0, -1, Device_ResetVertStartV, Device_ResetVertSpeed, Device_ResetVertAccel, 5 * Device_VertPulsPerDegree, Device_ResetVertSpeed / 5, 0, 100 * Device_VertPulsPerDegree)
    Do
        get_status 0, VertAxis, Status
        If Status = 0 Then
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    Ret = set_startv(0, VertAxis, Device_ResetVertStartV)
    Ret = set_speed(0, VertAxis, Device_ResetVertSpeed)
    Ret = set_acc(0, VertAxis, Device_ResetVertAccel)
    Puls = Device_VertAdjustmentDegree * Device_VertPulsPerDegree
    Ret = pmove(0, VertAxis, Puls)
    Do
        get_status 0, VertAxis, Status
        If Status = 0 Then
            Exit Do
        End If
        
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    set_command_pos 0, VertAxis, 0
    set_actual_pos 0, VertAxis, 0
    
End Sub

Sub GetPathXYByFeedPuls(ByVal cur_puls As Long, ByRef ux As Double, ByRef uy As Double)
    Dim I As Long, j As Long, total_puls As Long, d As Double, d1 As Double, d2 As Double, ds As Double
    Static start_id As Long
    
    cur_puls = cur_puls - Device_HeadDistance * FeedPulsPerMM
    If cur_puls <= 0 Then
        ux = -99999
        uy = -99999
        start_id = 1
        Exit Sub
    End If
    
    total_puls = TotalPathOutLength * FeedPulsPerMM
    d = 1# * cur_puls / total_puls
    
    For I = start_id To PathOutputPointCount - 1
        If PathOutputPoint(I).VertType <= 0 Then
            For j = I + 1 To PathOutputPointCount
                If PathOutputPoint(j).VertType <= 0 Then
                    Exit For
                End If
            Next
            
            d1 = (PathOutputPoint(I).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            d2 = (PathOutputPoint(j).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            If d >= d1 And d <= d2 Then
                ds = (d - d1) / (d2 - d1)
                ux = PathOutputPoint(I).ux + ds * (PathOutputPoint(j).ux - PathOutputPoint(I).ux)
                uy = PathOutputPoint(I).uy + ds * (PathOutputPoint(j).uy - PathOutputPoint(I).uy)
                start_id = I
                Exit Sub
            End If
        End If
    Next
    ux = -99999
    uy = -99999
End Sub

Sub GetPathXYByVertPuls(ByVal cur_puls As Long, ByRef ux As Double, ByRef uy As Double)
    Dim I As Long, j As Long, total_puls As Long, d As Double, d1 As Double, d2 As Double, ds As Double
    Static start_id As Long
    
    cur_puls = cur_puls - Device_HeadDistance * FeedPulsPerMM
    If cur_puls <= 0 Then
        ux = -99999
        uy = -99999
        start_id = 1
        Exit Sub
    End If
    
    total_puls = TotalPathOutLength * FeedPulsPerMM
    d = 1# * cur_puls / total_puls
    
    For I = start_id To PathOutputPointCount - 1
    'For I = 1 To PathOutputPointCount - 1
        If PathOutputPoint(I).VertType <= 0 Then
            For j = I + 1 To PathOutputPointCount
                If PathOutputPoint(j).VertType <= 0 Then
                    Exit For
                End If
            Next
            
            d1 = (PathOutputPoint(I).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            d2 = (PathOutputPoint(j).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            If d >= d1 And d <= d2 Then
                ds = (d - d1) / (d2 - d1)
                ux = PathOutputPoint(I).ux + ds * (PathOutputPoint(j).ux - PathOutputPoint(I).ux)
                uy = PathOutputPoint(I).uy + ds * (PathOutputPoint(j).uy - PathOutputPoint(I).uy)
                start_id = I
                Exit Sub
            End If
        End If
    Next
    ux = -99999
    uy = -99999
End Sub

Sub PushOut(ByVal deg As Double, ByVal bUseDO As Boolean)
    If deg = 0 Then  '当参数Device_CutDepth设置为0时， 系统不做切断处理，只有切槽
        Exit Sub
    End If
    'StopRunning = False
    IsRunning = True
    
    If bUseDO = False Then
        VertAngle -deg, False
    Else
        If CtrlCardType = 1 Then
            WriteIoBit_9030 0, 1, MagnetClampPort + 1
        ElseIf CtrlCardType = 4 Then
            WriteOutBit hDmc, MagnetClampPort, 1
        End If
        IsRunning = False
    End If
End Sub
Sub PullBack(ByVal deg As Double, ByVal bUseDO As Boolean)
    'StopRunning = False
    IsRunning = True
    
    
    If bUseDO = False Then
        VertAngle -deg
    Else
         If CtrlCardType = 1 Then
            WriteIoBit_9030 0, 0, MagnetClampPort + 1
        ElseIf CtrlCardType = 4 Then
            WriteOutBit hDmc, MagnetClampPort, 0
        End If
    End If
        IsRunning = False
    
End Sub
Sub VertAngle(ByVal deg As Double, Optional check_done As Boolean = True)
    Dim Ret As Long, cur_pos As Long, Puls As Long, Status As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
        
    Do While PauseRunning = True
        DoEvents
    Loop
    
    IsRunning = True
    
    
    '------------------------------------------------------------
    If CtrlCardType = 1 Then
        Do
            Status = ReadAxisState_9030(0, VertUpDownAxis)
            'cur_pos = ReadAxisPos_9030(0, VertUpDownAxis)
            If Status <> 1 Then    '摆刀之前，查询VertUpDownAxis轴状态，确认VertUpDownAxis轴已经停止运动,然后退出循环执行摆刀
                Exit Do
            End If
       
            DoEvents
        Loop
    End If
    '-------------------------------------------------------------
    
    If CtrlCardType = 0 Then
        Ret = set_startv(0, VertAxis, Device_VertStartV)
        Ret = set_speed(0, VertAxis, Device_VertSpeed)
        Ret = set_acc(0, VertAxis, Device_VertAccel)
            
        'TmrBend.Enabled = True
        
        Ret = get_command_pos(0, VertAxis, cur_pos)
        
        Puls = deg * Device_VertPulsPerDegree - cur_pos
        Ret = pmove(0, VertAxis, Puls)
    ElseIf CtrlCardType = 4 Then
        'Ret = set_startv(0, VertAxis, Device_VertStartV)
        Ret = SetVel(hDmc, VertAxis, Device_VertSpeed)
        Ret = SetAcc(hDmc, VertAxis, Device_VertAccel)
            
                
        Puls = deg * Device_VertPulsPerDegree
        Ret = PosMoveAbs(hDmc, VertAxis, Puls)
    Else
        Ret = SetAxisStartVel_9030(0, VertAxis, Device_VertStartV)
        Ret = SetAxisVel_9030(0, VertAxis, Device_VertSpeed)
        Ret = SetAxisAcc_9030(0, VertAxis, Device_VertAccel)
        Ret = SetAxisDec_9030(0, VertAxis, Device_VertAccel)
            
        'TmrBend.Enabled = True
        
        cur_pos = ReadAxisPos_9030(0, VertAxis)
        
        Puls = deg * Device_VertPulsPerDegree '+ cur_pos
        Ret = SetAxisPos_9030(0, VertAxis, Puls)
        Ret = StartAxis_9030(0, VertAxis)
    End If
    
    If CtrlCardType = 1 Then
    'Sleep (TimToChgSta) '等待检测铣刀偏转轴轴完全停止
    Wait 0.005
    Status = ReadAxisState_9030(0, VertAxis)
    Wait 0.005
    Status = ReadAxisState_9030(0, VertAxis)
    End If
    '------------------------------------------------------------
    'Do
    '    If CtrlCardType = 0 Then
    '    Else
    '        status = ReadAxisState_9030(0, VertAxis)
    '        If status = 1 Then     '查询VertAxis轴状态，确认轴已经开始运动,然后退出等待运动结束
    '            Exit Do
    '        End If
    '    End If
    '    DoEvents
    'Loop
    '-------------------------------------------------------------
    If check_done = True Then
        Do
            If StopRunning = True Then
                Exit Do
            End If
            
            Do While PauseRunning = True
                DoEvents
            Loop
            
            If CtrlCardType = 0 Then
                get_status 0, VertAxis, Status
                If Status = 0 Then
                    Exit Do
                End If
            ElseIf CtrlCardType = 4 Then
                Status = GetStatus(hDmc, VertAxis)
                If Status = 0 Then
                    Exit Do
                End If
            Else
                'Sleep (1)
                Status = ReadAxisState_9030(0, VertAxis)
                If Status <> 1 Then     '等待VertAxis停止
                    Exit Do
                End If
            End If
            
            DoEvents
        Loop
    End If
    IsRunning = False
End Sub

Sub VertOuterAngle(ByVal deg As Double) '手动铣外角
    deg = Abs(deg)
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If deg > Device_VertMaxOuterAngle Then
        deg = Device_VertMaxOuterAngle
    ElseIf deg < Device_VertKnifeDegree Then
        deg = Device_VertKnifeDegree
    End If
    
    If VersionNmb = 81 Then
    
        VertUpDownReset_V81
    Else
        VertMoveDown 'Up
        VertAngle -180 - (deg - Device_VertKnifeDegree) / 2
        
        If CtrlCardType = 0 Then
            write_bit 0, VertMotorPort, 1         '铣刀旋转
        Else
            WriteIoBit_9030 0, 1, VertMotorPort + 1
        End If
        
        Wait 0.005
        
        VertMoveUp 'Down   used by func VertOuterAngle
        VertAngle -180 + (deg - Device_VertKnifeDegree) / 2
        VertMoveDown 'Up
    End If
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0         '铣刀停转
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
    
    Wait 0.01
End Sub

Sub VertOuterAngle_prev(ByVal deg As Double, ByVal bTurn As Boolean)
    deg = Abs(deg)
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If deg > Device_VertMaxOuterAngle Then
        deg = Device_VertMaxOuterAngle
    ElseIf deg < Device_VertKnifeDegree Then
        deg = Device_VertKnifeDegree
    End If
    
    If bTurn = False Then
        VertMoveDown False   'Up
    End If
    
    VertThreadStep = 100
    VertThreadAngle = -180 - (deg - Device_VertKnifeDegree) / 2
    FrmMain.TmrVertThread.Enabled = True
End Sub

Sub VertOuterAngle_done(ByVal deg As Double, ByVal bTurn As Boolean)
    deg = Abs(deg)
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If deg > Device_VertMaxOuterAngle Then
        deg = Device_VertMaxOuterAngle
    ElseIf deg < Device_VertKnifeDegree Then
        deg = Device_VertKnifeDegree
    End If
    
'Debug.Print ">>>VertOuterAngle_done"
    
    Do While VertThreadStep <> 103
        DoEvents
    Loop
    
    VertThreadStep = 0
    
    '铣外角铣刀开启
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 1
        write_bit 0, VertMotorPort, 1
        write_bit 0, VertMotorPort, 1
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 1
    Else
        WriteIoBit_9030 0, 1, VertMotorPort + 1
    End If
    'Wait 0.01
    '
    If VersionNmb = 81 Then
    
        VertUpDownReset_V81
    Else
        VertMoveUp 'Down  used by VertOuterAngle_done
        If bTurn = True Then
            VertAngle -180 + (deg - Device_VertKnifeDegree) / 2
        End If
        VertMoveDown 'Up
    End If
    PullBack 0, UseMagnetDO
    
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertMotorPort, 0
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
    
    FrmMain.LblVertMotorMode.BackColor = RGB(0, 255, 0)
    'Wait 1
'Debug.Print "<<<VertOuterAngle_done"
End Sub

Sub VertInnerAngle(ByVal deg As Double, ByVal IsCutoff As Boolean)
    'deg = Abs(deg)
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
'    If Abs(deg) > Device_VertMaxInnerAngle Then
'        deg = Device_VertMaxInnerAngle
'    ElseIf Abs(deg) < Device_VertKnifeDegree Then
'        deg = Device_VertKnifeDegree
'    End If
    
    VertMoveDown 'Up
    VertAngle deg
    
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 1
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 1
        If IsCutoff = True Then
            WriteOutBit hDmc, MagnetClampPort + 1, 1
        End If
    Else
        WriteIoBit_9030 0, 1, VertMotorPort + 1
        If IsCutoff = True Then
            WriteIoBit_9030 0, 1, MagnetClampPort + 1
        End If
        
    End If
    
    'Wait 0.5
    If VersionNmb = 81 Then
    
        VertUpDownReset_V81
    Else
        VertMoveUp 'Down ,used by VertInnengle
        VertMoveDown 'Up
        
        VertAngle 0
    End If
    
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 0
        If IsCutoff = True Then
            WriteOutBit hDmc, MagnetClampPort + 1, 0
        End If
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
        WriteIoBit_9030 0, 0, MagnetClampPort + 1
    End If
    'Wait 0.1
End Sub

Sub VertInnerAngle_prev(ByVal deg As Double, ByVal bTurn As Boolean)
    deg = Abs(deg)
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If deg > Device_VertMaxInnerAngle Then
        deg = Device_VertMaxInnerAngle
    ElseIf deg < Device_VertKnifeDegree Then '
        deg = Device_VertKnifeDegree
    End If
    
    If bTurn = True Then
        VertMoveDown False
    End If
            
    VertThreadStep = 100
    VertThreadAngle = -(deg - Device_VertKnifeDegree) / 2
    FrmMain.TmrVertThread.Enabled = True
End Sub

Sub VertInnerAngle_done(ByVal deg As Double, ByVal bTurnAngle As Boolean)
    deg = Abs(deg)
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If deg > Device_VertMaxInnerAngle Then
        deg = Device_VertMaxInnerAngle
    ElseIf deg < Device_VertKnifeDegree Then
        deg = Device_VertKnifeDegree
    End If
    
'Debug.Print ">>>VertInnerAngle_done"

    VertThreadStep = 103
    Do While VertThreadStep <> 103
        DoEvents
    Loop
    VertThreadStep = 0
    
    '铣刀马达开启
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 1
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 1
    Else
        WriteIoBit_9030 0, 1, VertMotorPort + 1
    End If
    'Wait 0.1    '马达启动延时
    '
    If VersionNmb = 81 Then
    
        VertUpDownReset_V81
    Else
        VertMoveUp 'Down, used by VertInnerAngle_done
        If bTurnAngle = True Then
            VertAngle (deg - Device_VertKnifeDegree) / 2
        End If
        VertMoveDown 'Up
    End If
    PullBack 0, UseMagnetDO
    
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertMotorPort, 0
    Else
        'WriteIoBit_9030 0, 0, VertMotorPort + 1    '加工过程中不停转
        'WriteIoBit_9030 0, PortBit(2), VertClosePort + 1
        'WriteIoBit_9030 0, PortBit(3), VertMoveUpPort + 1
        'WriteIoBit_9030 0, PortBit(4), VertMoveDownPort + 1
    End If
    
    FrmMain.LblVertMotorMode.BackColor = RGB(0, 255, 0)
    'Wait 1
'Debug.Print "<<<VertInnerAngle_done"
End Sub

Sub VertEndAngle(ByVal mode As Long)
    Dim deg As Double
    
    If mode = 0 Then
        deg = Device_VertKnifeDegree / 2
    Else
        deg = -Device_VertKnifeDegree / 2
    End If
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
            
    VertMoveDown 'Up
    'VertAngle deg
    
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 1
    Else
        WriteIoBit_9030 0, 1, VertMotorPort + 1
    End If
    'Wait 0.5
    If VersionNmb = 81 Then
    
        VertUpDownReset_V81
    Else
        VertMoveUp 'Down; used by VertEndAngle
        'VertAngle deg
        VertMoveDown 'Up
    End If
    
    '铣刀马达停止
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
    'Wait 0.1
End Sub

Sub VertEndAngle_prev(ByVal mode As Long, ByVal bTurn As Boolean)
    Dim deg As Double
    
    If mode = 0 Then
        deg = Device_VertKnifeDegree / 2
    Else
        deg = -Device_VertKnifeDegree / 2
    End If
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
            
    If bTurn = True Then
        VertMoveDown False
    End If
    
    VertThreadStep = 100
    VertThreadAngle = deg
    FrmMain.TmrVertThread.Enabled = True
End Sub

Sub VertEndAngle_done(ByVal mode As Long, ByVal bTurnAngle As Boolean)
    Dim deg As Double
    
    If mode = 0 Then
        deg = Device_VertKnifeDegree / 2
    Else
        deg = -Device_VertKnifeDegree / 2
    End If
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
            
    Do While VertThreadStep <> 103
        DoEvents
    Loop
    
    VertThreadStep = 0
    
    '铣刀马达开启
    If CtrlCardType = 0 Then
    write_bit 0, VertMotorPort, 1
    write_bit 0, VertMotorPort, 1
    write_bit 0, VertMotorPort, 1
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 1
    Else
        WriteIoBit_9030 0, 1, VertMotorPort + 1
    End If
    'Wait 0.1
    '
    If VersionNmb = 81 Then
    
        VertUpDownReset_V81
    Else
        VertMoveUp  ' used by VertEndAngle_done
        If bTurnAngle = True Then
            VertAngle deg
        End If
        VertMoveDown
    End If
    
    '铣刀马达停止
    If CtrlCardType = 0 Then
    write_bit 0, VertMotorPort, 0
    write_bit 0, VertMotorPort, 0
    write_bit 0, VertMotorPort, 0
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
    
    FrmMain.LblVertMotorMode.BackColor = RGB(0, 255, 0)
    'Wait 1
End Sub

Sub VertMoveClose()
    Dim t As Double, t0 As Double, b As Long
    
    
    write_bit 0, VertClosePort, 1        '铣刀进刀
    
    t0 = Timer
    Do
        b = read_bit(0, 18)
        If b = 0 Then
            write_bit 0, VertClosePort, 0
             Exit Do
        End If
        
        t = Timer
        If TimeDiff(t, t0) > 5 Then
            Exit Do
        End If
        
        DoEvents
    Loop

End Sub

Sub VertMoveFar()
    Dim t As Double, t0 As Double, b As Long
    
    write_bit 0, VertClosePort, 1
    
    t0 = Timer
    Do
        b = read_bit(0, 19)
        If b = 0 Then
            write_bit 0, VertClosePort, 0
             Exit Do
        End If
        
        t = Timer
        If TimeDiff(t, t0) > 5 Then
            Exit Do
        End If
        
        DoEvents
    Loop
End Sub

Sub VertMoveUp(Optional check_done As Boolean = True)
    Dim t As Double, t0 As Double, b As Long
    Dim Ret As Long, cur_pos As Long, Puls As Long, Status As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If VertUpDownByDCMotor = True Then
        If CtrlCardType = 0 Then
            write_bit 0, VertMoveDownPort, 0
            'Wait 0.1
            write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
            
            If check_done = True Then
                FrmMain.TmrDevicePortChecking.Enabled = True
                t0 = Timer
                Do
                    If StopRunning = True Then
                        Exit Do
                    End If
                
                    Do While PauseRunning = True
                        DoEvents
                    Loop
                    
                    b = read_bit(0, 23)
                    If b = 0 Then
                        FrmMain.LblVertHighSensor.BackColor = RGB(255, 0, 0)
                        write_bit 0, VertMoveUpPort, 0
                         Exit Do
                    End If
                    
                    t = Timer
                    If TimeDiff(t, t0) > 5 Then
                        Exit Do
                    End If
                Loop
            End If
        Else
            WriteIoBit_9030 0, 0, VertMoveDownPort
            'Wait 0.1
            WriteIoBit_9030 0, 1, VertMoveUpPort      '铣刀向上运动
            
            If check_done = True Then
                FrmMain.TmrDevicePortChecking.Enabled = True
                t0 = Timer
                Do
                    If StopRunning = True Then
                        Exit Do
                    End If
                
                    Do While PauseRunning = True
                        DoEvents
                    Loop
                    
                    b = ReadIOBit_9030(0, TopSwitchIn)
                    If b = 1 Then
                        FrmMain.LblVertHighSensor.BackColor = RGB(255, 0, 0)
                        WriteIoBit_9030 0, 0, VertMoveUpPort
                         Exit Do
                    End If
                    
                    t = Timer
                    If TimeDiff(t, t0) > 5 Then
                        Exit Do
                    End If
                Loop
            End If
        End If
    Else
        IsRunning = True
        
        'Debug.Print "Move Up!"
        If CtrlCardType = 0 Then
            Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
            Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed)
            Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
                
            Ret = get_command_pos(0, VertUpDownAxis, cur_pos)
            
            'Puls = -cur_pos
            If Device_AmericanMaterial = False Or VertUpToMiddleWay = False Then
                Puls = -(Device_VertUpDownMM * Device_VertUpDownPulsPerMM - cur_pos)
            Else
                Puls = -(Device_VertUpDownMM_A * Device_VertUpDownPulsPerMM - cur_pos)
            End If
            
            Ret = pmove(0, VertUpDownAxis, Puls)
        ElseIf CtrlCardType = 4 Then
            'Ret = SetAxisStartVel_9030(0, VertUpDownAxis, Device_VertUpDownStartV)
            Ret = SetVel(hDmc, VertUpDownAxis, Device_VertUpDownSpeed)
            Ret = SetAcc(0, VertUpDownAxis, Device_VertUpDownAccel)
            Ret = SetDec(0, VertUpDownAxis, Device_VertUpDownAccel * 5)
            'Ret = SetAxisStopDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 25)
                
            
            If IsCutoff = True Then
                Puls = Device_CutoffHeight * Device_VertUpDownPulsPerMM             'Puls值在两种控制卡算法不同
            ElseIf Device_AmericanMaterial = True Or VertUpToMiddleWay = False Then
                Puls = Device_VertUpDownMM * Device_VertUpDownPulsPerMM             'Puls值在两种控制卡算法不同
            Else
                Puls = Device_VertUpDownMM_A * Device_VertUpDownPulsPerMM
            End If
            Ret = PosMoveAbs(hDmc, VertUpDownAxis, Puls)
            
        Else
            Ret = SetAxisStartVel_9030(0, VertUpDownAxis, Device_VertUpDownStartV)
            Ret = SetAxisVel_9030(0, VertUpDownAxis, Device_VertUpDownSpeed)
            Ret = SetAxisAcc_9030(0, VertUpDownAxis, Device_VertUpDownAccel)
            Ret = SetAxisDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 5)
            Ret = SetAxisStopDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 25)
                
            cur_pos = ReadAxisPos_9030(0, VertUpDownAxis)
            'cur_pos = 0 '
            
            'Puls = Device_VertUpDownMM * Device_VertUpDownPulsPerMM - cur_pos
            'Puls = -cur_pos
            'If Device_AmericanMaterial = False Or VertUpToMiddleWay = False Then
            If IsCutoff = True Then
                Puls = Device_CutoffHeight * Device_VertUpDownPulsPerMM             'Puls值在两种控制卡算法不同
            ElseIf Device_AmericanMaterial = True Or VertUpToMiddleWay = False Then
                Puls = Device_VertUpDownMM * Device_VertUpDownPulsPerMM             'Puls值在两种控制卡算法不同
            Else
                Puls = Device_VertUpDownMM_A * Device_VertUpDownPulsPerMM
            End If
            Ret = SetAxisPos_9030(0, VertUpDownAxis, Puls)
            Ret = StartAxis_9030(0, VertUpDownAxis)
        End If
        
        Sleep (TimToChgSta) '等待检测升降轴上升完全停止
        
        '------------------------------------------------------------
        'Do
        '    If CtrlCardType = 0 Then
        '    Else
        '        status = ReadAxisState_9030(0, VertUpDownAxis)
        '        If status = 1 Then     '查询轴状态，确认轴已经开始运动
        '            Exit Do
        '        End If
        '    End If
        '    DoEvents
        'Loop
        '-------------------------------------------------------------
        
        If check_done = True Then
            Do
                If StopRunning = True Then
                    Exit Do
                End If
            
                If CtrlCardType = 0 Then
                    get_status 0, VertUpDownAxis, Status
                    If Status = 0 Then
                        Exit Do
                    End If
                ElseIf CtrlCardType = 4 Then
                    Status = GetStatus(hDmc, VertUpDownAxis)
                    If Status = 0 Then
                        Exit Do
                    End If
                Else
                    Sleep (10)
                    Status = ReadAxisState_9030(0, VertUpDownAxis)
                    'cur_pos = ReadAxisPos_9030(0, VertUpDownAxis)
                    If Status <> 1 Then   '等待VertUpDownAxis停止
                        Exit Do
                    End If
                End If
                DoEvents
            Loop
        End If
        IsRunning = False
    End If
End Sub

Sub VertMoveDown(Optional check_done As Boolean = True)
    Dim t As Double, t0 As Double, b As Long
    Dim Ret As Long, cur_pos As Long, Puls As Long, Status As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If VertUpDownByDCMotor = True Then      'Public Const VertUpDownByDCMotor = False
        If CtrlCardType = 0 Then
            write_bit 0, VertMoveUpPort, 0       '铣刀向下运动
            Wait 0.1
            write_bit 0, VertMoveDownPort, 1
            
            If check_done = True Then
                FrmMain.TmrDevicePortChecking.Enabled = True
                t0 = Timer
                Do
                    If StopRunning = True Then
                        Exit Do
                    End If
                
                    Do While PauseRunning = True
                        DoEvents
                    Loop
                    
                    b = read_bit(0, 17)
                    If b = 0 Then
                        FrmMain.LblVertLowSensor.BackColor = RGB(255, 0, 0)
                        write_bit 0, VertMoveDownPort, 0
                        Exit Do
                    End If
                    DoEvents
                Loop
            End If
        Else
            
            WriteIoBit_9030 0, 0, VertMoveUpPort
            'Wait 0.1
            WriteIoBit_9030 0, 1, VertMoveDownPort      '铣刀向上运动
            
            If check_done = True Then
                FrmMain.TmrDevicePortChecking.Enabled = True
                t0 = Timer
                Do
                    If StopRunning = True Then
                        Exit Do
                    End If
                
                    Do While PauseRunning = True
                        DoEvents
                    Loop
                    
                    b = ReadIOBit_9030(0, BottomSwitchIn)
                    If b = 1 Then
                        FrmMain.LblVertHighSensor.BackColor = RGB(255, 0, 0)
                        WriteIoBit_9030 0, 0, VertMoveDownPort
                         Exit Do
                    End If
                    
                    t = Timer
                    If TimeDiff(t, t0) > 5 Then
                        Exit Do
                    End If
                Loop
            End If
        End If
    Else
        IsRunning = True
        If CtrlCardType = 0 Then
            Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
            Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed * 2)
            Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
            
                
            Ret = get_command_pos(0, VertUpDownAxis, cur_pos)
            
            'Puls = Device_VertUpDownMM * Device_VertUpDownPulsPerMM - cur_pos
            Puls = -cur_pos
            Ret = pmove(0, VertUpDownAxis, Puls)    '当前位置增量反向当前位置，绝对位置为0
        ElseIf CtrlCardType = 4 Then
            'Ret = SetAxisStartVel_9030(0, VertUpDownAxis, Device_VertUpDownStartV)
            Ret = SetVel(hDmc, VertUpDownAxis, Device_VertUpDownSpeed * 2)
            Ret = SetAcc(hDmc, VertUpDownAxis, Device_VertUpDownAccel)
            Ret = SetDec(hDmc, VertUpDownAxis, Device_VertUpDownAccel * 5)
            'Ret = SetAxisStopDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 25)
            Puls = 0
            'Puls = -cur_pos
            Ret = PosMoveAbs(hDmc, VertUpDownAxis, Puls)
            
        Else
            Ret = SetAxisStartVel_9030(0, VertUpDownAxis, Device_VertUpDownStartV)
            Ret = SetAxisVel_9030(0, VertUpDownAxis, Device_VertUpDownSpeed * 2)
            Ret = SetAxisAcc_9030(0, VertUpDownAxis, Device_VertUpDownAccel)
            Ret = SetAxisDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 5)
            Ret = SetAxisStopDec_9030(0, VertUpDownAxis, Device_VertUpDownAccel * 25)
                
            cur_pos = ReadAxisPos_9030(0, VertUpDownAxis)
            
            Puls = 0
            'Puls = -cur_pos
            Ret = SetAxisPos_9030(0, VertUpDownAxis, Puls)
            Ret = StartAxis_9030(0, VertUpDownAxis)
        End If
        
        Sleep (TimToChgSta) '等待检测升降轴下降完全停止
        '------------------------------------------------------------
        'Do
        '    If CtrlCardType = 0 Then
        '    Else
        '        status = ReadAxisState_9030(0, VertUpDownAxis)
        '        If status = 1 Then     '查询轴状态，确认轴已经开始运动
        '            Exit Do
        '        End If
        '    End If
        '    DoEvents
        'Loop
        '-------------------------------------------------------------
        If check_done = True Then
            Do
                If StopRunning = True Then
                    Exit Do
                End If
            
                If CtrlCardType = 0 Then
                    get_status 0, VertUpDownAxis, Status
                    If Status = 0 Then
                        Exit Do
                    End If
                ElseIf CtrlCardType = 4 Then
                    Status = GetStatus(hDmc, VertUpDownAxis)
                    If Status = 0 Then
                        Exit Do
                    End If
                Else
                    Sleep (1)
                    Status = ReadAxisState_9030(0, VertUpDownAxis)
                    cur_pos = ReadAxisPos_9030(0, VertUpDownAxis)
                    If Status <> 1 Then     '等待VertUpDownAxis停止
                    'If cur_pos < 10 * Device_VertUpDownPulsPerMM Then   '等待VertUpDownAxis停止
                        Exit Do
                    End If
                End If
                DoEvents
            Loop
        End If
        IsRunning = False
    End If
End Sub

Sub TurnAngle(ByVal deg As Double)
'    ElevatorUp
    
    If Abs(deg) > 90 Then
    End If
    
    BeatAngle deg, True
'    ElevatorDown
End Sub

Sub PosMove_9030(ByVal pos As Long)
    Dim Ret As Integer
    Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
    Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedSpeed)
    Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
    Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
    SetAxisPos_9030 0, FeedAxis, pos
    StartAxis_9030 0, FeedAxis
End Sub

Sub WaitAxisRunComplete(ByVal axisNo As Integer)
    Dim state As Integer
    state = 1
    Do
        Sleep 5
        state = ReadAxisState_9030(0, axisNo)
        If state <> 1 Then
            Exit Do
        End If
        DoEvents
    Loop
End Sub
