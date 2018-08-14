Attribute VB_Name = "MdlDeviceV8"
Option Explicit

'====================================================================
Public IsRunning As Boolean
Public StopRunning As Boolean
Public PauseRunning As Boolean
Public IsFeeding As Boolean

'====================================================================
Public Const FeedAxis = 1
Public Const BendAxis = 2
Public Const VertAxis = 3
Public Const VertUpDownAxis = 4

Public Const VertOrgSensor = 22
Public Const VertLowSensor = 17
Public Const VertHighSensor = 23
Public Const KniftOrgSensor = 16
Public Const FeedOrgSensor = 5

Public Const VertMotorPort = 0
Public Const VertMoveUpPort = 1
Public Const VertMoveDownPort = 2
Public Const VertClosePort = 3
Public Const FeedFWPort = 4
Public Const FeedBWPort = 5
Public Const FeedFWPort2 = 6
Public Const FeedBWPort2 = 7

Public Const ElevatorUpSensor = 10
Public Const ElevatorDownSensor = 11

'Public Const ElevatorUpPort = 6
'Public Const ElevatorDownPort = 7

Public LastBendDir As Long '0-right, 1-left

Public PortBit(6) As Long
Public FeedPulsePerMM As Double

'====================================================================

Public Device_PulsePerMM As Double
Public Device_EncoderPulsePerMM As Double
Public Device_UseEncoder As Boolean

Public Device_PulsePerDegree As Double

Public Device_AdjustmentDegree As Double
Public Device_EmptyDegree As Double

'Public Device_AdjustmentDegree2 As Double
Public Device_EmptyDegree2 As Double

'Public Device_WaitUpTime As Double
'Public Device_WaitDownTime As Double

Public Device_VertMotorDrive As Boolean
Public Device_VertAllHigh As Boolean
Public Device_VertNoTurn As Boolean
    
Public Device_VertPulsePerMM As Double
Public Device_VertAdjustmentMM As Double
    
Public Device_VertUpDownPulsePerMM As Double
Public Device_VertUpDownAdjustmentMM As Double
    
Public Device_HeadDistance As Double
Public Device_DoneDistance As Double
Public Device_DoneWaitingTime As Double
Public Device_ExtendMM As Double

Public Device_CurMaterial As String
Public Device_MaterialName(10) As String

Public Device_FeedStartV As Double
Public Device_FeedSpeed As Double
Public Device_FeedAccel As Double
Public Device_FeedOffset As Double

Public Device_ManualFeedStartV As Double
Public Device_ManualFeedSpeed As Double
Public Device_ManualFeedAccel As Double
Public Device_ManualFeedOffset As Double

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

Public Device_TurnPointOffsetMM As Double
Public Device_VertKnifeDegree As Double

Public Device_VertMaxInnerAngle As Double
Public Device_VertMaxOuterAngle As Double

Public Device_InnerAngleAdjustMM As Double
Public Device_OuterAngleAdjustMM As Double

Public Device_InnerLineTerminalAdjustMM As Double
Public Device_OuterLineTerminalAdjustMM As Double

Public Device_BenderBacklash As Double
Public Device_BenderSpringback As Double

Public Device_FastSpeedMinLenMM As Double

Public Device_AmericanMaterial As Boolean
Public Device_TailVertAngle As Double
Public Device_VertUpDownMM_A As Double
Public Device_KareanMaterial As Boolean

Public Device_MaterialThickMM As Double

Public Device_StartPointAdjustMM As Double
Public Device_EndPointAdjustMM As Double

Public Device_VertMaxMM As Double
Public Device_VertStep As Long

Public VertThreadStep As Long
Public VertThreadAngle As Double
Public VertThreadTime As Double

Public Device_VertMotorZoneMM As Double
Public FeedIntoVertMotorZone As Boolean
Public VertUpToMiddleWay As Boolean
    
Sub GetDeviceParameters()
    Dim v As Double, i As Long
    
    Device_PulsePerMM = GetValueFromINI("Device", "PulsePerMM", "100", App.Path & "\Parameters.ini")
    Device_EncoderPulsePerMM = GetValueFromINI("Device", "EncoderPulsePerMM", "100", App.Path & "\Parameters.ini")
    
    If FeedByDCMotor = True Then
        Device_UseEncoder = True
    Else
        v = GetValueFromINI("Device", "UseEncoder", "0", App.Path & "\Parameters.ini")
        Device_UseEncoder = IIf(v = 0, False, True)
    End If
    
    Device_PulsePerDegree = GetValueFromINI("Device", "PulsePerDegree", "100", App.Path & "\Parameters.ini")
    
    Device_AdjustmentDegree = GetValueFromINI("Device", "AdjustmentDegree", "0", App.Path & "\Parameters.ini")
    Device_EmptyDegree = GetValueFromINI("Device", "EmptyDegree", "0", App.Path & "\Parameters.ini")
    
    'Device_AdjustmentDegree2 = GetValueFromINI("Device", "AdjustmentDegree2", "0", App.Path & "\Parameters.ini")
    Device_EmptyDegree2 = GetValueFromINI("Device", "EmptyDegree2", "0", App.Path & "\Parameters.ini")
    
    'Device_WaitUpTime = GetValueFromINI("Device", "WaitUpTime", "0", App.Path & "\Parameters.ini")
    'Device_WaitDownTime = GetValueFromINI("Device", "WaitDownTime", "0", App.Path & "\Parameters.ini")
    
    v = GetValueFromINI("Device", "VertMotorDrive", "0", App.Path & "\Parameters.ini")
    Device_VertMotorDrive = IIf(v = 0, False, True)
    v = GetValueFromINI("Device", "VertAllHigh", "0", App.Path & "\Parameters.ini")
    Device_VertAllHigh = IIf(v = 0, False, True)
    'v = GetValueFromINI("Device", "VertNoTurn", "0", App.Path & "\Parameters.ini")
    Device_VertNoTurn = True 'IIf(v = 0, False, True)
    
    Device_VertPulsePerMM = GetValueFromINI("Device", "VertPulsePerMM", "100", App.Path & "\Parameters.ini")
    Device_VertAdjustmentMM = GetValueFromINI("Device", "VertAdjustmentMM", "0", App.Path & "\Parameters.ini")
        
    Device_VertUpDownPulsePerMM = GetValueFromINI("Device", "VertUpDownPulsePerMM", "100", App.Path & "\Parameters.ini")
    Device_VertUpDownAdjustmentMM = GetValueFromINI("Device", "VertUpDownAdjustmentMM", "0", App.Path & "\Parameters.ini")
    Device_VertUpDownMM = GetValueFromINI("Device", "VertUpDownMM", "100", App.Path & "\Parameters.ini")
        
    Device_VertMaxMM = GetValueFromINI("Device", "VertMaxMM", "1", App.Path & "\Parameters.ini")
    Device_VertStep = GetValueFromINI("Device", "VertStep", "3", App.Path & "\Parameters.ini")
    
    Device_HeadDistance = GetValueFromINI("Device", "HeadDistance", "0", App.Path & "\Parameters.ini")
    Device_DoneDistance = GetValueFromINI("Device", "DoneDistance", "0", App.Path & "\Parameters.ini")
    Device_DoneWaitingTime = GetValueFromINI("Device", "DoneWaitingTime", "0", App.Path & "\Parameters.ini")
    Device_ExtendMM = GetValueFromINI("Device", "ExtendMM", "0", App.Path & "\Parameters.ini")
    
    Device_CurMaterial = GetStringFromINI("Device", "CurMaterial", "Material00", App.Path & "\Parameters.ini")
    For i = 1 To 10
        Device_MaterialName(i) = GetStringFromINI("MaterialName", str(i), "#" & Trim(str(i)), App.Path & "\Parameters.ini")
    Next
    
    Device_FeedStartV = GetValueFromINI("Device", "FeedStartV", "1000", App.Path & "\Parameters.ini")
    Device_FeedSpeed = GetValueFromINI("Device", "FeedSpeed", "2000", App.Path & "\Parameters.ini")
    Device_FeedAccel = GetValueFromINI("Device", "FeedAccel", "1000", App.Path & "\Parameters.ini")
    Device_FeedOffset = GetValueFromINI("Device", "FeedOffset", "0", App.Path & "\Parameters.ini")
    
    Device_ManualFeedStartV = GetValueFromINI("Device", "ManualFeedStartV", "1000", App.Path & "\Parameters.ini")
    Device_ManualFeedSpeed = GetValueFromINI("Device", "ManualFeedSpeed", "2000", App.Path & "\Parameters.ini")
    Device_ManualFeedAccel = GetValueFromINI("Device", "ManualFeedAccel", "1000", App.Path & "\Parameters.ini")
    Device_ManualFeedOffset = GetValueFromINI("Device", "ManualFeedOffset", "10", App.Path & "\Parameters.ini")
    
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
    Device_TurnPointOffsetMM = GetValueFromINI("Device", "TurnPointOffsetMM", "3", App.Path & "\Parameters.ini")
    
    Device_VertKnifeDegree = GetValueFromINI("Device", "VertKnifeDegree", "45", App.Path & "\Parameters.ini")
    
    Device_VertMaxInnerAngle = GetValueFromINI("Device", "VertMaxInnerAngle", "120", App.Path & "\Parameters.ini")
    Device_VertMaxOuterAngle = GetValueFromINI("Device", "VertMaxOuterAngle", "80", App.Path & "\Parameters.ini")
    
    Device_InnerAngleAdjustMM = GetValueFromINI("Device", "InnerAngleAdjustMM", "0", App.Path & "\Parameters.ini")
    Device_OuterAngleAdjustMM = GetValueFromINI("Device", "OuterAngleAdjustMM", "0", App.Path & "\Parameters.ini")
    
    Device_InnerLineTerminalAdjustMM = GetValueFromINI("Device", "InnerLineTerminalAdjustMM", "0", App.Path & "\Parameters.ini")
    Device_OuterLineTerminalAdjustMM = GetValueFromINI("Device", "OuterLineTerminalAdjustMM", "0", App.Path & "\Parameters.ini")
    
    Device_BenderBacklash = GetValueFromINI("Device", "BenderBacklash", "0", App.Path & "\Parameters.ini")
    Device_BenderSpringback = GetValueFromINI("Device", "BenderSpringback", "0.5", App.Path & "\Parameters.ini")
    
    Device_FastSpeedMinLenMM = GetValueFromINI("Device", "FastSpeedMinLenMM", "20", App.Path & "\Parameters.ini")
    Device_VertMotorZoneMM = GetValueFromINI("Device", "VertMotorZoneMM", "50", App.Path & "\Parameters.ini")
    
    v = GetValueFromINI("Device", "AmericanMaterial", "0", App.Path & "\Parameters.ini")
    Device_AmericanMaterial = IIf(v = 0, False, True)
    
    Device_TailVertAngle = GetValueFromINI("Device", "TailVertAngle", "130", App.Path & "\Parameters.ini")
    Device_VertUpDownMM_A = GetValueFromINI("Device", "VertUpDownMM_A", "100", App.Path & "\Parameters.ini")
    
    Device_MaterialThickMM = GetValueFromINI("MaterialThickMM", Device_CurMaterial, "0.8", App.Path & "\Parameters.ini")
    
    '----------------------------------------------------------------------
    FrmMain.ChkStartPointVert90.Visible = Not Device_AmericanMaterial
    FrmMain.ChkEndPointVert90.Visible = Not Device_AmericanMaterial
End Sub

Sub SetDeviceParameters()
    Dim i As Long
    
    WritePrivateProfileString "Device", "PulsePerMM", str(Device_PulsePerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EncoderPulsePerMM", str(Device_EncoderPulsePerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "UseEncoder", IIf(Device_UseEncoder = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "PulsePerDegree", str(Device_PulsePerDegree), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "AdjustmentDegree", str(Device_AdjustmentDegree), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EmptyDegree", str(Device_EmptyDegree), App.Path & "\Parameters.ini"
    
    'WritePrivateProfileString "Device", "AdjustmentDegree2", Str(Device_AdjustmentDegree2), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "EmptyDegree2", str(Device_EmptyDegree2), App.Path & "\Parameters.ini"
    
    'WritePrivateProfileString "Device", "WaitUpTime", Str(Device_WaitUpTime), App.Path & "\Parameters.ini"
    'WritePrivateProfileString "Device", "WaitDownTime", Str(Device_WaitDownTime), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertMotorDrive", IIf(Device_VertMotorDrive = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertAllHigh", IIf(Device_VertAllHigh = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertNoTurn", IIf(Device_VertNoTurn = True, "1", "0"), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertUpDownPulsePerMM", str(Device_VertUpDownPulsePerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownAdjustmentMM", str(Device_VertUpDownAdjustmentMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownMM", str(Device_VertUpDownMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertMaxMM", str(Device_VertMaxMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertStep", str(Device_VertStep), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertPulsePerMM", str(Device_VertPulsePerMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertAdjustmentMM", str(Device_VertAdjustmentMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "HeadDistance", str(Device_HeadDistance), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "DoneDistance", str(Device_DoneDistance), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "DoneWaitingTime", str(Device_DoneWaitingTime), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ExtendMM", str(Device_ExtendMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "CurMaterial", Device_CurMaterial, App.Path & "\Parameters.ini"
    For i = 1 To 10
        WritePrivateProfileString "MaterialName", str(i), Device_MaterialName(i), App.Path & "\Parameters.ini"
    Next
    
    WritePrivateProfileString "Device", "FeedStartV", str(Device_FeedStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "FeedSpeed", str(Device_FeedSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "FeedAccel", str(Device_FeedAccel), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "FeedOffset", str(Device_FeedOffset), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "ManualFeedStartV", str(Device_ManualFeedStartV), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ManualFeedSpeed", str(Device_ManualFeedSpeed), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ManualFeedAccel", str(Device_ManualFeedAccel), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "ManualFeedOffset", str(Device_ManualFeedOffset), App.Path & "\Parameters.ini"
    
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
    WritePrivateProfileString "Device", "TurnPointOffsetMM", str(Device_TurnPointOffsetMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertKnifeDegree", str(Device_VertKnifeDegree), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "VertMaxInnerAngle", str(Device_VertMaxInnerAngle), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertMaxOuterAngle", str(Device_VertMaxOuterAngle), App.Path & "\Parameters.ini"

    WritePrivateProfileString "Device", "InnerAngleAdjustMM", str(Device_InnerAngleAdjustMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "OuterAngleAdjustMM", str(Device_OuterAngleAdjustMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "InnerLineTerminalAdjustMM", str(Device_InnerLineTerminalAdjustMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "OuterLineTerminalAdjustMM", str(Device_OuterLineTerminalAdjustMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "BenderBacklash", str(Device_BenderBacklash), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "BenderSpringback", str(Device_BenderSpringback), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "FastSpeedMinLenMM", str(Device_FastSpeedMinLenMM), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertMotorZoneMM", str(Device_VertMotorZoneMM), App.Path & "\Parameters.ini"
    
    WritePrivateProfileString "Device", "AmericanMaterial", IIf(Device_AmericanMaterial = True, "1", "0"), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "TailVertAngle", str(Device_TailVertAngle), App.Path & "\Parameters.ini"
    WritePrivateProfileString "Device", "VertUpDownMM_A", str(Device_VertUpDownMM_A), App.Path & "\Parameters.ini"
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
    Dim Ret As Long, cur_pos As Long, puls As Long, status As Long
    Dim Feedpuls As Long
    
    Dim GapPos As Long
    
    IsRunning = True
    
    If deg > 0 Then
        GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsePerDegree
    Else
        GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsePerDegree
    End If
    
    Ret = set_startv(0, BendAxis, Device_BendStartV)
    Ret = set_speed(0, BendAxis, Device_BendSpeed)
    Ret = set_acc(0, BendAxis, Device_BendAccel)
        
    FrmMain.TmrBend.Enabled = True
    
    Ret = get_command_pos(0, BendAxis, cur_pos)
    puls = deg * Device_PulsePerDegree - cur_pos
    Ret = pmove(0, BendAxis, puls)
    Do
        get_status 0, BendAxis, status
        If status = 0 Then
            Wait 0.1
            Exit Do
        End If
        DoEvents
    Loop
    
    puls = deg * Device_PulsePerDegree - GapPos

    If called_by_turn = False Then
        Ret = pmove(0, BendAxis, -puls)
    Else
        Ret = set_startv(0, FeedAxis, Device_TurnFeedStartV)
        Ret = set_speed(0, FeedAxis, Device_TurnFeedSpeed)
        Ret = set_acc(0, FeedAxis, Device_TurnFeedAccel)
        
        Feedpuls = Device_TurnFeedMM * Device_PulsePerMM
        'ret = inp_move2(0, BendAxis, -puls, FeedAxis, Feedpuls)
        Ret = pmove(0, BendAxis, -puls)
        Ret = pmove(0, FeedAxis, Feedpuls)
    End If
    Do
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    IsRunning = False
End Sub

Sub BeatRealAngle(ByVal real_deg As Double, ByVal Dis As Double)
    Dim Ret As Long, deg As Double, p As Long, q As Long, status As Long

    Dim CurPos As Long
    Dim GapPos As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    If real_deg = 0 Then
        Exit Sub
    End If
    
    Do While StopRunning = False
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    
    If real_deg > 0 Then
        GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsePerDegree
    Else
        GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsePerDegree
    End If
    
    get_command_pos 0, BendAxis, CurPos
            
    '------拍弧-----------------------------------------------
    IsRunning = True
    
    Ret = set_startv(0, BendAxis, Device_BendStartV)
    Ret = set_speed(0, BendAxis, Device_BendSpeed)
    Ret = set_acc(0, BendAxis, Device_BendAccel)
        
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
    
    p = deg * Device_PulsePerDegree - CurPos
    pmove 0, BendAxis, p
    Do While StopRunning = False
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
        End If
        
        FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
    
    '------复位--------------------------------------------
    If StopRunning = True Then
        Exit Sub
    End If
    
    q = -deg * Device_PulsePerDegree + GapPos
    pmove 0, BendAxis, q
    Do While StopRunning = False
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
        End If
        
        FrmMain.TmrBend_Timer
        FrmMain.ShowFeedMarkPoint
        FrmMain.ShowVertMarkPoint
        DoEvents
    Loop
End Sub

Sub TurnRealAngle(ByVal real_deg As Double)
    Dim Ret As Long, deg As Double, p As Long, q As Long, status As Long

    Dim CurPos As Long
    Dim GapPos As Long
    
    Dim Feedpuls As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    If real_deg = 0 Then
        Exit Sub
    End If
    
    Do While StopRunning = False
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    
    If real_deg > 0 Then
        GapPos = (1 - Device_BenderSpringback) * Device_EmptyDegree2 * Device_PulsePerDegree
    Else
        GapPos = -(1 - Device_BenderSpringback) * Device_EmptyDegree * Device_PulsePerDegree
    End If
    
    get_command_pos 0, BendAxis, CurPos
            
    IsRunning = True
    
    '--- 折角 --------------------------------
    Ret = set_startv(0, BendAxis, Device_BendStartV)
    Ret = set_speed(0, BendAxis, Device_BendSpeed)
    Ret = set_acc(0, BendAxis, Device_BendAccel)
        
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
    
    p = deg * Device_PulsePerDegree - CurPos
    Ret = pmove(0, BendAxis, p)
'Debug.Print "p="; p, "ret="; ret
    Do While StopRunning = False
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
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
    
    Ret = set_startv(0, FeedAxis, Device_TurnFeedStartV)
    Ret = set_speed(0, FeedAxis, Device_TurnFeedSpeed)
    Ret = set_acc(0, FeedAxis, Device_TurnFeedAccel)
        
    q = -deg * Device_PulsePerDegree + GapPos
    Ret = pmove(0, BendAxis, q)
    
    Feedpuls = Device_TurnFeedMM * Device_PulsePerMM
    Ret = pmove(0, FeedAxis, Feedpuls)
    Do While StopRunning = False
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
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
    Dim Ret As Long, cur_pos As Long, puls As Long, status As Long, backlash_puls As Long, dir As Long
    Static dir0 As Long
        
    IsRunning = True
    
    If ByManual = True Then
        Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
        Ret = set_speed(0, BendAxis, Device_ManualBendSpeed)
        Ret = set_acc(0, BendAxis, Device_ManualBendAccel)
    Else
        Ret = set_startv(0, BendAxis, Device_BendStartV)
        Ret = set_speed(0, BendAxis, Device_BendSpeed)
        Ret = set_acc(0, BendAxis, Device_BendAccel)
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
            
    Ret = get_command_pos(0, BendAxis, cur_pos)
    
    puls = deg * Device_PulsePerDegree - cur_pos + backlash_puls
    Ret = pmove(0, BendAxis, puls)
    
    Do
        get_status 0, BendAxis, status
        If status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    IsRunning = False
    
    dir0 = dir
End Sub

Sub BendAngleByRadius(ByVal radius As Double, Optional check_done As Boolean = False)
    Dim Ret As Long, deg0 As Double, deg As Double, cur_pos As Long, puls As Long, status As Long
    
    IsRunning = True
    
sudden_stop 0, BendAxis
Do
    get_status 0, BendAxis, status
    If status = 0 Then
        Exit Do
    End If
    DoEvents
Loop
    
    Ret = set_startv(0, BendAxis, Device_BendStartV)
    Ret = set_speed(0, BendAxis, Device_BendSpeed)
    Ret = set_acc(0, BendAxis, Device_BendAccel)
        
    'TmrBend.Enabled = True
    FrmMain.TmrBend_Timer
    
    Ret = get_command_pos(0, BendAxis, cur_pos)
    
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
    
    puls = deg * Device_PulsePerDegree - cur_pos
    Ret = pmove(0, BendAxis, puls)
    
    If check_done = True Then
        Do
            get_status 0, BendAxis, status
            If status = 0 Then
                Exit Do
            End If
        
        '    TmrBend_Timer
        '    ShowFeedMarkPoint
        '    ShowVertMarkPoint
            DoEvents
        Loop
    End If
    IsRunning = False
End Sub

Sub BendReset()
    Dim Ret As Long, puls As Long, status As Long

    Ret = set_startv(0, BendAxis, Device_ResetBendStartV)
    Ret = set_speed(0, BendAxis, Device_ResetBendSpeed)
    Ret = set_acc(0, BendAxis, Device_ResetBendAccel)
    puls = -20 * Device_PulsePerDegree
    Ret = pmove(0, BendAxis, puls)
    Do
        get_status 0, BendAxis, status
        If status = 0 Then
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
    
    Ret = home1(0, BendAxis, 0, 0, -1, Device_ResetBendStartV, Device_ResetBendSpeed, Device_ResetBendAccel, 3 * Device_PulsePerDegree, Device_ResetBendSpeed / 8, 0, 360 * Device_PulsePerDegree)
    Do
        get_status 0, BendAxis, status
        If status = 0 Then
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
    
    Ret = set_startv(0, BendAxis, Device_ResetBendStartV)
    Ret = set_speed(0, BendAxis, Device_ResetBendSpeed)
    Ret = set_acc(0, BendAxis, Device_ResetBendAccel)
    puls = Device_AdjustmentDegree * Device_PulsePerDegree
    Ret = pmove(0, BendAxis, puls)
    Do
        get_status 0, BendAxis, status
        If status = 0 Then
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

Sub FeedV(Optional dr As Long = 1)
    Dim Ret As Long
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
    
    FrmMain.LblSpeedMode.BackColor = RGB(0, 255, 0)
        
    If FeedByDCMotor = True Then
        'Debug.Print "FeedV"
        DCMotorFeedFWOff2
        DCMotorFeedFWOn
    Else
        Ret = set_startv(0, FeedAxis, Device_FeedStartV / 5)
        Ret = set_speed(0, FeedAxis, Device_FeedStartV)
        Ret = set_acc(0, FeedAxis, Device_FeedAccel / 5)
            
        Ret = pmove(0, FeedAxis, dr * 10000000)
    End If
End Sub

Sub FeedV2(Optional dr As Long = 1)
    Dim Ret As Long
    
    IsRunning = True
    IsFeeding = True
    'StopRunning = False
        
    FrmMain.LblSpeedMode.BackColor = RGB(255, 0, 0)

    If FeedByDCMotor = True Then
        'Debug.Print "FeedV2"
        DCMotorFeedFWOff
        DCMotorFeedFWOn2
    Else
        Ret = set_startv(0, FeedAxis, Device_FeedStartV)
        Ret = set_speed(0, FeedAxis, Device_FeedSpeed)
        Ret = set_acc(0, FeedAxis, Device_FeedAccel)
            
        Ret = pmove(0, FeedAxis, dr * 10000000)
    End If
End Sub

Sub StopFeedV()
    FrmMain.LblSpeedMode.BackColor = RGB(240, 240, 240)
    IsFeeding = False

    If FeedByDCMotor = True Then
        'Debug.Print "StopFeedV"
        DCMotorFeedFWOff
        DCMotorFeedBWOff
        DCMotorFeedFWOff2
        DCMotorFeedBWOff2
    Else
        reset_fifo 0
        sudden_stop 0, FeedAxis
    End If
End Sub

Sub FeedMM(ByVal mm As Double, ByVal use_encoder As Boolean, ByVal by_manual As Boolean, ByVal wait_sec As Double, Optional ShowText As Boolean = True)
    Dim Ret As Long, status As Long, cur_pos As Long, feed_puls As Double, cur_feed_puls As Double, t0 As Double, t As Double
    Dim obj As Object, stk As Long
    
    Dim feed_startv As Long
    Dim feed_speed As Long
    Dim feed_accel As Long
    Dim feed_offset As Long
    
    Dim nLogPos As Long                   '逻辑位置
    Dim nActPos As Long                   '实际位置
    Dim nSpeed As Long                    '运行速度
    
    Dim nLogPos0 As Long
    
    'On Error Resume Next
    
    IsRunning = True
    StopRunning = False
        
    If by_manual = True Then
        feed_startv = Device_ManualFeedStartV
        feed_speed = Device_ManualFeedSpeed
        feed_accel = Device_ManualFeedAccel
        feed_offset = Device_ManualFeedOffset
    Else
        feed_startv = Device_FeedStartV
        feed_speed = Device_FeedSpeed
        feed_accel = Device_FeedAccel
        feed_offset = Device_FeedOffset
    End If
    
    Ret = set_startv(0, FeedAxis, feed_startv)
    Ret = set_speed(0, FeedAxis, feed_speed)
    Ret = set_acc(0, FeedAxis, feed_accel)
        
    get_command_pos 0, FeedAxis, nLogPos0
    If use_encoder = False Then
        feed_puls = mm * Device_PulsePerMM
        cur_pos = nLogPos0
        'FrmMain.ShowFeedPos
        Ret = pmove(0, FeedAxis, feed_puls)
    Else
        feed_puls = Abs(mm) * Device_EncoderPulsePerMM
        get_actual_pos 0, FeedAxis, cur_pos
        'FrmMain.ShowFeedPos
        Ret = pmove(0, FeedAxis, IIf(mm > 0, 10000000, -10000000))
    End If
    
    FrmMain.LblFeedStatus.BackColor = RGB(255, 0, 0)
    stk = 0
    t0 = Timer
    Do
        'get_status 0, FeedAxis, status
        'If status = 0 Then
        '    Exit Do
        'End If
        '
        'If StopRunning = True Then
        '    Exit Do
        'End If
        '
        'If use_encoder = True Then
        '    get_actual_pos 0, FeedAxis, nActPos
        '    cur_feed_puls = nActPos - cur_pos
        '
        '    If Abs(cur_feed_puls) >= feed_puls - feed_offset Then
        '        sudden_stop 0, FeedAxis
        '        Exit Do
        '    End If
        'End If
        
        'ShowFeedPos
        'DoEvents
        
    
        Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
        
        If stk Mod 1000 = 0 Then
        '    StatusBar1.Panels.Item(2).Text = "Pos:" + Str(nLogPos) + " /" + Str(Round(nLogPos / Device_PulsePerMM, 2)) + " mm"
        '    StatusBar1.Panels.Item(3).Text = "EncPos:" + Str(nActPos) + " /" + Str(Round(nActPos / Device_EncoderPulsePerMM, 2)) + " mm"
        '    StatusBar1.Panels.Item(4).Text = "Speed:" + Str(nSpeed)
        End If
        
        stk = stk + 1
        If nSpeed = 0 Then
            'Debug.Print "stk,stk/s="; stk; stk / (Timer - t0)
            cur_feed_puls = nLogPos - cur_pos
            'If FrmTestVisible = True And ShowText Then
            '    FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "采样次数：" + str(stk) + " 当前脉冲：" + str(cur_feed_puls) + vbCrLf
            'End If
            cur_pos = nLogPos
            Exit Do
        End If
        
        If use_encoder = True Then
            cur_feed_puls = nActPos - cur_pos
        
            If Abs(cur_feed_puls) >= feed_puls - feed_offset Then
                'sudden_stop 0, FeedAxis
                dec_stop 0, FeedAxis
                
                'Debug.Print "采样次数："; stk; " 编码器理论脉冲："; feed_puls; " 脉冲提前量："; feed_offset; " 当前脉冲："; cur_feed_puls; " 误差脉冲："; Abs(cur_feed_puls) - (feed_puls - feed_offset)
                'If FrmTestVisible = True And ShowText Then
                '    FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "采样次数：" + str(stk) + " 编码器理论脉冲：" + str(feed_puls) + " 脉冲提前量：" + str(feed_offset) + " 当前脉冲：" + str(cur_feed_puls) + " 误差脉冲：" + str(Round(Abs(cur_feed_puls) - (feed_puls - feed_offset), 3)) + " 当前电机脉冲：" + str(nLogPos) + "(" + str(nLogPos - nLogPos0) + ")" + vbCrLf
                'End If
                cur_pos = nActPos
                nLogPos0 = nLogPos
                Exit Do
            End If
        End If
        
        DoEvents
    Loop
    FrmMain.LblFeedStatus.BackColor = RGB(220, 220, 220)
    
    t0 = Timer
    Do
        t = Timer
        If TimeDiff(t, t0) > wait_sec Then
            Exit Do
        End If
        
        'FrmMain.ShowFeedPos
        DoEvents
    Loop
    
    Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
    'If use_encoder = False Then
    '    Debug.Print "停止所用脉冲："; nLogPos - cur_pos
    '    If FrmTestVisible = True And ShowText Then
    '        FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "停止所用脉冲：" + str(nLogPos - cur_pos) + " 刹车距离：" + str(str(Round((nLogPos - cur_pos) / Device_EncoderPulsePerMM, 4))) + " mm" + vbCrLf
    '    End If
    'Else
    '    Debug.Print "停止所用脉冲："; nActPos - cur_pos; " 刹车距离："; str(Round((nActPos - cur_pos) / Device_EncoderPulsePerMM, 2)) + " mm"
    '    If FrmTestVisible = True And ShowText Then
    '        FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "停止所用脉冲：" + str(nActPos - cur_pos) + " 刹车距离：" + str(str(Round((nActPos - cur_pos) / Device_EncoderPulsePerMM, 4))) + " mm" + " 当前电机脉冲：" + str(nLogPos) + "(" + str(nLogPos - nLogPos0) + ")" + vbCrLf
    '    End If
    'End If
    
    IsRunning = False
End Sub

Sub FeedMMByDCMotor(ByVal mm As Double, ByVal wait_sec As Double, Optional ShowText As Boolean = True)
    Dim Ret As Long, status As Long, cur_pos As Long, cur_pos0 As Long, feed_puls As Double, cur_feed_puls As Double, t0 As Double, t As Double
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
        
    feed_puls = Abs(mm) * Device_EncoderPulsePerMM
    feed_puls_before_change_speed = feed_puls - Device_FastSpeedMinLenMM * Device_EncoderPulsePerMM
    get_actual_pos 0, FeedAxis, cur_pos
    cur_pos0 = cur_pos
    
    'FrmMain.ShowFeedPos
    If mm > Device_FastSpeedMinLenMM Then
        DCMotorFeedBWOff2
        DCMotorFeedFWOn2
    ElseIf mm > 0 Then
        DCMotorFeedBWOff
        DCMotorFeedFWOn
    ElseIf mm < -Device_FastSpeedMinLenMM Then
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
                If mm > 0 Then
                    DCMotorFeedFWOff2
                    DCMotorFeedFWOn
                Else
                    DCMotorFeedBWOff2
                    DCMotorFeedBWOn
                End If
            End If
        End If
        
        If Abs(cur_feed_puls) >= feed_puls - feed_offset Then
            If mm > 0 Then
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
    FrmMain.TxtStatistics.Text = "停止所需脉冲：" + vbCrLf + str(nActPos - cur_pos) + vbCrLf + vbCrLf + "刹车距离：" + vbCrLf + str(Round((nActPos - cur_pos) / Device_EncoderPulsePerMM, 2)) + " mm" + vbCrLf + vbCrLf + "运行总脉冲：" + vbCrLf + str(nActPos - cur_pos0) + vbCrLf + vbCrLf + "进料总距离：" + vbCrLf + str(Round((nActPos - cur_pos0) / Device_EncoderPulsePerMM, 2)) + " mm" + vbCrLf + vbCrLf
    'If FrmTestVisible = True And ShowText Then
    '    FrmTest.TxtTest.Text = FrmTest.TxtTest.Text + "停止所用脉冲：" + str(nActPos - cur_pos) + " 刹车距离：" + str(str(Round((nActPos - cur_pos) / Device_EncoderPulsePerMM, 4))) + " mm" + " 当前电机脉冲：" + str(nLogPos) + "(" + str(nLogPos - nLogPos0) + ")" + vbCrLf
    'End If
    
    IsRunning = False
End Sub

Public Sub Vert(ByVal low_or_high As Integer, ByVal motor As Integer)
    Dim Sensor As Long
    Dim Ret As Long, cur_puls As Long, puls As Long, status As Long
    
    If motor = 1 Then
        PortBit(1) = 1
        PortBit(2) = 1
        write_bit 0, VertMotorPort, 1      '铣刀旋转马达开
        'write_bit 0, VertClosePort, 1          '铣刀靠紧
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
        
        puls = Device_VertAdjustmentMM * Device_VertPulsePerMM - cur_puls
            
        Ret = pmove(0, VertAxis, puls)
        Do
            get_status 0, VertAxis, status
            If status = 0 Then
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
            get_status 0, VertAxis, status
            If status = 0 Then
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
        write_bit 0, VertMotorPort, 0      '铣刀旋转马达关
    End If
    
    Wait 2
    
End Sub

Sub VertUp(ByVal low_or_high As Integer, ByVal motor As Integer)
    Dim Sensor As Long
    Dim Ret As Long, cur_puls As Long, puls As Long, status As Long
    
    If motor = 1 Then
        write_bit 0, VertMotorPort, 1      '铣刀旋转马达开
        write_bit 0, VertClosePort, 1          '铣刀靠紧
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
        
            puls = Device_VertAdjustmentMM * Device_VertPulsePerMM - cur_puls
            
        Ret = pmove(0, VertAxis, puls)
        Do
            get_status 0, VertAxis, status
            If status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        
        'If low_or_high = 0 Then
        '    puls = -Device_VertLowMM * Device_VertPulsePerMM
        'Else
        '    puls = -Device_VertAdjustmentMM * Device_VertPulsePerMM
        'End If
        '
        'ret = pmove(0, VertAxis, puls)
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
        write_bit 0, VertMotorPort, 0      '铣刀旋转马达关
        write_bit 0, VertClosePort, 0      '铣刀离开
    End If
End Sub

Sub VertReset()
     
    FrmMain.CmdCompressOff_Click
    'CmdVertUp_Click
    'CmdVertDown_Click
    'Exit Sub
    
    
    Dim Sensor As Long
    Dim Ret As Long, puls As Long, status As Long
        
    If VertUpDownByDCMotor = True Then
        write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
        Wait 0.1
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
    Else
        Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed)
        Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
        puls = 5 * Device_VertUpDownPulsePerMM
        Ret = pmove(0, VertUpDownAxis, puls)
        Do
            get_status 0, VertUpDownAxis, status
            If status = 0 Then
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
        
        Ret = home1(0, VertUpDownAxis, 1, 0, -1, Device_VertUpDownStartV, Device_VertUpDownSpeed, Device_VertUpDownAccel, 1 * Device_VertUpDownPulsePerMM, Device_VertUpDownSpeed / 5, 0, 10 * Device_VertUpDownPulsePerMM)
        Do
            get_status 0, VertUpDownAxis, status
            If status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        
        Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed)
        Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
        puls = -Device_VertUpDownAdjustmentMM * Device_VertUpDownPulsePerMM
        Ret = pmove(0, VertUpDownAxis, puls)
        Do
            get_status 0, VertUpDownAxis, status
            If status = 0 Then
                Exit Do
            End If
            
            If StopRunning = True Then
                Exit Sub
            End If
            DoEvents
        Loop
        
        set_command_pos 0, VertUpDownAxis, 0
        set_actual_pos 0, VertUpDownAxis, 0
    End If
    
    '================================================================================================================================
    Ret = set_startv(0, VertAxis, Device_ResetVertStartV)
    Ret = set_speed(0, VertAxis, Device_ResetVertSpeed)
    Ret = set_acc(0, VertAxis, Device_ResetVertAccel)
    puls = 0.2 * Device_VertPulsePerMM
    Ret = pmove(0, VertAxis, puls)
    Do
        get_status 0, VertAxis, status
        If status = 0 Then
            Exit Do
        End If '

        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    If StopRunning = True Then
        Exit Sub
    End If
    
    'Ret = home1(0, VertAxis, 1, 0, -1, Device_VertStartV, Device_VertSpeed, Device_VertAccel, 5 * Device_VertPulsePerMM, Device_VertSpeed / 5, 0, 100 * Device_VertPulsePerMM)
    Ret = home1(0, VertAxis, 0, 0, -1, Device_ResetVertStartV, Device_ResetVertSpeed, Device_ResetVertAccel, 0.1 * Device_VertPulsePerMM, Device_ResetVertSpeed / 5, 0, 0)
    Do
        get_status 0, VertAxis, status
        If status = 0 Then
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
    puls = -Device_VertAdjustmentMM * Device_VertPulsePerMM
    Ret = pmove(0, VertAxis, puls)
    Do
        get_status 0, VertAxis, status
        If status = 0 Then
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

Sub GetPathXYByFeedpuls(ByVal cur_puls As Long, ByRef ux As Double, ByRef uy As Double)
    Dim i As Long, j As Long, total_puls As Long, d As Double, d1 As Double, d2 As Double, ds As Double
    Static start_id As Long
    
    cur_puls = cur_puls - Device_HeadDistance * FeedPulsePerMM
    If cur_puls <= 0 Then
        ux = -99999
        uy = -99999
        start_id = 1
        Exit Sub
    End If
    
    total_puls = TotalPathOutLength * FeedPulsePerMM
    d = 1# * cur_puls / total_puls
    
    For i = start_id To PathOutputPointCount - 1
        If PathOutputPoint(i).VertType <= 0 Then
            For j = i + 1 To PathOutputPointCount
                If PathOutputPoint(j).VertType <= 0 Then
                    Exit For
                End If
            Next
            
            d1 = (PathOutputPoint(i).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            d2 = (PathOutputPoint(j).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            If d >= d1 And d <= d2 Then
                ds = (d - d1) / (d2 - d1)
                ux = PathOutputPoint(i).ux + ds * (PathOutputPoint(j).ux - PathOutputPoint(i).ux)
                uy = PathOutputPoint(i).uy + ds * (PathOutputPoint(j).uy - PathOutputPoint(i).uy)
                start_id = i
                Exit Sub
            End If
        End If
    Next
    ux = -99999
    uy = -99999
End Sub

Sub GetPathXYByVertpuls(ByVal cur_puls As Long, ByRef ux As Double, ByRef uy As Double)
    Dim i As Long, j As Long, total_puls As Long, d As Double, d1 As Double, d2 As Double, ds As Double
    Static start_id As Long
    
    cur_puls = cur_puls - Device_HeadDistance * FeedPulsePerMM
    If cur_puls <= 0 Then
        ux = -99999
        uy = -99999
        start_id = 1
        Exit Sub
    End If
    
    total_puls = TotalPathOutLength * FeedPulsePerMM
    d = 1# * cur_puls / total_puls
    
    For i = start_id To PathOutputPointCount - 1
    'For I = 1 To PathOutputPointCount - 1
        If PathOutputPoint(i).VertType <= 0 Then
            For j = i + 1 To PathOutputPointCount
                If PathOutputPoint(j).VertType <= 0 Then
                    Exit For
                End If
            Next
            
            d1 = (PathOutputPoint(i).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            d2 = (PathOutputPoint(j).LengthFromStart - Device_HeadDistance) / TotalPathOutLength
            If d >= d1 And d <= d2 Then
                ds = (d - d1) / (d2 - d1)
                ux = PathOutputPoint(i).ux + ds * (PathOutputPoint(j).ux - PathOutputPoint(i).ux)
                uy = PathOutputPoint(i).uy + ds * (PathOutputPoint(j).uy - PathOutputPoint(i).uy)
                start_id = i
                Exit Sub
            End If
        End If
    Next
    ux = -99999
    uy = -99999
End Sub

Sub VertAngle(ByVal deg As Double, Optional check_done As Boolean = True)
    Dim Ret As Long, cur_pos As Long, puls As Long, status As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
        
    Do While PauseRunning = True
        DoEvents
    Loop
    
    IsRunning = True
    
    Ret = set_startv(0, VertAxis, Device_VertStartV)
    Ret = set_speed(0, VertAxis, Device_VertSpeed)
    Ret = set_acc(0, VertAxis, Device_VertAccel)
        
    'TmrBend.Enabled = True
    
    Ret = get_command_pos(0, VertAxis, cur_pos)
    
    puls = deg * Device_VertPulsePerMM - cur_pos
    Ret = pmove(0, VertAxis, puls)
    
    If check_done = True Then
        Do
            If StopRunning = True Then
                Exit Do
            End If
            
            Do While PauseRunning = True
                DoEvents
            Loop
        
            get_status 0, VertAxis, status
            If status = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
    End If
    IsRunning = False
End Sub

Sub VertOuterAngle(ByVal deg As Double)
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
    
    VertMoveDown 'Up
    VertAngle -180 - (deg - Device_VertKnifeDegree) / 2
    
    write_bit 0, VertMotorPort, 1         '铣刀旋转
    Wait 2
    
    VertMoveUp 'Down
    VertAngle -180 + (deg - Device_VertKnifeDegree) / 2
    VertMoveDown 'Up
    
    write_bit 0, VertMotorPort, 0         '铣刀停转
    Wait 1
End Sub


Sub VertMoveBack()
    Dim Ret As Long, cur_pos As Long, status As Long
    
    IsRunning = True
    
    Ret = set_startv(0, VertAxis, Device_VertStartV)
    Ret = set_speed(0, VertAxis, Device_VertSpeed)
    Ret = set_acc(0, VertAxis, Device_VertAccel)
        
    Ret = get_command_pos(0, VertAxis, cur_pos)
    
    Ret = pmove(0, VertAxis, -cur_pos)
    Do
        If StopRunning = True Then
            Exit Do
        End If
    
        get_status 0, VertAxis, status
        If status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    IsRunning = False
End Sub

Sub VertMoveFW(ByVal mm As Double)
    Dim Ret As Long, cur_pos As Long, pulse As Long, status As Long
    
    IsRunning = True
    
    Ret = set_startv(0, VertAxis, Device_VertStartV)
    Ret = set_speed(0, VertAxis, Device_VertSpeed)
    Ret = set_acc(0, VertAxis, Device_VertAccel)
        
    Ret = get_command_pos(0, VertAxis, cur_pos)
cur_pos = 0

    pulse = -(mm * Device_VertPulsePerMM - cur_pos)
    
    Ret = pmove(0, VertAxis, pulse)
    Do
        If StopRunning = True Then
            Exit Do
        End If
    
        get_status 0, VertAxis, status
        If status = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    IsRunning = False
End Sub

Sub VertMoveDown(Optional check_done As Boolean = True)
    Dim t As Double, t0 As Double, b As Long
    Dim Ret As Long, cur_pos As Long, puls As Long, status As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If VertUpDownByDCMotor = True Then
        write_bit 0, VertMoveDownPort, 0
        Wait 0.1
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
        IsRunning = True
        
        'Debug.Print "Move Up!"

        Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = set_speed(0, VertUpDownAxis, Device_VertUpDownSpeed)
        Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
            
        Ret = get_command_pos(0, VertUpDownAxis, cur_pos)
        
        'puls = -cur_pos
        If Device_AmericanMaterial = False Or VertUpToMiddleWay = False Then
            puls = (Device_VertUpDownMM * Device_VertUpDownPulsePerMM - cur_pos)
        Else
            puls = (Device_VertUpDownMM_A * Device_VertUpDownPulsePerMM - cur_pos)
        End If
        
        Ret = pmove(0, VertUpDownAxis, puls)
        
        If check_done = True Then
            Do
                If StopRunning = True Then
                    Exit Do
                End If
            
                get_status 0, VertUpDownAxis, status
                If status = 0 Then
                    Exit Do
                End If
                DoEvents
            Loop
        End If
        IsRunning = False
    End If
End Sub

Sub VertMoveUp(Optional check_done As Boolean = True)
    Dim t As Double, t0 As Double, b As Long
    Dim Ret As Long, cur_pos As Long, puls As Long, status As Long
    
    If StopRunning = True Then
        Exit Sub
    End If
    
    Do While PauseRunning = True
        DoEvents
    Loop
    
    If VertUpDownByDCMotor = True Then
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
        IsRunning = True
        
        Ret = set_startv(0, VertUpDownAxis, Device_VertUpDownStartV)
        Ret = set_speed(0, VertUpDownAxis, 3 * Device_VertUpDownSpeed)
        Ret = set_acc(0, VertUpDownAxis, Device_VertUpDownAccel)
            
        Ret = get_command_pos(0, VertUpDownAxis, cur_pos)
        
        'puls = Device_VertUpDownMM * Device_VertUpDownPulsePerMM - cur_pos
        puls = -cur_pos
        Ret = pmove(0, VertUpDownAxis, puls)
        
        If check_done = True Then
            Do
                If StopRunning = True Then
                    Exit Do
                End If
            
                get_status 0, VertUpDownAxis, status
                If status = 0 Then
                    Exit Do
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


Sub MillSlot()
    Dim i As Long, d As Double
    
    StopRunning = False
    
    CompressOn
    
    d = Device_VertMaxMM / Device_VertStep
    
    VertMoveBack
    VertMoveUp
    
    For i = 1 To Device_VertStep
        VertMoveFW d
        VertMoveDown
        VertMoveUp
    Next
    VertMoveBack
    
    CompressOff
End Sub

Sub CompressOn()
    write_bit 0, VertMotorPort, 1
    Wait 1
End Sub

Sub CompressOff()
    write_bit 0, VertMotorPort, 0
End Sub
