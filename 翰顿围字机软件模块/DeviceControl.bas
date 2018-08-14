Attribute VB_Name = "DeviceControl"
Option Explicit

Const MaxAxisNum = 4
Public AxisCount As Integer

Public Enum ModeType
    PointsAndLines = 0
    PointsOnly = 1
End Enum

    
'----------------------------------------------------
Public axis As Integer
Public OutputStatus(15) As Byte

'Public ResetKnifeHolder As Boolean
Public StopFeed As Boolean
'Public StopRun As Boolean

'Public MeasuringGap As Boolean

Public Const MaxBendDisNo = 5
Public BendDis(MaxBendDisNo) As Double
Public SupplementKeyCount As Long
Public KeyAngle(100) As Double
Public RealAngle(MaxBendDisNo, 100) As Double

'----------------------------------------------------

Public Device_DirX As Integer
Public Device_DirY As Integer
Public Device_DirZ As Integer

Public Device_Mode As ModeType
Public Device_CoordinateMode As Integer

Public Device_UserSize(MaxAxisNum) As Double
Public Device_PulsePerDM(MaxAxisNum) As Double
Public Device_MotorDir(MaxAxisNum) As Boolean
Public Device_Encoder(MaxAxisNum) As Boolean

'Public Device_ConSpeed(MaxAxisNum) As Double
'Public Device_HighSpeed(MaxAxisNum) As Double
'Public Device_Accel(MaxAxisNum) As Double

Public Device_CurUserPos(MaxAxisNum) As Double
Public Device_CurPulsePos(MaxAxisNum) As Long
Public Device_CurSpeed(MaxAxisNum) As Long
Public Device_ReadSpeed(MaxAxisNum) As Long
Public Device_CurAxisNum As Long

Dim CardNum As Long

Public Const WriteDebugFile As Boolean = False
Public DebugFileNo As Long
Public DebugFileName As String
Public DebugCounter(1) As Long

Public Ratio(4) As Double

Public ResetDone As Boolean
Public KnifeSensorGetSignal As Boolean

Public Sub Device_SetMode()
    WritePrivateProfileString "Device", "Mode", str(Device_Mode), App.Path & "\" & App.EXEName & ".ini"
    WritePrivateProfileString "Device", "CoordinateMode", str(Device_CoordinateMode), App.Path & "\" & App.EXEName & ".ini"
    FrmMain.Show
End Sub

'******************************************
'基础参数
Public Sub Device_SetBasicParam()
    Dim I As Integer
        
    For I = 1 To MaxAxisNum
        WritePrivateProfileString "UserSize", str(I), str(Device_UserSize(I)), App.Path & "\Parameters.ini"
    Next
End Sub

Public Sub SaveColorMode()
    WritePrivateProfileString "Platform", "Color", str(ColorMode), App.Path & "\Parameters.ini"
End Sub

Public Sub LoadParameters()
    Dim I As Long, t As Long
    
    For I = 1 To SupplementKeyCount
        KeyAngle(I) = 0
        For t = 1 To MaxBendDisNo
            RealAngle(t, I) = 0
        Next
    Next
    
    Device_CurMaterial = GetStringFromINI("Device", "CurMaterial", "Material00", App.Path & "\Parameters.ini")
        
    For t = 1 To MaxBendDisNo
        BendDis(t) = GetVFromINI_A("Gap" & Trim(str(t)))
    Next
    
    SupplementKeyCount = GetVFromINI_A("SupplementKeyCount")
    For I = 1 To SupplementKeyCount
        KeyAngle(I) = GetVFromINI_A("Key" & Trim(str(I)))
        For t = 1 To MaxBendDisNo
            RealAngle(t, I) = GetVFromINI_A("Real" & Trim(str(I)) & "_" & Trim(str(t)))
            'SupAngle(t, i) = KeyAngle(i) - RealAngle(t, i)
        Next
    Next
    
    TotalWorkLength = GetVFromINI("TotalWorkLength")
    TotalWorkBendCount = GetVFromINI("TotalWorkBendCount")
    TotalWorkCount = GetVFromINI("TotalWorkCount")
    TotalWorkTime = GetVFromINI("TotalWorkTime")
End Sub

Public Sub Device_GetBasicParam()
    Dim I As Integer
    
    'Device_Controller = GetValueFromINI("Device", "Controller", Controller, App.Path & "\" & App.EXEName & ".ini")

    Device_Mode = GetValueFromINI("Device", "Mode", "0", App.Path & "\" & App.EXEName & ".ini")
    Device_CoordinateMode = GetValueFromINI("Device", "CoordinateMode", "0", App.Path & "\" & App.EXEName & ".ini")
        
    If Device_CoordinateMode = 0 Then
        FrmMain.MnuOrgLeftDown.Checked = True
        FrmMain.MnuOrgLeftUp.Checked = False
    Else
        FrmMain.MnuOrgLeftDown.Checked = False
        FrmMain.MnuOrgLeftUp.Checked = True
    End If
    
    ToolMax = Tool_Max - Device_Mode * 10
    LayerMax = Layer_Max - Device_Mode * 2
    If Device_Mode = 0 Then
        FrmMain.MnuPointAndLine.Checked = True
        FrmMain.MnuPointOnly.Checked = False
    Else
        FrmMain.MnuPointAndLine.Checked = False
        FrmMain.MnuPointOnly.Checked = True
    End If
    
    For I = 1 To MaxAxisNum
        Device_UserSize(I) = GetValueFromINI("UserSize", str(I), "1000", App.Path & "\Parameters.ini")
    Next
    
    LoadParameters
    
    ColorMode = GetValueFromINI("Platform", "Color", "1", App.Path & "\Parameters.ini")
    If ColorMode = 0 Then
        FrmMain.MnuWhite.Checked = True
        FrmMain.MnuBlack.Checked = False
    Else
        FrmMain.MnuWhite.Checked = False
        FrmMain.MnuBlack.Checked = True
    End If
    
End Sub

Public Sub ConvertUserTopuls(ByVal ux As Double, ByVal uy As Double, ByVal uz As Double, px As Long, py As Long, pz As Long)
    px = Device_DirX * Int(ux * Device_PulsePerDM(1) / 100)
    py = Device_DirY * Int(uy * Device_PulsePerDM(2) / 100)
    pz = Device_DirZ * Int(uz * Device_PulsePerDM(3) / 100)
End Sub

Public Sub ConvertpulsToUser(ByVal px As Long, ByVal py As Long, ByVal pz As Long, ux As Double, uy As Double, uz As Double)
    ux = Device_DirX * px * 100 / Device_PulsePerDM(1)
    uy = Device_DirY * py * 100 / Device_PulsePerDM(2)
    uz = Device_DirZ * pz * 100 / Device_PulsePerDM(3)
End Sub

Public Sub ConvertpulsToPath(ByVal px As Long, ByVal py As Long, X As Single, Y As Single)
    Dim ux As Double, uy As Double, uz As Double
    
    ux = Device_DirX * px * 100 / Device_PulsePerDM(1)
    uy = Device_DirY * py * 100 / Device_PulsePerDM(2)
    
    ConvertUserToPath ux, uy, X, Y
End Sub

Function GetBeatAngleByRealAngle(ByVal Ang As Double, ByVal Dis As Double) As Double
    Dim t As Long, I As Long, i0 As Long, t0 As Long, t1 As Long, k0 As Double, k1 As Double
    Dim BendStartCol As Long, BendEndCol As Long
    
    'Ang = Abs(Ang)
    Dis = Abs(Dis)
    If Dis = 0 Then Dis = 10000
    
    If Ang < 0 Then
        BendStartCol = 3
        BendEndCol = 3
    Else
        BendStartCol = 4
        BendEndCol = 4
    End If
    
    Ang = Abs(Ang)
    
    t0 = BendStartCol '从第t0列开始
    t1 = t0
    For t = BendStartCol To BendEndCol 'MaxBendDisNo
        If BendDis(t) = 0 Then
            If t > BendStartCol Then
                t0 = t - 1
                t1 = t0
            End If
            Exit For
        End If
            
        If BendDis(t) >= Abs(Dis) Then
            If t > BendStartCol Then
                t0 = t - 1
            End If
            t1 = t
            Exit For
        End If
    Next
    If t > BendEndCol Then 'MaxBendDisNo Then
        t0 = BendEndCol 'MaxBendDisNo
        t1 = t0
    End If
            
    i0 = 1
    For I = 1 To SupplementKeyCount
        If RealAngle(t0, I) > 0 Then
            If Ang <= RealAngle(t0, I) Then
                Exit For
            Else
                i0 = I
            End If
        End If
    Next
    If I <= SupplementKeyCount Then
        If RealAngle(t0, i0) = 0 Then
            k0 = KeyAngle(I)
        Else
            k0 = KeyAngle(i0) + (KeyAngle(I) - KeyAngle(i0)) * (Ang - RealAngle(t0, i0)) / (RealAngle(t0, I) - RealAngle(t0, i0))
        End If
    Else
        k0 = KeyAngle(i0)
    End If

'    If BendDis(t1) <> BendDis(t0) Then 'including t1<>t0
'        i0 = 1
'        For I = 1 To SupplementKeyCount
'            If RealAngle(t1, I) > 0 Then
'                If Ang <= RealAngle(t1, I) Then
'                    Exit For
'                Else
'                    i0 = I
'                End If
'            End If
'        Next
'        If I <= SupplementKeyCount Then
'            If RealAngle(t1, i0) = 0 Then
'                k1 = KeyAngle(I)
'            Else
'                k1 = KeyAngle(i0) + (KeyAngle(I) - KeyAngle(i0)) * (Ang - RealAngle(t1, i0)) / (RealAngle(t1, I) - RealAngle(t1, i0))
'            End If
'        Else
'            k1 = KeyAngle(i0)
'        End If
'
'        GetBeatAngleByRealAngle = k0 + (k1 - k0) * (Dis - BendDis(t0)) / (BendDis(t1) - BendDis(t0))
'    Else
        GetBeatAngleByRealAngle = k0
'    End If
End Function

Function GetBendAngleByRadius(ByVal radius As Double) As Double
    Dim I As Long, i0 As Long, c As Long
        
    'radius = Abs(radius)
    
    If radius < 0 Then
        '取自第1列，此处RealAngle为半径
        c = 1
        radius = -radius
    Else
        '取自第2列，此处RealAngle为半径
        c = 2
    End If
    
    '取自第1列，此处RealAngle为半径
    
    i0 = 1
    For I = 1 To SupplementKeyCount
        If RealAngle(c, I) > 0 Then
            If radius >= RealAngle(c, I) Then
                Exit For
            Else
                i0 = I
            End If
        End If
    Next
    
    If i0 = I Then
        GetBendAngleByRadius = I
        Exit Function
    End If
    
    If I <= SupplementKeyCount Then
        If RealAngle(c, i0) = 0 Then
            GetBendAngleByRadius = KeyAngle(I)
        Else
            GetBendAngleByRadius = KeyAngle(i0) + (KeyAngle(I) - KeyAngle(i0)) * (radius - RealAngle(c, i0)) / (RealAngle(c, I) - RealAngle(c, i0))
        End If
    Else
        GetBendAngleByRadius = KeyAngle(i0)
    End If
End Function

Function GetTurnAngleByRealAngle(ByVal deg As Double) As Double
    Dim I As Long, i0 As Long
        
    deg = Abs(deg)
    
    '取自第3列
    
    i0 = 1
    For I = 1 To SupplementKeyCount
        If RealAngle(3, I) > 0 Then
            If deg <= RealAngle(3, I) Then
                Exit For
            Else
                i0 = I
            End If
        End If
    Next
    If I <= SupplementKeyCount Then
        If RealAngle(3, i0) = 0 Then
            GetTurnAngleByRealAngle = KeyAngle(I)
        Else
            GetTurnAngleByRealAngle = KeyAngle(i0) + (KeyAngle(I) - KeyAngle(i0)) * (deg - RealAngle(3, i0)) / (RealAngle(3, I) - RealAngle(3, i0))
        End If
    Else
        GetTurnAngleByRealAngle = KeyAngle(i0)
    End If
End Function


