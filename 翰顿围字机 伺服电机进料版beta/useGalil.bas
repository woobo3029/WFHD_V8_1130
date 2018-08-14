Attribute VB_Name = "useGalil"
'设置加速度
Public Function SetAcc(ByVal hDmc As Long, ByVal axis As Integer, ByVal acc As Long)
Dim rc As Long
Dim Response As String * 256
Dim Cmd As String
    If axis = 0 Then
        Cmd = "AC" + str(acc)
    ElseIf axis = 1 Then
        Cmd = "AC," + str(acc)
    ElseIf axis = 2 Then
        Cmd = "AC,," + str(acc)
    ElseIf axis = 3 Then
        Cmd = "AC,,," + str(acc)
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'设置减速度
Public Function SetDec(ByVal hDmc As Long, ByVal axis As Integer, ByVal dec As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "DCX=" + str(dec)
    ElseIf axis = 1 Then
        Cmd = "DCY=" + str(dec)
    ElseIf axis = 2 Then
        Cmd = "DCZ=" + str(dec)
    ElseIf axis = 3 Then
        Cmd = "DCW=" + str(dec)
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'设置速度
Public Function SetVel(ByVal hDmc As Long, ByVal axis As Integer, ByVal velocity As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "SPX=" + str(velocity)
    ElseIf axis = 1 Then
        Cmd = "SPY=" + str(velocity)
    ElseIf axis = 2 Then
        Cmd = "SPZ=" + str(velocity)
    ElseIf axis = 3 Then
        Cmd = "SPW=" + str(velocity)
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'相对位置运动
Public Function PosMoveRel(ByVal hDmc As Long, ByVal axis As Integer, ByVal pos As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "PRX=" + str(pos) + ";BGX"
    ElseIf axis = 1 Then
        Cmd = "PRY=" + str(pos) + ";BGY"
    ElseIf axis = 2 Then
        Cmd = "PRZ=" + str(pos) + ";BGZ"
    ElseIf axis = 3 Then
        Cmd = "PRW=" + str(pos) + ";BGW"
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'绝对位置运动
Public Function PosMoveAbs(ByVal hDmc As Long, ByVal axis As Integer, ByVal pos As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "PAX=" + str(pos) + ";BGX"
    ElseIf axis = 1 Then
        Cmd = "PAY=" + str(pos) + ";BGY"
    ElseIf axis = 2 Then
        Cmd = "PAZ=" + str(pos) + ";BGZ"
    ElseIf axis = 3 Then
        Cmd = "PAW=" + str(pos) + ";BGW"
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'连续运动
Public Function ContinousMove(ByVal hDmc As Long, ByVal axis As Integer, ByVal spd As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "JGX=" + str(spd) + ";BGX"
    ElseIf axis = 1 Then
        Cmd = "JGY=" + str(spd) + ";BGY"
    ElseIf axis = 2 Then
        Cmd = "JGZ=" + str(spd) + ";BGZ"
    ElseIf axis = 3 Then
        Cmd = "JGW=" + str(spd) + ";BGW"
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'指定当前位置
Public Function DefinePos(ByVal hDmc As Long, ByVal axis As Integer, ByVal pos As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "DPX=" + str(pos)
    ElseIf axis = 1 Then
        Cmd = "DPY=" + str(pos)
    ElseIf axis = 2 Then
        Cmd = "DPZ=" + str(pos)
    ElseIf axis = 3 Then
        Cmd = "DPW=" + str(pos)
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'指定编码器位置
Public Function DefineEnc(ByVal hDmc As Long, ByVal axis As Integer, ByVal pos As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "DEX=" + str(pos)
    ElseIf axis = 1 Then
        Cmd = "DEY=" + str(pos)
    ElseIf axis = 2 Then
        Cmd = "DEZ=" + str(pos)
    ElseIf axis = 3 Then
        Cmd = "DEW=" + str(pos)
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'使能轴
Public Function EnableAxis(ByVal hDmc As Long, ByVal axis As Integer)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "SHX"
    ElseIf axis = 1 Then
        Cmd = "SHY"
    ElseIf axis = 2 Then
        Cmd = "SHZ"
    ElseIf axis = 3 Then
        Cmd = "SHW"
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'停止轴
Public Function StopAxis(ByVal hDmc As Long, ByVal axis As Integer)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "STX"
    ElseIf axis = 1 Then
        Cmd = "STY"
    ElseIf axis = 2 Then
        Cmd = "STZ"
    ElseIf axis = 3 Then
        Cmd = "STW"
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'获取当前轴位置
Public Function GetPos(ByVal hDmc As Long, ByVal axis As Integer) As Long
    Dim pos As Long
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        rc = DMCCommand(hDmc, "TDX", Response, 256)
        pos = val(Response)
    ElseIf axis = 1 Then
        rc = DMCCommand(hDmc, "TDY", Response, 256)
        pos = val(Response)
    ElseIf axis = 2 Then
        rc = DMCCommand(hDmc, "TDZ", Response, 256)
        pos = val(Response)
    ElseIf axis = 3 Then
        rc = DMCCommand(hDmc, "TDW", Response, 256)
        pos = val(Response)
    End If
    GetPos = pos
End Function
'获取当前轴位置
Public Function GetVel(ByVal hDmc As Long, ByVal axis As Integer) As Long
    Dim vel As Long
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        rc = DMCCommand(hDmc, "SP?", Response, 256)
        vel = val(Response)
    ElseIf axis = 1 Then
        rc = DMCCommand(hDmc, "TVY", Response, 256)
        vel = val(Response)
    ElseIf axis = 2 Then
        rc = DMCCommand(hDmc, "TVZ", Response, 256)
        vel = val(Response)
    ElseIf axis = 3 Then
        rc = DMCCommand(hDmc, "TVW", Response, 256)
        vel = val(Response)
    End If
    GetVel = vel
End Function
'获取当前轴编码器位置
Public Function GetPosEnc(ByVal hDmc As Long, ByVal axis As Integer) As Long
    Dim pos As Long
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    
    If axis = 0 Then
        rc = DMCCommand(hDmc, "TPX", Response, 256)
        pos = val(Response)
    ElseIf axis = 1 Then
        rc = DMCCommand(hDmc, "TPY", Response, 256)
        pos = val(Response)
    ElseIf axis = 2 Then
        rc = DMCCommand(hDmc, "TPZ", Response, 256)
        pos = val(Response)
    ElseIf axis = 3 Then
        rc = DMCCommand(hDmc, "TPW", Response, 256)
        pos = val(Response)
    End If
    GetPosEnc = pos
End Function
'获取当前轴状态
Public Function GetStatus(ByVal hDmc As Long, ByVal axis As Integer) As Integer
    Dim sta As Integer
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    
    If axis = 0 Then
        rc = DMCCommand(hDmc, "TSX", Response, 256)
        
    ElseIf axis = 1 Then
        rc = DMCCommand(hDmc, "TSY", Response, 256)
        
    ElseIf axis = 2 Then
        rc = DMCCommand(hDmc, "TSZ", Response, 256)
        
    ElseIf axis = 3 Then
        rc = DMCCommand(hDmc, "TSW", Response, 256)
        
    End If
    sta = val(Response)
    GetStatus = (sta And &H80) / 2 ^ 7
End Function
'输出点控制
Public Function WriteOutBit(ByVal hDmc As Long, ByVal port, ByVal val)
    Dim sta As Integer
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    Cmd = "OB " + str(port) + "," + str(val)
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'读输入点值
Public Function ReadInBit(ByVal hDmc As Long, ByVal port) As Integer
    Dim sta As Integer
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    Dim v As Integer
    Cmd = "MG @IN[" + str(port) + "]"
    rc = DMCCommand(hDmc, Cmd, Response, 256)
    
    v = val(Response)
    ReadInBit = v
End Function
'读当前输出点值
Public Function ReadOutBit(ByVal hDmc As Long, ByVal port) As Integer
    Dim sta As Integer
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    Dim v As Integer
    Cmd = "MG @OUT[" + str(port) + "]"
    rc = DMCCommand(hDmc, Cmd, Response, 256)
    
    v = val(Response)
    ReadOutBit = v
End Function

'回零运动
Public Function GoHome(ByVal hDmc As Long, ByVal axis As Integer, ByVal spd As Long)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        
        Cmd = "JGX=" + str(spd) + ";HMX;BGX"
        'Cmd = "HMX;BGX"
    ElseIf axis = 1 Then
        'Cmd = "JGY=" + Str(spd) + ";BGY;HMY;BGY"
        Cmd = "JGY=" + str(spd) + ";HMY;BGY"
    ElseIf axis = 2 Then
        Cmd = "JGZ=" + str(spd) + ";HMZ;BGZ"
    ElseIf axis = 3 Then
        Cmd = "JGW=" + str(spd) + ";HMW;BGW"
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function

'获取当前轴原点状态， 0有效
Public Function GetHMStatus(ByVal hDmc As Long, ByVal axis As Integer) As Integer
    Dim sta As Integer
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    
    If axis = 0 Then
        rc = DMCCommand(hDmc, "TSX", Response, 256)
        
    ElseIf axis = 1 Then
        rc = DMCCommand(hDmc, "TSY", Response, 256)
        
    ElseIf axis = 2 Then
        rc = DMCCommand(hDmc, "TSZ", Response, 256)
        
    ElseIf axis = 3 Then
        rc = DMCCommand(hDmc, "TSW", Response, 256)
        
    End If
    sta = val(Response)
    GetHMStatus = (sta And &H2) / 2 ^ 1
End Function
'设置FE方向
Public Function SetFEdir(ByVal hDmc As Long, ByVal dir As Integer) '1=正向找 ，-1= 负向找
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    
        Cmd = "CN," + str(dir)
    
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
'设置轴输出方向,mode=2,mode = -2
Public Function SetAxisOutMode(ByVal hDmc As Long, ByVal axis As Integer, ByVal mode As Integer)
    Dim rc As Long
    Dim Response As String * 256
    Dim Cmd As String
    If axis = 0 Then
        Cmd = "MTX=" + str(mode)
    ElseIf axis = 1 Then
        Cmd = "MTY=" + str(mode)
    ElseIf axis = 2 Then
        Cmd = "MTZ=" + str(mode)
    ElseIf axis = 3 Then
        Cmd = "MTW=" + str(mode)
    End If
    rc = DMCCommand(hDmc, Cmd, Response, 256)
End Function
