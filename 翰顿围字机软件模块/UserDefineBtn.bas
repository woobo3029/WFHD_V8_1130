Attribute VB_Name = "UserDefineBtn"
Option Explicit

Public OpenFan As Boolean

Public Sub FanSwitch()
    Static tm0 As Double, tm As Double
    
    tm = Timer
    If Abs(tm - tm0) < 1 Then
        Exit Sub
    End If
    tm0 = tm
    
    If OpenFan = False Then
        Device_OutPort 40, IIf(Device_OutportSignalUpset = 1, 0, 1) 'D40 - 风扇开
        Device_OutPort 45, IIf(Device_OutportSignalUpset = 1, 0, 1) 'D45 - 风扇开
        OpenFan = True
    Else
        Device_OutPort 40, IIf(Device_OutportSignalUpset = 1, 1, 0) 'D40 - 风扇关
        Device_OutPort 45, IIf(Device_OutportSignalUpset = 1, 1, 0) 'D45 - 风扇关
        OpenFan = False
    End If
End Sub
