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
        Device_OutPort 40, IIf(Device_OutportSignalUpset = 1, 0, 1) 'D40 - ���ȿ�
        Device_OutPort 45, IIf(Device_OutportSignalUpset = 1, 0, 1) 'D45 - ���ȿ�
        OpenFan = True
    Else
        Device_OutPort 40, IIf(Device_OutportSignalUpset = 1, 1, 0) 'D40 - ���ȹ�
        Device_OutPort 45, IIf(Device_OutportSignalUpset = 1, 1, 0) 'D45 - ���ȹ�
        OpenFan = False
    End If
End Sub
