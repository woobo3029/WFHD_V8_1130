Attribute VB_Name = "MdlCommon"
Option Explicit

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_Flags = &H2 Or &H1

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Type POINT_API
    x As Long
    y As Long
End Type

Public MousePos As POINT_API
    
Public Type ScreenPoint
    x As Double
    y As Double
End Type

Sub Wait(ByVal sec As Double)
    Dim tm As Double
    
    tm = Timer
    Do
        If Timer < tm Or Timer - tm > sec Then
            Exit Do
        End If
        DoEvents
    Loop
End Sub

Sub MATSet(ByRef m() As Double, ByRef a() As Double)
    Dim I As Integer, j As Integer
    
    For I = 0 To 3
        For j = 0 To 3
            m(I, j) = a(I, j)
        Next
    Next
End Sub

Sub MATLeftMultiply(ByRef m() As Double, ByRef a() As Double)
    Dim I As Integer, j As Integer, k As Integer
    Dim c(4, 4) As Double
    
    For I = 0 To 3
        For j = 0 To 3
            c(I, j) = 0
            For k = 0 To 3
                c(I, j) = c(I, j) + a(I, k) * m(k, j)
            Next
        Next
    Next
    
    For I = 0 To 3
        For j = 0 To 3
            m(I, j) = c(I, j)
        Next
    Next
End Sub

Function MouseInWindow(ByRef Win As Object) As Boolean
    GetCursorPos MousePos
    ScreenToClient Win.hwnd, MousePos
    
    If MousePos.x >= 0 And MousePos.x < Win.ScaleWidth And MousePos.y >= 0 And MousePos.y < Win.ScaleHeight Then
        MouseInWindow = True
    Else
        MouseInWindow = False
    End If
End Function

Public Sub Rotate_Z(x0 As Single, y0 As Single, angle As Double, x1 As Single, y1 As Single)
    Dim CS As Double, SN As Double
    
    CS = Cos(angle)
    SN = Sin(angle)
    
    x1 = (CS * x0) - (SN * y0)
    y1 = (SN * x0) + (CS * y0)
End Sub

Function Max(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Function Min(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Public Function GetStringFromINI(ByVal AppName As String, ByVal KeyName As String, ByVal Default As String, ByVal INIFile As String) As String
    Dim s As String, s_len As Long
    
    s = String(256, "*")
    s_len = Len(s)
    
    GetPrivateProfileString AppName, KeyName, vbNullString, s, s_len, INIFile
    s = left(s, IIf(InStr(s, Chr(0)) = 0, 0, InStr(s, Chr(0)) - 1))
    If s = "" Then s = Default
    GetStringFromINI = Trim(s)
End Function

Public Function GetValueFromINI(ByVal AppName As String, ByVal KeyName As String, ByVal Default As String, ByVal INIFile As String) As Double
    Dim s As String, s_len As Long
    
    s = String(256, "*")
    s_len = Len(s)
    
    GetPrivateProfileString AppName, KeyName, vbNullString, s, s_len, INIFile
    s = left(s, IIf(InStr(s, Chr(0)) = 0, 0, InStr(s, Chr(0)) - 1))
    If s = "" Then s = Default
    GetValueFromINI = Val(Trim(s))
End Function

Public Sub ShowMsg(ByVal txt As String)
    FrmMsgDlg.LblMessage.caption = txt
    FrmMsgDlg.Show
End Sub

Function TimeDiff(Time1 As Variant, Time0 As Variant) As Double
    TimeDiff = Time1 - Time0
    If TimeDiff < 0 Then
        TimeDiff = TimeDiff + 86400 '24*60*60 -- one day
    End If
End Function

Function GetAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Double
    Dim p12 As Double, p23 As Double, p123 As Double, cosa As Double
    Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double, d As Double
    
    p12 = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    p23 = Sqr((x3 - x2) ^ 2 + (y3 - y2) ^ 2)
    p123 = (x2 - x1) * (x3 - x2) + (y2 - y1) * (y3 - y2)
    
    If p12 = 0 Or p23 = 0 Then
        GetAngle = 0
        Exit Function
    End If
    
    cosa = p123 / (p12 * p23)
    
    If cosa = 1 Then
        GetAngle = 0
        Exit Function
    ElseIf cosa = -1 Then
        GetAngle = 180
        Exit Function
    End If
    
    vx1 = x2 - x1
    vy1 = y2 - y1
    vx2 = x3 - x1
    vy2 = y3 - y1
        
    d = Sgn(vx1 * vy2 - vx2 * vy1)
    
    'GetAngle = d * ArcCos(cosa) * 180 / PI
    If Abs(1 - cosa * cosa) < 0.00000001 Then
        GetAngle = 0
    Else
        GetAngle = d * (Atn(-cosa / Sqr(1 - cosa * cosa)) + PI_2) / PI_180
    End If
End Function

Function GetFromINI(strVariableName As String) As String
    Dim strReturn As String
    
    strReturn = String(255, Chr(0))
    GetFromINI = left$(strReturn, GetPrivateProfileString("System", ByVal strVariableName, "", strReturn, Len(strReturn), App.Path & "\Parameters.ini"))
End Function

Function GetVFromINI(strVariableName As String) As String
    Dim strReturn As String
    
    strReturn = String(255, Chr(0))
    GetVFromINI = Val(left$(strReturn, GetPrivateProfileString("System", ByVal strVariableName, "", strReturn, Len(strReturn), App.Path & "\Parameters.ini")))
End Function

Function WriteToINI(strVariableName As String, strValue As String) As Integer
    WriteToINI = WritePrivateProfileString("System", strVariableName, strValue, App.Path & "\Parameters.ini")
End Function

Function GetVFromINI_A(strVariableName As String) As String
    Dim strReturn As String
    
    strReturn = String(255, Chr(0))
    GetVFromINI_A = Val(left$(strReturn, GetPrivateProfileString("Angle_List_of_" + Device_CurMaterial, ByVal strVariableName, "", strReturn, Len(strReturn), App.Path & "\Parameters.ini")))
    'GetVFromINI_A = Val(left$(strReturn, GetPrivateProfileString("Angle", ByVal strVariableName, "", strReturn, Len(strReturn), App.Path & "\Parameters.ini")))
End Function

Function WriteToINI_A(strVariableName As String, strValue As String) As Integer
    WriteToINI_A = WritePrivateProfileString("Angle_List_of_" + Device_CurMaterial, strVariableName, strValue, App.Path & "\Parameters.ini")
End Function


Sub SetDigiPad(ByVal form_name As String, ByVal obj_name As String)
    Dim obj As Object
    
    FrmDigiPad.Tag = form_name
    FrmDigiPad.TxtEdit.Tag = obj_name
    
    If form_name = "FrmMain" Then
        For Each obj In FrmMain
            If obj.Name = obj_name Then
                If TypeOf obj Is TextBox Then
                    FrmDigiPad.TxtEdit.Text = obj.Text
                    
                ElseIf TypeOf obj Is MSFlexGrid Then
                    FrmDigiPad.PanButton1.Tag = str(obj.Row)
                    FrmDigiPad.PanButton2.Tag = str(obj.Col)
                    
                    FrmDigiPad.TxtEdit.Text = obj.TextMatrix(obj.Row, obj.Col)
                    
                End If
                Exit For
            End If
        Next
        
    ElseIf form_name = "FormSettings" Then
        For Each obj In FormSettings
            If obj.Name = obj_name Then
                If TypeOf obj Is TextBox Then
                    FrmDigiPad.TxtEdit.Text = obj.Text
                    
                ElseIf TypeOf obj Is MSFlexGrid Then
                    FrmDigiPad.PanButton1.Tag = str(obj.Row)
                    FrmDigiPad.PanButton2.Tag = str(obj.Col)
                    
                    FrmDigiPad.TxtEdit.Text = obj.TextMatrix(obj.Row, obj.Col)
                    
                End If
                Exit For
            End If
        Next
    ElseIf form_name = "FormGetPulse" Then
        For Each obj In FormGetPulse
            If obj.Name = obj_name Then
                If TypeOf obj Is TextBox Then
                    FrmDigiPad.TxtEdit.Text = obj.Text
                    
                ElseIf TypeOf obj Is MSFlexGrid Then
                    FrmDigiPad.PanButton1.Tag = str(obj.Row)
                    FrmDigiPad.PanButton2.Tag = str(obj.Col)
                    
                    FrmDigiPad.TxtEdit.Text = obj.TextMatrix(obj.Row, obj.Col)
                    
                End If
                Exit For
            End If
        Next
        
    End If
    FrmDigiPad.Show

End Sub
