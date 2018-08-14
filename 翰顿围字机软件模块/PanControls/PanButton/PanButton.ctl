VERSION 5.00
Begin VB.UserControl PanButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   435
      Left            =   1140
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   420
      Top             =   720
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "caption"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   300
      TabIndex        =   1
      Top             =   1440
      Width           =   630
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   465
      Left            =   1560
      Top             =   1200
      Width           =   420
   End
End
Attribute VB_Name = "PanButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private mPos As POINTAPI

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Enum BorderStyleEnum
    BorderStyleDark = 0
    BorderStyleLight = 1
    BorderStyle3D = 2
End Enum

Public Event Click()
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mPicture As IPictureDisp
Private mCaption As String
Private mBorderStyle As Integer
Private mEnabled As Boolean
Private mFontName As String
Private mFontSize As Integer

Sub ShowButton(Optional MouseAction As Integer = 0)
    'MouseAction: 0-no action, 1-mouse down, 2-mouse up
    Dim r As Integer, l As Integer, w As Integer, clr As Long, cw As Integer, ch As Integer, I As Integer
    Dim BorderClr(8) As Long
    Dim x0 As Integer, y0 As Integer, X As Integer, Y As Integer, bkclr As Long
        
    Dim PixelRed As Long, PixelGreen As Long, PixelBlue As Long
    Dim PixelGray As Long, mclr As Single
    
    On Error Resume Next
    
    UserControl.BorderStyle = 0
    If mBorderStyle = 0 Then
        BorderClr(1) = RGB(128, 128, 128)
        BorderClr(2) = RGB(128, 128, 128)
        BorderClr(3) = RGB(128, 128, 128)
        BorderClr(4) = RGB(128, 128, 128)
        BorderClr(5) = RGB(128, 128, 128)
        BorderClr(6) = RGB(128, 128, 128)
        BorderClr(7) = RGB(128, 128, 128)
        BorderClr(8) = RGB(128, 128, 128)
    ElseIf mBorderStyle = 1 Then
        BorderClr(1) = RGB(200, 200, 255)
        BorderClr(2) = RGB(200, 200, 255)
        BorderClr(3) = RGB(200, 200, 255)
        BorderClr(4) = RGB(200, 200, 255)
        BorderClr(5) = RGB(200, 200, 255)
        BorderClr(6) = RGB(200, 200, 255)
        BorderClr(7) = RGB(200, 200, 255)
        BorderClr(8) = RGB(200, 200, 255)
    Else
        BorderClr(1) = RGB(200, 200, 255)
        BorderClr(2) = RGB(128, 128, 255)
        BorderClr(3) = RGB(128, 128, 255)
        BorderClr(4) = RGB(0, 0, 255)
        BorderClr(5) = RGB(0, 0, 128)
        BorderClr(6) = RGB(128, 128, 255)
        BorderClr(7) = RGB(200, 200, 255)
        BorderClr(8) = RGB(230, 230, 255)
    End If
    
    UserControl.AutoRedraw = True
    UserControl.Cls
    UserControl.DrawMode = 13
    UserControl.DrawWidth = 1
    UserControl.Line (2, 0)-Step(UserControl.ScaleWidth - 5, 0), BorderClr(1)
    UserControl.Line -(UserControl.ScaleWidth - 1, 2), BorderClr(2)
    UserControl.Line -Step(0, UserControl.ScaleHeight - 5), BorderClr(3)
    UserControl.Line -(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1), BorderClr(4)
    UserControl.Line -(2, UserControl.ScaleHeight - 1), BorderClr(5)
    UserControl.Line -(0, UserControl.ScaleHeight - 3), BorderClr(6)
    UserControl.Line -(0, 2), BorderClr(7)
    UserControl.Line -(2, 0), BorderClr(8)
    
    For r = 1 To UserControl.ScaleHeight - 2
        If r = 1 Or r = UserControl.ScaleHeight - 2 Then
            l = 2
            w = UserControl.ScaleWidth - 4
        Else
            l = 1
            w = UserControl.ScaleWidth - 2
        End If
        If MouseAction <> 1 Then
            If r < UserControl.ScaleHeight / 10 Then
                clr = RGB(255, 255, 255)
            ElseIf r < 2 * UserControl.ScaleHeight / 10 Then
                clr = RGB(247, 247, 247)
            ElseIf r < 3 * UserControl.ScaleHeight / 10 Then
                clr = RGB(242, 242, 242)
            ElseIf r < 4 * UserControl.ScaleHeight / 10 Then
                clr = RGB(239, 239, 239)
            ElseIf r < 5 * UserControl.ScaleHeight / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 6 * UserControl.ScaleHeight / 10 Then
                clr = RGB(233, 233, 233)
            ElseIf r < 7 * UserControl.ScaleHeight / 10 Then
                clr = RGB(230, 230, 230)
            ElseIf r < 8 * UserControl.ScaleHeight / 10 Then
                clr = RGB(227, 227, 227)
            ElseIf r < 9 * UserControl.ScaleHeight / 10 Then
                clr = RGB(224, 224, 224)
            Else
                clr = RGB(220, 220, 220)
            End If
        Else
            clr = RGB(200, 200, 200)
        End If
        UserControl.Line (l, r)-Step(w, 0), clr
    Next
    
    Picture1.AutoSize = True
    If Not mPicture Is Nothing Then
        Set Picture1.Picture = mPicture
    Else
        Picture1.Width = 0
    End If
    
    LblCaption.caption = mCaption
    LblCaption.FontName = mFontName
    LblCaption.FontSize = mFontSize
    For I = 1 To 5
        If Picture1.Width + LblCaption.Width + 2 <= UserControl.ScaleWidth Then
            Exit For
        End If
        
        LblCaption.FontSize = LblCaption.FontSize - 1
    Next
    
    cw = LblCaption.Width 'UserControl.TextWidth(mCaption)
    ch = LblCaption.Height 'UserControl.TextHeight(mCaption)
    
    If Not mPicture Is Nothing Then
        'If mEnabled = True Then
        '    Set Image1.Picture = mPicture
        '    Image1.Move (UserControl.ScaleWidth - Image1.Width - cw) / 2, (UserControl.ScaleHeight - Image1.Height) / 2
        '    Image1.Visible = True
        '
        '    UserControl.CurrentX = Image1.left + Image1.Width + 2
        'Else
            Image1.Visible = False
            
            x0 = (UserControl.ScaleWidth - Picture1.Width - cw) / 2
            y0 = (UserControl.ScaleHeight - Picture1.Height) / 2
                        
            mclr = 0.5 ' 0 -> 1
            bkclr = Picture1.Point(0, 0)
            For Y = 0 To Picture1.Height - 1
                For X = 0 To Picture1.Width - 1
                    clr = Picture1.Point(X, Y)
                    If clr <> bkclr Then
                        If mEnabled = False Then
                            PixelBlue = Int(clr / 65536)
                            PixelGreen = Int((clr - PixelBlue * 65536) / 256)
                            PixelRed = Int(clr - PixelBlue * 65536 - PixelGreen * 256)
                        
                            'PixelBlue = mclr * PixelBlue
                            'If PixelBlue > 255 Then PixelBlue = 255
                            
                            'PixelGreen = mclr * PixelGreen
                            'If PixelGreen > 255 Then PixelGreen = 255
                            
                            'PixelRed = mclr * PixelRed
                            'If PixelRed > 255 Then PixelRed = 255
            
                            'clr = RGB(PixelRed, PixelGreen, PixelBlue)
                            
                            PixelGray = (30 * PixelRed + 59 * PixelGreen + 11 * PixelBlue) / 100
                            PixelGray = PixelGray + mclr * (255 - PixelGray)
                            If PixelGray > 255 Then PixelGray = 255
                            
                            clr = RGB(PixelGray, PixelGray, PixelGray)
                        End If
                        UserControl.PSet (x0 + X + 1, y0 + Y + 1), clr
                    End If
                Next
            Next
        
            UserControl.CurrentX = x0 + Picture1.Width + 2
        'End If
    Else
        UserControl.CurrentX = (UserControl.ScaleWidth - cw) / 2
    End If
    'UserControl.CurrentX = Image1.Left + Image1.Width + 5
    UserControl.CurrentY = (UserControl.ScaleHeight - ch) / 2
    
    'UserControl.FontName = "宋体"
    'UserControl.FontSize = 9
    'UserControl.ForeColor = IIf(mEnabled = True, RGB(0, 0, 0), RGB(180, 180, 180))
    'UserControl.Print mCaption
    LblCaption.left = Int(UserControl.CurrentX)
    LblCaption.top = Int(UserControl.CurrentY)
    LblCaption.ForeColor = IIf(mEnabled = True, RGB(0, 0, 0), RGB(180, 180, 180))

    If MouseAction <> 0 Then
        UserControl.DrawMode = 7
        UserControl.DrawWidth = 2
        
        UserControl.Line (2, 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), RGB(20, 60, 160), B
        UserControl.Line (2, 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), RGB(160, 60, 20), B
    End If
End Sub

Private Sub LblCaption_Click()
    UserControl_Click
End Sub

Private Sub LblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, LblCaption.left + X / Screen.TwipsPerPixelX, LblCaption.top + Y / Screen.TwipsPerPixelY)
End Sub

Private Sub LblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub LblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, LblCaption.left + X / Screen.TwipsPerPixelX, LblCaption.top + Y / Screen.TwipsPerPixelY)
End Sub

Private Sub Timer1_Timer()
    If mEnabled = False Then
        Exit Sub
    End If
    
    Call GetCursorPos(mPos)

    If Not WindowFromPoint(mPos.X, mPos.Y) = UserControl.hWnd Then
        MouseMoveButton 0
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Click()
    If mEnabled = False Then
        Exit Sub
    End If
    
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    If mEnabled = False Then
        Exit Sub
    End If
    
    GetFocusButton 1
    'LblCaption.FontUnderline = True
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If mEnabled = False Then
        Exit Sub
    End If
    
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_LostFocus()
    If mEnabled = False Then
        Exit Sub
    End If
    
    GetFocusButton 0
    'LblCaption.FontUnderline = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then
        Exit Sub
    End If
    
    ShowButton 1
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then
        Exit Sub
    End If
    
    MouseMoveButton 1
    Timer1.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then
        Exit Sub
    End If
    
    ShowButton 2
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    'UserControl.FontName = "宋体"
    'UserControl.FontName = "Arial"
    'UserControl.FontName = "Courier New"
    'UserControl.FontSize = 9
    
    ShowButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mPicture = PropBag.ReadProperty("Picture", Nothing)
    mCaption = PropBag.ReadProperty("Caption", "")
    mBorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    mFontName = PropBag.ReadProperty("FontName", "宋体")
    mFontSize = PropBag.ReadProperty("FontSize", 9)
    
    ShowButton
End Sub

Private Sub UserControl_Resize()
    UserControl.BackColor = RGB(255, 0, 0) 'Extender.Parent.BackColor
    UserControl.BackColor = Extender.Parent.BackColor
    ShowButton
End Sub

Sub MouseMoveButton(ByVal status As Integer)
    Static s As Integer
    
    If s <> status Then
        s = status
        UserControl.DrawMode = 7
        UserControl.DrawWidth = 2
        
        UserControl.Line (2, 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), RGB(20, 60, 160), B
    End If
End Sub

Sub GetFocusButton(ByVal status As Integer)
    Static s As Integer
    
    If s <> status Then
        s = status
        UserControl.DrawMode = 7
        UserControl.DrawWidth = 2
        
        UserControl.Line (2, 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), RGB(160, 60, 20), B
    End If
End Sub

Public Property Get Picture() As IPictureDisp
    Set Picture = mPicture
End Property

Public Property Set Picture(NewPicture As IPictureDisp)
    Set mPicture = NewPicture
    
    ShowButton
    PropertyChanged "Picture"
End Property

Public Property Get caption() As String
    caption = mCaption
End Property

Public Property Let caption(NewCaption As String)
    mCaption = NewCaption
    
    ShowButton
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(NewValue As Boolean)
    UserControl.Enabled = NewValue
    mEnabled = NewValue
    
    ShowButton
    PropertyChanged "Enabled"
End Property

Public Property Get BorderStyle() As BorderStyleEnum
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(NewValue As BorderStyleEnum)
    mBorderStyle = NewValue
    
    ShowButton
    PropertyChanged "BorderStyle"
End Property

Public Property Get FontName() As String
    FontName = mFontName
End Property

Public Property Let FontName(NewValue As String)
    mFontName = NewValue
    
    ShowButton
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Integer
    FontSize = mFontSize
End Property

Public Property Let FontSize(NewValue As Integer)
    mFontSize = NewValue
    
    ShowButton
    PropertyChanged "FontSize"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Picture", mPicture, Nothing
    PropBag.WriteProperty "Caption", mCaption, ""
    PropBag.WriteProperty "BorderStyle", mBorderStyle, 0
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "FontName", mFontName, "宋体"
    PropBag.WriteProperty "FontSize", mFontSize, 9
End Sub
