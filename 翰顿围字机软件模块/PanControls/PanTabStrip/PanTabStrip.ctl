VERSION 5.00
Begin VB.UserControl PanTabStrip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   226
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   300
      Top             =   600
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "caption"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   435
      Index           =   0
      Left            =   300
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "PanTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Event TabClick(Index As Integer)
Public Event TabScrollLeft()
Public Event TabScrollRight()
Public Event KeyDown(KeyCode As Integer)

Private mPos As POINTAPI
Private mTabCount As Integer
Private mTabMinWidth As Integer
Private mTabCurIndex As Integer
Private mPicture() As IPictureDisp
Private mCaption() As String
Private mEnabled As Boolean
Private mGotFocus As Boolean

Private mTabScroll As Boolean
Private mTabScrollLeftTab As Integer
Private mTabScrollCount As Integer
Private mTabScrollTabWidth As Integer
Private mTabScrollBtnWidth As Integer

Private Sub LblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, LblCaption(Index).Left + X / Screen.TwipsPerPixelX, LblCaption(Index).Top + Y / Screen.TwipsPerPixelY
End Sub

Private Sub LblCaption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, LblCaption(Index).Left + X / Screen.TwipsPerPixelX, LblCaption(Index).Top + Y / Screen.TwipsPerPixelY
End Sub

Private Sub Timer1_Timer()
    Call GetCursorPos(mPos)
    
    If Not WindowFromPoint(mPos.X, mPos.Y) = UserControl.hWnd Then
        MouseMoveTabStrip 0, 0
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_GotFocus()
    mGotFocus = True
    
    ShowTabStripFocus
End Sub

Private Sub UserControl_Initialize()
    mTabCount = 1
    mTabCurIndex = 0
    
    mTabMinWidth = 100
    mTabScrollLeftTab = 0
End Sub

Public Property Get TabCount() As Integer
    TabCount = mTabCount
End Property

Public Property Let TabCount(ByVal NewValue As Integer)
    Dim I As Integer
    
    mTabCount = NewValue
    
    ReDim Preserve mPicture(mTabCount)
    ReDim Preserve mCaption(mTabCount)
    
    If LblCaption.Count < mTabCount Then
        For I = LblCaption.Count To mTabCount - 1
            Load LblCaption(I)
            LblCaption(I).AutoSize = True
        Next
    ElseIf LblCaption.Count > mTabCount Then
        For I = LblCaption.Count - 1 To mTabCount Step -1
            Unload LblCaption(I)
        Next
    End If
    
    ShowTabStrip
    PropertyChanged "TabCount"
End Property

Sub ShowTabStrip()
    Dim w0 As Integer, w As Integer, d As Integer, I As Integer, cw As Integer, ch As Integer
    Dim k As Integer
    
    On Error Resume Next
        
    If mGotFocus = True Then
        k = 1
        mGotFocus = False
        ShowTabStripFocus
    End If
    
    '---------- auto check mTabMinWidth ------------
    w = mTabMinWidth
    For I = 0 To mTabCount - 1
        LblCaption(I).Caption = mCaption(I)
        If LblCaption(I).Width > w Then
            w = LblCaption(I).Width
        End If
    Next
    mTabMinWidth = w
    '------------------------------------------------
    
    If UserControl.ScaleWidth / mTabCount >= mTabMinWidth Then
        mTabScroll = False
        mTabScrollCount = mTabCount
    Else
        mTabScroll = True
        mTabScrollCount = Int((UserControl.ScaleWidth - 2 * 15) / mTabMinWidth)
    End If
    
    If mTabScroll Then d = 15 Else d = 0
    mTabScrollTabWidth = Int((UserControl.ScaleWidth - 2 * d) / mTabScrollCount)
    If mTabScroll Then
        mTabScrollBtnWidth = (UserControl.ScaleWidth - mTabScrollCount * mTabScrollTabWidth) / 2
    Else
        mTabScrollBtnWidth = 0
    End If
    
    For I = mTabCount To Image1.Count - 1
        Unload Image1(I)
    Next
    For I = Image1.Count To mTabCount - 1
        Load Image1(I)
        'Image1(i).Visible = True
    Next
    
    ShowTabs

    'UserControl.FontName = "ËÎÌו"
    'UserControl.FontSize = 9
    
    'UserControl.ForeColor = IIf(mEnabled, RGB(0, 0, 0), RGB(160, 160, 160))
    
    For I = 0 To mTabCount - 1
        If I <= UBound(mPicture) Then
            If I >= mTabScrollLeftTab And I < mTabScrollLeftTab + mTabScrollCount Then
                cw = UserControl.TextWidth(mCaption(I))
                ch = UserControl.TextHeight(mCaption(I))
                
                If Not mPicture(I) Is Nothing Then
                    Set Image1(I).Picture = mPicture(I)
                    Image1(I).Move mTabScrollBtnWidth + (I - mTabScrollLeftTab) * mTabScrollTabWidth + (mTabScrollTabWidth - Image1(I).Width - 5 - cw) / 2, (UserControl.ScaleHeight - Image1(I).Width) / 2
                    Image1(I).Visible = True
                
                    UserControl.CurrentX = Image1(I).Left + Image1(I).Width + 5
                Else
                    Image1(I).Visible = False
                    UserControl.CurrentX = mTabScrollBtnWidth + (I - mTabScrollLeftTab) * mTabScrollTabWidth + (mTabScrollTabWidth - cw) / 2
                End If
                UserControl.CurrentY = (UserControl.ScaleHeight - ch) / 2
                'UserControl.Print mCaption(i)
                
                LblCaption(I).Left = UserControl.CurrentX
                LblCaption(I).Top = UserControl.CurrentY
                LblCaption(I).ForeColor = IIf(mEnabled, RGB(0, 0, 0), RGB(160, 160, 160))
                LblCaption(I).Caption = mCaption(I)
                LblCaption(I).Visible = True
            Else
                Image1(I).Visible = False
                LblCaption(I).Visible = False
            End If
        End If
    Next
    
    If k = 1 Then
        mGotFocus = True
        ShowTabStripFocus
    End If
    UserControl.Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If mEnabled = False Then
        Exit Sub
    End If

    'RaiseEvent KeyDown(KeyCode)
    If KeyCode = 37 Then
        UserControl_MouseUp 0, 0, 0, -9998
    ElseIf KeyCode = 39 Then
        UserControl_MouseUp 0, 0, 0, -9999
    End If
End Sub

Private Sub UserControl_LostFocus()
    mGotFocus = False
    
    ShowTabStripFocus
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    
    If Not mEnabled Then Exit Sub
    
    If X >= mTabScrollBtnWidth And X < mTabScrollBtnWidth + mTabScrollCount * mTabScrollTabWidth Then
        For I = 0 To mTabScrollCount - 1
            If X >= mTabScrollBtnWidth + I * mTabScrollTabWidth And X < mTabScrollBtnWidth + (I + 1) * mTabScrollTabWidth Then Exit For
        Next
        
        MouseMoveTabStrip 1, I
        Timer1.Enabled = True
        
    ElseIf mTabScroll Then
        If X < mTabScrollBtnWidth Then
            MouseMoveTabStrip 1, -1
            Timer1.Enabled = True
        
        ElseIf X >= mTabScrollBtnWidth + mTabScrollCount * mTabScrollTabWidth Then
            MouseMoveTabStrip 1, mTabScrollCount
            Timer1.Enabled = True
        
        End If
    End If
    
    MainFormUserEventRaised = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    
    If Not mEnabled Then Exit Sub
    
    'LblCaption(mTabCurIndex).FontUnderline = False
    
    If Y = -9998 Or Y = -9999 Then
        If Not mTabScroll Then
            If Y = -9998 Then
                If mTabCurIndex > 0 Then
                    mTabCurIndex = mTabCurIndex - 1
                    RaiseEvent TabClick(mTabCurIndex)
                End If
            ElseIf Y = -9999 Then
                If mTabCurIndex < mTabCount - 1 Then
                    mTabCurIndex = mTabCurIndex + 1
                    RaiseEvent TabClick(mTabCurIndex)
                End If
            End If
        Else
            If Y = -9998 Then
                If mTabCurIndex > 0 Then
                    If mTabCurIndex > mTabScrollLeftTab Then
                        mTabCurIndex = mTabCurIndex - 1
                        RaiseEvent TabClick(mTabCurIndex)
                    ElseIf mTabScrollLeftTab > 0 Then
                        mTabScrollLeftTab = mTabScrollLeftTab - 1
                        mTabCurIndex = mTabCurIndex - 1
                        RaiseEvent TabClick(mTabCurIndex)
                        RaiseEvent TabScrollLeft
                    End If
                End If
            ElseIf Y = -9999 Then
                If mTabCurIndex < mTabCount - 1 Then
                    If mTabCurIndex < mTabScrollLeftTab + mTabScrollCount - 1 Then
                        mTabCurIndex = mTabCurIndex + 1
                        RaiseEvent TabClick(mTabCurIndex)
                    ElseIf mTabScrollLeftTab + mTabScrollCount < mTabCount Then
                        mTabScrollLeftTab = mTabScrollLeftTab + 1
                        'If mTabCurIndex < mTabScrollLeftTab Then
                            mTabCurIndex = mTabCurIndex + 1
                            RaiseEvent TabClick(mTabCurIndex)
                        'End If
                        RaiseEvent TabScrollRight
                    End If
                End If
            End If
        End If
                    
        MouseMoveTabStrip 0, 0
        ShowTabStrip
        
        'LblCaption(mTabCurIndex).FontUnderline = True
        
        Exit Sub
    End If
    
    If X >= mTabScrollBtnWidth And X < mTabScrollBtnWidth + mTabScrollCount * mTabScrollTabWidth Then
        For I = 0 To mTabScrollCount - 1
            If X >= mTabScrollBtnWidth + I * mTabScrollTabWidth And X < mTabScrollBtnWidth + (I + 1) * mTabScrollTabWidth Then Exit For
        Next
        
        mTabCurIndex = mTabScrollLeftTab + I
        RaiseEvent TabClick(mTabCurIndex)
        
        MouseMoveTabStrip 0, 0
        ShowTabStrip
    
    ElseIf mTabScroll Then
        If X < mTabScrollBtnWidth Then
            If mTabScrollLeftTab >= mTabScrollCount Then
                mTabScrollLeftTab = mTabScrollLeftTab - mTabScrollCount
                mTabCurIndex = mTabScrollLeftTab
                RaiseEvent TabClick(mTabCurIndex)
                RaiseEvent TabScrollLeft
                    
                MouseMoveTabStrip 0, 0
                ShowTabStrip
            ElseIf mTabScrollLeftTab > 0 Then
                mTabScrollLeftTab = 0
                mTabCurIndex = 0
                RaiseEvent TabClick(mTabCurIndex)
                RaiseEvent TabScrollLeft
                
                MouseMoveTabStrip 0, 0
                ShowTabStrip
            End If
        ElseIf X >= mTabScrollBtnWidth + mTabScrollCount * mTabScrollTabWidth Then
            If mTabScrollLeftTab + 2 * mTabScrollCount < mTabCount Then
                mTabScrollLeftTab = mTabScrollLeftTab + mTabScrollCount + 1
                mTabCurIndex = mTabScrollLeftTab + mTabScrollCount - 1
                RaiseEvent TabClick(mTabCurIndex)
                RaiseEvent TabScrollRight
                
                MouseMoveTabStrip 0, 0
                ShowTabStrip
            ElseIf mTabScrollLeftTab + mTabScrollCount < mTabCount Then
                mTabScrollLeftTab = mTabCount - mTabScrollCount
                mTabCurIndex = mTabCount - 1
                RaiseEvent TabClick(mTabCurIndex)
                RaiseEvent TabScrollRight
                
                MouseMoveTabStrip 0, 0
                ShowTabStrip
            End If
        End If
        
    End If
    
    'LblCaption(mTabCurIndex).FontUnderline = True
End Sub

Private Sub UserControl_Paint()
    'ShowTabStrip 'will cause a dump loop
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim I As Integer
    
    On Error Resume Next
    
    mEnabled = PropBag.ReadProperty("Enabled", True)
    mTabCount = PropBag.ReadProperty("TabCount", 10)
    
    ReDim Preserve mPicture(mTabCount)
    For I = 0 To mTabCount - 1
        Set mPicture(I) = PropBag.ReadProperty("Picture" & Trim(Str(I)), Nothing)
    Next
    ReDim Preserve mCaption(mTabCount)
    For I = 0 To mTabCount - 1
        mCaption(I) = PropBag.ReadProperty("Caption" & Trim(Str(I)), "")
    Next
    
    If LblCaption.Count < mTabCount Then
        For I = LblCaption.Count To mTabCount - 1
            Load LblCaption(I)
        Next
    ElseIf LblCaption.Count > mTabCount Then
        For I = LblCaption.Count - 1 To mTabCount Step -1
            Unload LblCaption(I)
        Next
    End If
    
    mTabMinWidth = PropBag.ReadProperty("TabMinWidth", 100)
    mTabCurIndex = PropBag.ReadProperty("TabCutIndex", 0)
    
    Set UserControl.font = PropBag.ReadProperty("TabFont", UserControl.font)
    
    ShowTabStrip
End Sub

Private Sub UserControl_Resize()
    ShowTabStrip
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim I As Integer
    
    On Error Resume Next
    
    PropBag.WriteProperty "Enabled", mEnabled
    PropBag.WriteProperty "TabCount", mTabCount
    For I = 0 To mTabCount - 1
        PropBag.WriteProperty "Picture" & Trim(Str(I)), mPicture(I), Nothing
    Next
    For I = 0 To mTabCount - 1
        PropBag.WriteProperty "Caption" & Trim(Str(I)), mCaption(I), ""
    Next
    PropBag.WriteProperty "TabMinWidth", mTabMinWidth, 100
    PropBag.WriteProperty "TabCurIndex", mTabCurIndex
    
    PropBag.WriteProperty "TabFont", UserControl.font
End Sub

Sub ShowTabs()
    Dim I As Integer, r As Integer, l As Integer, h As Integer, clr As Long
    Dim dx As Integer, dy As Integer, dw As Integer
    Dim BorderClr(8) As Long
    Dim AutoRedraw0 As Boolean
    
    h = UserControl.ScaleHeight
    
    AutoRedraw0 = UserControl.AutoRedraw
    UserControl.AutoRedraw = True
    UserControl.Cls
    UserControl.BorderStyle = 0
    UserControl.BackColor = UserControl.Parent.BackColor
    UserControl.DrawMode = 13
    UserControl.DrawWidth = 1
    
    If mTabScroll Then
        BorderClr(1) = RGB(200, 200, 255)
        BorderClr(2) = RGB(128, 128, 128)
        BorderClr(3) = RGB(128, 128, 128)
        BorderClr(4) = RGB(180, 180, 255)
        BorderClr(5) = RGB(200, 200, 255)
        BorderClr(6) = RGB(200, 200, 255)
    
        dx = 0
        dy = 0
        
        UserControl.Line (dx + 2, dy)-Step(mTabScrollBtnWidth - 3, 0), BorderClr(1)
        UserControl.Line -Step(0, h - 1), BorderClr(2)
        UserControl.Line -(dx + 2, h - 1), BorderClr(3)
        UserControl.Line -(dx, h - 3), BorderClr(4)
        UserControl.Line -(dx, 2), BorderClr(5)
        UserControl.Line -(dx + 2, 0), BorderClr(6)
        
        For r = 1 To h - 2
            If r = 1 Or r = h - 2 Then
                l = 2
                dw = mTabScrollBtnWidth - 3
            Else
                l = 1
                dw = mTabScrollBtnWidth - 2
            End If
            If r < h / 10 Then
                clr = RGB(255, 255, 255)
            ElseIf r < 2 * h / 10 Then
                clr = RGB(247, 247, 247)
            ElseIf r < 3 * h / 10 Then
                clr = RGB(242, 242, 242)
            ElseIf r < 4 * h / 10 Then
                clr = RGB(239, 239, 239)
            ElseIf r < 5 * h / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 6 * h / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 7 * h / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 8 * h / 10 Then
                clr = RGB(233, 233, 233)
            ElseIf r < 9 * h / 10 Then
                clr = RGB(230, 230, 230)
            Else
                clr = RGB(227, 227, 227)
            End If
            UserControl.Line (dx + l, r)-Step(dw, 0), clr
        Next
        
        dx = mTabScrollBtnWidth / 2 - 3
        dy = UserControl.ScaleHeight / 2 - 5
        
        If mTabScrollLeftTab = 0 Or mEnabled = False Then
            clr = RGB(220, 220, 220)
        Else
            clr = RGB(0, 0, 128)
        End If
        
        UserControl.Line (dx + 4, dy + 0)-Step(1, 0), clr
        UserControl.Line (dx + 3, dy + 1)-Step(2, 0), clr
        UserControl.Line (dx + 2, dy + 2)-Step(3, 0), clr
        UserControl.Line (dx + 1, dy + 3)-Step(4, 0), clr
        UserControl.Line (dx + 0, dy + 4)-Step(5, 0), clr
        UserControl.Line (dx + 1, dy + 5)-Step(4, 0), clr
        UserControl.Line (dx + 2, dy + 6)-Step(3, 0), clr
        UserControl.Line (dx + 3, dy + 7)-Step(2, 0), clr
        UserControl.Line (dx + 4, dy + 8)-Step(1, 0), clr
        
        '-----------------------------------------------------------
        
        BorderClr(1) = RGB(200, 200, 255)
        BorderClr(2) = RGB(200, 200, 200)
        BorderClr(3) = RGB(128, 128, 128)
        BorderClr(4) = RGB(128, 128, 128)
        BorderClr(5) = RGB(128, 128, 128)
        BorderClr(6) = RGB(200, 200, 200)
    
        dx = mTabScrollBtnWidth + mTabScrollCount * mTabScrollTabWidth
        dy = 0
        
        UserControl.Line (dx, dy)-Step(mTabScrollBtnWidth - 3, 0), BorderClr(1)
        UserControl.Line -(dx + mTabScrollBtnWidth - 1, 2 + dy), BorderClr(2)
        UserControl.Line -Step(0, h - 5), BorderClr(3)
        UserControl.Line -Step(-2, 2), BorderClr(4)
        UserControl.Line -(dx, h - 1), BorderClr(5)
        UserControl.Line -(dx, 0), BorderClr(6)
        
        For r = 1 To h - 2
            If r = 1 Or r = h - 2 Then
                l = 2
                dw = mTabScrollBtnWidth - 4
            Else
                l = 1
                dw = mTabScrollBtnWidth - 2
            End If
            If r < h / 10 Then
                clr = RGB(255, 255, 255)
            ElseIf r < 2 * h / 10 Then
                clr = RGB(247, 247, 247)
            ElseIf r < 3 * h / 10 Then
                clr = RGB(242, 242, 242)
            ElseIf r < 4 * h / 10 Then
                clr = RGB(239, 239, 239)
            ElseIf r < 5 * h / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 6 * h / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 7 * h / 10 Then
                clr = RGB(236, 236, 236)
            ElseIf r < 8 * h / 10 Then
                clr = RGB(233, 233, 233)
            ElseIf r < 9 * h / 10 Then
                clr = RGB(230, 230, 230)
            Else
                clr = RGB(227, 227, 227)
            End If
            UserControl.Line (dx + l, r)-Step(dw, 0), clr
        Next
        
        dx = dx + mTabScrollBtnWidth / 2 - 2
        dy = UserControl.ScaleHeight / 2 - 5
        
        If mTabScrollLeftTab + mTabScrollCount >= mTabCount Or mEnabled = False Then
            clr = RGB(220, 220, 220)
        Else
            clr = RGB(0, 0, 128)
        End If
        
        UserControl.Line (dx, dy + 0)-Step(1, 0), clr
        UserControl.Line (dx, dy + 1)-Step(2, 0), clr
        UserControl.Line (dx, dy + 2)-Step(3, 0), clr
        UserControl.Line (dx, dy + 3)-Step(4, 0), clr
        UserControl.Line (dx, dy + 4)-Step(5, 0), clr
        UserControl.Line (dx, dy + 5)-Step(4, 0), clr
        UserControl.Line (dx, dy + 6)-Step(3, 0), clr
        UserControl.Line (dx, dy + 7)-Step(2, 0), clr
        UserControl.Line (dx, dy + 8)-Step(1, 0), clr
    End If
    
    For I = 0 To mTabScrollCount - 1
        If mTabScrollLeftTab + I <> mTabCurIndex Then
            dx = mTabScrollBtnWidth + I * mTabScrollTabWidth
            dy = 3
                
            BorderClr(1) = RGB(128, 128, 128)
            BorderClr(2) = RGB(128, 128, 128)
            BorderClr(3) = RGB(128, 128, 128)
            BorderClr(4) = RGB(128, 128, 128)
            BorderClr(5) = RGB(180, 180, 180)
            BorderClr(6) = RGB(128, 128, 128)
            
            UserControl.Line (dx + 2, dy)-Step(mTabScrollTabWidth - 5, 0), BorderClr(1)
            UserControl.Line -(dx + mTabScrollTabWidth - 1, 2 + dy), BorderClr(2)
            UserControl.Line -Step(0, h - 4), BorderClr(3)
            UserControl.Line -Step(-mTabScrollTabWidth + 1, 0), BorderClr(4)
            UserControl.Line -(dx, 2 + dy), BorderClr(5)
            UserControl.Line -(dx + 2, dy), BorderClr(6)
    
            For r = 1 + dy To h - 2
                If r = 1 + dy Then
                    l = 2
                    dw = mTabScrollTabWidth - 4
                Else
                    l = 1
                    dw = mTabScrollTabWidth - 2
                End If
                If r < h / 10 Then
                    clr = RGB(190, 190, 190)
                ElseIf r < 2 * h / 10 Then
                    clr = RGB(202, 202, 202)
                ElseIf r < 3 * h / 10 Then
                    clr = RGB(215, 215, 215)
                ElseIf r < 4 * h / 10 Then
                    clr = RGB(225, 225, 225)
                ElseIf r < 5 * h / 10 Then
                    clr = RGB(230, 230, 230)
                ElseIf r < 6 * h / 10 Then
                    clr = RGB(225, 225, 225)
                ElseIf r < 7 * h / 10 Then
                    clr = RGB(230, 230, 230)
                ElseIf r < 8 * h / 10 Then
                    clr = RGB(230, 230, 230)
                ElseIf r < 9 * h / 10 Then
                    clr = RGB(230, 230, 230)
                Else
                    clr = RGB(230, 230, 230)
                End If
                UserControl.Line (dx + l, r)-Step(dw, 0), clr
            Next
            
            UserControl.Line (dx, h - 5)-Step(mTabScrollTabWidth, 0), RGB(200, 200, 255)
            UserControl.Line (dx, h - 4)-Step(mTabScrollTabWidth, 0), RGB(247, 247, 247)
            UserControl.Line (dx, h - 3)-Step(mTabScrollTabWidth, 0), RGB(239, 239, 239)
            UserControl.Line (dx, h - 2)-Step(mTabScrollTabWidth, 0), RGB(200, 200, 200)
            UserControl.Line (dx, h - 1)-Step(mTabScrollTabWidth, 0), RGB(128, 128, 128)
            
            If I = 0 Then
                UserControl.Line (dx, h - 4)-Step(0, 2), RGB(200, 200, 255)
            ElseIf I = mTabScrollCount - 1 Then
                UserControl.Line (dx + mTabScrollTabWidth - 1, h - 5)-Step(0, 4), RGB(128, 128, 128)
            End If
        End If
    Next
    
    I = mTabCurIndex - mTabScrollLeftTab
    dx = mTabScrollBtnWidth + I * mTabScrollTabWidth
    dy = 0
                
    BorderClr(1) = RGB(200, 200, 255)
    BorderClr(2) = RGB(128, 128, 128)
    BorderClr(3) = RGB(128, 128, 128)
    BorderClr(4) = RGB(200, 200, 255)
    BorderClr(5) = RGB(200, 200, 255)
    BorderClr(6) = RGB(200, 200, 255)

    UserControl.Line (dx + 2, dy)-Step(mTabScrollTabWidth - 5, 0), BorderClr(1)
    UserControl.Line -(dx + mTabScrollTabWidth - 1, 2 + dy), BorderClr(2)
    UserControl.Line -Step(0, h - 4), BorderClr(3)
    UserControl.Line -Step(-mTabScrollTabWidth + 1, 0), BorderClr(4)
    UserControl.Line -(dx, 2 + dy), BorderClr(5)
    UserControl.Line -(dx + 2, dy), BorderClr(6)
    
    For r = 1 To h - 2
        If r = 1 Then
            l = 2
            dw = mTabScrollTabWidth - 4
        Else
            l = 1
            dw = mTabScrollTabWidth - 2
        End If
        If r < h / 10 Then
            clr = RGB(255, 255, 255)
        ElseIf r < 2 * h / 10 Then
            clr = RGB(247, 247, 247)
        ElseIf r < 3 * h / 10 Then
            clr = RGB(242, 242, 242)
        ElseIf r < 4 * h / 10 Then
            clr = RGB(239, 239, 239)
        ElseIf r < 5 * h / 10 Then
            clr = RGB(236, 236, 236)
        ElseIf r < 6 * h / 10 Then
            clr = RGB(236, 236, 236)
        ElseIf r < 7 * h / 10 Then
            clr = RGB(236, 236, 236)
        ElseIf r < 8 * h / 10 Then
            clr = RGB(236, 236, 236)
        ElseIf r < 9 * h / 10 Then
            clr = RGB(236, 236, 236)
        Else
            clr = RGB(236, 236, 236)
        End If
        UserControl.Line (dx + l, r)-Step(dw, 0), clr
    Next
    
    UserControl.Line (dx, h - 3)-Step(mTabScrollTabWidth, 0), RGB(239, 239, 239)
    UserControl.Line (dx, h - 2)-Step(mTabScrollTabWidth, 0), RGB(200, 200, 200)
    UserControl.Line (dx, h - 1)-Step(mTabScrollTabWidth, 0), RGB(128, 128, 128)
    
    If I = 0 Then
        UserControl.Line (mTabScrollBtnWidth, h - 4)-Step(0, 2), RGB(200, 200, 255)
    End If
    If I = mTabScrollCount - 1 Then
        UserControl.Line (dx + mTabScrollTabWidth - 1, h - 5)-Step(0, 4), RGB(128, 128, 128)
    End If
    
    UserControl.AutoRedraw = AutoRedraw0
    UserControl.Refresh
End Sub

Public Property Get TabCurIndex() As Integer
    TabCurIndex = mTabCurIndex
End Property

Public Property Let TabCurIndex(ByVal NewValue As Integer)
    On Error Resume Next
    
    'LblCaption(mTabCurIndex).FontUnderline = False
    mTabCurIndex = NewValue
    
    If mTabCurIndex - mTabScrollLeftTab >= mTabScrollCount Then
        mTabScrollLeftTab = mTabCurIndex - mTabScrollCount + 1
    ElseIf mTabCurIndex < mTabScrollLeftTab Then
        mTabScrollLeftTab = mTabCurIndex
    End If

    ShowTabStrip
    'LblCaption(mTabCurIndex).FontUnderline = True
    PropertyChanged "TabCurIndex"
End Property

Public Property Get TabMinWidth() As Integer
    TabMinWidth = mTabMinWidth
End Property

Public Property Let TabMinWidth(ByVal NewValue As Integer)
    mTabMinWidth = NewValue
    
    ShowTabStrip
    PropertyChanged "TabMinWidth"
End Property

Public Property Get TabCurPicture() As IPictureDisp
    On Error Resume Next
    
    Set TabCurPicture = mPicture(mTabCurIndex)
End Property

Public Property Set TabCurPicture(NewPicture As IPictureDisp)
    Set mPicture(mTabCurIndex) = NewPicture
    
    ShowTabStrip
    PropertyChanged "Picture" + Trim(Str(mTabCurIndex))
End Property

Public Property Get TabCurCaption() As String
    On Error Resume Next
    
    TabCurCaption = mCaption(mTabCurIndex)
End Property

Public Property Let TabCurCaption(NewCaption As String)
    mCaption(mTabCurIndex) = NewCaption
    
    ShowTabStrip
    PropertyChanged "Caption" + Trim(Str(mTabCurIndex))
End Property

Public Property Get TabFont() As font
    On Error Resume Next
    
    Set TabFont = UserControl.font
End Property

Public Property Set TabFont(newfont As font)
    Set UserControl.font = newfont
    
    ShowTabStrip
    PropertyChanged "TabFont"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(NewValue As Boolean)
    UserControl.Enabled = NewValue
    mEnabled = NewValue
    
    ShowTabStrip
    PropertyChanged "Enabled"
End Property

Sub MouseMoveTabStrip(ByVal EnterStatus As Integer, ByVal CurIndex As Integer)
    Static s As Integer, idx As Integer
    
    If s <> EnterStatus Or idx <> CurIndex Then
        
        UserControl.DrawMode = 7
        'UserControl.DrawWidth = 2
        
        If s = EnterStatus Or EnterStatus = 0 Then
            If idx >= 0 And idx < mTabScrollCount Then
                UserControl.DrawWidth = 1
                If idx + mTabScrollLeftTab = mTabCurIndex Then
                    UserControl.Line (mTabScrollBtnWidth + idx * mTabScrollTabWidth + 2, 1)-Step(mTabScrollTabWidth - 4, 0), RGB(20, 60, 160)
                    UserControl.Line (mTabScrollBtnWidth + idx * mTabScrollTabWidth + 1, 2)-Step(mTabScrollTabWidth - 2, 0), RGB(20, 60, 160)
                Else
                    UserControl.Line (mTabScrollBtnWidth + idx * mTabScrollTabWidth + 2, 4)-Step(mTabScrollTabWidth - 4, 0), RGB(40, 120, 200)
                    UserControl.Line (mTabScrollBtnWidth + idx * mTabScrollTabWidth + 1, 5)-Step(mTabScrollTabWidth - 2, 0), RGB(40, 120, 200)
                End If
            ElseIf idx = -1 Then
                UserControl.DrawWidth = 2
                UserControl.Line (2, 2)-Step(mTabScrollBtnWidth - 4, UserControl.ScaleHeight - 4), RGB(20, 60, 160), B
            Else
                UserControl.DrawWidth = 2
                UserControl.Line (mTabScrollBtnWidth + idx * mTabScrollTabWidth + 2, 2)-Step(mTabScrollBtnWidth - 4, UserControl.ScaleHeight - 4), RGB(20, 60, 160), B
            End If
        End If
        
        If EnterStatus = 1 Then
            If CurIndex >= 0 And CurIndex < mTabScrollCount Then
                UserControl.DrawWidth = 1
                If CurIndex + mTabScrollLeftTab = mTabCurIndex Then
                    UserControl.Line (mTabScrollBtnWidth + CurIndex * mTabScrollTabWidth + 2, 1)-Step(mTabScrollTabWidth - 4, 0), RGB(20, 60, 160)
                    UserControl.Line (mTabScrollBtnWidth + CurIndex * mTabScrollTabWidth + 1, 2)-Step(mTabScrollTabWidth - 2, 0), RGB(20, 60, 160)
                Else
                    UserControl.Line (mTabScrollBtnWidth + CurIndex * mTabScrollTabWidth + 2, 4)-Step(mTabScrollTabWidth - 4, 0), RGB(40, 120, 200)
                    UserControl.Line (mTabScrollBtnWidth + CurIndex * mTabScrollTabWidth + 1, 5)-Step(mTabScrollTabWidth - 2, 0), RGB(40, 120, 200)
                End If
            ElseIf CurIndex = -1 Then
                UserControl.DrawWidth = 2
                UserControl.Line (2, 2)-Step(mTabScrollBtnWidth - 4, UserControl.ScaleHeight - 4), RGB(20, 60, 160), B
            Else
                UserControl.DrawWidth = 2
                UserControl.Line (mTabScrollBtnWidth + CurIndex * mTabScrollTabWidth + 2, 2)-Step(mTabScrollBtnWidth - 4, UserControl.ScaleHeight - 4), RGB(20, 60, 160), B
            End If
        End If
        
        s = EnterStatus
        idx = CurIndex
    End If
End Sub

Sub ShowTabStripFocus()
    Dim d As Integer
    Static s As Boolean
    
    d = 3
    
    If s <> mGotFocus Then
        s = mGotFocus
        UserControl.DrawMode = 7
        UserControl.DrawWidth = d
        
        UserControl.Line (d, UserControl.ScaleHeight - d)-(UserControl.ScaleWidth - d, UserControl.ScaleHeight - d), RGB(160, 60, 20)
    End If
End Sub

