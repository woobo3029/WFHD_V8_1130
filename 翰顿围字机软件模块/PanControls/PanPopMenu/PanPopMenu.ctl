VERSION 5.00
Begin VB.UserControl PanPopMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
   Begin VB.VScrollBar VScroll1 
      Height          =   2175
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   780
      Top             =   1200
   End
   Begin VB.Image ImgIcon 
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   420
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "PanPopMenu"
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

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private mMenuIndex As Integer
Private mParentMenuIndex As Integer
Private mItemCount As Integer
Private mItemCaption() As String
Private mItemHotKey() As String
Private mItemIconPath() As String
Private mItemTop() As Integer
Private mItemCurIndex As Integer
Private mItemMaxWidth As Integer
Private mLineH As Integer, mLineH2 As Integer
Private mCurItem As Integer
Private mCurItemTop As Integer
'Private mEnabled As Boolean

Private mHasVScrollBar As Boolean
Private mTopItemIndex As Integer
Private mItemsShow As Integer
Private mDisableTimer As Boolean

Public Event MouseOnItem(MenuIndex As Integer, ItemIndex As Integer, ItemTop As Integer)
Public Event ClickOnItem(MenuIndex As Integer, ItemIndex As Integer)
Public Event SubMenuSelected(MenuIndex As Integer, Param As Integer, SelectMode As Integer)

Private Sub PicTemp_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift
End Sub

Private Sub Timer1_Timer()
    Dim LButtonState As Long
    
    LButtonState = GetKeyState(VK_LBUTTON)
    
    If DisablePopMenuTimer = True Then
        Exit Sub
    End If
    
    If (LButtonState = -127 Or LButtonState = -128) Then
        Timer1.Enabled = False
                
        'ShowWindow UserControl.hwnd, SW_HIDE
        UserControl.Width = Screen.TwipsPerPixelX
        UserControl.Height = Screen.TwipsPerPixelY
            
        If mMenuIndex > 0 Then
            DestroyWindow UserControl.hWnd
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    mParentMenuIndex = -1
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer, Y As Single, k As Integer
    
    On Error Resume Next
    
    If KeyCode = 13 Then
        Y = mItemTop(mCurItem) + 8
        UserControl_MouseDown 0, 0, 30, Y
        
    ElseIf KeyCode = 37 Then
        If mParentMenuIndex > -1 Then
            ShowMenu
            RaiseEvent SubMenuSelected(mMenuIndex, mParentMenuIndex, 2)
        End If
        
    ElseIf KeyCode = 38 Then
        k = 0
        I = mCurItem
        Do
            If I > 0 Then
                I = I - 1
            ElseIf mTopItemIndex > 0 Then
                mTopItemIndex = mTopItemIndex - 1
                VScroll1.Value = mTopItemIndex
                MoveMenuItem mTopItemIndex
            Else
                I = mItemsShow - 1
            End If
            
            k = k + 1
            If k > mItemCount Then
                Exit Do
            End If
        Loop Until mItemCaption(I + mTopItemIndex) <> "-" And Left(mItemCaption(I + mTopItemIndex), 1) <> "@"
        
        Y = mItemTop(I) + 8
        UserControl_MouseMove 0, -1, 30, Y
        UserControl.SetFocus
        
    ElseIf KeyCode = 39 Then
        RaiseEvent SubMenuSelected(mMenuIndex, mCurItem + mTopItemIndex, 1)
        
    ElseIf KeyCode = 40 Then
        k = 0
        I = mCurItem
        Do
            If I < mItemsShow - 1 Then
                I = I + 1
            ElseIf mTopItemIndex < VScroll1.Max Then
                mTopItemIndex = mTopItemIndex + 1
                VScroll1.Value = mTopItemIndex
                MoveMenuItem mTopItemIndex
            Else
                I = 0
            End If
            
            k = k + 1
            If k > mItemCount Then
                Exit Do
            End If
        Loop Until mItemCaption(I + mTopItemIndex) <> "-" And Left(mItemCaption(I + mTopItemIndex), 1) <> "@"
        
        Y = mItemTop(I) + 8
        UserControl_MouseMove 0, -1, 30, Y
        UserControl.SetFocus
        
    ElseIf KeyCode = 27 Then
        RaiseEvent SubMenuSelected(mMenuIndex, mCurItem + mTopItemIndex, 3)
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mCurItem >= 0 Then
        RaiseEvent ClickOnItem(mMenuIndex, mCurItem + mTopItemIndex)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static k As Integer, t As Integer, x0 As Single, y0 As Single, shift0 As Integer
    
    On Error Resume Next
    
    If PicTemp.Visible = True Then 'for the vscroll1 LostFocus event
        PicTemp.SetFocus
    End If
    
    If Button <> 0 Then Exit Sub
    
    If X = x0 And Y = y0 Then
        Exit Sub
    End If
    x0 = X
    y0 = Y
    
    MainFormUserEventRaised = True
    
    If shift0 = -1 And Shift = 0 Then
        shift0 = Shift
        Exit Sub
    End If
    shift0 = Shift

    If t = 0 Or mCurItem = -2 Then
        k = -9
        t = 1
    End If
    
    PointMenu X, Y

    If k <> mCurItem Then
        k = mCurItem
        RaiseEvent MouseOnItem(mMenuIndex, mCurItem, mCurItemTop)
    End If
    
    'DisablePopMenuTimer = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim I As Integer
    
    mItemCount = PropBag.ReadProperty("ItemCount", 0)
    If mItemCount > 0 Then
        ReDim mItemCaption(mItemCount), mItemHotKey(mItemCount), mItemIconPath(mItemCount), mItemTop(mItemCount)
        For I = 0 To mItemCount - 1
            If I > 0 Then
                Load ImgIcon(I)
            End If
            mItemCaption(I) = PropBag.ReadProperty("ItemCaption" & Trim(Str(I)), "")
            mItemHotKey(I) = PropBag.ReadProperty("ItemHotKey" & Trim(Str(I)), "")
            mItemIconPath(I) = PropBag.ReadProperty("ItemIconPath" & Trim(Str(I)), "")
            If mItemIconPath(I) <> "" Then
                Set ImgIcon(I).Picture = LoadPicture(mItemIconPath(I))
            Else
                Set ImgIcon(I).Picture = Nothing
            End If
        Next
    End If
    
    'ShowMenu
End Sub

Private Sub UserControl_Terminate()
    Dim I As Integer
    
    For I = 0 To mItemCount - 1
        Set ImgIcon(I).Picture = Nothing
        If I > 0 Then
            Unload ImgIcon(I)
        End If
    Next
    ReDim mItemCaption(0), mItemHotKey(0), mItemIconPath(0), mItemTop(0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim I As Integer
    
    PropBag.WriteProperty "ItemCount", mItemCount, 0
    If mItemCount > 0 Then
        For I = 0 To mItemCount - 1
            PropBag.WriteProperty "ItemCaption" & Trim(Str(I)), mItemCaption(I), ""
            PropBag.WriteProperty "ItemHotKey" & Trim(Str(I)), mItemHotKey(I), ""
            PropBag.WriteProperty "ItemIconPath" & Trim(Str(I)), mItemIconPath(I), ""
        Next
    End If
End Sub

Public Property Get MenuIndex() As Integer
    MenuIndex = mMenuIndex
End Property

Public Property Let MenuIndex(ByVal NewValue As Integer)
    mMenuIndex = NewValue
    
    PropertyChanged "MenuIndex"
End Property

Public Property Get ParentMenuIndex() As Integer
    ParentMenuIndex = mParentMenuIndex
End Property

Public Property Let ParentMenuIndex(ByVal NewValue As Integer)
    mParentMenuIndex = NewValue
    
    PropertyChanged "ParentMenuIndex"
End Property

Public Property Get ItemCount() As Integer
    ItemCount = mItemCount
End Property

Public Property Let ItemCount(ByVal NewValue As Integer)
    Dim I As Integer
    
    mItemCount = NewValue
    
    ReDim Preserve mItemCaption(mItemCount), mItemHotKey(mItemCount), mItemIconPath(mItemCount), mItemTop(mItemCount)
    For I = 0 To ImgIcon.Count - 1
        Set ImgIcon(I).Picture = Nothing
        If I > 0 Then
            Unload ImgIcon(I)
        End If
    Next
    
    If NewValue > 0 Then
        For I = ImgIcon.Count To NewValue - 1
            Load ImgIcon(I)
        Next
    End If
    'ShowMenu
    PropertyChanged "ItemCount"
End Property

Public Property Get ItemCurIndex() As Integer
    ItemCurIndex = mItemCurIndex
End Property

Public Property Let ItemCurIndex(ByVal NewValue As Integer)
    If NewValue < mItemCount Then
        mItemCurIndex = NewValue
        
        'ShowMenu
        PropertyChanged "ItemCurIndex"
    End If
End Property

Public Property Get ItemCaption() As String
    If mItemCurIndex < 0 Or mItemCurIndex >= mItemCount Then
        Exit Sub
    End If
    
    ItemCaption = mItemCaption(mItemCurIndex)
End Property

Public Property Let ItemCaption(ByVal NewValue As String)
    mItemCaption(mItemCurIndex) = NewValue
    
    'ShowMenu
    PropertyChanged "ItemCaption"
End Property

Public Property Get ItemHotKey() As String
    If mItemCurIndex < 0 Or mItemCurIndex >= mItemCount Then
        Exit Property
    End If
    
    ItemHotKey = mItemHotKey(mItemCurIndex)
End Property

Public Property Let ItemHotKey(ByVal NewValue As String)
    mItemHotKey(mItemCurIndex) = NewValue
    
    'ShowMenu
    PropertyChanged "ItemHotKey"
End Property

Public Property Get ItemIconPath() As String
    If mItemCurIndex >= mItemCount Then
        Exit Property
    End If
    
    ItemIconPath = mItemIconPath(mItemCurIndex)
End Property

Public Property Let ItemIconPath(ByVal NewValue As String)
    mItemIconPath(mItemCurIndex) = NewValue
    If mItemIconPath(mItemCurIndex) <> "" Then
        Set ImgIcon(mItemCurIndex).Picture = LoadPicture(mItemIconPath(mItemCurIndex))
    Else
        Set ImgIcon(mItemCurIndex).Picture = Nothing
    End If
    
    'ShowMenu
    PropertyChanged "ItemIconPath"
End Property

Public Property Get ItemTop() As Integer
    If mItemCurIndex >= mItemCount Then
        Exit Property
    End If
    
    ItemTop = mCurItemTop
End Property

'Public Property Let ItemTop(ByVal NewValue As String)
'    mItemTop(mItemCurIndex) = NewValue
'
'    'ShowMenu
'    PropertyChanged "ItemTop"
'End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Sub PopUp(Optional MaxItems As Integer = 0)
    'UserControl.BackStyle = 1
    'UserControl.Refresh
    On Error Resume Next
    
    Timer1.Enabled = True
    
    ShowWindow UserControl.hWnd, SW_SHOW
    ShowMenu MaxItems
    UserControl.SetFocus

    Timer1.Enabled = True
End Sub

Private Sub ShowMenu(Optional MaxItems As Integer = 0)
    Dim n As Integer, w As Integer, h As Integer, icon_w As Integer, LineMaxW As Integer
    Dim I As Integer, k As Integer, s As String
    
    On Error Resume Next
    
    n = mItemCount
    k = 0
    For I = 0 To n - 1
        If mItemIconPath(I) <> "" Then
            k = 1
        End If
    Next
    If k = 1 Then
        icon_w = 16
    Else
        icon_w = 0
    End If
    
    With UserControl
        .FontName = "Times New Roman" '"·ÂËÎ_GB2312"
        .FontSize = 10 '11
       
        .FontBold = True
        LineMaxW = 0
        For I = 0 To n - 1
            If LineMaxW < MixTextWidth(mItemCaption(I)) + 10 + MixTextWidth(mItemHotKey(I)) Then
                LineMaxW = MixTextWidth(mItemCaption(I)) + 10 + MixTextWidth(mItemHotKey(I))
            End If
        Next
        .FontBold = False
        
        mLineH = 19
        w = 15 + icon_w + 8 + LineMaxW + 10
        If n > 20 Then
            LineMaxW = IIf(LineMaxW < 60, 60, LineMaxW)
            w = 15 + icon_w + 8 + LineMaxW + 10
        End If
       
        If MaxItems = 0 Or n <= MaxItems Then
            mHasVScrollBar = False
            mLineH2 = mLineH / 2
            mItemsShow = n
            
            k = 0
            For I = 0 To n - 1
                If mItemCaption(I) = "-" Then k = k + 1
            Next
            h = (n - k) * mLineH + k * mLineH2 + 12
            If n = 0 Then
                w = 60
                h = 26
            End If
        Else
            mHasVScrollBar = True
            mLineH2 = mLineH
            mItemsShow = MaxItems
            
            h = MaxItems * mLineH + 12
            w = w + 18
        End If
        mTopItemIndex = 0
        
        .Width = w * Screen.TwipsPerPixelX
        .Height = h * Screen.TwipsPerPixelY
        .AutoRedraw = True
        .BackColor = RGB(220, 220, 220)
        .Cls
        DrawFrame 0, 0, w, h, RGB(230, 230, 230), RGB(0, 0, 0)
        DrawFrame 1, 1, w - 2, h - 2, RGB(255, 255, 255), RGB(120, 120, 120)
        
        If mHasVScrollBar = True Then
            VScroll1.Top = 6
            VScroll1.Left = w - 20
            VScroll1.Height = h - 12
            VScroll1.Min = 0
            VScroll1.Max = n - MaxItems
            VScroll1.Value = 0
            VScroll1.Visible = True
            PicTemp.Move VScroll1.Left, VScroll1.Top
            PicTemp.Visible = True
        Else
            VScroll1.Visible = False
            VScroll1.Min = 0
            VScroll1.Max = 0
            PicTemp.Visible = False
        End If
       
        .CurrentY = 5
        For I = 0 To mItemsShow - 1
            mItemTop(I) = .CurrentY
            If mItemCaption(I) = "-" Then
                .CurrentY = .CurrentY + mLineH2
            Else
                .CurrentY = .CurrentY + mLineH
            End If
        Next
       
        For I = 0 To mItemsShow - 1
            If mItemCaption(I) = "-" Then
                UserControl.Line (6, mItemTop(I) + Int(mLineH2 / 2) + 1)-Step(w - 12, 0), RGB(255, 255, 255)
                UserControl.Line (6, mItemTop(I) + Int(mLineH2 / 2))-Step(w - 12, 0), RGB(100, 100, 100)
                
            ElseIf Left(mItemCaption(I), 1) = "@" Then
            
                If mItemIconPath(I) <> "" Then
                    ImgIcon(I).Move 8, mItemTop(I) + 2
                    ImgIcon(I).Visible = True
                End If
                
                .CurrentX = 10 + icon_w + 8 + 1
                .CurrentY = mItemTop(I) + 2
                PrintMix Mid$(mItemCaption(I), 2), RGB(255, 255, 255)
                
                .CurrentX = 10 + icon_w + 8
                .CurrentY = mItemTop(I) + 1
                PrintMix Mid$(mItemCaption(I), 2), RGB(120, 120, 120)
                
                If mItemHotKey(I) <> "" Then
                    .CurrentX = w - MixTextWidth(mItemHotKey(I)) - 11
                    .CurrentY = mItemTop(I) + 2
                    PrintMix mItemHotKey(I), RGB(255, 255, 255)
                    
                    .CurrentX = w - MixTextWidth(mItemHotKey(I)) - 12
                    .CurrentY = mItemTop(I) + 1
                    PrintMix mItemHotKey(I), RGB(120, 120, 120)
                End If
                
            ElseIf Right(mItemCaption(I), 1) = "&" Then
            
                If mItemIconPath(I) <> "" Then
                    ImgIcon(I).Move 8, mItemTop(I) + 2
                    ImgIcon(I).Visible = True
                End If
                
                s = mItemCaption(I)
                s = Mid(s, 1, Len(s) - 1)
                .CurrentX = 10 + icon_w + 8
                .CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(0, 0, 0)
                
                If mItemHotKey(I) <> "" Then
                    .CurrentX = w - MixTextWidth(mItemHotKey(I)) - 12
                    .CurrentY = mItemTop(I) + 1
                    PrintMix mItemHotKey(I), RGB(0, 0, 0)
                End If
                
                UserControl.Line (w - 12, mItemTop(I) + 6)-Step(1, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 7)-Step(2, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 8)-Step(3, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 9)-Step(4, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 10)-Step(5, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 11)-Step(4, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 12)-Step(3, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 13)-Step(2, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 14)-Step(1, 0), RGB(0, 0, 0)
                
            Else
                If mItemIconPath(I) <> "" Then
                    ImgIcon(I).Move 8, mItemTop(I) + 2
                    ImgIcon(I).Visible = True
                End If
                
                .CurrentX = 10 + icon_w + 8
                .CurrentY = mItemTop(I) + 1
                PrintMix mItemCaption(I), RGB(0, 0, 0)
                
                If mItemHotKey(I) <> "" Then
                    .CurrentX = w - MixTextWidth(mItemHotKey(I)) - 12
                    .CurrentY = mItemTop(I) + 1
                    PrintMix mItemHotKey(I), RGB(0, 0, 0)
                End If
                
            End If
        Next
        
        .Refresh
    End With
    
    mCurItem = -1
End Sub

Private Function MixTextWidth(ByVal txt As String) As Integer
    Dim I As Integer, c As String, Left As Integer
    Dim font As String, font0 As String
    Dim w As Integer
    
    On Error Resume Next
    
    font0 = UserControl.FontName
    font = font0
    For I = 1 To Len(txt)
        c = Mid$(txt, I, 1)
        If Asc(c) < 0 Then
           If font <> "MS Sans Serif" Then
               font = "MS Sans Serif"
               UserControl.FontName = font
           End If
          w = w + UserControl.TextWidth(c)
        Else
           If font <> "Times New Roman" Then
               font = "Times New Roman"
               UserControl.FontName = font
           End If
           w = w + UserControl.TextWidth(c)
        End If
    Next
    UserControl.FontName = font0
    MixTextWidth = w
End Function

Private Sub PrintMix(ByVal txt As String, ByVal color As Long, Optional w As Single = 0, Optional h As Single = 0)
    Dim I As Integer, c As String, Left As Integer
    Dim font As String
    Dim Y As Long, d As Long
    
    On Error Resume Next
    
    Left = UserControl.CurrentX
    Y = UserControl.CurrentY
    UserControl.ForeColor = color
    font = UserControl.FontName
    For I = 1 To Len(txt)
        If h > 0 And Y > h - UserControl.TextHeight(txt) * 1.1 Then Exit For
        c = Mid$(txt, I, 1)
        If Asc(c) < 0 Then
            If font <> "MS Sans Serif" Then
                font = "MS Sans Serif"
                UserControl.FontName = font
            End If
            d = 0
        Else
            If font <> "Times New Roman" Then
                font = "Times New Roman"
                UserControl.FontName = font
            End If
            d = 1
        End If
        If w > 0 Then
           If UserControl.CurrentX > w - Left - UserControl.TextWidth(c) * 1.5 And c <> ")" Then
              UserControl.CurrentX = Left
              Y = Y + UserControl.TextHeight(txt) * 1.3
           End If
        End If
        UserControl.CurrentY = Y + d
        UserControl.Print c;
    Next I
    UserControl.CurrentY = Y
    UserControl.Print
End Sub

Sub DrawFrame(ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal c As Long, ByVal b As Long)
    Dim Y As Integer

    UserControl.Line (Left, Top)-Step(Width, 0), c
    UserControl.Line (Left, Top)-Step(0, Height), c
    
    Y = Top + Height - 1
    UserControl.Line (Left + Width - 1, Y)-Step(0, -Height), b
    UserControl.Line (Left + Width - 1, Y)-Step(-Width, 0), b
    UserControl.Line (Left, Y)-Step(Width, 0), b
End Sub

Public Sub PointMenu(ByVal X As Single, ByVal Y As Single)
    Dim I As Integer, k As Integer, b As Boolean, icon_w As Integer, w As Integer, s As String
    
    On Error Resume Next
    
    k = 0
    For I = 0 To mItemCount - 1
        If mItemIconPath(I) <> "" Then
            k = 1
        End If
    Next
    If k = 1 Then
        icon_w = 16
    Else
        icon_w = 0
    End If
    
    mCurItemTop = -1
    k = 0
    For I = 0 To mItemsShow - 1
        If mItemCaption(I + mTopItemIndex) = "-" Or Left(mItemCaption(I + mTopItemIndex), 1) = "@" Then
            If Y > mItemTop(I) And Y <= mItemTop(I) + mLineH2 Then
                mCurItemTop = mItemTop(I)
                Exit For
            End If
        Else
            If Y > mItemTop(I) And Y <= mItemTop(I) + mLineH Then
                mCurItemTop = mItemTop(I)
                k = 1
                Exit For
            End If
        End If
    Next I
    
    If I = mCurItem Then
        Exit Sub
    End If
    
    With UserControl
        If mCurItem >= 0 And mCurItem < mItemsShow Then
            UserControl.Line (4, mItemTop(mCurItem))-Step(.ScaleWidth - 8, mLineH), RGB(220, 220, 220), BF
            s = mItemCaption(mCurItem + mTopItemIndex)
            If Right(s, 1) = "&" Then
                s = Mid(s, 1, Len(s) - 1)
                
                w = .ScaleWidth
                UserControl.Line (w - 12, mItemTop(mCurItem) + 6)-Step(1, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 7)-Step(2, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 8)-Step(3, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 9)-Step(4, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 10)-Step(5, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 11)-Step(4, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 12)-Step(3, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 13)-Step(2, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(mCurItem) + 14)-Step(1, 0), RGB(0, 0, 0)
            End If
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(mCurItem) + 1
            UserControl.FontBold = False
            PrintMix s, RGB(0, 0, 0)
            
            s = mItemHotKey(mCurItem + mTopItemIndex)
            If s <> "" Then
                w = .ScaleWidth
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(mCurItem) + 1
                PrintMix s, RGB(0, 0, 0)
            End If
        End If
        mCurItem = -1
        If k = 1 Then
            UserControl.Line (4, mItemTop(I))-Step(.ScaleWidth - 8, mLineH), RGB(120, 120, 220), BF
            UserControl.Line (4, mItemTop(I) + mLineH / 2)-Step(.ScaleWidth - 8, 5), RGB(118, 118, 220), BF
            UserControl.Line (4, mItemTop(I) + mLineH / 2 + 1)-Step(.ScaleWidth - 8, 3), RGB(116, 116, 220), BF
            UserControl.Line (4, mItemTop(I) + mLineH / 2 + 2)-Step(.ScaleWidth - 8, 1), RGB(113, 113, 220), BF
            
            UserControl.Line (4, mItemTop(I))-Step(.ScaleWidth - 8, 0), RGB(100, 100, 100), BF
            UserControl.Line (4, mItemTop(I))-Step(0, mLineH), RGB(100, 100, 100), BF
            UserControl.Line (4, mItemTop(I) + mLineH)-Step(.ScaleWidth - 8, 0), RGB(255, 255, 255), BF
            UserControl.Line -Step(0, -mLineH), RGB(255, 255, 255), BF
            
            s = mItemCaption(I + mTopItemIndex)
            If Right(s, 1) = "&" Then
                s = Mid(s, 1, Len(s) - 1)
                
                'w = UserControl.Width / Screen.TwipsPerPixelX
                w = .ScaleWidth
                UserControl.Line (w - 12, mItemTop(I) + 6)-Step(1, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 7)-Step(2, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 8)-Step(3, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 9)-Step(4, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 10)-Step(5, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 11)-Step(4, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 12)-Step(3, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 13)-Step(2, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 14)-Step(1, 0), RGB(255, 255, 255)
                
                'RaiseEvent MouseOnItem(mMenuIndex, I, mItemTop(I))
            End If
            
            UserControl.FontBold = True
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(I)
            PrintMix s, RGB(0, 0, 0)
            
            UserControl.CurrentX = 10 + icon_w + 8 - 1
            UserControl.CurrentY = mItemTop(I) + 1
            PrintMix s, RGB(0, 0, 0)
            
            UserControl.CurrentX = 10 + icon_w + 8 + 1
            UserControl.CurrentY = mItemTop(I) + 1
            PrintMix s, RGB(0, 0, 0)
            
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(I) + 2
            PrintMix s, RGB(0, 0, 0)
                    
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(I) + 1
            PrintMix s, RGB(255, 255, 255)
            
            s = mItemHotKey(I + mTopItemIndex)
            If s <> "" Then
                w = .ScaleWidth
                UserControl.FontBold = True
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(I)
                PrintMix s, RGB(0, 0, 0)
                
                UserControl.CurrentX = w - MixTextWidth(s) - 12 - 1
                UserControl.CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(0, 0, 0)
                
                UserControl.CurrentX = w - MixTextWidth(s) - 12 + 1
                UserControl.CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(0, 0, 0)
                
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(I) + 2
                PrintMix s, RGB(0, 0, 0)
                        
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(255, 255, 255)
            End If
            
            mCurItem = I
            
            UserControl.FontBold = False
        End If
        
        .Refresh
    End With
End Sub

Private Sub MoveMenuItem(ByVal VScrollValue As Integer)
    Dim n As Integer, w As Integer, h As Integer, icon_w As Integer, LineMaxW As Integer
    Dim I As Integer, k As Integer, s As String
    Static mTopItemIndex0 As Integer
    
    On Error Resume Next
    
    mTopItemIndex = VScrollValue
    
    n = mItemCount
    k = 0
    For I = 0 To n - 1
        If mItemIconPath(I) <> "" Then
            k = 1
        End If
    Next
    If k = 1 Then
        icon_w = 16
    Else
        icon_w = 0
    End If
    
    With UserControl
        .Cls
        w = .Width / tx
        h = .Height / ty
        DrawFrame 0, 0, w, h, RGB(230, 230, 230), RGB(0, 0, 0)
        DrawFrame 1, 1, w - 2, h - 2, RGB(255, 255, 255), RGB(120, 120, 120)
               
        .CurrentY = 5
        For I = 0 To mItemsShow - 1
            mItemTop(I) = .CurrentY
            If mItemCaption(I + mTopItemIndex) = "-" Then
                .CurrentY = .CurrentY + mLineH2
            Else
                .CurrentY = .CurrentY + mLineH
            End If
        Next
       
        For I = 0 To mItemsShow - 1
            If mItemCaption(I + mTopItemIndex) = "-" Then
                UserControl.Line (6, mItemTop(I) + Int(mLineH2 / 2) + 1)-Step(w - 12, 0), RGB(255, 255, 255)
                UserControl.Line (6, mItemTop(I) + Int(mLineH2 / 2))-Step(w - 12, 0), RGB(100, 100, 100)
                
            ElseIf Left(mItemCaption(I + mTopItemIndex), 1) = "@" Then
            
                If mItemIconPath(I + mTopItemIndex) <> "" Then
                    ImgIcon(I + mTopItemIndex).Move 8, mItemTop(I) + 2
                    ImgIcon(I + mTopItemIndex).Visible = True
                End If
                
                .CurrentX = 10 + icon_w + 8 + 1
                .CurrentY = mItemTop(I) + 2
                PrintMix Mid$(mItemCaption(I + mTopItemIndex), 2), RGB(255, 255, 255)
                
                .CurrentX = 10 + icon_w + 8
                .CurrentY = mItemTop(I) + 1
                PrintMix Mid$(mItemCaption(I + mTopItemIndex), 2), RGB(120, 120, 120)
                
                If mItemHotKey(I + mTopItemIndex) <> "" Then
                    .CurrentX = w - MixTextWidth(mItemHotKey(I + mTopItemIndex)) - 11
                    .CurrentY = mItemTop(I) + 2
                    PrintMix mItemHotKey(I + mTopItemIndex), RGB(255, 255, 255)
                    
                    .CurrentX = w - MixTextWidth(mItemHotKey(I + mTopItemIndex)) - 12
                    .CurrentY = mItemTop(I) + 1
                    PrintMix mItemHotKey(I + mTopItemIndex), RGB(120, 120, 120)
                End If
                
            ElseIf Right(mItemCaption(I + mTopItemIndex), 1) = "&" Then
            
                If mItemIconPath(I + mTopItemIndex) <> "" Then
                    ImgIcon(I).Move 8, mItemTop(I) + 2
                    ImgIcon(I).Visible = True
                End If
                
                s = mItemCaption(I + mTopItemIndex)
                s = Mid(s, 1, Len(s) - 1)
                .CurrentX = 10 + icon_w + 8
                .CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(0, 0, 0)
                
                If mItemHotKey(I + mTopItemIndex) <> "" Then
                    .CurrentX = w - MixTextWidth(mItemHotKey(I)) - 12
                    .CurrentY = mItemTop(I) + 1
                    PrintMix mItemHotKey(I + mTopItemIndex), RGB(0, 0, 0)
                End If
                
                UserControl.Line (w - 12, mItemTop(I) + 6)-Step(1, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 7)-Step(2, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 8)-Step(3, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 9)-Step(4, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 10)-Step(5, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 11)-Step(4, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 12)-Step(3, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 13)-Step(2, 0), RGB(0, 0, 0)
                UserControl.Line (w - 12, mItemTop(I) + 14)-Step(1, 0), RGB(0, 0, 0)
                
            Else
                If mItemIconPath(I + mTopItemIndex) <> "" Then
                    ImgIcon(I + mTopItemIndex).Move 8, mItemTop(I) + 2
                    ImgIcon(I + mTopItemIndex).Visible = True
                End If
                
                .CurrentX = 10 + icon_w + 8
                .CurrentY = mItemTop(I) + 1
                PrintMix mItemCaption(I + mTopItemIndex), RGB(0, 0, 0)
                
                If mItemHotKey(I) <> "" Then
                    .CurrentX = w - MixTextWidth(mItemHotKey(I)) - 12
                    .CurrentY = mItemTop(I) + 1
                    PrintMix mItemHotKey(I + mTopItemIndex), RGB(0, 0, 0)
                End If
                
            End If
        Next
        
        I = mCurItem - (mTopItemIndex - mTopItemIndex0)
        If I >= 0 And I < mItemsShow Then
            UserControl.Line (4, mItemTop(I))-Step(.ScaleWidth - 8, mLineH), RGB(120, 120, 220), BF
            UserControl.Line (4, mItemTop(I) + mLineH / 2)-Step(.ScaleWidth - 8, 5), RGB(118, 118, 220), BF
            UserControl.Line (4, mItemTop(I) + mLineH / 2 + 1)-Step(.ScaleWidth - 8, 3), RGB(116, 116, 220), BF
            UserControl.Line (4, mItemTop(I) + mLineH / 2 + 2)-Step(.ScaleWidth - 8, 1), RGB(113, 113, 220), BF
            
            UserControl.Line (4, mItemTop(I))-Step(.ScaleWidth - 8, 0), RGB(100, 100, 100), BF
            UserControl.Line (4, mItemTop(I))-Step(0, mLineH), RGB(100, 100, 100), BF
            UserControl.Line (4, mItemTop(I) + mLineH)-Step(.ScaleWidth - 8, 0), RGB(255, 255, 255), BF
            UserControl.Line -Step(0, -mLineH), RGB(255, 255, 255), BF
            
            s = mItemCaption(I + mTopItemIndex)
            If Right(s, 1) = "&" Then
                s = Mid(s, 1, Len(s) - 1)
                
                'w = UserControl.Width / Screen.TwipsPerPixelX
                w = .ScaleWidth
                UserControl.Line (w - 12, mItemTop(I) + 6)-Step(1, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 7)-Step(2, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 8)-Step(3, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 9)-Step(4, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 10)-Step(5, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 11)-Step(4, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 12)-Step(3, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 13)-Step(2, 0), RGB(255, 255, 255)
                UserControl.Line (w - 12, mItemTop(I) + 14)-Step(1, 0), RGB(255, 255, 255)
                
                'RaiseEvent MouseOnItem(mMenuIndex, I, mItemTop(I))
            End If
            
            UserControl.FontBold = True
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(I)
            PrintMix s, RGB(0, 0, 0)
            
            UserControl.CurrentX = 10 + icon_w + 8 - 1
            UserControl.CurrentY = mItemTop(I) + 1
            PrintMix s, RGB(0, 0, 0)
            
            UserControl.CurrentX = 10 + icon_w + 8 + 1
            UserControl.CurrentY = mItemTop(I) + 1
            PrintMix s, RGB(0, 0, 0)
            
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(I) + 2
            PrintMix s, RGB(0, 0, 0)
                    
            UserControl.CurrentX = 10 + icon_w + 8
            UserControl.CurrentY = mItemTop(I) + 1
            PrintMix s, RGB(255, 255, 255)
            
            s = mItemHotKey(I + mTopItemIndex)
            If s <> "" Then
                w = .ScaleWidth
                UserControl.FontBold = True
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(I)
                PrintMix s, RGB(0, 0, 0)
                
                UserControl.CurrentX = w - MixTextWidth(s) - 12 - 1
                UserControl.CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(0, 0, 0)
                
                UserControl.CurrentX = w - MixTextWidth(s) - 12 + 1
                UserControl.CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(0, 0, 0)
                
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(I) + 2
                PrintMix s, RGB(0, 0, 0)
                        
                UserControl.CurrentX = w - MixTextWidth(s) - 12
                UserControl.CurrentY = mItemTop(I) + 1
                PrintMix s, RGB(255, 255, 255)
            End If
            
            UserControl.FontBold = False
        End If
        
        mCurItem = I
        mTopItemIndex0 = mTopItemIndex
            
        .Refresh
    End With
End Sub

Private Sub VScroll1_Change()
    MoveMenuItem VScroll1.Value
End Sub

Private Sub VScroll1_GotFocus()
    DisablePopMenuTimer = True
End Sub

Private Sub VScroll1_LostFocus()
    DisablePopMenuTimer = False
End Sub

Private Sub VScroll1_Scroll()
    MoveMenuItem VScroll1.Value
End Sub
