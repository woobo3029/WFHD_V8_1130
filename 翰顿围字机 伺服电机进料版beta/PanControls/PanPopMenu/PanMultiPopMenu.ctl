VERSION 5.00
Begin VB.UserControl PanMultiPopMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin BizCall.PanPopMenu PanPopMenu 
      Height          =   1275
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   2249
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1500
      Top             =   1620
   End
End
Attribute VB_Name = "PanMultiPopMenu"
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

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private mMenuCount As Integer
Private mMenuCurIndex As Integer
Private mMaxItemCount As Integer
Private mMaxItemShow As Integer

Private mItemCount(20) As Integer
Private mItemCaption() As String
Private mItemIconPath() As String
Private mItemSubMenuIndex() As Integer
Private mItemSubMenuShowing() As Boolean
Private mItemCurIndex As Integer

Public Event ClickOnItem(MenuIndex As Integer, ItemIndex As Integer)
Public Event KeyPress(KeyAscii As Integer)

Public Sub PopUp(ByVal X As Integer, ByVal Y As Integer, Optional MaxItemShow As Integer = 0)
    Dim I As Integer
    
    MoveWindow UserControl.hWnd, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, 0
    
    If mMenuCount = 0 Then
        Exit Sub
    End If
    
    mMaxItemShow = MaxItemShow
    
    Wait 0.1 'wait for the timer1_timer event
    
    PopMenu X, Y, 0
    SetFocusOnMenu 0
    Timer1.Enabled = True
End Sub

Private Sub PopMenu(ByVal X As Integer, ByVal Y As Integer, ByVal MenuIndex As Integer)
    Dim m As Integer, I As Integer
    Dim Left As Integer, Top As Integer
    
    If MenuIndex = 0 Then
        m = 0
    Else
        For I = 0 To PanPopMenu.Count - 1
            If PanPopMenu(I).MenuIndex = MenuIndex Then
                Exit Sub
            End If
        Next
        m = PanPopMenu.Count
        Load PanPopMenu(m)
        PanPopMenu(m).ParentMenuIndex = mMenuCurIndex
    End If
    
    mMenuCurIndex = MenuIndex
    PanPopMenu(m).MenuIndex = MenuIndex
    PanPopMenu(m).ItemCount = mItemCount(mMenuCurIndex)
    For I = 0 To mItemCount(mMenuCurIndex) - 1
        PanPopMenu(m).ItemCurIndex = I
        PanPopMenu(m).ItemCaption = mItemCaption(mMenuCurIndex, I)
        PanPopMenu(m).ItemIconPath = mItemIconPath(mMenuCurIndex, I)
    Next
    
    On Error Resume Next
    
    PanPopMenu(m).PopUp mMaxItemShow

    If X + PanPopMenu(m).Width < UserControl.Parent.Width / Screen.TwipsPerPixelX - 8 Then
        Left = X
    Else
        Left = UserControl.Parent.Width / Screen.TwipsPerPixelX - PanPopMenu(m).Width - 8
    End If
    
    If X + PanPopMenu(m).Width > UserControl.Parent.Width / Screen.TwipsPerPixelX + 30 Then
        Y = Y + 20
    End If
        
    Top = UserControl.Parent.Height / Screen.TwipsPerPixelY - PanPopMenu(m).Height - 45
    Top = IIf(Y < Top, Y, Top)
    
    PanPopMenu(m).Move Left, Top
    PanPopMenu(m).ZOrder 0
    PanPopMenu(m).Visible = True
End Sub

Private Sub SetFocusOnMenu(ByVal MenuIndex As Integer, Optional KeyCode As Integer = 40)
    Dim I As Integer
    
    On Error Resume Next
    
    For I = 0 To PanPopMenu.Count - 1
        If PanPopMenu(I).MenuIndex = MenuIndex Then
            PanPopMenu(I).SetFocus
            PanPopMenu(I).UserControl_KeyDown KeyCode, 0
            Exit Sub
        End If
    Next
End Sub

Private Sub PanPopMenu_ClickOnItem(Index As Integer, MenuIndex As Integer, ItemIndex As Integer)
    RaiseEvent ClickOnItem(MenuIndex, ItemIndex)
End Sub

Private Sub PanPopMenu_MouseOnItem(Index As Integer, MenuIndex As Integer, ItemIndex As Integer, ItemTop As Integer)
    Dim I As Integer, k As Integer, z As Boolean, m As Integer, t As Integer
    Dim X As Integer, Y As Integer, w As Integer
    
    For I = 0 To mItemCount(MenuIndex) - 1
        If I <> ItemIndex Then
            z = mItemSubMenuShowing(MenuIndex, I)
            If z Then
                k = mItemSubMenuIndex(MenuIndex, I)
                For m = 1 To PanPopMenu.Count - 1
                    If PanPopMenu(m).MenuIndex = k Then
                        For t = m To PanPopMenu.Count - 1
                            mMenuCurIndex = PanPopMenu(t).ParentMenuIndex
                            Unload PanPopMenu(t)
                        Next
                        Exit For
                    End If
                    DoEvents
                Next
                mItemSubMenuShowing(MenuIndex, I) = False
                Exit For
            End If
        End If
        
        DoEvents
    Next
    
    If ItemIndex <> -1 Then
        For m = 0 To PanPopMenu.Count - 1
            If PanPopMenu(m).MenuIndex = MenuIndex Then
                X = PanPopMenu(m).Left
                Y = PanPopMenu(m).Top
                w = PanPopMenu(m).Width
                Exit For
            End If
        Next
        
        If mItemSubMenuIndex(MenuIndex, ItemIndex) > 0 Then
            PopMenu X + w - 6, Y + ItemTop - 4, mItemSubMenuIndex(MenuIndex, ItemIndex)
            mItemSubMenuShowing(MenuIndex, ItemIndex) = True
        End If
    End If
    
End Sub

Private Sub HideMenu(ByVal MenuIndex As Integer, ByVal ItemIndex As Integer)
    Dim I As Integer
    
    If ItemIndex > -1 Then
        I = mItemSubMenuIndex(MenuIndex, ItemIndex)
        If I > 0 Then
            HideMenu I, -1
        End If
    End If
    
    For I = 0 To PanPopMenu.Count - 1
        If PanPopMenu(I).MenuIndex = MenuIndex Then
            If I = 0 Then
                PanPopMenu(I).Visible = False
            Else
                Unload PanPopMenu(I)
            End If
            Exit For
        End If
    Next
    
    UserControl.Refresh
End Sub

Private Sub PanPopMenu_SubMenuSelected(Index As Integer, MenuIndex As Integer, Param As Integer, SelectMode As Integer)
    Dim I As Integer, k As Integer, Top As Integer
    
    If SelectMode = 1 And Param >= 0 Then 'go into submenu
        k = mItemSubMenuIndex(MenuIndex, Param)
        If k > 0 Then
            For I = 0 To PanPopMenu.Count - 1
                If PanPopMenu(I).MenuIndex = k Then
                    SetFocusOnMenu k
                    Exit For
                End If
            Next
            If I = PanPopMenu.Count Then
                Top = PanPopMenu(Index).ItemTop
                PanPopMenu_MouseOnItem Index, MenuIndex, Param, Top
                SetFocusOnMenu k
            End If
        End If
    ElseIf SelectMode = 2 Then 'return from submenu
        For I = 0 To PanPopMenu.Count - 1
            If PanPopMenu(I).MenuIndex = MenuIndex Then
                Unload PanPopMenu(I)
                Exit For
            End If
        Next
        
        SetFocusOnMenu Param, 0
        UserControl.Refresh
    ElseIf SelectMode = 3 Then
        HideMenu MenuIndex, Param
    End If
End Sub

Private Sub Timer1_Timer()
    Dim LButtonState As Long, I As Integer, t As Variant

    On Error Resume Next
    
    LButtonState = GetKeyState(VK_LBUTTON)

    If DisablePopMenuTimer = True Then
        Exit Sub
    End If
    
    If (LButtonState = -127 Or LButtonState = -128) Then
        Timer1.Enabled = False

        For I = 0 To PanPopMenu.Count - 1
            ShowWindow PanPopMenu(I).hWnd, SW_HIDE
            PanPopMenu(I).Visible = False
        Next

        'Wait 0.2

        For I = 0 To PanPopMenu.Count - 1
            If I = 0 Then
                PanPopMenu(0).ItemCount = 1
            Else
                Unload PanPopMenu(I)
            End If
        Next
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackStyle = 0
    UserControl.BorderStyle = 0
    PanPopMenu(0).Visible = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim I As Integer, j As Integer
    
    mMenuCount = PropBag.ReadProperty("MenuCount", 0)
    mMaxItemCount = 1
    If mMenuCount > 0 Then
        'ReDim mItemCount(mMenuCount)
        For I = 0 To mMenuCount - 1
            mItemCount(I) = PropBag.ReadProperty("ItemCount" & Trim(Str(I)), 0)
            If mMaxItemCount < mItemCount(I) Then
                mMaxItemCount = mItemCount(I)
            End If
        Next
    End If
    
    'ReDim Preserve mItemCaption(mMenuCount, mMaxItemCount), mItemPath(mMenuCount, mMaxItemCount), mItemSubMenuIndex(mMenuCount, mMaxItemCount)
    ReDim Preserve mItemCaption(20, mMaxItemCount), mItemIconPath(20, mMaxItemCount), mItemSubMenuIndex(20, mMaxItemCount), mItemSubMenuShowing(20, mMaxItemCount)
    
    If mMenuCount > 0 Then
        For I = 0 To mMenuCount - 1
            For j = 0 To mItemCount(I) - 1
                mItemCaption(I, j) = PropBag.ReadProperty("ItemCaption" & Trim(Str(I)) & "_" & Trim(Str(j)), "")
                mItemIconPath(I, j) = PropBag.ReadProperty("ItemIconPath" & Trim(Str(I)) & "_" & Trim(Str(j)), "")
                mItemSubMenuIndex(I, j) = PropBag.ReadProperty("ItemSubMenuIndex" & Trim(Str(I)) & "_" & Trim(Str(j)), -1)
            Next
        Next
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim I As Integer, j As Integer
    
    PropBag.WriteProperty "MenuCount", mMenuCount, 0
    If mMenuCount > 0 Then
        For I = 0 To mMenuCount - 1
            PropBag.WriteProperty "ItemCount" & Trim(Str(I)), mItemCount(I), 0
        Next
        
        For I = 0 To mMenuCount - 1
            For j = 0 To mItemCount(I) - 1
                PropBag.WriteProperty "ItemCaption" & Trim(Str(I)) & "_" & Trim(Str(j)), mItemCaption(I, j), ""
                PropBag.WriteProperty "ItemIconPath" & Trim(Str(I)) & "_" & Trim(Str(j)), mItemIconPath(I, j), ""
                PropBag.WriteProperty "ItemSubMenuIndex" & Trim(Str(I)) & "_" & Trim(Str(j)), mItemSubMenuIndex(I, j), -1
            Next
        Next
    End If
End Sub

Public Property Get MenuCount() As Integer
    MenuCount = mMenuCount
End Property

Public Property Let MenuCount(ByVal NewValue As Integer)
    mMenuCount = NewValue
            
    If mMaxItemCount = 0 Then mMaxItemCount = 1
    'ReDim Preserve mItemCaption(mMenuCount, mMaxItemCount), mItemPath(mMenuCount, mMaxItemCount), mItemSubMenuIndex(mMenuCount, mMaxItemCount)
    ReDim Preserve mItemCaption(20, mMaxItemCount), mItemIconPath(20, mMaxItemCount), mItemSubMenuIndex(20, mMaxItemCount), mItemSubMenuShowing(20, mMaxItemCount)
    PropertyChanged "MenuCount"
End Property

Public Property Get MenuCurIndex() As Integer
    MenuCurIndex = mMenuCurIndex
End Property

Public Property Let MenuCurIndex(ByVal NewValue As Integer)
    If NewValue < mMenuCount Then
        mMenuCurIndex = NewValue
        mItemCurIndex = 0
        
        PropertyChanged "MenuCurIndex"
    End If
End Property

Public Property Get ItemCount() As Integer
    If mMenuCurIndex < mMenuCount Then
        ItemCount = mItemCount(mMenuCurIndex)
    End If
End Property

Public Property Let ItemCount(ByVal NewValue As Integer)
    If mMenuCurIndex < mMenuCount Then
        mItemCount(mMenuCurIndex) = NewValue
    
        If mMaxItemCount < NewValue Then
            mMaxItemCount = NewValue
            'ReDim Preserve mItemCaption(mMenuCount, mMaxItemCount), mItemPath(mMenuCount, mMaxItemCount), mItemSubMenuIndex(mMenuCount, mMaxItemCount)
            ReDim Preserve mItemCaption(20, mMaxItemCount), mItemIconPath(20, mMaxItemCount), mItemSubMenuIndex(20, mMaxItemCount), mItemSubMenuShowing(20, mMaxItemCount)
        End If
    End If
    PropertyChanged "ItemCount"
End Property

Public Property Get ItemCurIndex() As Integer
    ItemCurIndex = mItemCurIndex
End Property

Public Property Let ItemCurIndex(ByVal NewValue As Integer)
    If NewValue < mItemCount(mMenuCurIndex) Then
        mItemCurIndex = NewValue
        
        PropertyChanged "ItemCurIndex"
    End If
End Property

Public Property Get ItemCaption() As String
    If mMenuCurIndex >= mMenuCount Then
        Exit Property
    End If
    If mItemCurIndex >= mItemCount(mMenuCurIndex) Then
        Exit Property
    End If
    
    ItemCaption = mItemCaption(mMenuCurIndex, mItemCurIndex)
End Property

Public Property Let ItemCaption(ByVal NewValue As String)
    mItemCaption(mMenuCurIndex, mItemCurIndex) = NewValue
    
    PropertyChanged "ItemCaption"
End Property

Public Property Get ItemIconPath() As String
    If mMenuCurIndex >= mMenuCount Then
        Exit Property
    End If
    If mItemCurIndex >= mItemCount(mMenuCurIndex) Then
        Exit Property
    End If
    
    ItemIconPath = mItemIconPath(mMenuCurIndex, mItemCurIndex)
End Property

Public Property Let ItemIconPath(ByVal NewValue As String)
    mItemIconPath(mMenuCurIndex, mItemCurIndex) = NewValue
    
    PropertyChanged "ItemIconPath"
End Property

Public Property Get ItemSubMenuIndex() As Integer
    If mMenuCurIndex >= mMenuCount Then
        Exit Property
    End If
    If mItemCurIndex >= mItemCount(mMenuCurIndex) Then
        Exit Property
    End If
        
    ItemSubMenuIndex = mItemSubMenuIndex(mMenuCurIndex, mItemCurIndex)
End Property

Public Property Let ItemSubMenuIndex(ByVal NewValue As Integer)
    mItemSubMenuIndex(mMenuCurIndex, mItemCurIndex) = NewValue
    
    PropertyChanged "ItemSubMenuIndex"
End Property


