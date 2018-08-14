VERSION 5.00
Begin VB.UserControl PanTextBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   139
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   780
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "PanTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ScrollBarType
    None = 0
    Horizontal = 1
    Vertical = 2
    Both = 3
End Enum

Private mBorderColor As OLE_COLOR
Private mMultiLine As Boolean
Public Text As TextBox
Public hWnd As Long

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Change()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub Text1_Change()
    MainFormUserEventRaised = True
    RaiseEvent Change
End Sub

Private Sub Text1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    MainFormUserEventRaised = True
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    MainFormUserEventRaised = True
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    MainFormUserEventRaised = True
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static x0 As Single, y0 As Single
    
    If MsgFormIsShowing = True Then
        Exit Sub
    End If
    
    If x0 <> X Or y0 <> Y Then
'        Text1.SetFocus
        
        x0 = X
        y0 = Y
    
        MainFormUserEventRaised = True
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub Text2_Change()
    MainFormUserEventRaised = True
    RaiseEvent Change
End Sub

Private Sub Text2_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    MainFormUserEventRaised = True
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    MainFormUserEventRaised = True
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    MainFormUserEventRaised = True
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static x0 As Single, y0 As Single
    
    If MsgFormIsShowing = True Then
        Exit Sub
    End If
    
    If x0 <> X Or y0 <> Y Then
'        Text2.SetFocus
        
        x0 = X
        y0 = Y
        
        MainFormUserEventRaised = True
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_Initialize()
    hWnd = UserControl.hWnd
End Sub

Private Sub UserControl_InitProperties()
    mBorderColor = RGB(128, 128, 255)
End Sub

Private Sub UserControl_Paint()
    ShowTextBox
End Sub

Sub ShowTextBox()
    Dim BorderClr As Long
    
    On Error Resume Next
    
    UserControl.BorderStyle = 0
    UserControl.BackColor = Extender.Parent.BackColor
    
    BorderClr = mBorderColor
    
    UserControl.AutoRedraw = True
    UserControl.Cls
    UserControl.DrawMode = 13
    UserControl.DrawWidth = 1
    UserControl.Line (2, 0)-Step(UserControl.ScaleWidth - 5, 0), BorderClr
    UserControl.Line -(UserControl.ScaleWidth - 1, 2), BorderClr
    UserControl.Line -Step(0, UserControl.ScaleHeight - 5), BorderClr
    UserControl.Line -(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1), BorderClr
    UserControl.Line -(2, UserControl.ScaleHeight - 1), BorderClr
    UserControl.Line -(0, UserControl.ScaleHeight - 3), BorderClr
    UserControl.Line -(0, 2), BorderClr
    UserControl.Line -(2, 0), BorderClr

    If Text Is Nothing Then
        Set Text = Text1
        Text2.Visible = False
    End If
    Text.Move 2, 2, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    mBorderColor = PropBag.ReadProperty("BorderColor", RGB(128, 128, 255))
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Text1.BackColor = PropBag.ReadProperty("BackColor")
    Text1.ForeColor = PropBag.ReadProperty("ForeColor")
    Text2.BackColor = Text1.BackColor
    Text2.ForeColor = Text1.ForeColor
    Text2.Enabled = Text1.Enabled
    
    mMultiLine = PropBag.ReadProperty("MultiLine")
    If mMultiLine = False Then
        Set Text = Text1
        Text1.Visible = True
        Text2.Visible = False
    Else
        Set Text = Text2
        Text1.Visible = False
        Text2.Visible = True
    End If
    
    Text1.font = PropBag.ReadProperty("font")
    Text2.font = Text1.font
    
    ShowTextBox
End Sub

Private Sub UserControl_Resize()
    ShowTextBox
End Sub

Public Property Get Enabled() As Boolean
    On Error Resume Next
    Enabled = Text.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    Text.Enabled = NewValue
    
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Text.Enabled, True
    PropBag.WriteProperty "BackColor", Text.BackColor
    PropBag.WriteProperty "ForeColor", Text.ForeColor
    PropBag.WriteProperty "BorderColor", mBorderColor
    PropBag.WriteProperty "MultiLine", mMultiLine
    PropBag.WriteProperty "font", Text1.font
End Sub

Public Property Get BackColor() As OLE_COLOR
    If Not Text Is Nothing Then
        BackColor = Text.BackColor
    End If
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    Text.BackColor = NewValue
    
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    If Not Text Is Nothing Then
        ForeColor = Text.ForeColor
    End If
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    Text.ForeColor = NewValue
    
    PropertyChanged "ForeColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
    mBorderColor = NewValue
    
    ShowTextBox
    PropertyChanged "BorderColor"
End Property

Public Property Get MultiLine() As Boolean
    MultiLine = mMultiLine
End Property

Public Property Let MultiLine(ByVal NewValue As Boolean)
    mMultiLine = NewValue
    
    If mMultiLine = False Then
        Text1.BackColor = Text2.BackColor
        Text1.ForeColor = Text2.ForeColor
        Text1.Enabled = Text2.Enabled
        
        Set Text = Text1
        Text1.Visible = True
        Text2.Visible = False
    Else
        Text2.BackColor = Text1.BackColor
        Text2.ForeColor = Text1.ForeColor
        Text2.Enabled = Text1.Enabled
        
        Set Text = Text2
        Text1.Visible = False
        Text2.Visible = True
    End If
    
    ShowTextBox
    PropertyChanged "MultiLine"
End Property

Public Property Get font() As font
    Set font = Text1.font
End Property

Public Property Set font(newfont As font)
    Set Text1.font = newfont
    Set Text2.font = newfont
    
    ShowTextBox
    PropertyChanged "font"
End Property

Public Property Get SelStart() As Integer
    On Error Resume Next
    SelStart = Text.SelStart
End Property

Public Property Let SelStart(ByVal NewValue As Integer)
    Text.SelStart = NewValue
    
    PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Integer
    On Error Resume Next
    SelLength = Text.SelLength
End Property

Public Property Let SelLength(ByVal NewValue As Integer)
    Text.SelLength = NewValue
    
    PropertyChanged "SelLength"
End Property

Public Property Get SelText() As String
    On Error Resume Next
    SelText = Text.SelText
End Property

Public Property Let SelText(ByVal NewValue As String)
    Text.SelText = NewValue
    
    PropertyChanged "SelText"
End Property


