VERSION 5.00
Begin VB.Form FrmUserParameter 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "平台参数设置"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "FrmParameter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8265
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox TxtMinPathStep 
      Height          =   300
      Left            =   3240
      TabIndex        =   34
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox TxtPlatformHMM 
      Height          =   285
      Left            =   4440
      TabIndex        =   30
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox TxtPlatformWMM 
      Height          =   285
      Left            =   1800
      TabIndex        =   29
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox TxtSPLineD 
      Height          =   300
      Left            =   4440
      TabIndex        =   28
      Top             =   2760
      Width           =   855
   End
   Begin VB.CheckBox ChkShowAux 
      BackColor       =   &H8000000A&
      Caption         =   "显示辅助线"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox ChkAux 
      BackColor       =   &H8000000A&
      Caption         =   "对齐辅助线"
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TxtAuxYLine 
      Height          =   2535
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TxtAuxXline 
      Height          =   2535
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TxtCornerR 
      Height          =   315
      Left            =   4440
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox TxtArcStepFactor 
      Height          =   315
      Left            =   4440
      TabIndex        =   16
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   6840
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox TxtHVTrapWidth 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox TxtTrapWidth 
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox TxtPointSize 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox TxtSubGridY 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox TxtSubGridX 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtMainGridY 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox TxtMainGridX 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "最小步距(mm)"
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Y 轴方向高度(mm)"
      Height          =   255
      Left            =   2880
      TabIndex        =   32
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X 轴方向宽度(mm)"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "样条插值参数"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "辅助线"
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "倒角半径(mm)"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "圆弧分角参数"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "平直对齐距离(象素)"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1635
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "节点捕捉半径(象素)"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "节点尺寸(象素)"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "子网格Y轴间距(mm)"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "子网格X轴间距(mm)"
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "主网格Y轴间距(mm)"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "主网格X轴间距(mm)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1635
   End
End
Attribute VB_Name = "FrmUserParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ListIndex As Long

Private Sub ChkAux_Click()
    AuxLineEnabled = IIf(ChkAux.value = 1, True, False)
End Sub

Private Sub ChkShowAux_Click()
    AuxLineVisible = IIf(ChkShowAux.value = 1, True, False)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim Count As Long, I As Long, s As String, c As Long, v As Variant
    Dim ItemName As String, PointTime As String, UserSpeed As String
    
    Device_UserSize(1) = Val(TxtPlatformWMM.Text)
    Device_UserSize(2) = Val(TxtPlatformHMM.Text)
    
    WritePrivateProfileString "UserSize", Str(1), Str(Device_UserSize(1)), App.Path & "\" & App.EXEName & ".ini"
    WritePrivateProfileString "UserSize", Str(2), Str(Device_UserSize(2)), App.Path & "\" & App.EXEName & ".ini"
    
    '--------------------------------------------------
    
    MainGridX = Val(TxtMainGridX.Text)
    MainGridY = Val(TxtMainGridY.Text)
    SubGridX = Val(TxtSubGridX.Text)
    SubGridY = Val(TxtSubGridY.Text)
    
    PointSize = Max(1, Val(TxtPointSize.Text))
    TrapWidth = Max(1, Val(TxtTrapWidth.Text))
    HVTrapWidth = Max(1, Val(TxtHVTrapWidth.Text))
    CornerR = Max(1, Val(TxtCornerR.Text))
    ArcStepFactor = Max(1, Val(TxtArcStepFactor.Text))
    MinPathStep = Abs(Val(TxtMinPathStep.Text))
    SPLine_SegmentBetweenPoints = Max(1, Val(TxtSPLineD.Text))
        
    utw = GetUserDistance(0, 0, TrapWidth, 0)
    
    s = TxtAuxXline.Text
    v = Split(s, vbCrLf)
    AuxXLineCount = 0
    For I = 0 To UBound(v)
        If Trim(v(I)) <> "" Then
            ReDim Preserve AuxXLine(AuxXLineCount)
            AuxXLine(AuxXLineCount) = Trim(v(I))
            AuxXLineCount = AuxXLineCount + 1
        End If
    Next
            
    s = TxtAuxYLine.Text
    v = Split(s, vbCrLf)
    AuxYLineCount = 0
    For I = 0 To UBound(v)
        If Trim(v(I)) <> "" Then
            ReDim Preserve AuxYLine(AuxYLineCount)
            AuxYLine(AuxYLineCount) = Trim(v(I))
            AuxYLineCount = AuxYLineCount + 1
        End If
    Next
        
    Unload Me
    
    ViewMinX = 0
    ViewMaxX = Device_UserSize(1)
    ViewMinY = 0
    ViewMaxY = Device_UserSize(2)
    ViewMargin = 0.03
    
    FrmMain.FormResize
    FrmMain.PicPathCls
    DrawAll
    
    WriteUserParameter
End Sub

Private Sub Form_Deactivate()
    Me.ZOrder 0
End Sub

Private Sub Form_Load()
    Dim Count As Long, I As Long, s As String, v As Variant
    
    Set Me.Icon = FrmMain.Icon
    
    TxtPlatformWMM.Text = GetStringFromINI("UserSize", Str(1), "1000", App.Path & "\" & App.EXEName & ".ini")
    TxtPlatformHMM.Text = GetStringFromINI("UserSize", Str(2), "1000", App.Path & "\" & App.EXEName & ".ini")

    TxtMainGridX.Text = MainGridX
    TxtMainGridY.Text = MainGridY
    TxtSubGridX.Text = SubGridX
    TxtSubGridY.Text = SubGridY
    
    TxtPointSize.Text = PointSize
    TxtTrapWidth.Text = TrapWidth
    TxtHVTrapWidth.Text = HVTrapWidth
    TxtCornerR.Text = CornerR
    TxtArcStepFactor.Text = ArcStepFactor
    TxtMinPathStep.Text = MinPathStep
    TxtSPLineD.Text = SPLine_SegmentBetweenPoints
    
    TxtAuxXline.Text = ""
    For I = 0 To AuxXLineCount - 1
        TxtAuxXline.Text = TxtAuxXline.Text & Str(AuxXLine(I)) & vbCrLf
    Next
    
    TxtAuxYLine.Text = ""
    For I = 0 To AuxYLineCount - 1
        TxtAuxYLine.Text = TxtAuxYLine.Text & Str(AuxYLine(I)) & vbCrLf
    Next
    
    ChkAux.value = IIf(AuxLineEnabled = True, 1, 0)
    ChkShowAux.value = IIf(AuxLineVisible = True, 1, 0)
    
    If Device_Mode = 1 Then
        TxtHVTrapWidth.Visible = False
        TxtCornerR.Visible = False
        TxtArcStepFactor.Visible = False
        TxtSPLineD.Visible = False
        
        Label9.Visible = False
        Label13.Visible = False
        Label15.Visible = False
        Label19.Visible = False
    End If
End Sub

