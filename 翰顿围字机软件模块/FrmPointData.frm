VERSION 5.00
Begin VB.Form FrmPointData 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   Icon            =   "FrmPointData.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtVDown 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtVSpeed 
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxtCornerR 
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   1740
      Width           =   975
   End
   Begin VB.TextBox TxtYPlus 
      Height          =   285
      Left            =   3960
      TabIndex        =   21
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox TxtXPlus 
      Height          =   285
      Left            =   3960
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox CmbZLayer 
      Height          =   315
      Left            =   3960
      TabIndex        =   17
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox TxtStayTime 
      Height          =   285
      Left            =   3960
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtPointType 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   60
      Width           =   975
   End
   Begin VB.ComboBox CmbAction 
      Height          =   315
      ItemData        =   "FrmPointData.frx":000C
      Left            =   1200
      List            =   "FrmPointData.frx":0019
      TabIndex        =   11
      Text            =   "无"
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   2220
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定"
      Height          =   375
      Left            =   2940
      TabIndex        =   8
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox TxtZPlus 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox TxtZPos 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox TxtYPos 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox TxtXPos 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "下降高度(mm)"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "矢量速度(mm/s)"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2100
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "倒角半径 (mm)"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Y 轴补偿 (μm)"
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "X 轴补偿 (μm)"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "停留时间 (ms)"
      Height          =   255
      Left            =   2820
      TabIndex        =   15
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Z 轴分层"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "该点类型"
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "喷头动作"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2340
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Z 轴补偿 (μm)"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Z 轴位置 (mm)"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Y 轴位置 (mm)"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "X 轴位置 (mm)"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPointData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbZLayer_Click()
    If CmbZLayer.Text <> Str(0) Then
        TxtZPos.Text = Format(LayerZValue(Val(CmbZLayer.Text)), "#0.0##")
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim I As Long, j As Long
    
    PointList(CurPointIndex).Layer = Val(CmbZLayer.Text)
    PointList(CurPointIndex).X = Val(TxtXPos.Text)
    PointList(CurPointIndex).Y = Val(TxtYPos.Text)
    PointList(CurPointIndex).z = Val(TxtZPos.Text)
    PointList(CurPointIndex).xp = Val(TxtXPlus.Text)
    PointList(CurPointIndex).yp = Val(TxtYPlus.Text)
    PointList(CurPointIndex).zp = Val(TxtZPlus.Text)
    
    Select Case CmbAction.Text
        Case "无"
            PointList(CurPointIndex).action = No_Action
        Case "开胶"
            PointList(CurPointIndex).action = StartDropping
        Case "过胶"
            PointList(CurPointIndex).action = Dropping
        Case "闭胶"
            PointList(CurPointIndex).action = StopDropping
        'Case "点胶"
        '    PointList(CurPointIndex).action = PointDropping
    End Select
    
    If PointList(CurPointIndex).action = StartDropping Then
        For I = 1 To OutputStartPointList.count
            If OutputStartPointList.point_id(I) = CurPointIndex Then
                Exit For
            End If
        Next
        If I > OutputStartPointList.count Then
            OutputStartPointList.count = OutputStartPointList.count + 1
            ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
            ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.count)
            ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.count)
            
            OutputStartPointList.point_id(OutputStartPointList.count) = CurPointIndex
        End If
        
        SaveUndo
        
    Else
        For I = 1 To OutputStartPointList.count
            If OutputStartPointList.point_id(I) = CurPointIndex Then
                Exit For
            End If
        Next
        If I <= OutputStartPointList.count Then
            For j = I To OutputStartPointList.count - 1
                OutputStartPointList.point_id(j) = OutputStartPointList.point_id(j + 1)
            Next
            OutputStartPointList.count = OutputStartPointList.count - 1
            ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
        End If
        
        SaveUndo
        
    End If
    
    PointList(CurPointIndex).stay_time = Val(TxtStayTime.Text)
    PointList(CurPointIndex).v_down = Val(TxtVDown.Text)
    'PointList(CurPointIndex).v_up = Val(TxtVUp.Text)
        
    If PointList(CurPointIndex).method = PointMethod.RoundedCorner Then
        'ArcList(PointList(CurPointIndex).arc_id).a = Val(TxtCornerR.Text)
        RoundCorner CurPointIndex, Val(TxtCornerR.Text)
    End If
    
    Unload Me
    FrmMain.PicPathCls
    DrawAll
End Sub

Private Sub CmbAction_Click()
    If CmbAction.Text <> "无" Then
        TxtStayTime.Visible = True
        Label8.Visible = True
    Else
        TxtStayTime.Visible = False
        Label8.Visible = False
    End If
End Sub

Public Sub Form_Load()
    Select Case PointList(CurPointIndex).Type
        Case PointType.NormalPoint
            FrmPointData.TxtPointType.Text = "普通"
        Case PointType.BoxPoint
            FrmPointData.TxtPointType.Text = "矩形顶点"
        Case PointType.ArcPoint
            FrmPointData.TxtPointType.Text = "圆弧点"
        Case PointType.SPLinePoint
            FrmPointData.TxtPointType.Text = "样条线点"
    End Select
    
    CmbZLayer.Clear
    If Device_Mode = 0 Then
        CmbZLayer.AddItem Str(0)
        CmbZLayer.AddItem Str(1)
        CmbZLayer.AddItem Str(2)
        CmbZLayer.AddItem Str(3)
        CmbZLayer.AddItem Str(4)
        CmbZLayer.AddItem Str(5)
        CmbZLayer.AddItem Str(6)
        CmbZLayer.AddItem Str(7)
                    
        If PointList(CurPointIndex).method = PointMethod.RoundedCorner Then
            TxtCornerR.Text = Format(ArcList(PointList(CurPointIndex).arc_id).a, "#0.0##")
            TxtCornerR.Visible = True
            Label11.Visible = True
        Else
            TxtCornerR.Visible = False
            Label11.Visible = False
        End If
        
    ElseIf Device_Mode = 1 Then
        CmbZLayer.AddItem Str(1)
        CmbZLayer.AddItem Str(2)
        CmbZLayer.AddItem Str(3)
        CmbZLayer.AddItem Str(4)
        CmbZLayer.AddItem Str(5)
        
    End If
    
    CmbAction.Clear
    CmbAction.AddItem "无"
    CmbAction.AddItem "开胶"
    CmbAction.AddItem "过胶"
    CmbAction.AddItem "闭胶"
    'CmbAction.AddItem "点胶"
    Select Case PointList(CurPointIndex).action
        Case ActionType.StartDropping
            CmbAction.Text = "开胶"
        Case ActionType.Dropping
            CmbAction.Text = "过胶"
        Case ActionType.StopDropping
            CmbAction.Text = "闭胶"
        'Case ActionType.PointDropping
        '    CmbAction.Text = "点胶"
        Case Else
            CmbAction.Text = "无"
    End Select
    
    CmbZLayer.Text = Str(PointList(CurPointIndex).Layer)
    TxtXPos.Text = Format(PointList(CurPointIndex).X, "#0.0##")
    TxtYPos.Text = Format(PointList(CurPointIndex).Y, "#0.0##")
    TxtZPos.Text = Format(PointList(CurPointIndex).z, "#0.0##")
    TxtXPlus.Text = Format(PointList(CurPointIndex).xp, "#0.0##")
    TxtYPlus.Text = Format(PointList(CurPointIndex).yp, "#0.0##")
    TxtZPlus.Text = Format(PointList(CurPointIndex).zp, "#0.0##")
    TxtVDown.Text = Format(PointList(CurPointIndex).v_down, "#0.0##")
    'TxtVUp.Text = Format(PointList(CurPointIndex).v_up, "#0.0##")
    
    If PointList(CurPointIndex).action <> No_Action Then
        TxtStayTime.Text = Str(PointList(CurPointIndex).stay_time)
        TxtStayTime.Visible = True
        Label8.Visible = True
    Else
        TxtStayTime.Visible = False
        Label8.Visible = False
    End If
    
    If Device_Mode = 0 Then
'        Label13.Visible = False
'        Label14.Visible = False
        
'        TxtVDown.Visible = False
'        TxtVUp.Visible = False
        
    ElseIf Device_Mode = 1 Then
        Label5.Visible = False
        CmbAction.Visible = False
        
        Label6.Visible = False
        TxtPointType.Visible = False
    
        Label11.Visible = False
        TxtCornerR.Visible = False
'        Label12.Visible = False
        
'        TxtVSpeed.Visible = False
    
'        Label13.top = Label12.top
'        Label14.top = Label11.top
        
 '       TxtVDown.top = TxtVSpeed.top
 '       TxtVUp.top = TxtCornerR.top
    End If
    
    If Device_ZAxisControlMode = ZAxisControlMode.Switch Then
        Label3.Visible = False
        Label4.Visible = False
        Label7.Visible = False
        Label13.Visible = False
        CmbZLayer.Visible = False
        TxtZPos.Visible = False
        TxtZPlus.Visible = False
        TxtVDown.Visible = False
    End If
End Sub
