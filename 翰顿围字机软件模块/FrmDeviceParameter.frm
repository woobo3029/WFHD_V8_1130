VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDeviceParameter 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设备参数设置"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17145
   Icon            =   "FrmDeviceParameter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   17145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtTotalWorkBendCount 
      Height          =   285
      Left            =   2880
      TabIndex        =   120
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清空"
      Height          =   420
      Left            =   6000
      TabIndex        =   118
      Top             =   8760
      Width           =   735
   End
   Begin VB.TextBox TxtTotalWorkTime 
      Height          =   285
      Left            =   4800
      TabIndex        =   114
      Top             =   9000
      Width           =   975
   End
   Begin VB.TextBox TxtTotalWorkCount 
      Height          =   285
      Left            =   4800
      TabIndex        =   113
      Top             =   8640
      Width           =   975
   End
   Begin VB.TextBox TxtTotalWorkLength 
      Height          =   285
      Left            =   2880
      TabIndex        =   112
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   11160
      TabIndex        =   104
      Top             =   3600
      Width           =   135
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000A&
      Caption         =   "弯刀头运动角度/实现角度"
      Height          =   8295
      Left            =   11400
      TabIndex        =   79
      Top             =   240
      Width           =   5535
      Begin VB.TextBox TxtBendDis 
         Height          =   285
         Index           =   4
         Left            =   3600
         TabIndex        =   111
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   110
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   109
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   108
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   285
         Index           =   5
         Left            =   4320
         TabIndex        =   107
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   1800
         TabIndex        =   105
         Top             =   7920
         Width           =   255
      End
      Begin VB.CommandButton CmdSortAngleTable 
         Caption         =   "排序"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   7800
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid GrdAngleTable 
         Height          =   6975
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   12303
         _Version        =   393216
         Rows            =   500
         Cols            =   7
         ScrollBars      =   2
         BorderStyle     =   0
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "间距(mm)"
         Height          =   255
         Left            =   360
         TabIndex        =   106
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Caption         =   "设备尺寸"
      Height          =   2655
      Left            =   240
      TabIndex        =   56
      Top             =   5880
      Width           =   5415
      Begin VB.TextBox TxtCutWMM 
         Height          =   285
         Left            =   4440
         TabIndex        =   80
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox TxtHoleWMM3 
         Height          =   285
         Left            =   4440
         TabIndex        =   71
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox TxtHoleWMM2 
         Height          =   285
         Left            =   4440
         TabIndex        =   70
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TxtHoleWMM1 
         Height          =   285
         Left            =   4440
         TabIndex        =   69
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtLenMM 
         Height          =   285
         Left            =   960
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtHoleMM1 
         Height          =   285
         Left            =   3360
         TabIndex        =   61
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtHoleMM2 
         Height          =   285
         Left            =   3360
         TabIndex        =   60
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TxtHoleMM3 
         Height          =   285
         Left            =   3360
         TabIndex        =   59
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox TxtCutMM 
         Height          =   285
         Left            =   3360
         TabIndex        =   57
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "注：此处的距离均从弯刀口开始"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "到中心线距离(mm)"
         Height          =   495
         Left            =   3360
         TabIndex        =   68
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "宽度(mm)"
         Height          =   255
         Left            =   4440
         TabIndex        =   67
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "到输送器前端距离(mm)"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "桥位孔"
         Height          =   255
         Left            =   2640
         TabIndex        =   64
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "直切孔"
         Height          =   255
         Left            =   2640
         TabIndex        =   63
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "鹰嘴孔"
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "剪刀"
         Height          =   255
         Left            =   2880
         TabIndex        =   58
         Top             =   2040
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      Caption         =   "开孔控制"
      Height          =   1815
      Left            =   240
      TabIndex        =   49
      Top             =   3960
      Width           =   5415
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   3840
         TabIndex        =   74
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   3840
         TabIndex        =   73
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox TxtStartV3 
         Height          =   285
         Left            =   2280
         TabIndex        =   52
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtSpeed3 
         Height          =   285
         Left            =   2280
         TabIndex        =   51
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtAccl3 
         Height          =   285
         Left            =   2280
         TabIndex        =   50
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "鹰嘴孔"
         Height          =   255
         Left            =   4200
         TabIndex        =   76
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "直切孔"
         Height          =   255
         Left            =   4200
         TabIndex        =   75
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "剪断时选择:"
         Height          =   255
         Left            =   3600
         TabIndex        =   72
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(p/s)"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(p/s)"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(p/ss)"
         Height          =   255
         Left            =   1200
         TabIndex        =   53
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "弯刀控制"
      Height          =   8295
      Left            =   5880
      TabIndex        =   24
      Top             =   240
      Width           =   5175
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000A&
         Caption         =   "弯刀头间隙测试"
         Height          =   3735
         Left            =   240
         TabIndex        =   86
         Top             =   4320
         Width           =   4695
         Begin VB.CommandButton CmdKnifeBack 
            Caption         =   "刀片后退20mm"
            Height          =   375
            Left            =   1680
            TabIndex        =   103
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton CmdKnifeFeed 
            Caption         =   "刀片前进20mm"
            Height          =   375
            Left            =   120
            TabIndex        =   102
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton CmdReset2 
            Caption         =   "无偏移复位"
            Height          =   495
            Left            =   3240
            TabIndex        =   100
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtBendHeadOffsetUpdatedD 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox TxtLeftGapUpdatedD 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox TxtRightGapUpdatedD 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   3000
            Width           =   735
         End
         Begin VB.CommandButton CmdSetGap 
            Caption         =   "设置"
            Height          =   495
            Left            =   3240
            TabIndex        =   93
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton CmdTurnPositive 
            Caption         =   "弯刀头正转"
            Height          =   495
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton CmdTurnNegative 
            Caption         =   "弯刀头反转"
            Height          =   495
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox TxtLeftGapTestD 
            Height          =   285
            Left            =   2040
            TabIndex        =   88
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtRightGapTestD 
            Height          =   285
            Left            =   2040
            TabIndex        =   87
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "(复位时不对中心偏移进行校正)"
            Height          =   255
            Left            =   480
            TabIndex        =   101
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "中心偏移量(°)"
            Height          =   255
            Left            =   720
            TabIndex        =   99
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "校正后正向间隙角度(°)"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "校正后负向间隙角度(°)"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "正向间隙角度(°)"
            Height          =   255
            Left            =   600
            TabIndex        =   90
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "负向间隙角度(°)"
            Height          =   255
            Left            =   600
            TabIndex        =   89
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBendHeadOffsetD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   85
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBreakingAngle 
         Height          =   285
         Left            =   4200
         TabIndex        =   78
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox TxtRightGapD 
         Height          =   285
         Left            =   4200
         TabIndex        =   46
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox TxtLeftGapD 
         Height          =   285
         Left            =   4200
         TabIndex        =   45
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtMaxSpeedDPS 
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtAccl2 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TxtSpeed2 
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox TxtStartV2 
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TxtBackStartV2 
         Height          =   285
         Left            =   4200
         TabIndex        =   28
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TxtBackSpeed2 
         Height          =   285
         Left            =   4200
         TabIndex        =   27
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtBackAccl2 
         Height          =   285
         Left            =   4200
         TabIndex        =   26
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox TxtPPD 
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "中心偏移量(°)"
         Height          =   255
         Left            =   3000
         TabIndex        =   84
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "挠断往复角度(°)"
         Height          =   255
         Left            =   2760
         TabIndex        =   77
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "负向间隙角度(°)"
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "正向间隙角度(°)"
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "最大速度(度/秒)(°/s)"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(°/ss)"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(°/s)"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(°/s)"
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(°/s)"
         Height          =   255
         Left            =   3120
         TabIndex        =   38
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(°/s)"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(°/ss)"
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "弯刀速度"
         Height          =   255
         Left            =   1920
         TabIndex        =   35
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "空回速度"
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "每度脉冲数(p/°)"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "刀片输送控制"
      Height          =   3615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5415
      Begin VB.TextBox TxtFeedMaxMM 
         Height          =   285
         Left            =   4440
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtAccl 
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox TxtSpeed 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtStartV 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TxtBackStartV 
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TxtBackSpeed 
         Height          =   285
         Left            =   4440
         TabIndex        =   8
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtBackAccl 
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox TxtHoldDelay 
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtMaxSpeedMMPS 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtPP100MM 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "输送行程(mm)"
         Height          =   255
         Left            =   3240
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(mm/ss)"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(mm/s)"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(mm/s)"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(mm/s)"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(mm/s)"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(mm/ss)"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "刀片速度"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "空回速度"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "夹紧延时(ms)"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "最大速度(毫米/秒)(mm/s)"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "每100毫米脉冲数(p/100mm)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.CheckBox ChkLocked 
      BackColor       =   &H8000000A&
      Caption         =   "锁定参数"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定"
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "总折弯数"
      Height          =   255
      Left            =   1920
      TabIndex        =   119
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "总时数(h)"
      Height          =   255
      Left            =   3960
      TabIndex        =   117
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "总件数"
      Height          =   255
      Left            =   3960
      TabIndex        =   116
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "总长度(m)"
      Height          =   255
      Left            =   1920
      TabIndex        =   115
      Top             =   8640
      Width           =   855
   End
End
Attribute VB_Name = "FrmDeviceParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurRow As Long
Dim CurCol As Long

Dim Reset2WithoutOffset As Boolean

Private Sub ChkLocked_Click()
    Dim obj As Object
    If ChkLocked.value = 1 Then
        For Each obj In Me
            If (TypeOf obj Is TextBox) Or (TypeOf obj Is CommandButton) Or (TypeOf obj Is OptionButton) Or (TypeOf obj Is MSFlexGrid) Then
                obj.Enabled = False
            End If
        Next
        CmdSave.Enabled = True
        CmdCancel.Enabled = True
    Else
        If MsgBox("错误的参数设置将导致设备运行异常。请确定是否放弃该操作？ ", vbQuestion + vbYesNo + vbSystemModal, "") = vbNo Then
            For Each obj In Me
                obj.Enabled = True
            Next
        Else
            ChkLocked.value = 1
        End If
    End If

End Sub


Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdKnifeBack_Click()
    StopRun = False
    KnifeBackMM 20
End Sub

Private Sub CmdKnifeFeed_Click()
    StopRun = False
    KnifeFeedMM 20
End Sub

Private Sub CmdReset2_Click()
    Dim BendHeadOffsetD0 As Double
    
    BendHeadOffsetD0 = BendHeadOffsetD
    BendHeadOffsetD = 0
    Reset2
    BendHeadOffsetD = BendHeadOffsetD0
    
    Reset2WithoutOffset = True
End Sub

Private Sub CmdSave_Click()
    Dim I As Long, t As Long
    
    WriteToINI "PP100MM", TxtPP100MM.Text
    WriteToINI "MaxSpeedMMPS", TxtMaxSpeedMMPS.Text
    
    WriteToINI "FeedMaxMM", TxtFeedMaxMM.Text
    WriteToINI "HoldDelay", TxtHoldDelay.Text
    
    WriteToINI "StartV", TxtStartV.Text
    WriteToINI "Speed", TxtSpeed.Text
    WriteToINI "Accl", TxtAccl.Text
    
    WriteToINI "BackStartV", TxtBackStartV.Text
    WriteToINI "BackSpeed", TxtBackSpeed.Text
    WriteToINI "BackAccl", TxtBackAccl.Text
    
    '------------------------------------------------------
    
    WriteToINI "PPD", TxtPPD.Text
    WriteToINI "MaxSpeedDPS", TxtMaxSpeedDPS.Text
    
    WriteToINI "BendHeadOffsetD", TxtBendHeadOffsetD.Text
    WriteToINI "LeftGapD", TxtLeftGapD.Text
    WriteToINI "RightGapD", TxtRightGapD.Text
    
    'WriteToINI "SubFactor", TxtSubFactor.Text
    WriteToINI "BreakingAngle", TxtBreakingAngle.Text
    
    WriteToINI "StartV2", TxtStartV2.Text
    WriteToINI "Speed2", TxtSpeed2.Text
    WriteToINI "Accl2", TxtAccl2.Text
    
    WriteToINI "BackStartV2", TxtBackStartV2.Text
    WriteToINI "BackSpeed2", TxtBackSpeed2.Text
    WriteToINI "BackAccl2", TxtBackAccl2.Text
    
    '------------------------------------------------------
    
    WriteToINI "StartV3", TxtStartV3.Text
    WriteToINI "Speed3", TxtSpeed3.Text
    WriteToINI "Accl3", TxtAccl3.Text
    
    WriteToINI "CutHole", IIf(Option2.value = True, "2", "3")
    '------------------------------------------------------

    WriteToINI "LenMM", TxtLenMM.Text
    WriteToINI "CutMM", TxtCutMM.Text
    WriteToINI "CutWMM", TxtCutWMM.Text
    WriteToINI "HoleMM1", TxtHoleMM1.Text
    WriteToINI "HoleMM2", TxtHoleMM2.Text
    WriteToINI "HoleMM3", TxtHoleMM3.Text
    WriteToINI "HoleWMM1", TxtHoleWMM1.Text
    WriteToINI "HoleWMM2", TxtHoleWMM2.Text
    WriteToINI "HoleWMM3", TxtHoleWMM3.Text

    CmdSortAngleTable_Click
    
    For t = 1 To MaxBendDisNo
        BendDis(t) = Val(TxtBendDis(t).Text)
    Next
    
    SupplementKeyCount = 0
    For I = 1 To GrdAngleTable.Rows - 1
        If Val(GrdAngleTable.TextMatrix(I, 1)) <> 0 Then
            SupplementKeyCount = I
            KeyAngle(I) = Val(GrdAngleTable.TextMatrix(I, 1))
            For t = 1 To MaxBendDisNo
                RealAngle(t, I) = Val(GrdAngleTable.TextMatrix(I, t + 1))
                SupAngle(t, I) = KeyAngle(I) - RealAngle(t, I)
            Next
        Else
            Exit For
        End If
    Next
    
    For t = 1 To MaxBendDisNo
        WriteToINI_A "Gap" & Trim(Str(t)), Str(BendDis(t))
    Next
    
    WriteToINI_A "SupplementKeyCount", Str(SupplementKeyCount)
    For I = 1 To SupplementKeyCount
        WriteToINI_A "Key" & Trim(Str(I)), Str(KeyAngle(I))
        For t = 1 To MaxBendDisNo
            WriteToINI_A "Real" & Trim(Str(I)) & "_" & Trim(Str(t)), Str(RealAngle(t, I))
        Next
    Next
           
    WriteToINI "TotalWorkLength", Str(TotalWorkLength)
    WriteToINI "TotalWorkBendCount", Str(TotalWorkBendCount)
    WriteToINI "TotalWorkCount", Str(TotalWorkCount)
    WriteToINI "TotalWorkTime", Str(TotalWorkTime)
           
    Unload Me
End Sub

Private Sub CmdSetGap_Click()
    If Reset2WithoutOffset = False Then
        MsgBox "请先进行[无偏移复位]，然后再执行本功能。  ", vbInformation + vbOKOnly + vbSystemModal, ""
        Exit Sub
    End If
       
    TxtBendHeadOffsetD.Text = TxtBendHeadOffsetUpdatedD.Text
    TxtLeftGapD.Text = TxtLeftGapUpdatedD.Text
    TxtRightGapD.Text = TxtRightGapUpdatedD.Text
End Sub

Private Sub CmdSortAngleTable_Click()
    Dim r As Long, r2 As Long, c As Long, c2 As Long, a As Double, b As Double, s As String
    
    For c = 1 To MaxBendDisNo
        If Val(TxtBendDis(c).Text) = 0 Then
            TxtBendDis(c).Text = Str(10000 + c)
        End If
    Next
    
    For c = 1 To MaxBendDisNo - 1
        For c2 = c + 1 To MaxBendDisNo
            a = Val(TxtBendDis(c).Text)
            b = Val(TxtBendDis(c2).Text)
                
            If a > b Then
                s = TxtBendDis(c).Text
                TxtBendDis(c).Text = TxtBendDis(c2).Text
                TxtBendDis(c2).Text = s
                
                For r = 1 To GrdAngleTable.Rows - 1
                    s = GrdAngleTable.TextMatrix(r, c + 1)
                    GrdAngleTable.TextMatrix(r, c + 1) = GrdAngleTable.TextMatrix(r, c2 + 1)
                    GrdAngleTable.TextMatrix(r, c2 + 1) = s
                Next
            End If
        Next
    Next
    
    For c = 1 To MaxBendDisNo
        If Val(TxtBendDis(c).Text) > 10000 Then
            TxtBendDis(c).Text = ""
        End If
    Next
    
    
    For r = 1 To GrdAngleTable.Rows - 2
        For r2 = r + 1 To GrdAngleTable.Rows - 1
            a = Val(GrdAngleTable.TextMatrix(r, 1))
            b = Val(GrdAngleTable.TextMatrix(r2, 1))
            
            If a > b Then
                s = GrdAngleTable.TextMatrix(r, 1)
                GrdAngleTable.TextMatrix(r, 1) = GrdAngleTable.TextMatrix(r2, 1)
                GrdAngleTable.TextMatrix(r2, 1) = s
                
                s = GrdAngleTable.TextMatrix(r, 2)
                GrdAngleTable.TextMatrix(r, 2) = GrdAngleTable.TextMatrix(r2, 2)
                GrdAngleTable.TextMatrix(r2, 2) = s
                
                s = GrdAngleTable.TextMatrix(r, 3)
                GrdAngleTable.TextMatrix(r, 3) = GrdAngleTable.TextMatrix(r2, 3)
                GrdAngleTable.TextMatrix(r2, 3) = s
                
                s = GrdAngleTable.TextMatrix(r, 4)
                GrdAngleTable.TextMatrix(r, 4) = GrdAngleTable.TextMatrix(r2, 4)
                GrdAngleTable.TextMatrix(r2, 4) = s
                
                s = GrdAngleTable.TextMatrix(r, 5)
                GrdAngleTable.TextMatrix(r, 5) = GrdAngleTable.TextMatrix(r2, 5)
                GrdAngleTable.TextMatrix(r2, 5) = s
                
                s = GrdAngleTable.TextMatrix(r, 6)
                GrdAngleTable.TextMatrix(r, 6) = GrdAngleTable.TextMatrix(r2, 6)
                GrdAngleTable.TextMatrix(r2, 6) = s
            End If
        Next
    Next
    
    For r = GrdAngleTable.Rows - 1 To 1 Step -1
        If Val(GrdAngleTable.TextMatrix(r, 1)) = 0 Then
            Exit For
        End If
    Next
    If r < GrdAngleTable.Rows - 1 And r > 0 Then
        r2 = r
        
        For r = r2 + 1 To GrdAngleTable.Rows - 1
            s = GrdAngleTable.TextMatrix(r, 1)
            GrdAngleTable.TextMatrix(r, 1) = GrdAngleTable.TextMatrix(r - r2, 1)
            GrdAngleTable.TextMatrix(r - r2, 1) = s
            
            s = GrdAngleTable.TextMatrix(r, 2)
            GrdAngleTable.TextMatrix(r, 2) = GrdAngleTable.TextMatrix(r - r2, 2)
            GrdAngleTable.TextMatrix(r - r2, 2) = s
        
            s = GrdAngleTable.TextMatrix(r, 3)
            GrdAngleTable.TextMatrix(r, 3) = GrdAngleTable.TextMatrix(r - r2, 3)
            GrdAngleTable.TextMatrix(r - r2, 3) = s
            
            s = GrdAngleTable.TextMatrix(r, 4)
            GrdAngleTable.TextMatrix(r, 4) = GrdAngleTable.TextMatrix(r - r2, 4)
            GrdAngleTable.TextMatrix(r - r2, 4) = s
            
            s = GrdAngleTable.TextMatrix(r, 5)
            GrdAngleTable.TextMatrix(r, 5) = GrdAngleTable.TextMatrix(r - r2, 5)
            GrdAngleTable.TextMatrix(r - r2, 5) = s
        
            s = GrdAngleTable.TextMatrix(r, 6)
            GrdAngleTable.TextMatrix(r, 6) = GrdAngleTable.TextMatrix(r - r2, 6)
            GrdAngleTable.TextMatrix(r - r2, 6) = s
        Next
    End If
    
    For r = 1 To GrdAngleTable.Rows - 1
        If Val(GrdAngleTable.TextMatrix(r, 1)) = 0 Then
            GrdAngleTable.TextMatrix(r, 0) = ""
            GrdAngleTable.TextMatrix(r, 2) = ""
            GrdAngleTable.TextMatrix(r, 3) = ""
            GrdAngleTable.TextMatrix(r, 4) = ""
            GrdAngleTable.TextMatrix(r, 5) = ""
            GrdAngleTable.TextMatrix(r, 6) = ""
        Else
            GrdAngleTable.TextMatrix(r, 0) = Str(r)
        End If
    Next
End Sub


Private Sub CmdTurnNegative_Click()
    Dim d As Double
    
    If Reset2WithoutOffset = False Then
        MsgBox "请先进行[无偏移复位]，然后再执行本功能。  ", vbInformation + vbOKOnly + vbSystemModal, ""
        Exit Sub
    End If
       
    CmdTurnNegative.Enabled = False
    CmdTurnNegative.BackColor = RGB(255, 0, 0)
    d = -Val(TxtRightGapTestD.Text)
    KnifeBendHeadMoveD d
    CmdTurnNegative.BackColor = CmdSave.BackColor
    CmdTurnNegative.Enabled = True
End Sub

Private Sub CmdTurnPositive_Click()
    Dim d As Double
    
    If Reset2WithoutOffset = False Then
        MsgBox "请先进行[无偏移复位]，然后再执行本功能。  ", vbInformation + vbOKOnly + vbSystemModal, ""
        Exit Sub
    End If
       
    CmdTurnPositive.Enabled = False
    CmdTurnPositive.BackColor = RGB(255, 0, 0)
    d = Val(TxtLeftGapTestD.Text)
    KnifeBendHeadMoveD d
    CmdTurnPositive.BackColor = CmdSave.BackColor
    CmdTurnPositive.Enabled = True
End Sub

Private Sub Command1_Click()
    If Me.Width <= 762 * Screen.TwipsPerPixelX Then
        Me.Width = 1150 * Screen.TwipsPerPixelX
    Else
        Me.Width = 762 * Screen.TwipsPerPixelX
    End If
    Me.Move (Screen.Width - Me.Width) / 2
End Sub

Private Sub Command2_Click()
    FrmShowCurve.Show
End Sub

Private Sub Command3_Click()
    TotalWorkLength = 0
    TotalWorkBendCount = 0
    TotalWorkCount = 0
    TotalWorkTime = 0
    
    TxtTotalWorkLength.Text = Format(TotalWorkLength / 1000, "0.0")
    TxtTotalWorkBendCount.Text = Str(TotalWorkBendCount)
    TxtTotalWorkCount.Text = Str(TotalWorkCount)
    TxtTotalWorkTime.Text = Format(TotalWorkTime / 3600, "0.0")
End Sub

Private Sub Form_Load()
    Dim I As Long, t As Long
    
    LoadParameters

    TxtPP100MM.Text = Str(PP100MM)
    TxtMaxSpeedMMPS.Text = Str(MaxSpeedMMPS)
    
    'TxtFeedMaxMM.Text = Str(FeedMaxMM)
    TxtHoldDelay.Text = Str(HoldDelay)
    
    TxtStartV.Text = Str(startv)
    TxtSpeed.Text = Str(speed)
    TxtAccl.Text = Str(Accl)

    TxtBackStartV.Text = Str(BackStartV)
    TxtBackSpeed.Text = Str(BackSpeed)
    TxtBackAccl.Text = Str(BackAccl)

    TxtPPD.Text = Str(PPD)
    TxtMaxSpeedDPS.Text = Str(MaxSpeedDPS)
    
    TxtBendHeadOffsetD.Text = Format(BendHeadOffsetD, " 0.0###")
    TxtLeftGapD.Text = Str(LeftGapD)
    TxtRightGapD.Text = Str(RightGapD)
'    TxtSubFactor.Text = Format(SubFactor, " 0.0##")
    TxtBreakingAngle.Text = Str(BreakingAngle)
    
    TxtStartV2.Text = Str(startv2)
    TxtSpeed2.Text = Str(speed2)
    TxtAccl2.Text = Str(Accl2)
    
    TxtBackStartV2.Text = Str(BackStartV2)
    TxtBackSpeed2.Text = Str(BackSpeed2)
    TxtBackAccl2.Text = Str(BackAccl2)
    
    TxtStartV3.Text = Str(startv3)
    TxtSpeed3.Text = Str(speed3)
    TxtAccl3.Text = Str(Accl3)
    Option2.value = IIf(CutHole = 2, True, False)
    Option3.value = IIf(CutHole = 3, True, False)
    
    TxtLenMM.Text = Str(LenMM)
    TxtCutMM.Text = Str(CutMM)
    TxtCutWMM.Text = Str(CutWMM)
    TxtHoleMM1.Text = Str(HoleMM1)
    TxtHoleMM2.Text = Str(HoleMM2)
    TxtHoleMM3.Text = Str(HoleMM3)
    TxtHoleWMM1.Text = Str(HoleWMM1)
    TxtHoleWMM2.Text = Str(HoleWMM2)
    TxtHoleWMM3.Text = Str(HoleWMM3)

    '----------------------------------------------------------------
    
    TxtLeftGapTestD.Text = GetFromINI("LeftGapTestD")
    TxtRightGapTestD.Text = GetFromINI("RightGapTestD")

    '----------------------------------------------------------------
    
    For t = 1 To MaxBendDisNo
        TxtBendDis(t).Text = IIf(BendDis(t) = 0, "", Str(BendDis(t)))
    Next
    
    GrdAngleTable.Cols = 2 + MaxBendDisNo
    GrdAngleTable.Rows = 101
    GrdAngleTable.Clear
    GrdAngleTable.ColWidth(0) = 27 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(1) = 62 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(2) = 48 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(3) = 48 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(4) = 48 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(5) = 48 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(6) = 48 * Screen.TwipsPerPixelX
    GrdAngleTable.RowHeightMin = 18 * Screen.TwipsPerPixelY
    
    GrdAngleTable.ColAlignment(1) = 1
    GrdAngleTable.ColAlignment(2) = 1
    
    GrdAngleTable.TextMatrix(0, 0) = "No."
    GrdAngleTable.TextMatrix(0, 1) = "运动量"
    GrdAngleTable.TextMatrix(0, 2) = "实现量1"
    GrdAngleTable.TextMatrix(0, 3) = "实现量2"
    GrdAngleTable.TextMatrix(0, 4) = "实现量3"
    GrdAngleTable.TextMatrix(0, 5) = "实现量4"
    GrdAngleTable.TextMatrix(0, 6) = "实现量5"
    
    For I = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(I, 0) = Str(I)
        GrdAngleTable.TextMatrix(I, 1) = Format(KeyAngle(I), " 0.0###")
        For t = 1 To MaxBendDisNo
            GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0 And BendDis(t) = 0, "", Format(RealAngle(t, I), " 0.0###"))
        Next
    Next
    '----------------------------------------------------------------
    
    TxtTotalWorkLength.Text = Format(TotalWorkLength / 1000, "0.0")
    TxtTotalWorkBendCount.Text = Str(TotalWorkBendCount)
    TxtTotalWorkCount.Text = Str(TotalWorkCount)
    TxtTotalWorkTime.Text = Format(TotalWorkTime / 3600, "0.0")
    
    '----------------------------------------------------------------
    ChkLocked.value = 1
    Reset2WithoutOffset = False
    
    Me.Width = 762 * Screen.TwipsPerPixelX
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_Flags
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim pos As Long
    
    If Reset2WithoutOffset = True Then
        get_command_pos 0, 2, pos
        set_command_pos 0, 2, -BendHeadOffsetD * PPD + pos
    End If
End Sub

Private Sub GrdAngleTable_EnterCell()
    CurRow = GrdAngleTable.Row
    CurCol = GrdAngleTable.Col

    GrdAngleTable.RowSel = CurRow
    GrdAngleTable.ColSel = CurCol
    
    GrdAngleTable.ForeColorSel = RGB(255, 255, 255)
End Sub

Private Sub GrdAngleTable_KeyPress(KeyAscii As Integer)
    Dim txt As String
    
    CurRow = GrdAngleTable.Row
    CurCol = GrdAngleTable.Col

    GrdAngleTable.RowSel = CurRow
    GrdAngleTable.ColSel = CurCol
    
    txt = GrdAngleTable.TextMatrix(CurRow, CurCol)
    If KeyAscii = 8 Then
        If Len(txt) > 0 Then
            txt = Mid(txt, 1, Len(txt) - 1)
        End If
    ElseIf KeyAscii = 13 Then
        If CurCol < 2 Then
            GrdAngleTable.Col = CurCol + 1
        ElseIf CurRow < GrdAngleTable.Rows - 1 Then
            GrdAngleTable.Row = CurRow + 1
            GrdAngleTable.Col = 1
        End If
        Exit Sub
    Else
        txt = txt & Chr(KeyAscii)
    End If
    GrdAngleTable.TextMatrix(CurRow, CurCol) = txt
End Sub

Private Sub GrdAngleTable_LeaveCell()
    CurRow = GrdAngleTable.Row
    CurCol = GrdAngleTable.Col
    
    If Val(GrdAngleTable.TextMatrix(CurRow, 1)) < 2 And Trim(GrdAngleTable.TextMatrix(CurRow, CurCol)) <> "" Then
        If CurCol = 1 Then
            'PosList(CurRow - 1).d1 = Val(GrdAngleTable.TextMatrix(CurRow, CurCol)) - DPx
        ElseIf CurCol = 2 Then
            'PosList(CurRow - 1).d2 = Val(GrdAngleTable.TextMatrix(CurRow, CurCol)) - DPy
        End If
    End If
End Sub

Private Sub TxtLeftGapTestD_Change()
    Dim lv As Double, rv As Double
    
    lv = Val(TxtLeftGapTestD.Text)
    rv = Val(TxtRightGapTestD.Text)
    
    TxtBendHeadOffsetUpdatedD.Text = Str((lv - rv) / 2)
    TxtLeftGapUpdatedD.Text = Str(lv - (lv - rv) / 2)
    TxtRightGapUpdatedD.Text = Str(rv + (lv - rv) / 2)
    
    WriteToINI "LeftGapTestD", TxtLeftGapTestD.Text
End Sub

Private Sub TxtRightGapTestD_Change()
    Dim lv As Double, rv As Double
    
    lv = Val(TxtLeftGapTestD.Text)
    rv = Val(TxtRightGapTestD.Text)
    
    TxtBendHeadOffsetUpdatedD.Text = Str((lv - rv) / 2)
    TxtLeftGapUpdatedD.Text = Str(lv - (lv - rv) / 2)
    TxtRightGapUpdatedD.Text = Str(rv + (lv - rv) / 2)
    
    WriteToINI "RightGapTestD", TxtRightGapTestD.Text
End Sub
