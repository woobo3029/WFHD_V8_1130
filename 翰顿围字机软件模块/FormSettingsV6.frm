VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14940
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   996
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton OptSymmetryTest 
      Caption         =   "对称性(不含空程角)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12780
      TabIndex        =   117
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CheckBox ChkLocked 
      Caption         =   "锁定参数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   116
      Top             =   7380
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.OptionButton OptBeatR 
      Caption         =   "右拍弧"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11820
      TabIndex        =   113
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton CmdCalculateAhead 
      Caption         =   "计算提前量(pulse)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6660
      TabIndex        =   111
      Top             =   3960
      Width           =   1755
   End
   Begin VB.OptionButton OptBendR 
      Caption         =   "右弯弧"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9900
      TabIndex        =   110
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton CmdMake 
      Caption         =   "制作样本"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13260
      TabIndex        =   63
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox TxtFeedMM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12180
      TabIndex        =   62
      Top             =   7440
      Width           =   495
   End
   Begin VB.TextBox TxtAngleDeg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10380
      TabIndex        =   61
      Top             =   7440
      Width           =   495
   End
   Begin VB.OptionButton OptBeatL 
      Caption         =   "左拍弧"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10860
      TabIndex        =   60
      Top             =   7080
      Width           =   1095
   End
   Begin VB.OptionButton OptTurn 
      Caption         =   "折角"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12660
      TabIndex        =   59
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton OptBendL 
      Caption         =   "左弯弧"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8940
      TabIndex        =   58
      Top             =   7080
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "运动角度(deg)/弯弧半径(mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   9000
      TabIndex        =   50
      Top             =   60
      Width           =   5775
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   140
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   139
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   138
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   137
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   136
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   135
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtMaterialThickMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   134
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtMaterialName 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   114
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox CmbMaterial 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FormSettingsV6.frx":0000
         Left            =   1200
         List            =   "FormSettingsV6.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   360
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   71
         Top             =   6420
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox TxtR 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2760
            TabIndex        =   77
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton CmdR 
            Caption         =   "半径"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2160
            TabIndex        =   76
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox TxtL 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   75
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox TxtA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   73
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "(mm)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3330
            TabIndex        =   78
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label23 
            Caption         =   "弦长"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   74
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "弧长"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.CommandButton CmdSortAngleTable 
         Caption         =   "排序"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   6540
         Width           =   675
      End
      Begin VB.CommandButton CmdShowCurve 
         Caption         =   "显示参数曲线"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   51
         Top             =   6540
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid GrdAngleTable 
         Height          =   5115
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   9022
         _Version        =   393216
         Rows            =   500
         Cols            =   7
         ScrollBars      =   2
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "型材厚度(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "编辑名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   115
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "选择型材"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "不保存"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   7380
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
      Begin VB.Frame Frame4 
         Height          =   3015
         Left            =   7020
         TabIndex        =   156
         Top             =   240
         Width           =   1635
         Begin VB.CheckBox ChkAmericanMaterial 
            Caption         =   "A-型材"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   161
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox TxtTailVertAngle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   160
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox TxtVertUpDownMM_A 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   159
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox TxtExtendMM 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   158
            Top             =   1200
            Width           =   675
         End
         Begin VB.CheckBox ChkKareanMaterial 
            Caption         =   "K-型材"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   157
            Top             =   2280
            Width           =   915
         End
         Begin VB.Label Label45 
            Caption         =   "末端角度(Deg)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   164
            Top             =   1500
            Width           =   1275
         End
         Begin VB.Label Label52 
            Caption         =   "升降(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   163
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "拼接段(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   162
            Top             =   1020
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtInnerLineTerminalAdjustMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   153
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TxtOuterLineTerminalAdjustMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   151
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TxtBenderSpringback 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   142
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtBenderBacklash 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   130
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtOuterAngleAdjustMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   128
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox TxtInnerAngleAdjustMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   126
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox TxtEmptyDegree2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6120
         TabIndex        =   112
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtVertUpDownMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   107
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtVertUpDownAdjustmentMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   104
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox TxtVertUpDownPulsPerMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   102
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtTurnPointOffsetMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8340
         TabIndex        =   94
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtVertAdjustmentDegree 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   86
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox TxtVertPulsPerDegree 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   84
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox ChkVertNoTurn 
         Caption         =   "铣角(否则为折角)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   79
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtEmptyDegree 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   27
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtHeadDistance 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   25
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtAdjustmentDegree 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox ChkUseEncoder 
         Caption         =   "使用编码器"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtEncoderPulsPerMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   8535
         Begin VB.TextBox TxtVertMotorZoneMM 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   155
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox TxtDoneWaitingTime 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6480
            TabIndex        =   148
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox TxtResetVertAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5160
            TabIndex        =   146
            Top             =   2100
            Width           =   975
         End
         Begin VB.TextBox TxtResetVertSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   144
            Top             =   2100
            Width           =   975
         End
         Begin VB.TextBox TxtResetVertStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   143
            Top             =   2100
            Width           =   975
         End
         Begin VB.TextBox TxtFastSpeedMinLenMM 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   132
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox TxtVertMaxOuterAngle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5400
            TabIndex        =   100
            Top             =   3420
            Width           =   735
         End
         Begin VB.TextBox TxtVertMaxInnerAngle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   97
            Top             =   3420
            Width           =   735
         End
         Begin VB.TextBox TxtVertKnifeDegree 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5400
            TabIndex        =   96
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox TxtVertAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   91
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox TxtVertSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3240
            TabIndex        =   89
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox TxtVertStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   88
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox TxtDoneDistance 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   82
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox TxtCutRadiusMM 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6480
            TabIndex        =   81
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox TxtTurnFeedAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   70
            Top             =   1500
            Width           =   975
         End
         Begin VB.TextBox TxtTurnFeedSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3240
            TabIndex        =   68
            Top             =   1500
            Width           =   975
         End
         Begin VB.TextBox TxtTurnFeedStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   67
            Top             =   1500
            Width           =   975
         End
         Begin VB.TextBox TxtTurnFeedMM 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1860
            TabIndex        =   57
            Top             =   4980
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtVertMinDistance 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   55
            Top             =   3060
            Width           =   735
         End
         Begin VB.TextBox TxtFeedOffset 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6480
            TabIndex        =   44
            Top             =   900
            Width           =   855
         End
         Begin VB.TextBox TxtBeatMaxRadius 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5400
            TabIndex        =   43
            Top             =   3060
            Width           =   735
         End
         Begin VB.TextBox TxtVertMinAngle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   40
            Top             =   2700
            Width           =   735
         End
         Begin VB.TextBox TxtResetBendAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   38
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtResetBendSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3240
            TabIndex        =   37
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtManualFeedAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   -540
            TabIndex        =   34
            Top             =   2340
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtManualFeedSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   -720
            TabIndex        =   33
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtManualFeedStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   -660
            TabIndex        =   32
            Top             =   2340
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtResetBendStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   31
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtManualBendAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   30
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox TxtManualBendSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3240
            TabIndex        =   29
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox TxtManualBendStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   28
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox TxtBendAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtBendSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3240
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtBendStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtFeedAccel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   -720
            TabIndex        =   14
            Top             =   2340
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtFeedSpeed 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   -780
            TabIndex        =   13
            Top             =   2340
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtFeedStartV 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   -540
            TabIndex        =   12
            Top             =   2280
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label51 
            Caption         =   "铣刀提前启动距离(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6420
            TabIndex        =   154
            Top             =   2700
            Width           =   1935
         End
         Begin VB.Label Label47 
            Caption         =   "分段等待时间(s)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6420
            TabIndex        =   149
            Top             =   2220
            Width           =   1875
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            Caption         =   "铣角器复位速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   147
            Top             =   2160
            Width           =   1515
            WordWrap        =   -1  'True
         End
         Begin VB.Label LblResetVertSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   145
            Top             =   2100
            Width           =   975
         End
         Begin VB.Label Label44 
            Caption         =   "进料减速距离(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6420
            TabIndex        =   131
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "铣角器升降速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   660
            TabIndex        =   101
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "铣外角最大角度(Degree)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   99
            Top             =   3420
            Width           =   2235
         End
         Begin VB.Label Label31 
            Caption         =   "铣内角最大角度(Degree)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   98
            Top             =   3420
            Width           =   2055
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "铣刀刀锋角度(Degree)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   95
            Top             =   2700
            Width           =   2235
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "铣角器旋转速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            TabIndex        =   92
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label LblVertSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   90
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "分段间隔/尾端进料(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6420
            TabIndex        =   83
            Top             =   1740
            Width           =   1995
         End
         Begin VB.Label Label25 
            Caption         =   "雕刻刀半径(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6420
            TabIndex        =   80
            Top             =   3180
            Width           =   2055
         End
         Begin VB.Label LblTurnFeedSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   69
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "折角进料速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -660
            TabIndex        =   66
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "折角后进料距离(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   4980
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "最短铣角间距(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   54
            Top             =   3060
            Width           =   2115
         End
         Begin VB.Label LblResetBendSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   49
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label LblManualBendSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   48
            Top             =   900
            Width           =   975
         End
         Begin VB.Label LblBendSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   47
            Top             =   600
            Width           =   975
         End
         Begin VB.Label LblManualFeedSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -660
            TabIndex        =   46
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LblFeedSpeed 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -840
            TabIndex        =   45
            Top             =   2400
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "弯弧最小半径(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3060
            TabIndex        =   42
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "铣角起始角度(Degree)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   41
            Top             =   2640
            Width           =   2115
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "弯弧复位速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            TabIndex        =   39
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "人工弯弧速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            TabIndex        =   36
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "人工进料速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -540
            TabIndex        =   35
            Top             =   2460
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "加速度(pulse/s2)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5100
            TabIndex        =   21
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "驱动速度(pulse/s)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3420
            TabIndex        =   20
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "初始速度(pulse/s)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1500
            TabIndex        =   19
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "程序弯弧速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            TabIndex        =   18
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "程序进料速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -840
            TabIndex        =   11
            Top             =   2340
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtPulsPerDegree 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtPulsPerMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "闭合内轮廓端点补偿(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   152
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "闭合外轮廓端点补偿(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "弯弧器拍弧回弹系数(0-1)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "弯弧器反向间隙(Degree)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "铣外角长度补偿(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   127
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "铣内角长度补偿(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "升降行程(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   106
         Top             =   2295
         Width           =   1935
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "复位调整距离(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   105
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "铣角器升降每毫米脉冲数(Pulse/mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   103
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label29 
         Caption         =   "折角点相对弯弧点偏移(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8220
         TabIndex        =   93
         Top             =   3180
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "复位调整角度(Degree)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   87
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "铣角器旋转每角度脉冲数(Pulse/Degree)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   85
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "左、右弯弧空程角度(Degree)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   26
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "铣角点到弯弧点的距离(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   2295
         Width           =   2295
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "复位调整角度(Degree)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "编码器每毫米脉冲数(Pulse/mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "弯弧器每角度脉冲数(Pulse/Degree)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "进料电机每毫米脉冲数(Puls/mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Label LblString7 
      Caption         =   "步数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   124
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label LblString6 
      Caption         =   "弧长(mm)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   123
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Label LblString5 
      Caption         =   "右拍弧角度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15540
      TabIndex        =   122
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label LblString4 
      Caption         =   "左拍弧角度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15420
      TabIndex        =   121
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblString3 
      Caption         =   "右弯弧半径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15420
      TabIndex        =   120
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblString2 
      Caption         =   "左弯弧半径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   119
      Top             =   1500
      Width           =   1455
   End
   Begin VB.Label LblString1 
      Caption         =   "运动角度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15300
      TabIndex        =   118
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label LblFeedMM 
      Alignment       =   1  'Right Justify
      Caption         =   "弧长(mm)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11340
      TabIndex        =   65
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label LblAngle 
      Alignment       =   1  'Right Justify
      Caption         =   "运动角度(deg)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9060
      TabIndex        =   64
      Top             =   7440
      Width           =   1215
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkAmericanMaterial_Click()
    ChkKareanMaterial.value = 0
End Sub

Private Sub ChkKareanMaterial_Click()
    ChkAmericanMaterial.value = 0
End Sub

Private Sub ChkLocked_Click()
    Dim obj As Object
    If ChkLocked.value = 1 Then
        For Each obj In Me
            If (TypeOf obj Is TextBox) Or (TypeOf obj Is OptionButton) Or (TypeOf obj Is CommandButton) Or (TypeOf obj Is MSFlexGrid) Then
                obj.Enabled = False
            ElseIf TypeOf obj Is CheckBox Then
                If obj.Name <> "ChkLocked" Then
                    obj.Enabled = False
                End If
            End If
        Next
    Else
        If MsgBox("错误的参数设置将导致设备运行异常。请确定是否放弃该操作？ ", vbQuestion + vbYesNo + vbSystemModal, "") = vbNo Then
            For Each obj In Me
                obj.Enabled = True
            Next
            'ShowZAxisMode
            'ShowHeadMode
        Else
            ChkLocked.value = 1
        End If
    End If
End Sub

Private Sub CmbMaterial_Click()
    Dim i As Long, t As Long
    
    Device_CurMaterial = "Material" + Format(CmbMaterial.ListIndex, "00")
    TxtMaterialName.Text = CmbMaterial.List(CmbMaterial.ListIndex)
    
    WritePrivateProfileString "Device", "CurMaterial", Device_CurMaterial, App.Path & "\Parameters.ini"
    
    For i = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(i, 0) = ""
        GrdAngleTable.TextMatrix(i, 1) = ""
        For t = 1 To MaxBendDisNo
            GrdAngleTable.TextMatrix(i, t + 1) = ""
        Next
    Next
    
    LoadParameters
    For i = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(i, 0) = str(i)
        GrdAngleTable.TextMatrix(i, 1) = Format(KeyAngle(i), " 0.0###")
        For t = 1 To MaxBendDisNo
            'GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0 And BendDis(t) = 0, "", Format(RealAngle(t, I), " 0.0###"))
            GrdAngleTable.TextMatrix(i, t + 1) = IIf(RealAngle(t, i) = 0, "", Format(RealAngle(t, i), " 0.0###"))
        Next
    Next
    
    Device_MaterialThickMM = GetValueFromINI("MaterialThickMM", Device_CurMaterial, "0.8", App.Path & "\Parameters.ini")
    TxtMaterialThickMM.Text = Format(Device_MaterialThickMM, " 0.0###")
End Sub


Private Sub CmdCalculateAhead_Click()
    Dim i As Long, ahead As Double, ahead_sum As Double
    Dim Ret As Long, nActPos As Long, Pos0 As Long, FeedPulsCount As Long, n As Long
           
    FrmMsgDlg.LblMessage.caption = "请上好型材。本功能将自动进料 10 次"
    FrmMsgDlg.CmdClose.caption = "确定"
    FrmMsgDlg.Show
    
    n = 10
    
    Do While FrmMsgDlg.Visible = True
        DoEvents
    Loop
    
    If Device_FastSpeedMinLenMM <= 50 Then
        FeedPulsCount = Device_FastSpeedMinLenMM * Device_EncoderPulsPerMM
    Else
        FeedPulsCount = 50 * Device_EncoderPulsPerMM
    End If
    
    FrmMain.TxtStatistics.Text = "停止所需脉冲：" + vbCrLf + vbCrLf
    ahead_sum = 0
    For i = 1 To n
        get_actual_pos 0, FeedAxis, nActPos
        Pos0 = nActPos
        
        DCMotorFeedFWOn
        
        Do
            get_actual_pos 0, FeedAxis, nActPos
            If nActPos - Pos0 >= FeedPulsCount Then
                Pos0 = nActPos
                Exit Do
            End If
            DoEvents
        Loop
        
        DCMotorFeedFWOff
        
        Wait 2
        
        get_actual_pos 0, FeedAxis, nActPos
        ahead = nActPos - Pos0
        FrmMain.TxtStatistics.Text = FrmMain.TxtStatistics.Text + str(ahead) + vbCrLf

        ahead_sum = ahead_sum + ahead
    Next
    
    TxtFeedOffset.Text = str(Round(ahead_sum / n, 1))
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdMake_Click()
    Dim AngleDEG As Double, FeedMM As Double, n As Long, i As Long, s As Double, dw As Double
           
    AngleDEG = Val(TxtAngleDeg.Text)
    'AngleDEG = AngleDEG + IIf(AngleDEG > 0, Device_EmptyDegree, IIf(AngleDEG = 0, 0, -Device_EmptyDegree))
    FeedMM = Val(TxtFeedMM.Text)
    'n = Val(TxtN.Text)
    
    StopRunning = False
    IsRunning = True
    
    '复位
    BendReset
    ''铣槽-高位
    'FrmMain.Vert 1, 1
    VertInnerAngle 0
    
    If OptBendL.value = True Or OptBendR.value = True Then
    
        If OptBendL.value = True Then
            AngleDEG = -AngleDEG - Device_EmptyDegree
        Else
            AngleDEG = AngleDEG + Device_EmptyDegree2
        End If
        
        If FeedMM > Device_HeadDistance Then
            '进料
            FeedMMByDCMotor Device_HeadDistance, 0, False
            '弯弧
            BendAngle AngleDEG
            '进料
            FeedMMByDCMotor FeedMM - Device_HeadDistance, 0, False
            ''铣槽-高位
            'FrmMain.Vert 1, 1
            VertInnerAngle 0
            '进料
            FeedMMByDCMotor Device_HeadDistance, 0, False
            '弯弧
            BendAngle 0
        Else
            '进料
            FeedMMByDCMotor FeedMM, 0, False
            ''铣槽-高位
            'FrmMain.Vert 1, 1
            VertInnerAngle 0
            '进料
            FeedMMByDCMotor Device_HeadDistance - FeedMM, 0, False
            '弯弧
            BendAngle AngleDEG
            '进料
            FeedMMByDCMotor FeedMM, 0, False
            '弯弧
            BendAngle 0
        End If
        
    ElseIf OptBeatL.value = True Or OptBeatR.value = True Then
    
        If OptBeatL.value = True Then
            AngleDEG = -AngleDEG - Device_EmptyDegree
        Else
            AngleDEG = AngleDEG + Device_EmptyDegree2
        End If
        
        n = FeedMM
        FeedMM = 2
        s = n * FeedMM
        dw = 40
        
        'FrmMain.BendReset
        
        '进料
        FeedMMByDCMotor dw + s + dw, 0, False
        ''铣槽-高位
        'FrmMain.Vert 1, 1
        VertInnerAngle 0
        '进料
        FeedMMByDCMotor -s - dw + Device_HeadDistance - 10, 0, False
        Wait 1
        '拍弧
        For i = 1 To n
            BeatAngle AngleDEG
            FeedMMByDCMotor FeedMM, 0, False
        Next
        '进料
        FeedMMByDCMotor dw + 10, 0, False
        
    ElseIf OptTurn.value = True Then
        '进料
        FeedMMByDCMotor Device_HeadDistance / 4, 0, False
        ''铣槽-高位
        'FrmMain.Vert 1, 1
        VertInnerAngle 0
        '进料
        FeedMMByDCMotor Device_HeadDistance / 4, 0, False
        ''铣槽-高位
        'FrmMain.Vert 1, 1
        VertInnerAngle 0
        '进料
        FeedMMByDCMotor Device_HeadDistance - Device_HeadDistance / 4, 0, False
        BeatAngle AngleDEG, True
        '进料
        FeedMMByDCMotor Device_HeadDistance / 4, 0, False
        
    ElseIf OptSymmetryTest.value = True Then
        '进料
        FeedMMByDCMotor Device_HeadDistance, 0, False
        '铣槽
        VertInnerAngle 0
        '左弯弧
        BendAngle -AngleDEG
        '进料
        FeedMMByDCMotor Device_HeadDistance / 2, 0, False
        '右弯弧
        BendAngle AngleDEG
        '进料
        FeedMMByDCMotor Device_HeadDistance / 2, 0, False
        '弯弧
        BendAngle 0
    End If

    '进料
    FeedMMByDCMotor 50, 0, False
End Sub

Private Sub CmdR_Click()
    Dim a As Double, l As Double, r0 As Double, r1 As Double, r As Double, v0 As Double, v1 As Double, v As Double
    
    a = Val(TxtA.Text)
    l = Val(TxtL.Text)
    
    If a > l And l > 0 Then
        r0 = l / 3
        v0 = 2 * r0 * Sin(a / (2 * r0)) - l
        Do
            r1 = r1 + 0.0001
            v1 = 2 * r1 * Sin(a / (2 * r1)) - l
            If Sgn(v0) * Sgn(v1) <= 0 Then
                Exit Do
            End If
        Loop
'Debug.Print "v0,v1="; v0; v1

        Do While r < 2000
            r = (r0 + r1) / 2
            v = 2 * r * Sin(a / (2 * r)) - l
'Debug.Print "r,v="; r; v
            If Abs(v) <= 0.00001 Then
                Exit Do
            Else
                If Sgn(v0) = Sgn(v) Then
                    r0 = r
                Else
                    r1 = r
                End If
            End If
        Loop
        TxtR.Text = Trim(str(Round(r, 3)))
    End If
End Sub

Private Sub CmdSave_Click()
    Dim t As Long, i As Long
    
    Device_PulsPerMM = Val(TxtPulsPerMM.Text)
    Device_EncoderPulsPerMM = Val(TxtEncoderPulsPerMM.Text)
    Device_UseEncoder = IIf(ChkUseEncoder.value = 1, True, False)
    
    Device_PulsPerDegree = Val(TxtPulsPerDegree.Text)
    Device_AdjustmentDegree = Val(TxtAdjustmentDegree.Text)
    Device_EmptyDegree = Val(TxtEmptyDegree.Text)
    
    'Device_AdjustmentDegree2 = Val(TxtAdjustmentDegree2.Text)
    Device_EmptyDegree2 = Val(TxtEmptyDegree2.Text)
    
    'Device_VertMotorDrive = IIf(ChkVertMotorDrive.value = 1, True, False)
    'Device_VertAllHigh = IIf(ChkVertAllHigh.value = 1, True, False)
    Device_VertNoTurn = IIf(ChkVertNoTurn.value = 1, True, False)
    
    Device_VertUpDownPulsPerMM = Val(TxtVertUpDownPulsPerMM.Text)
    Device_VertUpDownAdjustmentMM = Val(TxtVertUpDownAdjustmentMM.Text)
    Device_VertUpDownMM = Val(TxtVertUpDownMM.Text)
    
    Device_VertPulsPerDegree = Val(TxtVertPulsPerDegree.Text)
    Device_VertAdjustmentDegree = Val(TxtVertAdjustmentDegree.Text)
    
    Device_HeadDistance = Val(TxtHeadDistance.Text)
    Device_DoneDistance = Val(TxtDoneDistance.Text)
    Device_DoneWaitingTime = Val(TxtDoneWaitingTime.Text)
    Device_ExtendMM = Val(TxtExtendMM.Text)
    
    'Device_WaitUpTime = Val(TxtWaitUpTime.Text)
    'Device_WaitDownTime = Val(TxtWaitDownTime.Text)
    
    Device_FeedStartV = Val(TxtFeedStartV.Text)
    Device_FeedSpeed = Val(TxtFeedSpeed.Text)
    Device_FeedAccel = Val(TxtFeedAccel.Text)
    Device_FeedOffset = Val(TxtFeedOffset.Text)
    
    'Device_ManualFeedStartV = Val(TxtManualFeedStartV.Text)
    'Device_ManualFeedSpeed = Val(TxtManualFeedSpeed.Text)
    'Device_ManualFeedAccel = Val(TxtManualFeedAccel.Text)
    'Device_ManualFeedOffset = Val(TxtManualFeedOffset.Text)
    
    Device_BendStartV = Val(TxtBendStartV.Text)
    Device_BendSpeed = Val(TxtBendSpeed.Text)
    Device_BendAccel = Val(TxtBendAccel.Text)
    
    Device_ManualBendStartV = Val(TxtManualBendStartV.Text)
    Device_ManualBendSpeed = Val(TxtManualBendSpeed.Text)
    Device_ManualBendAccel = Val(TxtManualBendAccel.Text)
    
    Device_ResetBendStartV = Val(TxtResetBendStartV.Text)
    Device_ResetBendSpeed = Val(TxtResetBendSpeed.Text)
    Device_ResetBendAccel = Val(TxtResetBendAccel.Text)
    
    'Device_TurnFeedStartV = Val(TxtTurnFeedStartV.Text)
    'Device_TurnFeedSpeed = Val(TxtTurnFeedSpeed.Text)
    'Device_TurnFeedAccel = Val(TxtTurnFeedAccel.Text)
    Device_VertUpDownStartV = Val(TxtTurnFeedStartV.Text)
    Device_VertUpDownSpeed = Val(TxtTurnFeedSpeed.Text)
    Device_VertUpDownAccel = Val(TxtTurnFeedAccel.Text)
    
    Device_TurnFeedStartV = Val(TxtVertStartV.Text)
    Device_TurnFeedSpeed = Val(TxtTurnFeedSpeed.Text)
    Device_TurnFeedAccel = Val(TxtTurnFeedAccel.Text)
    
    Device_VertStartV = Val(TxtVertStartV.Text)
    Device_VertSpeed = Val(TxtVertSpeed.Text)
    Device_VertAccel = Val(TxtVertAccel.Text)
    
    Device_ResetVertStartV = Val(TxtResetVertStartV.Text)
    Device_ResetVertSpeed = Val(TxtResetVertSpeed.Text)
    Device_ResetVertAccel = Val(TxtResetVertAccel.Text)
    
    Device_VertMinAngle = Val(TxtVertMinAngle.Text)
    Device_VertMinDistance = Val(TxtVertMinDistance.Text)
    Device_BeatMaxRadius = Val(TxtBeatMaxRadius.Text)
    
    Device_TurnFeedMM = Val(TxtTurnFeedMM.Text)
    Device_CutRadiusMM = Val(TxtCutRadiusMM.Text)
    
    Device_TurnPointOffsetMM = Val(TxtTurnPointOffsetMM.Text)
    Device_VertKnifeDegree = Val(TxtVertKnifeDegree.Text)

    Device_VertMaxOuterAngle = Val(TxtVertMaxOuterAngle.Text)
    Device_VertMaxInnerAngle = Val(TxtVertMaxInnerAngle.Text)
    
    Device_OuterAngleAdjustMM = Val(TxtOuterAngleAdjustMM.Text)
    Device_InnerAngleAdjustMM = Val(TxtInnerAngleAdjustMM.Text)
    
    Device_OuterLineTerminalAdjustMM = Val(TxtOuterLineTerminalAdjustMM.Text)
    Device_InnerLineTerminalAdjustMM = Val(TxtInnerLineTerminalAdjustMM.Text)
    
    Device_BenderBacklash = Val(TxtBenderBacklash.Text)
    Device_BenderSpringback = Val(TxtBenderSpringback.Text)
    
    Device_FastSpeedMinLenMM = Val(TxtFastSpeedMinLenMM.Text)
    Device_VertMotorZoneMM = Val(TxtVertMotorZoneMM.Text)
    
    Device_AmericanMaterial = IIf(ChkAmericanMaterial.value = 1, True, False)
    Device_TailVertAngle = Val(TxtTailVertAngle.Text)
    Device_VertUpDownMM_A = Val(TxtVertUpDownMM_A.Text)
    Device_KareanMaterial = IIf(ChkKareanMaterial.value = 1, True, False)
    
    SetDeviceParameters
        
    FrmMain.ChkStartPointVert90.Visible = Not Device_AmericanMaterial
    FrmMain.ChkEndPointVert90.Visible = Not Device_AmericanMaterial
    
    '------------------------------------------------------------------
    
    CmdSortAngleTable_Click
    
    For t = 1 To MaxBendDisNo
        BendDis(t) = Val(TxtBendDis(t).Text)
    Next
    
    SupplementKeyCount = 0
    For i = 1 To GrdAngleTable.Rows - 1
        If Val(GrdAngleTable.TextMatrix(i, 1)) <> 0 Then
            SupplementKeyCount = i
            KeyAngle(i) = Val(GrdAngleTable.TextMatrix(i, 1))
            For t = 1 To MaxBendDisNo
                RealAngle(t, i) = Val(GrdAngleTable.TextMatrix(i, t + 1))
                'SupAngle(t, i) = KeyAngle(i) - RealAngle(t, i)
            Next
        Else
            Exit For
        End If
    Next
    
    For t = 1 To MaxBendDisNo
        WriteToINI_A "Gap" & Trim(str(t)), str(BendDis(t))
    Next
    
    WriteToINI_A "SupplementKeyCount", str(SupplementKeyCount)
    For i = 1 To SupplementKeyCount
        WriteToINI_A "Key" & Trim(str(i)), str(KeyAngle(i))
        For t = 1 To MaxBendDisNo
            WriteToINI_A "Real" & Trim(str(i)) & "_" & Trim(str(t)), str(RealAngle(t, i))
        Next
    Next
               
    Unload Me
End Sub

Private Sub CmdShowCurve_Click()
    FrmShowCurve.Show
End Sub

Private Sub CmdShowTable_Click()
    If FormSettings.Width < 700 * Screen.TwipsPerPixelX Then
        FormSettings.Width = 910 * Screen.TwipsPerPixelX
    Else
        FormSettings.Width = 510 * Screen.TwipsPerPixelX
    End If
    
    FormSettings.left = (Screen.Width - FormSettings.Width) / 2
End Sub

Public Sub Form_Load()
    Dim t As Long, i As Long
    Dim obj As Object
    
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, SWP_Flags
    
    TxtPulsPerMM.Text = str(Device_PulsPerMM)
    TxtEncoderPulsPerMM.Text = str(Device_EncoderPulsPerMM)
    ChkUseEncoder.value = IIf(Device_UseEncoder = True, 1, 0)
    TxtPulsPerDegree.Text = str(Device_PulsPerDegree)
    
    TxtAdjustmentDegree.Text = Format(Device_AdjustmentDegree, " 0.0#######")
    TxtEmptyDegree.Text = Format(Device_EmptyDegree, " 0.0#######")
    
    'TxtAdjustmentDegree2.Text = Format(Device_AdjustmentDegree2, " 0.0#######")
    TxtEmptyDegree2.Text = Format(Device_EmptyDegree2, " 0.0#######")
    
    'ChkVertMotorDrive.value = IIf(Device_VertMotorDrive = True, 1, 0)
    'ChkVertAllHigh.value = IIf(Device_VertAllHigh = True, 1, 0)
    ChkVertNoTurn.value = IIf(Device_VertNoTurn = True, 1, 0)
    
    TxtVertUpDownPulsPerMM.Text = Format(Device_VertUpDownPulsPerMM, " 0.0#######")
    TxtVertUpDownAdjustmentMM.Text = Format(Device_VertUpDownAdjustmentMM, " 0.0#######")
    TxtVertUpDownMM.Text = Format(Device_VertUpDownMM, " 0.0#######")
    
    TxtVertPulsPerDegree.Text = Format(Device_VertPulsPerDegree, " 0.0#######")
    TxtVertAdjustmentDegree.Text = Format(Device_VertAdjustmentDegree, " 0.0#######")
    
    TxtHeadDistance.Text = Format(Device_HeadDistance, " 0.0###")
    TxtDoneDistance.Text = Format(Device_DoneDistance, " 0.0###")
    TxtDoneWaitingTime.Text = Format(Device_DoneWaitingTime, " 0.0###")
    TxtExtendMM.Text = Format(Device_ExtendMM, " 0.0###")
    
    'TxtWaitUpTime.Text = Str(Device_WaitUpTime)
    'TxtWaitDownTime.Text = Str(Device_WaitDownTime)
    
    TxtFeedStartV.Text = str(Device_FeedStartV)
    TxtFeedSpeed.Text = str(Device_FeedSpeed)
    TxtFeedAccel.Text = str(Device_FeedAccel)
    TxtFeedOffset.Text = str(Device_FeedOffset)
    
    'TxtManualFeedStartV.Text = str(Device_ManualFeedStartV)
    'TxtManualFeedSpeed.Text = str(Device_ManualFeedSpeed)
    'TxtManualFeedAccel.Text = str(Device_ManualFeedAccel)
    'TxtManualFeedOffset.Text = str(Device_ManualFeedOffset)
    
    TxtBendStartV.Text = str(Device_BendStartV)
    TxtBendSpeed.Text = str(Device_BendSpeed)
    TxtBendAccel.Text = str(Device_BendAccel)
    
    TxtManualBendStartV.Text = str(Device_ManualBendStartV)
    TxtManualBendSpeed.Text = str(Device_ManualBendSpeed)
    TxtManualBendAccel.Text = str(Device_ManualBendAccel)
    
    TxtResetBendStartV.Text = str(Device_ResetBendStartV)
    TxtResetBendSpeed.Text = str(Device_ResetBendSpeed)
    TxtResetBendAccel.Text = str(Device_ResetBendAccel)
    
    'TxtTurnFeedStartV.Text = Str(Device_TurnFeedStartV)
    'TxtTurnFeedSpeed.Text = Str(Device_TurnFeedSpeed)
    'TxtTurnFeedAccel.Text = Str(Device_TurnFeedAccel)
    
    TxtTurnFeedStartV.Text = str(Device_VertUpDownStartV)
    TxtTurnFeedSpeed.Text = str(Device_VertUpDownSpeed)
    TxtTurnFeedAccel.Text = str(Device_VertUpDownAccel)
    
    TxtVertStartV.Text = str(Device_VertStartV)
    TxtVertSpeed.Text = str(Device_VertSpeed)
    TxtVertAccel.Text = str(Device_VertAccel)
    
    TxtResetVertStartV.Text = str(Device_ResetVertStartV)
    TxtResetVertSpeed.Text = str(Device_ResetVertSpeed)
    TxtResetVertAccel.Text = str(Device_ResetVertAccel)
    
    TxtVertMinAngle.Text = Format(Device_VertMinAngle, " 0.0###")
    TxtVertMinDistance.Text = Format(Device_VertMinDistance, " 0.0###")
    TxtBeatMaxRadius.Text = Format(Device_BeatMaxRadius, " 0.0###")
    TxtTurnFeedMM.Text = Format(Device_TurnFeedMM, " 0.0###")
    TxtCutRadiusMM.Text = Format(Device_CutRadiusMM, " 0.0###")
    TxtTurnPointOffsetMM.Text = Format(Device_TurnPointOffsetMM)
    
    TxtVertKnifeDegree.Text = Format(Device_VertKnifeDegree, " 0.0###")
    TxtVertMaxOuterAngle.Text = Format(Device_VertMaxOuterAngle, " 0.0###")
    TxtVertMaxInnerAngle.Text = Format(Device_VertMaxInnerAngle, " 0.0###")
        
    TxtInnerAngleAdjustMM.Text = Format(Device_InnerAngleAdjustMM, " 0.0###")
    TxtOuterAngleAdjustMM.Text = Format(Device_OuterAngleAdjustMM, " 0.0###")
    
    TxtInnerLineTerminalAdjustMM.Text = Format(Device_InnerLineTerminalAdjustMM, " 0.0###")
    TxtOuterLineTerminalAdjustMM.Text = Format(Device_OuterLineTerminalAdjustMM, " 0.0###")
    
    TxtBenderBacklash.Text = Format(Device_BenderBacklash, " 0.0###")
    TxtBenderSpringback.Text = Format(Device_BenderSpringback, " 0.0###")
    
    TxtFastSpeedMinLenMM.Text = Format(Device_FastSpeedMinLenMM, " 0.0###")
    TxtVertMotorZoneMM.Text = Format(Device_VertMotorZoneMM, " 0.0###")
    
    ChkAmericanMaterial.value = IIf(Device_AmericanMaterial = True, 1, 0)
    TxtTailVertAngle.Text = Format(Device_TailVertAngle, " 0.0###")
    TxtVertUpDownMM_A.Text = Format(Device_VertUpDownMM_A, " 0.0###")
    ChkKareanMaterial.value = IIf(Device_KareanMaterial = True, 1, 0)
    
'-------------------------------------------------------------------
    
    CmbMaterial.Clear
    For t = 1 To 10
        CmbMaterial.AddItem Device_MaterialName(t)
    Next
    CmbMaterial.ListIndex = Val(Right(Device_CurMaterial, 2))
    
    'For t = 1 To MaxBendDisNo
    '    TxtBendDis(t).Text = IIf(BendDis(t) = 0, "", Str(BendDis(t)))
    'Next
    
    GrdAngleTable.Width = 370 * Screen.TwipsPerPixelX
    GrdAngleTable.Cols = 2 + MaxBendDisNo
    GrdAngleTable.Rows = 101
    GrdAngleTable.Clear
    GrdAngleTable.ColWidth(0) = 27 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(1) = 53 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(2) = 68 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(3) = 68 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(4) = 68 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(5) = 68 * Screen.TwipsPerPixelX
    GrdAngleTable.ColWidth(6) = 0 * Screen.TwipsPerPixelX
    GrdAngleTable.RowHeightMin = 18 * Screen.TwipsPerPixelY
    
    GrdAngleTable.ColAlignment(1) = 1
    GrdAngleTable.ColAlignment(2) = 1
    GrdAngleTable.ColAlignment(3) = 1
    GrdAngleTable.ColAlignment(4) = 1
    GrdAngleTable.ColAlignment(5) = 1
    
    GrdAngleTable.TextMatrix(0, 0) = "No."
    GrdAngleTable.TextMatrix(0, 1) = LblString1.caption
    GrdAngleTable.TextMatrix(0, 2) = LblString2.caption
    GrdAngleTable.TextMatrix(0, 3) = LblString3.caption
    GrdAngleTable.TextMatrix(0, 4) = LblString4.caption
    GrdAngleTable.TextMatrix(0, 5) = LblString5.caption
    'GrdAngleTable.TextMatrix(0, 6) = "拍弧角度"
    
    For i = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(i, 0) = str(i)
        GrdAngleTable.TextMatrix(i, 1) = Format(KeyAngle(i), " 0.0###")
        For t = 1 To MaxBendDisNo
            'GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0 And BendDis(t) = 0, "", Format(RealAngle(t, I), " 0.0###"))
            GrdAngleTable.TextMatrix(i, t + 1) = IIf(RealAngle(t, i) = 0, "", Format(RealAngle(t, i), " 0.0###"))
        Next
    Next

    ChkLocked_Click
End Sub

Private Sub OptBeat_Click()
    LblAngle.Enabled = True
    TxtAngleDeg.Enabled = True
    
    LblFeedMM.caption = "步距(mm)"
    LblFeedMM.Enabled = True
    TxtFeedMM.Enabled = True
End Sub

Private Sub OptBend_Click()
    LblAngle.Enabled = True
    TxtAngleDeg.Enabled = True
    
    LblFeedMM.caption = "弧长(mm)"
    LblFeedMM.Enabled = True
    TxtFeedMM.Enabled = True
End Sub


Private Sub GrdAngleTable_DblClick()
    'Debug.Print GrdAngleTable.Row, GrdAngleTable.Col
    If GrdAngleTable.Row >= 1 And GrdAngleTable.Col >= 1 Then
        SetDigiPad "FormSettings", "GrdAngleTable"
    End If
End Sub

Private Sub OptBeatL_Click()
    LblFeedMM.caption = LblString7.caption
    LblFeedMM.Visible = True
    TxtFeedMM.Visible = True
End Sub

Private Sub OptBeatR_Click()
    LblFeedMM.caption = LblString7.caption
    LblFeedMM.Visible = True
    TxtFeedMM.Visible = True
End Sub

Private Sub OptBendL_Click()
    LblFeedMM.caption = LblString6.caption
    LblFeedMM.Visible = True
    TxtFeedMM.Visible = True
End Sub

Private Sub OptBendR_Click()
    LblFeedMM.caption = LblString6.caption
    LblFeedMM.Visible = True
    TxtFeedMM.Visible = True
End Sub

Private Sub OptSymmetryTest_Click()
    LblFeedMM.Visible = False
    TxtFeedMM.Visible = False
End Sub

Private Sub OptTurn_Click()
    LblAngle.Enabled = True
    TxtAngleDeg.Enabled = True
    
    LblFeedMM.Enabled = False
    TxtFeedMM.Text = ""
    TxtFeedMM.Enabled = False
End Sub

Private Sub TxtAdjustmentDegree_DblClick()
    SetDigiPad "FormSettings", "TxtAdjustmentDegree"
End Sub

Private Sub TxtAngleDeg_DblClick()
    SetDigiPad "FormSettings", "TxtAngleDeg"
End Sub

Private Sub TxtBeatMaxRadius_DblClick()
    SetDigiPad "FormSettings", "TxtBeatMaxRadius"
End Sub

Private Sub TxtBendAccel_DblClick()
    SetDigiPad "FormSettings", "TxtBendAccel"
End Sub

Private Sub TxtBenderBacklash_DblClick()
    SetDigiPad "FormSettings", "TxtBenderBacklash"
End Sub

Private Sub TxtBendSpeed_Change()
    LblBendSpeed.caption = Format(Round(Val(TxtBendSpeed.Text) / Device_PulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtBendSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtBendSpeed"
End Sub

Private Sub TxtBendStartV_DblClick()
    SetDigiPad "FormSettings", "TxtBendStartV"
End Sub

Private Sub TxtCutRadiusMM_DblClick()
    SetDigiPad "FormSettings", "TxtCutRadiusMM"
End Sub

Private Sub TxtDoneDistance_DblClick()
    SetDigiPad "FormSettings", "TxtDoneDistance"
End Sub

Private Sub TxtDoneWaitingTime_DblClick()
    SetDigiPad "FormSettings", "TxtDoneWaitingTime"
End Sub

Private Sub TxtEmptyDegree_DblClick()
    SetDigiPad "FormSettings", "TxtEmptyDegree"
End Sub

Private Sub TxtEmptyDegree2_DblClick()
    SetDigiPad "FormSettings", "TxtEmptyDegree2"
End Sub

Private Sub TxtEncoderPulsPerMM_DblClick()
    SetDigiPad "FormSettings", "TxtEncoderPulsPerMM"
End Sub

Private Sub TxtExtendMM_DblClick()
    SetDigiPad "FormSettings", "TxtExtendMM"
End Sub

Private Sub TxtFastSpeedMinLenMM_DblClick()
    SetDigiPad "FormSettings", "TxtFastSpeedMinLenMM"
End Sub

Private Sub TxtFeedMM_DblClick()
    SetDigiPad "FormSettings", "TxtFeedMM"
End Sub

Private Sub TxtFeedOffset_DblClick()
    SetDigiPad "FormSettings", "TxtFeedOffset"
End Sub

Private Sub TxtFeedSpeed_Change()
    LblFeedSpeed.caption = Format(Round(60 * Val(TxtFeedSpeed.Text) / Device_PulsPerMM / 1000, 2), " 0.0## m/m")
End Sub

Private Sub TxtHeadDistance_DblClick()
    SetDigiPad "FormSettings", "TxtHeadDistance"
End Sub

Private Sub TxtInnerAngleAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtInnerAngleAdjustMM"
End Sub

Private Sub TxtManualBendAccel_DblClick()
    SetDigiPad "FormSettings", "TxtManualBendAccel"
End Sub

Private Sub TxtManualBendSpeed_Change()
    LblManualBendSpeed.caption = Format(Round(Val(TxtManualBendSpeed.Text) / Device_PulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtManualBendSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtManualBendSpeed"
End Sub

Private Sub TxtManualBendStartV_DblClick()
    SetDigiPad "FormSettings", "TxtManualBendStartV"
End Sub

Private Sub TxtManualFeedSpeed_Change()
    LblManualFeedSpeed.caption = Format(Round(60 * Val(TxtManualFeedSpeed.Text) / Device_PulsPerMM / 1000, 2), " 0.0## m/m")
End Sub

Private Sub TxtMaterialName_DblClick()
    SetDigiPad "FormSettings", "TxtMaterialName"
End Sub

Private Sub TxtMaterialName_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Idx As Long
    
    Idx = CmbMaterial.ListIndex + 1
    WritePrivateProfileString "MaterialName", str(Idx), TxtMaterialName.Text, App.Path & "\Parameters.ini"
    Device_MaterialName(Idx) = TxtMaterialName.Text
    
    CmbMaterial.Clear
    For Idx = 1 To 10
        CmbMaterial.AddItem Device_MaterialName(Idx)
    Next
    CmbMaterial.ListIndex = Val(Right(Device_CurMaterial, 2))
    
End Sub

Private Sub TxtMaterialThickMM_Change()
    Device_MaterialThickMM = Val(TxtMaterialThickMM.Text)
    WritePrivateProfileString "MaterialThickMM", Device_CurMaterial, str(Device_MaterialThickMM), App.Path & "\Parameters.ini"
End Sub

Private Sub TxtMaterialThickMM_DblClick()
    SetDigiPad "FormSettings", "TxtMaterialThickMM"
End Sub

Private Sub TxtOuterAngleAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtOuterAngleAdjustMM"
End Sub

Private Sub TxtPulsPerDegree_DblClick()
    SetDigiPad "FormSettings", "TxtPulsPerDegree"
End Sub

Private Sub TxtResetBendAccel_DblClick()
    SetDigiPad "FormSettings", "TxtResetBendAccel"
End Sub

Private Sub TxtResetBendSpeed_Change()
    LblResetBendSpeed.caption = Format(Round(Val(TxtResetBendSpeed.Text) / Device_PulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub CmdSortAngleTable_Click()
    Dim r As Long, r2 As Long, c As Long, c2 As Long, a As Double, b As Double, s As String
    
    For c = 1 To MaxBendDisNo
        If Val(TxtBendDis(c).Text) = 0 Then
            TxtBendDis(c).Text = str(10000 + c)
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
        
        a = Val(GrdAngleTable.TextMatrix(r, 1))
        GrdAngleTable.TextMatrix(r, 1) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = Val(GrdAngleTable.TextMatrix(r, 2))
        GrdAngleTable.TextMatrix(r, 2) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = Val(GrdAngleTable.TextMatrix(r, 3))
        GrdAngleTable.TextMatrix(r, 3) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = Val(GrdAngleTable.TextMatrix(r, 4))
        GrdAngleTable.TextMatrix(r, 4) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = Val(GrdAngleTable.TextMatrix(r, 5))
        GrdAngleTable.TextMatrix(r, 5) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = Val(GrdAngleTable.TextMatrix(r, 6))
        GrdAngleTable.TextMatrix(r, 6) = IIf(a = 0, "", Format(a, " 0.0###"))
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
            GrdAngleTable.TextMatrix(r, 0) = str(r)
        End If
    Next
End Sub

Private Sub GrdAngleTable_EnterCell()
    Dim CurRow As Long
    Dim CurCol As Long
    
    CurRow = GrdAngleTable.Row
    CurCol = GrdAngleTable.Col

    GrdAngleTable.RowSel = CurRow
    GrdAngleTable.ColSel = CurCol
    
    GrdAngleTable.ForeColorSel = RGB(255, 255, 255)
End Sub

Private Sub GrdAngleTable_KeyPress(KeyAscii As Integer)
    Dim CurRow As Long
    Dim CurCol As Long
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
    
    CmdShowCurve.Enabled = False
End Sub

Private Sub GrdAngleTable_LeaveCell()
    Dim CurRow As Long
    Dim CurCol As Long
    
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

Private Sub TxtResetBendSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtResetBendSpeed"
End Sub

Private Sub TxtResetBendStartV_DblClick()
    SetDigiPad "FormSettings", "TxtResetBendStartV"
End Sub

Private Sub TxtResetVertAccel_DblClick()
    SetDigiPad "FormSettings", "TxtResetVertAccel"
End Sub

Private Sub TxtResetVertSpeed_Change()
    LblResetVertSpeed.caption = Format(Round(Val(TxtResetVertSpeed.Text) / Device_VertPulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtResetVertSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtResetVertSpeed"
End Sub

Private Sub TxtResetVertStartV_DblClick()
    SetDigiPad "FormSettings", "TxtResetVertStartV"
End Sub

Private Sub TxtTailVertAngle_DblClick()
    SetDigiPad "FormSettings", "TxtTailVertAngle"
End Sub

Private Sub TxtTurnFeedAccel_DblClick()
    SetDigiPad "FormSettings", "TxtTurnFeedAccel"
End Sub

Private Sub TxtTurnFeedSpeed_Change()
    LblTurnFeedSpeed.caption = Format(Round(Val(TxtTurnFeedSpeed.Text) / Device_VertUpDownPulsPerMM, 2), " 0.0## mm/s")
End Sub

Private Sub TxtTurnFeedSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtTurnFeedSpeed"
End Sub

Private Sub TxtTurnFeedStartV_DblClick()
    SetDigiPad "FormSettings", "TxtTurnFeedStartV"
End Sub

Private Sub TxtVertAccel_DblClick()
    SetDigiPad "FormSettings", "TxtVertAccel"
End Sub

Private Sub TxtVertAdjustmentDegree_DblClick()
    SetDigiPad "FormSettings", "TxtVertAdjustmentDegree"
End Sub

Private Sub TxtVertKnifeDegree_DblClick()
    SetDigiPad "FormSettings", "TxtVertKnifeDegree"
End Sub

Private Sub TxtVertMaxInnerAngle_DblClick()
    SetDigiPad "FormSettings", "TxtVertMaxInnerAngle"
End Sub

Private Sub TxtVertMaxOuterAngle_DblClick()
    SetDigiPad "FormSettings", "TxtVertMaxOuterAngle"
End Sub

Private Sub TxtVertMinAngle_DblClick()
    SetDigiPad "FormSettings", "TxtVertMinAngle"
End Sub

Private Sub TxtVertMinDistance_DblClick()
    SetDigiPad "FormSettings", "TxtVertMinDistance"
End Sub

Private Sub TxtVertMotorZoneMM_DblClick()
    SetDigiPad "FormSettings", "TxtVertMotorZoneMM"
End Sub

Private Sub TxtVertPulsPerDegree_DblClick()
    SetDigiPad "FormSettings", "TxtVertPulsPerDegree"
End Sub

Private Sub TxtVertSpeed_Change()
    LblVertSpeed.caption = Format(Round(Val(TxtVertSpeed.Text) / Device_VertPulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtVertSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtVertSpeed"
End Sub

Private Sub TxtVertStartV_DblClick()
    SetDigiPad "FormSettings", "TxtVertStartV"
End Sub

Private Sub TxtVertUpDownAdjustmentMM_DblClick()
    SetDigiPad "FormSettings", "TxtVertUpDownAdjustmentMM"
End Sub

Private Sub TxtVertUpDownMM_A_DblClick()
    SetDigiPad "FormSettings", "TxtVertUpDownMM_A"
End Sub

Private Sub TxtVertUpDownMM_DblClick()
    SetDigiPad "FormSettings", "TxtVertUpDownMM"
End Sub

Private Sub TxtVertUpDownPulsPerMM_DblClick()
    SetDigiPad "FormSettings", "TxtVertUpDownPulsPerMM"
End Sub
