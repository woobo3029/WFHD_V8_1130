VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DFE8CC&
   Caption         =   "HD_WZ V8"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   -31830
   ClientWidth     =   11295
   FillColor       =   &H80000006&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   683
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   753
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkUseRemainder 
      Caption         =   "Use Ends"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2235
      TabIndex        =   50
      Top             =   7935
      Width           =   1230
   End
   Begin VB.TextBox TextBendLength 
      Height          =   315
      Left            =   4455
      TabIndex        =   136
      Text            =   "Text10"
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   129
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   126
      Text            =   "Text9"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   125
      Text            =   "Text8"
      Top             =   9360
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   124
      Text            =   "Text7"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   123
      Text            =   "Text6"
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   122
      Text            =   "Text5"
      Top             =   8640
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   117
      Text            =   "Text4"
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   116
      Text            =   "Text3"
      Top             =   9000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   115
      Text            =   "Text2"
      Top             =   8640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   114
      Text            =   "Text1"
      Top             =   8280
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6315
      Top             =   9030
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   3810
      TabIndex        =   67
      Top             =   990
      Visible         =   0   'False
      Width           =   3660
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Feed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   120
         TabIndex        =   101
         Top             =   180
         Width           =   3375
         Begin VB.TextBox TxtFeedMM 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   102
            Top             =   300
            Width           =   735
         End
         Begin HD_WZ_V8.PanButton CmdFeedBkV2 
            Height          =   435
            Left            =   120
            TabIndex        =   103
            Top             =   1140
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Fast BW"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdFeedFWV2 
            Height          =   435
            Left            =   1680
            TabIndex        =   104
            Top             =   1140
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Fast FW"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdRestAxisFeed 
            Height          =   255
            Left            =   3240
            TabIndex        =   105
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            FontName        =   ""
            FontSize        =   0
         End
         Begin HD_WZ_V8.PanButton CmdFeedFWV 
            Height          =   435
            Left            =   1680
            TabIndex        =   106
            Top             =   720
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Slow FW"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdFeedFWP 
            Height          =   495
            Left            =   2280
            TabIndex        =   107
            Top             =   180
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   873
            Caption         =   "FW"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdFeedBkP 
            Height          =   495
            Left            =   1320
            TabIndex        =   108
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   873
            Caption         =   "BW"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdFeedBkV 
            Height          =   435
            Left            =   120
            TabIndex        =   109
            Top             =   720
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Slow BW"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin VB.Label LblEncoderOS 
            BackStyle       =   0  'Transparent
            Caption         =   "编码器滞后"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   112
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LblEncoderOffset 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   111
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   930
            TabIndex        =   110
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bend"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   120
         TabIndex        =   91
         Top             =   1785
         Width           =   3375
         Begin VB.CheckBox CheckBendAbs 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Abs Angle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1560
            TabIndex        =   130
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox ChkAddEmptyDegree 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Add Idle Stroke"
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
            Left            =   120
            TabIndex        =   93
            Top             =   510
            Width           =   1455
         End
         Begin VB.TextBox TxtBendDeg 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   735
         End
         Begin HD_WZ_V8.PanButton CmdBenReset 
            Height          =   375
            Left            =   2385
            TabIndex        =   94
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "Home"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdResetAxisBend 
            Height          =   255
            Left            =   3240
            TabIndex        =   95
            Top             =   600
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            FontName        =   ""
            FontSize        =   0
         End
         Begin HD_WZ_V8.PanButton CmdBendRV 
            Height          =   495
            Left            =   1680
            TabIndex        =   96
            Top             =   1200
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   873
            Caption         =   "C.Bend Right"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdBendRP 
            Height          =   435
            Left            =   1680
            TabIndex        =   97
            Top             =   780
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Bend Right"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdBendLP 
            Height          =   435
            Left            =   120
            TabIndex        =   98
            Top             =   780
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Bend Left"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdBendLV 
            Height          =   495
            Left            =   120
            TabIndex        =   99
            Top             =   1200
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   873
            Caption         =   "C.Bend Left"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Deg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   900
            TabIndex        =   100
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mill Slot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   90
         TabIndex        =   75
         Top             =   3555
         Width           =   3405
         Begin VB.TextBox TxtVertDeg 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   735
         End
         Begin HD_WZ_V8.PanButton CmdVertRV 
            Height          =   435
            Left            =   3465
            TabIndex        =   77
            Top             =   1755
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   767
            Caption         =   "R"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdVertLV 
            Height          =   435
            Left            =   3465
            TabIndex        =   78
            Top             =   1215
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   767
            Caption         =   "L"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdVertInnerLine 
            Height          =   435
            Left            =   1665
            TabIndex        =   79
            Top             =   1125
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Slot"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdVertOutLine 
            Height          =   435
            Left            =   105
            TabIndex        =   80
            Top             =   1125
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Cut Off"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdMoveToOutLine 
            Height          =   435
            Left            =   120
            TabIndex        =   81
            Top             =   675
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Add Depth"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdMoveToInnerLine 
            Height          =   435
            Left            =   1665
            TabIndex        =   82
            Top             =   675
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "ResetDepth"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdVertDown 
            Height          =   435
            Left            =   1665
            TabIndex        =   83
            Top             =   1575
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Dn"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdVertUp 
            Height          =   435
            Left            =   120
            TabIndex        =   84
            Top             =   1575
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Up"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdVertReset 
            Height          =   495
            Left            =   2280
            TabIndex        =   85
            Top             =   180
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Caption         =   "Home"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdMotorStop 
            Height          =   435
            Left            =   1665
            TabIndex        =   86
            Top             =   2025
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Cutter Off"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin HD_WZ_V8.PanButton CmdMotorStart 
            Height          =   435
            Left            =   120
            TabIndex        =   139
            Top             =   2025
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   767
            Caption         =   "Cutter On"
            FontName        =   "Arial"
            FontSize        =   12
         End
         Begin VB.Label LblVertLowSensor 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   600
            TabIndex        =   90
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label LblVertHighSensor 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Deg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   930
            TabIndex        =   88
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "折角"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   69
         Top             =   8880
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CheckBox ChkAddEmptyDegree2 
            BackColor       =   &H00DFE8CC&
            Caption         =   "加上空程角"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   71
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtTurnDeg 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   735
         End
         Begin HD_WZ_V8.PanButton CmdTurnR 
            Height          =   255
            Left            =   960
            TabIndex        =   72
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "右折角"
            FontSize        =   0
         End
         Begin HD_WZ_V8.PanButton CmdTurnL 
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "左折角"
            FontSize        =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Deg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   74
            Top             =   240
            Width           =   615
         End
      End
      Begin HD_WZ_V8.PanButton CmdStopA 
         Height          =   615
         Left            =   360
         TabIndex        =   68
         Top             =   6240
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1085
         Caption         =   "Stop"
         FontName        =   "黑体"
         FontSize        =   30
      End
   End
   Begin VB.CheckBox ChkPathSmooth 
      BackColor       =   &H00DFE8CC&
      Caption         =   "Smoth path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   66
      Top             =   7200
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Timer TmrFeedV3Thread 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10200
      Top             =   9720
   End
   Begin VB.Timer TmrGetCurRunState 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6870
      Top             =   9030
   End
   Begin VB.Timer TmrFace 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   9120
      Top             =   9720
   End
   Begin HD_WZ_V8.PanButton CmdManualControl 
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Caption         =   "Show Manual Panel"
      FontName        =   "Arial"
      FontSize        =   12
   End
   Begin HD_WZ_V8.PanButton CmdVertInnerLineA 
      Height          =   675
      Left            =   11280
      TabIndex        =   64
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1191
      Caption         =   "铣内角"
      FontName        =   "黑体"
      FontSize        =   18
   End
   Begin HD_WZ_V8.PanButton CmdFeedFWV2A 
      Height          =   675
      Left            =   2040
      TabIndex        =   63
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1191
      Caption         =   "F-FW"
      FontName        =   "Arial"
      FontSize        =   18
   End
   Begin HD_WZ_V8.PanButton CmdFeedBkV2A 
      Height          =   675
      Left            =   240
      TabIndex        =   62
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1191
      Caption         =   "F-BW"
      FontName        =   "Arial"
      FontSize        =   18
   End
   Begin VB.PictureBox PicFace 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   8160
      ScaleHeight     =   6000
      ScaleWidth      =   9000
      TabIndex        =   61
      Top             =   2400
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.TextBox TxtEndPointAdjustMM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   58
      Text            =   "0"
      Top             =   7545
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox TxtStartPointAdjustMM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   57
      Text            =   "0"
      Top             =   7185
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Timer TmrVertThread 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   8580
      Top             =   9720
   End
   Begin VB.CheckBox ChkEndPointVert90 
      BackColor       =   &H00DFE8CC&
      Caption         =   "末端平铣"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4440
      TabIndex        =   54
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox ChkStartPointVert90 
      BackColor       =   &H00DFE8CC&
      Caption         =   "首端平铣"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4440
      TabIndex        =   53
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00DFE8CC&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   120
      TabIndex        =   42
      Top             =   1620
      Width           =   3615
      Begin HD_WZ_V8.PanButton CmdMovePoint 
         Height          =   375
         Left            =   1770
         TabIndex        =   60
         Top             =   1890
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         Caption         =   "ModifyPoint"
         FontName        =   "Arial"
         FontSize        =   12
      End
      Begin HD_WZ_V8.PanButton CmdAddPoint 
         Height          =   375
         Left            =   945
         TabIndex        =   51
         Top             =   1890
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         Caption         =   "AddPoint"
         FontName        =   "Arial"
         FontSize        =   12
      End
      Begin HD_WZ_V8.PanButton PanButton6 
         Height          =   375
         Left            =   2700
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "测面积"
         FontSize        =   0
      End
      Begin HD_WZ_V8.PanButton PanButton4 
         Height          =   375
         Left            =   1920
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "测距离"
         FontSize        =   0
      End
      Begin HD_WZ_V8.PanButton PanButton2 
         Height          =   645
         Left            =   135
         TabIndex        =   47
         Top             =   240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1138
         Caption         =   "Inline"
         FontName        =   "Arial"
         FontSize        =   22
      End
      Begin HD_WZ_V8.PanButton PanButton5 
         Height          =   810
         Left            =   2400
         TabIndex        =   43
         Top             =   1020
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1429
         Caption         =   "Clear"
         FontName        =   "Arial"
         FontSize        =   20
      End
      Begin HD_WZ_V8.PanButton PanButton3 
         Height          =   645
         Left            =   1845
         TabIndex        =   44
         Top             =   255
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1138
         Caption         =   "Outline"
         FontName        =   "Arial"
         FontSize        =   22
      End
      Begin HD_WZ_V8.PanButton PanButton1 
         Height          =   810
         Left            =   1290
         TabIndex        =   45
         Top             =   1020
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1429
         Caption         =   "EndPt"
         FontName        =   "Arial"
         FontSize        =   20
      End
      Begin HD_WZ_V8.PanButton PanButton11 
         Height          =   810
         Left            =   120
         TabIndex        =   46
         Top             =   1020
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1429
         Caption         =   "StartPt"
         FontName        =   "Arial"
         FontSize        =   20
      End
      Begin HD_WZ_V8.PanButton CmdDeletePoint 
         Height          =   375
         Left            =   135
         TabIndex        =   138
         Top             =   1890
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         Caption         =   "DeletePoint"
         FontName        =   "Arial"
         FontSize        =   12
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1680
      TabIndex        =   41
      Top             =   11400
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   360
      Scrolling       =   1
   End
   Begin VB.TextBox TxtRunCount 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3180
      TabIndex        =   40
      Top             =   4620
      Width           =   495
   End
   Begin VB.TextBox TxtRunN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3180
      TabIndex        =   37
      Text            =   "1"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton CmdToolBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   13560
      TabIndex        =   36
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox TxtStatistics 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   12240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   6000
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   1680
      TabIndex        =   30
      Top             =   11160
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin HD_WZ_V8.PanButton CmdResume 
      Height          =   855
      Left            =   2040
      TabIndex        =   29
      Top             =   6120
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1508
      Caption         =   "Resume"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin VB.Timer TmrBend 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4080
      Top             =   9240
   End
   Begin VB.PictureBox PicToolTip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      ScaleHeight     =   375
      ScaleWidth      =   1935
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdStop4 
      Caption         =   "Knife Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   11760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdStop3 
      Caption         =   "Stop 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   11280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdStop2 
      Caption         =   "Stop 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer TmrCheckMouse 
      Interval        =   100
      Left            =   5640
      Top             =   9240
   End
   Begin VB.Frame FraEdit 
      BackColor       =   &H00DFE8CC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   10320
      TabIndex        =   17
      Top             =   4800
      Width           =   1335
      Begin VB.CheckBox ChkEdit 
         BackColor       =   &H00DFE8CC&
         Caption         =   "自动优化"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin HD_WZ_V8.PanButton CmdEdit 
         Height          =   375
         Left            =   165
         TabIndex        =   19
         Top             =   1785
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "数据更新"
         FontSize        =   0
      End
      Begin VB.TextBox TxtEdit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblEdit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkReset 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   15
      Top             =   11520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox TxtTotalRun 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      TabIndex        =   13
      Text            =   "1"
      Top             =   11040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   315
      Left            =   9240
      Max             =   1
      Min             =   100
      TabIndex        =   12
      Top             =   11040
      Value           =   1
      Visible         =   0   'False
      Width           =   255
   End
   Begin HD_WZ_V8.PanButton CmdPause 
      Height          =   855
      Left            =   225
      TabIndex        =   11
      Top             =   6135
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1508
      Caption         =   "Pause"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton CmdResetOrg 
      Height          =   900
      Left            =   240
      TabIndex        =   10
      Top             =   4155
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1588
      Caption         =   "Home"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton CmdStop 
      Height          =   945
      Left            =   2040
      TabIndex        =   9
      Top             =   5100
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1667
      Caption         =   "Stop"
      FontName        =   "黑体"
      FontSize        =   32
   End
   Begin HD_WZ_V8.PanButton CmdTest 
      Height          =   435
      Left            =   300
      TabIndex        =   7
      Top             =   9780
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      Caption         =   "Pos Clr"
      FontName        =   "Arial"
      FontSize        =   16
   End
   Begin VB.Timer TmrDevicePortChecking 
      Interval        =   100
      Left            =   4680
      Top             =   9240
   End
   Begin VB.Timer TmrReadDevicePos 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5160
      Top             =   9240
   End
   Begin MSComctlLib.ImageList ImgLstToolDisabled 
      Left            =   12960
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":12C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1618
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":196A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":200E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2360
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":26B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":30A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":33FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":374C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4142
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4494
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":47E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgLstTool 
      Left            =   12840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":51E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5532
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5886
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":65D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6924
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":731A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":766C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":79BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":7D10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   12480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraToolBox 
      BackColor       =   &H00DFE8CC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   10320
      TabIndex        =   5
      Top             =   720
      Width           =   1635
      Begin HD_WZ_V8.PanButton CmdTool 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         FontSize        =   0
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8655
      Left            =   9960
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   8640
      Visible         =   0   'False
      Width           =   7455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   9930
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   16
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
            Object.Width           =   1720
            MinWidth        =   1720
            Text            =   "进料(长度)"
            TextSave        =   "进料(长度)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4233
            MinWidth        =   4233
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1720
            MinWidth        =   1720
            Text            =   "弯弧(半径)"
            TextSave        =   "弯弧(半径)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1720
            MinWidth        =   1720
            Text            =   "拍弧(角度)"
            TextSave        =   "拍弧(角度)"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1191
            MinWidth        =   1191
            Text            =   "铣槽"
            TextSave        =   "铣槽"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2910
            MinWidth        =   2910
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2910
            MinWidth        =   2910
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel15 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
         EndProperty
         BeginProperty Panel16 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicFrame 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   4035
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   0
      Top             =   795
      Width           =   7455
      Begin VB.PictureBox PicPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   1035
         ScaleHeight     =   287
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   351
         TabIndex        =   2
         Top             =   600
         Width           =   5295
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgLstTool"
      DisabledImageList=   "ImgLstToolDisabled"
      HotImageList    =   "ImgLstTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "This is a tip"
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "New"
            Object.ToolTipText     =   "New File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "打开AI文件/Open AI file"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ImportDXF"
            Object.ToolTipText     =   "打开DXF文件/Open DXF file"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Save file"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "SaveAs"
            Object.ToolTipText     =   "Save file as..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "撤销/Undo"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "重做/Redo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "SetAll"
            Object.ToolTipText     =   "自动设置路径"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Unset"
            Object.ToolTipText     =   "取消所有加工路径选择"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ShowCursorLines"
            Object.ToolTipText     =   "显示光标线"
            ImageIndex      =   10
            Style           =   1
            Object.Width           =   500
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowGridLines"
            Object.ToolTipText     =   "显示栅格/Show Grid Lines"
            ImageIndex      =   11
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowDirection"
            Object.ToolTipText     =   "显示方向/Show Direction"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowPoints"
            Object.ToolTipText     =   "显示节点/Show Points"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "SnapGrid"
            Object.ToolTipText     =   "对齐网格点"
            ImageIndex      =   14
            Style           =   1
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "CatchHVLines"
            Object.ToolTipText     =   "对齐平直线"
            ImageIndex      =   15
            Style           =   1
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "EditOnlyCurLayer"
            Object.ToolTipText     =   "仅编辑当前层"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Narrow"
            Object.ToolTipText     =   "平台变窄"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Wide"
            Object.ToolTipText     =   "平台变宽"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Low"
            Object.ToolTipText     =   "平台变高"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "High"
            Object.ToolTipText     =   "平台变低"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   500
         EndProperty
      EndProperty
      Begin VB.TextBox TxtCurTool 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox TxtCurData 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "X:   Y:"
         Top             =   0
         Width           =   2415
      End
      Begin VB.Image ImgLogo 
         Height          =   480
         Left            =   11880
         Picture         =   "FrmMain.frx":8062
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2820
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   11640
         Picture         =   "FrmMain.frx":E20C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2115
      End
   End
   Begin VB.CheckBox CheckXiaohu 
      BackColor       =   &H00DFE8CC&
      Caption         =   "Smoth bending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   131
      Top             =   7560
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin HD_WZ_V8.PanButton PanButtonRun 
      Height          =   405
      Left            =   1290
      TabIndex        =   8
      Top             =   6915
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "Run1"
      BorderStyle     =   2
      FontName        =   "黑体"
      FontSize        =   16
   End
   Begin HD_WZ_V8.PanButton CmdRun 
      Height          =   960
      Left            =   240
      TabIndex        =   137
      Top             =   5100
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1693
      Caption         =   "Run"
      BorderStyle     =   2
      FontName        =   "黑体"
      FontSize        =   32
   End
   Begin HD_WZ_V8.PanButton PanButtonBackSet 
      Height          =   435
      Left            =   1785
      TabIndex        =   87
      Top             =   9780
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      Caption         =   "BackSet"
      FontName        =   "Arial"
      FontSize        =   16
   End
   Begin VB.Label IN6 
      Height          =   255
      Left            =   3480
      TabIndex        =   135
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label IN9 
      Height          =   255
      Left            =   3480
      TabIndex        =   134
      Top             =   9480
      Width           =   255
   End
   Begin VB.Label Label17 
      Height          =   225
      Left            =   0
      TabIndex        =   133
      Top             =   0
      Width           =   255
   End
   Begin VB.Label IN5 
      Height          =   255
      Left            =   3480
      TabIndex        =   132
      Top             =   8760
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "Cutter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   128
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label LabelVertMotor 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   127
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "CutHeight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   121
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "CutDepth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   120
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "BendAng"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   119
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "送料"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   118
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "FeedLen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   113
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label LblVertMotorMode 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   59
      Top             =   5580
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "End_Comp(mm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   56
      Top             =   7575
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pre_Comp(mm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   55
      Top             =   7230
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblSpeedMode 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   52
      Top             =   5100
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Pieces"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   39
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Times"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   38
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label LblFeedStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   35
      Top             =   11160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LblReset3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1320
      TabIndex        =   34
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label LblReset2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   780
      TabIndex        =   33
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label LblReset1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   240
      TabIndex        =   32
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "循环前先复位"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   16
      Top             =   11520
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "循环次数"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   14
      Top             =   11040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File(&F)"
      Begin VB.Menu MnuNew 
         Caption         =   "New(&N)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "Open(&O)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar11 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuImportDXF 
         Caption         =   "Import DXF file(&D)"
      End
      Begin VB.Menu MnuImportAI 
         Caption         =   "Import AI file(&I)"
      End
      Begin VB.Menu mnubar12 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Save(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save As(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar13 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit(&E)"
      End
      Begin VB.Menu MnuBar14 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRecentFile 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "Edit(&E)"
      Begin VB.Menu MnuUndo 
         Caption         =   "Undo(&U)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "Redo(&R)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuBar21 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBarDelPoint 
         Caption         =   "DeletePoint"
      End
      Begin VB.Menu MnuBarAddPoint 
         Caption         =   "AddPoint"
      End
      Begin VB.Menu MnuBarMovePoint 
         Caption         =   "ModifyPoint"
      End
      Begin VB.Menu MnuBarBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBarMeasureDis 
         Caption         =   "MeasureDistance"
      End
      Begin VB.Menu MnuBarMeasureAera 
         Caption         =   "MeasureAera"
      End
      Begin VB.Menu MnuBYDrawingOrder 
         Caption         =   "按作图顺序切割(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar22 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMatrixByColumn 
         Caption         =   "矩阵逐列切割(&C)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMatrixByRow 
         Caption         =   "矩阵逐行切割(&R)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEraseDroppingSetting 
         Caption         =   "取消全部切割设置(&E)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "View(&V)"
      Begin VB.Menu MnuZoom 
         Caption         =   "Zoom"
         Index           =   0
      End
      Begin VB.Menu MnuBar31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPointList 
         Caption         =   "ShowPointList"
      End
      Begin VB.Menu MnuBar32 
         Caption         =   "-"
      End
      Begin VB.Menu MnuWhite 
         Caption         =   "White Background(&W)"
      End
      Begin VB.Menu MnuBlack 
         Caption         =   "Black Background(&B)"
      End
   End
   Begin VB.Menu MnuTool 
      Caption         =   "Tools(&T)"
      Visible         =   0   'False
      Begin VB.Menu MnuShiftToOrg 
         Caption         =   "Translate to Origin(&O)"
      End
      Begin VB.Menu MnuShift 
         Caption         =   "Translate(&T)"
      End
      Begin VB.Menu MnuMirror 
         Caption         =   "Mirror(&M)"
         Begin VB.Menu MnuMX 
            Caption         =   "沿X中心线(&X)"
         End
         Begin VB.Menu MnuMY 
            Caption         =   "沿Y中心线(&Y)"
         End
         Begin VB.Menu MnuM45 
            Caption         =   "沿 45°线(&D)"
         End
      End
      Begin VB.Menu MnuRotate 
         Caption         =   "旋转(&R)"
      End
      Begin VB.Menu MnuScale 
         Caption         =   "缩放(&S)"
         Begin VB.Menu MnuScaleM 
            Caption         =   "200 %"
            Index           =   2
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "300 %"
            Index           =   3
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "400 %"
            Index           =   4
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "500 %"
            Index           =   5
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "600 %"
            Index           =   6
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "700 %"
            Index           =   7
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "800 %"
            Index           =   8
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "900 %"
            Index           =   9
         End
         Begin VB.Menu MnuScaleM 
            Caption         =   "1000 %"
            Index           =   10
         End
         Begin VB.Menu MnuBar222 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSaleUserDefined 
            Caption         =   "自定义"
         End
      End
      Begin VB.Menu MnuBar51 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSetAll 
         Caption         =   "设置辅助线的所有交点(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEraseAll 
         Caption         =   "取消辅助线的所有交点(&E)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuMachine 
      Caption         =   "Machine(&M)"
      Begin VB.Menu MnuPointOnly 
         Caption         =   "点位方式(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuPointAndLine 
         Caption         =   "点线方式(&L)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar41 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOrgLeftUp 
         Caption         =   "原点位于左上角(&U)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOrgLeftDown 
         Caption         =   "原点位于左下角(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar42 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSetParam 
         Caption         =   "Parameters Setting(&S)"
      End
      Begin VB.Menu MnuBar999 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlatformParameter 
         Caption         =   "平台参数(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBar123 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuShowTest 
         Caption         =   "显示内部测试数据(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubar124 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCHN 
         Caption         =   "中文"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuENG 
         Caption         =   "English"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help(&H)"
      Begin VB.Menu MnuAbout 
         Caption         =   "About(&A)"
      End
      Begin VB.Menu MnuBar61 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRegistration 
         Caption         =   "User's Guide(&R)"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'============= Set the const value in  MdlDevice.bas ==============================
'Public Const Test_for_FeedByDCMotor_without_device = True/False
'==================================================================================

'防止弯弧进料时过度停止，采用比较弯弧角度增量与限定值大小的方法，来决定是否停止
Const IncBendangleLmt = 1.5

Dim IncAngle_g As Double
Dim OldBendangle_g As Double

Dim CurIndex As Integer
Dim BeatAngelPointLength As Double

'Const IsDemoVersion = True
Const IsDemoVersion = False

Dim VersionMark As String

Dim PicPathRatio As Double
Const WL = 258
Const WR = 164 '128
Const HT = 23
Const HB = 25

Const ErrLmt = 0.3

'定义拉直距离
Dim LazhiDis As Double

Dim languageType As Integer

'Dim ZoomFactor As Double
Dim tv() As Title_Value

Public CurX As Double
Public CurY As Double
Public CurZ As Double

Public DeviceMotion As Boolean
Dim ShowHead As Boolean
Dim ControlPadCheckDone As Boolean

'ControlPad
Dim pad_x0 As Single, pad_y0 As Single

'PicPath
Dim path_x0 As Single, path_y0 As Single, path_MouseDnX As Single, path_MouseDnY As Single
Dim path_dir0 As Integer, path_k0 As Integer, path_a0 As Double
Dim path_R0 As Double, path_D0 As Double
Dim MouseDnUX0 As Double, MouseDnUY0 As Double
Dim path_ShiftX0 As Single, path_ShiftY0 As Single

Dim XORPen As Boolean
Dim angle_adjust As Double

Dim CatchedElement As CatchedElementType
Dim CatchedID As Long
Dim CatchedElementDetial As CatchedElementType
Dim CatchedParam As Long
Dim CatchedOaram2 As Long

Dim CurFileName As String
Dim KeyDown As Integer

Dim RunningStartTime As Double

Const NumDigitsAfterDecimal = 2

'=================================================================

Public MotionCardOK As Boolean
Public FeedPulsCount As Long
Public FeedMark_x0 As Single
Public FeedMark_y0 As Single

Public VertMark_x0 As Single
Public VertMark_y0 As Single

Public VertFeedPulsCount As Long

Dim ResetOK As Boolean
Public FrmTestVisible As Boolean

Dim VT1 As Long, VT2 As Long, VTDir As Long

'打印铣刀点文件返回值
Dim fileNum_g As Integer

'定义连续弯弧长度
Dim Bendlength As Double
Dim posErr As Double
Dim PostoComp As Long
Dim VertCnt As Integer
Dim FirstRun As Boolean



Private Sub ChkPathSmooth_Click()
    PathSmooth = IIf(ChkPathSmooth.value = 1, True, False)
End Sub

Private Sub ChkUseRemainder_Click()
    FirstRun = IIf(ChkUseRemainder.value = 1, False, True)
End Sub

Private Sub CmdAddPoint_Click()
    CurTool = ToolType.BreakSegment
    TxtCurTool.Text = "加点"
End Sub

Private Sub CmdBeatL_Click()
   BeatAngle -val(TxtBendDeg.Text) - IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree)
End Sub

Private Sub CmdBeatR_Click()
    If val(TxtBendDeg.Text) > 0 Then
        BeatAngle val(TxtBendDeg.Text) + IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree2)
    Else
        BeatAngle val(TxtBendDeg.Text) - IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree)
    End If
End Sub

Private Sub CmdBendLP_Click()
    If CheckBendAbs.value = 0 Then
        BendAngle -val(TxtBendDeg.Text) - IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree), True
    Else
        BendAngleAbs -val(TxtBendDeg.Text) - IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree), True
    End If
End Sub

Private Sub CmdBendLV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long

    IsRunning = True
    TmrBend.Enabled = True
    If CtrlCardType = 0 Then
        Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
        Ret = set_speed(0, BendAxis, Device_ManualBendSpeed)
        Ret = set_acc(0, BendAxis, Device_ManualBendAccel)
        Ret = continue_move1(0, BendAxis, 1)
    ElseIf CtrlCardType = 4 Then
        'Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
        'Ret = SetVel(hDmc, BendAxis, Device_ManualBendSpeed)
        Ret = SetAcc(hDmc, BendAxis, Device_ManualBendAccel)
        Ret = SetDec(hDmc, BendAxis, Device_ManualBendAccel)
        Ret = ContinousMove(hDmc, BendAxis, -1 * Device_ManualBendSpeed)
    Else
        Ret = SetAxisStartVel_9030(0, BendAxis, Device_ManualBendStartV)
        'Ret = SetAxisVel_9030(0, BendAxis, Device_ManualBendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_ManualBendAccel)
        Ret = SetAxisDec_9030(0, BendAxis, Device_ManualBendAccel)
        Ret = StartAxisVel_9030(0, BendAxis, -1 * Device_ManualBendSpeed)
        
    End If
End Sub

Private Sub CmdBendLV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If CtrlCardType = 0 Then
        Ret = sudden_stop(0, BendAxis)
    ElseIf CtrlCardType = 4 Then
        Ret = SetDec(hDmc, BendAxis, 10 * Device_ManualBendAccel)
        Ret = StopAxis(hDmc, BendAxis)
    Else
        Ret = CeaseAxis_9030(0, BendAxis)
    End If
    
    TxtBendDeg.SetFocus

    IsRunning = False
End Sub

Private Sub CmdBendRP_Click()
    If CheckBendAbs.value = 0 Then
        BendAngle val(TxtBendDeg.Text) + IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree2), True
    Else
        BendAngleAbs val(TxtBendDeg.Text) + IIf(ChkAddEmptyDegree.value = 0, 0, Device_EmptyDegree2), True
    End If
End Sub

Private Sub CmdBendRV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long

    IsRunning = True
    TmrBend.Enabled = True

    If CtrlCardType = 0 Then
        Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
        Ret = set_speed(0, BendAxis, Device_ManualBendSpeed)
        Ret = set_acc(0, BendAxis, Device_ManualBendAccel)
        Ret = continue_move1(0, BendAxis, 0)
    ElseIf CtrlCardType = 4 Then
        'Ret = set_startv(0, BendAxis, Device_ManualBendStartV)
        'Ret = SetVel(hDmc, BendAxis, Device_ManualBendSpeed)
        Ret = SetAcc(hDmc, BendAxis, Device_ManualBendAccel)
        Ret = SetDec(hDmc, BendAxis, Device_ManualBendAccel)
        Ret = ContinousMove(hDmc, BendAxis, Device_ManualBendSpeed)
    Else
        Ret = SetAxisStartVel_9030(0, BendAxis, Device_ManualBendStartV)
        'Ret = SetAxisVel_9030(0, BendAxis, Device_ManualBendSpeed)
        Ret = SetAxisAcc_9030(0, BendAxis, Device_ManualBendAccel)
        Ret = SetAxisDec_9030(0, BendAxis, Device_ManualBendAccel)
        Ret = StartAxisVel_9030(0, BendAxis, Device_ManualBendSpeed)
        
    End If
End Sub

Private Sub CmdBendRV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long

    If CtrlCardType = 0 Then
        Ret = sudden_stop(0, BendAxis)
    ElseIf CtrlCardType = 4 Then
        Ret = SetDec(hDmc, BendAxis, 10 * Device_ManualBendAccel)
        StopAxis hDmc, BendAxis
    Else
        Ret = CeaseAxis_9030(0, BendAxis)
    End If
    TxtBendDeg.SetFocus

    IsRunning = False
End Sub

Private Sub CmdBenReset_Click()
    StopRunning = False
    If CtrlCardType = 0 Then
        BendReset
    Else
        If CtrlCardType = 1 Then
            BendReset_9030_V8
        ElseIf CtrlCardType = 4 Then
            BendReset_GALIL_V8
        End If
    End If
End Sub

Private Sub CmdDeletePoint_Click()
    CurTool = ToolType.DeleteElement_Point
    TxtCurTool.Text = "删点"
End Sub

Private Sub CmdEdit_Click()
    Dim pid As Long, sid As Long, body_id As Long
    Dim X As Double, Y As Double, w As Double, h As Double, ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double, dX As Double, dy As Double
    Dim cx As Double, cy As Double, r As Double, r1 As Double, r2 As Double, sa As Double, ea As Double, Ret As Boolean
    Dim I As Long, j As Long, k As Long, CS As Double, SN As Double, group_id As Long, n As Long, id0 As Long, id1 As Long
    Dim TempPoint1 As Path_Point, TempPoint2 As Path_Point, TempArc As Path_Arc
    
    FraEdit.Visible = False
    If FraEdit.Tag = "" Then
        Exit Sub
    End If
    FraEdit.Tag = ""
    
    If val(CmdEdit.Tag) = ToolType.SetPoint Then 'Add Point
        X = val(TxtEdit(0).Text)
        Y = val(TxtEdit(1).Text)
        
        If X <> PointList(PointCount).X Or Y <> PointList(PointCount).Y Then
            PointList(PointCount).X = X
            PointList(PointCount).Y = Y
            
            PopAllXORStack
            CloseXORStack
            PicPath.Cls
            DrawAll
            
            SaveUndo
        End If
        
    ElseIf val(CmdEdit.Tag) = ToolType.SetSegment Then
        X = val(TxtEdit(0).Text)
        Y = val(TxtEdit(1).Text)
        
        If X <> CurPoint.X Or Y <> CurPoint.Y Then
            CurPoint.X = X
            CurPoint.Y = Y
                    
            PointList(PointCount).X = X
            PointList(PointCount).Y = Y
            
            PopAllXORStack
            CloseXORStack
            PicPath.Cls
            DrawAll
            
            SaveUndo
        End If
    ElseIf val(CmdEdit.Tag) = ToolType.SetBox Then
        X = val(TxtEdit(0).Text)
        Y = val(TxtEdit(1).Text)
        w = val(TxtEdit(2).Text)
        h = val(TxtEdit(3).Text)
        
        If w <> 0 And h <> 0 Then
            'sid = BodyList(BodyCount).first_segment_id
            sid = CurBoxFirstSegmentID
            
            pid = SegmentList(sid).point0_id
            PointList(pid).X = X
            PointList(pid).Y = Y
            
            pid = SegmentList(sid + 1).point0_id
            PointList(pid).X = X
            PointList(pid).Y = Y + h
            
            pid = SegmentList(sid + 2).point0_id
            PointList(pid).X = X + w
            PointList(pid).Y = Y + h
            
            pid = SegmentList(sid + 3).point0_id
            PointList(pid).X = X + w
            PointList(pid).Y = Y
            
            PopAllXORStack
            CloseXORStack
            PicPath.Cls
            DrawAll
            
            SaveUndo
        End If
    ElseIf val(CmdEdit.Tag) = ToolType.SetSPLine Then
        X = val(TxtEdit(0).Text)
        Y = val(TxtEdit(1).Text)
        
        If X <> CurPoint.X Or Y <> CurPoint.Y Then
            CurPoint.X = X
            CurPoint.Y = Y
                    
            PointList(PointCount).X = X
            PointList(PointCount).Y = Y
            
            PopAllXORStack
            CloseXORStack
            PicPath.Cls
            DrawAll
            
            OpenXORStack
            DrawSPLine TempSPline, Not LayerColor(CurLayer)
            
            SaveUndo
        End If
    ElseIf val(CmdEdit.Tag) = ToolType.SetCircle Then
        CurArc.X = val(TxtEdit(0).Text)
        CurArc.Y = val(TxtEdit(1).Text)
        CurArc.a = val(TxtEdit(2).Text)
        CurArc.b = val(TxtEdit(2).Text)
        
        PointList(CurArc.pointm_id).X = CurArc.X
        PointList(CurArc.pointm_id).Y = CurArc.Y
        
        CurArc.start_angle = val(TxtEdit(3).Text) * Pi / 180
        CurArc.end_angle = CurArc.start_angle + val(TxtEdit(4).Text) * Pi / 180
        
        'PointList(CurArc.point0_id).X = Cos(CurArc.start_angle) * CurArc.a + CurArc.X
        'PointList(CurArc.point0_id).Y = Sin(CurArc.start_angle) * CurArc.B + CurArc.Y
        
        X = Cos(CurArc.start_angle) * CurArc.a
        Y = Sin(CurArc.start_angle) * CurArc.b
        CS = Cos(CurArc.ax_angle)
        SN = Sin(CurArc.ax_angle)
        PointList(CurArc.point0_id).X = (CS * X) - (SN * Y) + CurArc.X
        PointList(CurArc.point0_id).Y = (SN * X) + (CS * Y) + CurArc.Y
        
        'PointList(CurArc.point1_id).X = Cos(CurArc.end_angle) * CurArc.a + CurArc.X
        'PointList(CurArc.point1_id).Y = Sin(CurArc.end_angle) * CurArc.B + CurArc.Y
                    
        X = Cos(CurArc.end_angle) * CurArc.a
        Y = Sin(CurArc.end_angle) * CurArc.b
        CS = Cos(CurArc.ax_angle)
        SN = Sin(CurArc.ax_angle)
        PointList(CurArc.point1_id).X = (CS * X) - (SN * Y) + CurArc.X
        PointList(CurArc.point1_id).Y = (SN * X) + (CS * Y) + CurArc.Y
                    
        ArcList(CurArc.id) = CurArc
        
        'in case it is used for circle editting
        For I = 1 To OutputStartPointList.count
            If PointList(OutputStartPointList.point_id(I)).body_id = CurArc.body_id Then
                OutputStartPointList.leading_point0(I).X = PointList(OutputStartPointList.leading_point0(I).id).X
                OutputStartPointList.leading_point0(I).Y = PointList(OutputStartPointList.leading_point0(I).id).Y
                OutputStartPointList.leading_point1(I).X = PointList(OutputStartPointList.leading_point1(I).id).X
                OutputStartPointList.leading_point1(I).Y = PointList(OutputStartPointList.leading_point1(I).id).Y
            End If
        Next
        
        PopAllXORStack
        CloseXORStack
        PicPath.Cls
        DrawAll
        
        SaveUndo
        
'    ElseIf Val(CmdEdit.Tag) = ToolType.SetCircle_3p Then
'        LastPoint.X = Val(TxtEdit(0).Text)
'        LastPoint.Y = Val(TxtEdit(1).Text)
'        CurPoint.X = Val(TxtEdit(2).Text)
'        CurPoint.Y = Val(TxtEdit(3).Text)
'        X = Val(TxtEdit(4).Text)
'        Y = Val(TxtEdit(5).Text)
'
'        ret = GetCircleBy3Points(LastPoint.X, LastPoint.Y, CurPoint.X, CurPoint.Y, X, Y, cx, cy, r, sa, ea)
'        If ret = True Then
'            CurArc.X = cx
'            CurArc.Y = cy
'            CurArc.a = r
'            CurArc.B = r
'            If Abs(CurArc.end_angle - CurArc.start_angle) = PI2 Then '若原为整园则保持封闭
'                PointList(CurArc.point0_id).X = LastPoint.X
'                PointList(CurArc.point0_id).Y = LastPoint.Y
'                PointList(CurArc.point1_id).X = LastPoint.X
'                PointList(CurArc.point1_id).Y = LastPoint.Y
'            Else
'                CurArc.start_angle = sa
'                CurArc.end_angle = ea
'
'                PointList(CurArc.point0_id).X = LastPoint.X
'                PointList(CurArc.point0_id).Y = LastPoint.Y
'                PointList(CurArc.point1_id).X = X
'                PointList(CurArc.point1_id).Y = Y
'            End If
'
'            PointList(CurArc.pointm_id).X = CurPoint.X
'            PointList(CurArc.pointm_id).Y = CurPoint.Y
'
'            ArcList(CurArc.id) = CurArc
'
'            'in case it is used for circle editting
'            For I = 1 To OutputStartPointList.count
'                If PointList(OutputStartPointList.point_id(I)).body_id = CurArc.body_id Then
'                    OutputStartPointList.leading_point0(I).X = PointList(OutputStartPointList.leading_point0(I).id).X
'                    OutputStartPointList.leading_point0(I).Y = PointList(OutputStartPointList.leading_point0(I).id).Y
'                    OutputStartPointList.leading_point1(I).X = PointList(OutputStartPointList.leading_point1(I).id).X
'                    OutputStartPointList.leading_point1(I).Y = PointList(OutputStartPointList.leading_point1(I).id).Y
'                End If
'            Next
'
'            PopAllXORStack
'            CloseXORStack
'            PicPath.Cls
'            DrawAll
'
'            SaveUndo
'        Else
'            MsgBox " 点的数据错误，无法得出对应的圆形 ! ", vbExclamation + vbOKOnly
'        End If
'
    ElseIf val(CmdEdit.Tag) = ToolType.SetEllipse Then
        CurArc.X = val(TxtEdit(0).Text)
        CurArc.Y = val(TxtEdit(1).Text)
        CurArc.a = val(TxtEdit(2).Text)
        CurArc.b = val(TxtEdit(3).Text)
        
        PointList(CurArc.pointm_id).X = CurArc.X
        PointList(CurArc.pointm_id).Y = CurArc.Y
        
        CurArc.start_angle = val(TxtEdit(4).Text) * Pi / 180
        CurArc.end_angle = CurArc.start_angle + val(TxtEdit(5).Text) * Pi / 180
        
        'PointList(CurArc.point0_id).X = Cos(CurArc.start_angle) * CurArc.a + CurArc.X
        'PointList(CurArc.point0_id).Y = Sin(CurArc.start_angle) * CurArc.B + CurArc.Y
        
        X = Cos(CurArc.start_angle) * CurArc.a
        Y = Sin(CurArc.start_angle) * CurArc.b
        CS = Cos(CurArc.ax_angle)
        SN = Sin(CurArc.ax_angle)
        PointList(CurArc.point0_id).X = (CS * X) - (SN * Y) + CurArc.X
        PointList(CurArc.point0_id).Y = (SN * X) + (CS * Y) + CurArc.Y
        
        'PointList(CurArc.point1_id).X = Cos(CurArc.end_angle) * CurArc.a + CurArc.X
        'PointList(CurArc.point1_id).Y = Sin(CurArc.end_angle) * CurArc.B + CurArc.Y
                    
        X = Cos(CurArc.end_angle) * CurArc.a
        Y = Sin(CurArc.end_angle) * CurArc.b
        CS = Cos(CurArc.ax_angle)
        SN = Sin(CurArc.ax_angle)
        PointList(CurArc.point1_id).X = (CS * X) - (SN * Y) + CurArc.X
        PointList(CurArc.point1_id).Y = (SN * X) + (CS * Y) + CurArc.Y
            
        'in case it is used for circle editting
        For I = 1 To OutputStartPointList.count
            If PointList(OutputStartPointList.point_id(I)).body_id = CurArc.body_id Then
                OutputStartPointList.leading_point0(I).X = PointList(OutputStartPointList.leading_point0(I).id).X
                OutputStartPointList.leading_point0(I).Y = PointList(OutputStartPointList.leading_point0(I).id).Y
                OutputStartPointList.leading_point1(I).X = PointList(OutputStartPointList.leading_point1(I).id).X
                OutputStartPointList.leading_point1(I).Y = PointList(OutputStartPointList.leading_point1(I).id).Y
            End If
        Next
        
        ArcList(CurArc.id) = CurArc
    
        PopAllXORStack
        CloseXORStack
        PicPath.Cls
        DrawAll
        
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.RoundCornerByPoint Then
        Ret = RoundCorner(ArcList(CurArcIndex).point_id, val(TxtEdit(0).Text))
        
        PopAllXORStack
        CloseXORStack
        PicPath.Cls
        DrawAll
        
        SaveUndo
    
    ElseIf val(CmdEdit.Tag) = ToolType.EditElement Then
        If CurPointIndex > 0 Then
            X = val(TxtEdit(0).Text)
            Y = val(TxtEdit(1).Text)
            
            If MovePoint(CurPointIndex, X, Y) = True Then
                PopAllXORStack
                CloseXORStack
                PicPath.Cls
                DrawAll
                
                SaveUndo
            End If
            
        ElseIf CurSegmentIndex > 0 Then
            If SegmentOnBox(CurSegmentIndex, X, Y, w, h, sid, body_id) = True Then
                X = val(TxtEdit(0).Text)
                Y = val(TxtEdit(1).Text)
                w = val(TxtEdit(2).Text)
                h = val(TxtEdit(3).Text)
                
                If w <> 0 And h <> 0 Then
                    pid = SegmentList(sid).point0_id
                    MovePoint pid, X, Y
                    
                    pid = SegmentList(sid + 1).point0_id
                    MovePoint pid, X, Y + h
                    
                    pid = SegmentList(sid + 2).point0_id
                    MovePoint pid, X + w, Y + h
                    
                    pid = SegmentList(sid + 3).point0_id
                    MovePoint pid, X + w, Y
                    
                    PopAllXORStack
                    CloseXORStack
                    PicPath.Cls
                    DrawAll
                    
                    SaveUndo
                End If
            End If
        ElseIf CurArcIndex > 0 Then
        End If
        
    ElseIf val(CmdEdit.Tag) = ToolType.RotateElement Then
        RotateGroup CurGroupID, CurGroupCenterX, CurGroupCenterY, val(TxtEdit(0).Text) * PI_180
    
        PopAllXORStack
        CloseXORStack
        PicPath.Cls
        DrawAll
        
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.PieceArray Then
        Undo True
        
        PieceArray CurArrayedGroupID, val(TxtEdit(0).Text), val(TxtEdit(1).Text), val(TxtEdit(2).Text), val(TxtEdit(3).Text), val(TxtEdit(4).Text)
        
        PicPathCls
        DrawAll
        SaveUndo
        FraEdit.Visible = True
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateCircle Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        r = val(TxtEdit(2).Text)
        
        AddArc cx, cy, LayerZValue(CurLayer), r, r, 0, PI2, 0, 0, 0, CurLayer, ArcType.CircleCR
                
        ux0 = cx + r
        uy0 = cy
        ux1 = ux0
        uy1 = uy0
        
        AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).point0_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).pointm_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).point1_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        DrawArc ArcList(ArcCount)
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateArc Or val(CmdEdit.Tag) = ToolType.CreateSector Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        r1 = val(TxtEdit(2).Text)
        If val(CmdEdit.Tag) = ToolType.CreateSector Then
            r2 = 0
            ea = val(TxtEdit(3).Text) * PI_180
        Else
            r2 = val(TxtEdit(3).Text)
            ea = val(TxtEdit(4).Text) * PI_180
        End If
        
        If r1 <= r2 Or ea = 0 Then
            Beep
            Exit Sub
        End If
        
        sa = 0
        
        AddArc cx, cy, LayerZValue(CurLayer), r1, r1, sa, ea, 0, 0, 0, CurLayer, ArcType.CircleCR
                
        ux0 = cx + r1
        uy0 = cy
        ux1 = Cos(ea) * r1 + cx
        uy1 = Sin(ea) * r1 + cy
        
        AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).point0_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).pointm_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).point1_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        DrawArc ArcList(ArcCount)
        
        If r2 > 0 Then
            AddArc cx, cy, LayerZValue(CurLayer), r2, r2, ea, sa, 0, 0, 0, CurLayer, ArcType.CircleCR
                    
            ux0 = Cos(ea) * r2 + cx
            uy0 = Sin(ea) * r2 + cy
            ux1 = cx + r2
            uy1 = cy
            
            AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
            ArcList(ArcCount).point0_id = PointList(PointCount).id
            PointList(PointCount).arc_id = ArcList(ArcCount).id
            PointList(PointCount).body_id = ArcList(ArcCount).body_id
            PointList(PointCount).group_id = ArcList(ArcCount).group_id
            
            AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
            ArcList(ArcCount).pointm_id = PointList(PointCount).id
            PointList(PointCount).arc_id = ArcList(ArcCount).id
            PointList(PointCount).body_id = ArcList(ArcCount).body_id
            PointList(PointCount).group_id = ArcList(ArcCount).group_id
            
            AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
            ArcList(ArcCount).point1_id = PointList(PointCount).id
            PointList(PointCount).arc_id = ArcList(ArcCount).id
            PointList(PointCount).body_id = ArcList(ArcCount).body_id
            PointList(PointCount).group_id = ArcList(ArcCount).group_id
            
            DrawArc ArcList(ArcCount)
            
            AddSegment ArcList(ArcCount - 1).point1_id, ArcList(ArcCount).point0_id
            DrawSegment SegmentList(SegmentCount)
            AddSegment ArcList(ArcCount).point1_id, ArcList(ArcCount - 1).point0_id
            DrawSegment SegmentList(SegmentCount)
        Else
            AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
            AddSegment ArcList(ArcCount).point1_id, PointList(PointCount).id
            DrawSegment SegmentList(SegmentCount)
            AddSegment PointList(PointCount).id, ArcList(ArcCount).point0_id
            DrawSegment SegmentList(SegmentCount)
        End If
        
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateEllipse Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        r1 = val(TxtEdit(2).Text)
        r2 = val(TxtEdit(3).Text)
        
        AddArc cx, cy, LayerZValue(CurLayer), r1, r2, 0, PI2, 0, 0, 0, CurLayer, ArcType.Ellipse
                
        ux0 = cx + r1
        uy0 = cy
        ux1 = ux0
        uy1 = uy0
        
        AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).point0_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).pointm_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        ArcList(ArcCount).point1_id = PointList(PointCount).id
        PointList(PointCount).arc_id = ArcList(ArcCount).id
        PointList(PointCount).body_id = ArcList(ArcCount).body_id
        PointList(PointCount).group_id = ArcList(ArcCount).group_id
        
        DrawArc ArcList(ArcCount)
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateSquare Or val(CmdEdit.Tag) = ToolType.CreateRectangle Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        w = val(TxtEdit(2).Text)
        If val(CmdEdit.Tag) = ToolType.CreateSquare Then
            h = w
        Else
            h = val(TxtEdit(3).Text)
        End If
        
        TempPoint1.X = cx - w / 2
        TempPoint1.Y = cy - h / 2
        TempPoint2.X = cx + w / 2
        TempPoint2.Y = cy + h / 2
        
        AddBox TempPoint1, TempPoint2
        
        PicPath.Cls
        DrawAll
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateParallelRectangle Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        w = val(TxtEdit(2).Text)
        h = val(TxtEdit(3).Text)
        dX = val(TxtEdit(4).Text)
        
        X = cx - w / 2 + dX
        Y = cy + h / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        k = PointCount
        id0 = PointCount
        
        X = X + w
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        X = cx + w / 2
        Y = cy - h / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        X = cx - w / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        AddSegment id0, k
         
        PicPath.Cls
        DrawAll
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreatePolygon Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        n = val(TxtEdit(2).Text)
        r = val(TxtEdit(3).Text)
        
        ux0 = cx + r
        uy0 = cy
        
        AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        k = PointCount
        id0 = PointCount
        
        For I = 1 To n - 1
            ux1 = Cos(I * PI2 / n) * r + cx
            uy1 = Sin(I * PI2 / n) * r + cy
            
            AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
            id1 = PointCount
            
            AddSegment id0, id1
            id0 = id1
        Next
        id1 = k
        AddSegment id0, id1
        
        PicPath.Cls
        DrawAll
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateTriangle Or val(CmdEdit.Tag) = ToolType.CreateTriangle1 Or val(CmdEdit.Tag) = ToolType.CreateTriangle2 Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        w = val(TxtEdit(2).Text)
        If val(CmdEdit.Tag) = ToolType.CreateTriangle Then
            h = w * Sin(Pi / 3)
            X = cx
        ElseIf val(CmdEdit.Tag) = ToolType.CreateTriangle1 Then
            h = val(TxtEdit(3).Text)
            X = cx
        ElseIf val(CmdEdit.Tag) = ToolType.CreateTriangle2 Then
            h = val(TxtEdit(3).Text)
            X = cx - w / 2
        End If
        Y = cy + h / 2
        
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        k = PointCount
        id0 = PointCount
        
        X = cx + w / 2
        Y = cy - h / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        X = cx - w / 2
        Y = cy - h / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
                
        AddSegment id0, k
         
        PicPath.Cls
        DrawAll
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateTrapezoid Or val(CmdEdit.Tag) = ToolType.CreateTrapezoid1 Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        r1 = val(TxtEdit(2).Text)
        r2 = val(TxtEdit(3).Text)
        h = val(TxtEdit(4).Text)
        
        If val(CmdEdit.Tag) = ToolType.CreateTrapezoid Then
            X = cx - r1 / 2
        Else
            X = cx - r2 / 2
        End If
        Y = cy + h / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        k = PointCount
        id0 = PointCount
        
        If val(CmdEdit.Tag) = ToolType.CreateTrapezoid Then
            X = cx + r1 / 2
        Else
            X = X + r1
        End If
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        X = cx + r2 / 2
        Y = cy - h / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        X = cx - r2 / 2
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        AddSegment id0, id1
        id0 = id1
        
        AddSegment id0, k
         
        PicPath.Cls
        DrawAll
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.Create5PointStar Or val(CmdEdit.Tag) = ToolType.CreateMultiPointStar Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        If val(CmdEdit.Tag) = ToolType.Create5PointStar Then
            n = 5
            r = val(TxtEdit(2).Text)
            r1 = val(TxtEdit(3).Text)
        Else
            n = val(TxtEdit(2).Text)
            r = val(TxtEdit(3).Text)
            r1 = val(TxtEdit(4).Text)
        End If
        
        ux0 = cx + r * Cos(PI_2)
        uy0 = cy + r * Sin(PI_2)
        
        AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        k = PointCount
        id0 = PointCount
        
        If r1 = 0 Then
            r1 = r * Cos(PI2 / n) / Cos(Pi / n)
        End If
        
        For I = 1 To n - 1
            ux1 = Cos(PI_2 + (I - 0.5) * PI2 / n) * r1 + cx
            uy1 = Sin(PI_2 + (I - 0.5) * PI2 / n) * r1 + cy
            
            AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
            id1 = PointCount
            
            AddSegment id0, id1
            id0 = id1
            ux1 = Cos(PI_2 + I * PI2 / n) * r + cx
            uy1 = Sin(PI_2 + I * PI2 / n) * r + cy
            
            AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
            id1 = PointCount
            
            AddSegment id0, id1
            id0 = id1
        Next
        ux1 = Cos(PI_2 + (I - 0.5) * PI2 / n) * r1 + cx
        uy1 = Sin(PI_2 + (I - 0.5) * PI2 / n) * r1 + cy
        
        AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        id1 = PointCount
        
        AddSegment id0, id1
        id0 = id1
        
        id1 = k
        AddSegment id0, id1
        
        PicPath.Cls
        DrawAll
        SaveUndo
        
    ElseIf val(CmdEdit.Tag) = ToolType.CreateCurvePolygon Then
        cx = val(TxtEdit(0).Text)
        cy = val(TxtEdit(1).Text)
        n = val(TxtEdit(2).Text)
        r = val(TxtEdit(3).Text)
        r2 = val(TxtEdit(4).Text)
        
        If r = 0 Then
            Beep
            Exit Sub
        End If
        
        ux0 = cx + r * Cos(PI_2)
        uy0 = cy + r * Sin(PI_2)
        
        AddPoint ux0, uy0, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        k = PointCount
        id0 = PointCount
        
        r1 = r + r2
        
        For I = 1 To n - 1
            ux0 = Cos(PI_2 + (I - 0.5) * PI2 / n) * r1 + cx
            uy0 = Sin(PI_2 + (I - 0.5) * PI2 / n) * r1 + cy
            
            ux1 = Cos(PI_2 + I * PI2 / n) * r + cx
            uy1 = Sin(PI_2 + I * PI2 / n) * r + cy
            
            AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
            id1 = PointCount
            
            GetCircleBy3Points PointList(id0).X, PointList(id0).Y, ux0, uy0, ux1, uy1, X, Y, r2, sa, ea
            AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
            AddArc X, Y, LayerZValue(CurLayer), r2, r2, sa, ea, id0, id1, PointCount, CurLayer, ArcType.CircleCR
            
            id0 = id1
        Next
        ux0 = Cos(PI_2 + (I - 0.5) * PI2 / n) * r1 + cx
        uy0 = Sin(PI_2 + (I - 0.5) * PI2 / n) * r1 + cy
        
        'AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        'id1 = PointCount
        
        GetCircleBy3Points PointList(id0).X, PointList(id0).Y, ux0, uy0, PointList(k).X, PointList(k).Y, X, Y, r2, sa, ea
        AddPoint X, Y, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
        AddArc X, Y, LayerZValue(CurLayer), r2, r2, sa, ea, id0, k, PointCount, CurLayer, ArcType.CircleCR
            
        PicPath.Cls
        DrawAll
        SaveUndo
    End If
    
End Sub

Sub PieceArray(ByVal GroupID As Long, ByVal AreaW As Double, ByVal AreaH As Double, ByVal dX As Double, ByVal dy As Double, ByVal angle As Double)
    Dim ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double, w As Double, h As Double
    Dim I As Long, j As Long, k As Long, group_id As Long
    
    If angle <> 0 Then
        RotateGroup GroupID, 0, 0, angle * PI_180
    End If
    GetGroupScale GroupID, ux0, uy0, ux1, uy1
    w = ux1 - ux0
    h = uy1 - uy0

    If w = 0 Then w = 20
    If h = 0 Then h = 20
        
    k = 0
    I = 0
    Do
        I = I + 1
        If I * h + (I - 1) * dy > AreaH Then
            Exit Do
        End If
        
        If I = 1 Then
            MoveGroup GroupID, -ux0, -uy0
            group_id = GroupID
        Else
            group_id = CopyGroup(GroupID)
            MoveGroup group_id, 0, (I - 1) * (h + dy)
        End If
        k = k + 1
        
        j = 0
        Do
            j = j + 1
            If (j + 1) * w + j * dX > AreaW Then
                Exit Do
            End If
            
            group_id = CopyGroup(group_id)
            MoveGroup group_id, w + dX, 0
            k = k + 1
        Loop
    Loop
End Sub


Private Sub CmdElevatorDown_Click()
    ElevatorDown
End Sub

Private Sub CmdElevatorUp_Click()
    ElevatorUp
End Sub

Private Sub CmdFeedBkP_Click()
    write_bit 0, VertMotorPort, 0
    If FeedByDCMotor = True Then
        FeedMMByDCMotor -val(TxtFeedMM.Text), 0
    Else
        FeedMM -val(TxtFeedMM.Text), Device_UseEncoder, 0
    End If
        
    IsRunning = False
End Sub

Private Sub CmdFeedBkV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Dim ret As Long
    '
    'IsRunning = True
    '
    'TmrFeed.Enabled = True
    'ret = set_startv(0, FeedAxis, Device_FeedStartV)
    'ret = set_speed(0, FeedAxis, Device_FeedSpeed)
    'ret = set_acc(0, FeedAxis, Device_FeedAccel)
    ''ret = continue_move1(0, FeedAxis, 1)
    'ret = pmove(0, FeedAxis, -10000000)
    
    write_bit 0, VertMotorPort, 0
    If FeedByDCMotor = True Then
        DCMotorFeedBWOn
    Else
        'FeedMM -10000, Device_UseEncoder, True, 2
        FeedV -1
    End If
End Sub

Private Sub CmdFeedBkV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long, nActPos As Long
        
    On Error Resume Next
    
'    If Device_UseEncoder = True Then
'        If CtrlCardType = 0 Then
'            get_actual_pos 0, FeedAxis, nActPos
'            FeedPulsCount = nActPos
'        Else
'            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
'            FeedPulsCount = nActPos
'        End If
'    End If
                                
    If FeedByDCMotor = True Then
        DCMotorFeedBWOff
    Else
        If CtrlCardType = 0 Then
            Ret = sudden_stop(0, FeedAxis)
        ElseIf CtrlCardType = 4 Then
            Ret = StopAxis(hDmc, FeedAxis)
        Else
            Ret = SetAxisStopDec_9030(0, FeedAxis, Device_FeedAccel)
            Ret = StopAxis_9030(0, FeedAxis)
        End If
        TxtFeedMM.SetFocus
    End If
    
'    If Device_UseEncoder = True Then
'        Wait 1
'        If CtrlCardType = 0 Then
'            get_actual_pos 0, FeedAxis, nActPos
'        Else
'            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
'        End If
'        LblEncoderOffset.caption = nActPos - FeedPulsCount
'    End If
        
    IsRunning = False
End Sub

Private Sub CmdFeedBkV2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    write_bit 0, VertMotorPort, 0
    If FeedByDCMotor = True Then
        DCMotorFeedBWOn2
    Else
        'FeedMM -10000, Device_UseEncoder, True, 2
        FeedV2 -1
    End If
End Sub

Private Sub CmdFeedBkV2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long, nActPos As Long
        
    On Error Resume Next
    
    If Device_UseEncoder = True Then
        If CtrlCardType = 0 Then
            get_actual_pos 0, FeedAxis, nActPos
        ElseIf CtrlCardType = 4 Then
            Ret = StopAxis(hDmc, FeedAxis)
        Else
            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
        End If
        FeedPulsCount = nActPos
    End If
                                
    If FeedByDCMotor = True Then
        DCMotorFeedBWOff2
    Else
        If CtrlCardType = 0 Then
            Ret = sudden_stop(0, FeedAxis)
        ElseIf CtrlCardType = 4 Then
            StopAxis hDmc, FeedAxis
        Else
            Ret = SetAxisStopDec_9030(0, FeedAxis, Device_FeedAccel * 25)
            Ret = StopAxis_9030(0, FeedAxis)
        End If
        TxtFeedMM.SetFocus
    End If
    
'    If Device_UseEncoder = True Then
'        Wait 1
'
'        'get_actual_pos 0, FeedAxis, nActPos
'        If CtrlCardType = 0 Then
'            get_actual_pos 0, FeedAxis, nActPos
'        Else
'            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
'        End If
'        LblEncoderOffset.caption = nActPos - FeedPulsCount
'    End If
        
    IsRunning = False
End Sub

Private Sub CmdFeedBkV2A_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdFeedBkV2_MouseDown Button, Shift, X, Y
End Sub

Private Sub CmdFeedBkV2A_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdFeedBkV2_MouseUp Button, Shift, X, Y
End Sub

Private Sub CmdFeedFWP_Click()
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
    ElseIf CtrlCardType = 4 Then
        
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
    If FeedByDCMotor = True Then
        FeedMMByDCMotor val(TxtFeedMM.Text), 0
    Else
        FeedMM val(TxtFeedMM.Text), Device_UseEncoder, 0
    End If
        
    IsRunning = False
End Sub
    
Private Sub CmdFeedFWV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'write_bit 0, VertMotorPort, 0
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
    ElseIf CtrlCardType = 4 Then
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
    If FeedByDCMotor = True Then
        DCMotorFeedFWOn
    Else
        'FeedMM 10000, Device_UseEncoder, True, 2
        FeedV
    End If
End Sub

Private Sub CmdFeedFWV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long, nActPos As Long
    
    On Error Resume Next
    
    If Device_UseEncoder = True Then
        'get_actual_pos 0, FeedAxis, nActPos
        If CtrlCardType = 0 Then
            get_actual_pos 0, FeedAxis, nActPos
        ElseIf CtrlCardType = 4 Then
            Ret = StopAxis(hDmc, FeedAxis)
        Else
            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
        End If
        FeedPulsCount = nActPos
    End If
                                
    If FeedByDCMotor = True Then
        DCMotorFeedFWOff
    Else
        If CtrlCardType = 0 Then
            Ret = sudden_stop(0, FeedAxis)
        ElseIf CtrlCardType = 4 Then
            Ret = StopAxis(hDmc, FeedAxis)
        Else
            Ret = SetAxisStopDec_9030(0, FeedAxis, Device_FeedAccel)
            Ret = StopAxis_9030(0, FeedAxis)
        End If
        TxtFeedMM.SetFocus
    End If
    
'    If Device_UseEncoder = True Then
'        Wait 1
'
'        'get_actual_pos 0, FeedAxis, nActPos
'        If CtrlCardType = 0 Then
'            get_actual_pos 0, FeedAxis, nActPos
'        Else
'            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
'        End If
'        LblEncoderOffset.caption = nActPos - FeedPulsCount
'    End If
    
    IsRunning = False
End Sub


Private Sub CmdFeedFWV2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    write_bit 0, VertMotorPort, 0
    If FeedByDCMotor = True Then
        DCMotorFeedFWOn2
    Else
        'FeedMM 10000, Device_UseEncoder, True, 2
        FeedV2
    End If
End Sub

Private Sub CmdFeedFWV2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long, nActPos As Long
    
    On Error Resume Next
    
    If Device_UseEncoder = True Then
        'get_actual_pos 0, FeedAxis, nActPos
        If CtrlCardType = 0 Then
            get_actual_pos 0, FeedAxis, nActPos
        ElseIf CtrlCardType = 4 Then
            Ret = StopAxis(hDmc, FeedAxis)
        Else
            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
        End If
        FeedPulsCount = nActPos
    End If
                                
    If FeedByDCMotor = True Then
        DCMotorFeedFWOff2
    Else
        If CtrlCardType = 0 Then
            Ret = sudden_stop(0, FeedAxis)
        ElseIf CtrlCardType = 4 Then
            StopAxis hDmc, FeedAxis
        Else
            Ret = SetAxisStopDec_9030(0, FeedAxis, Device_FeedAccel * 25)
            Ret = StopAxis_9030(0, FeedAxis)
        End If
        TxtFeedMM.SetFocus
    End If
    
'    If Device_UseEncoder = True Then
'        Wait 1
'
'        'get_actual_pos 0, FeedAxis, nActPos
'        If CtrlCardType = 0 Then
'            get_actual_pos 0, FeedAxis, nActPos
'        Else
'            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
'        End If
'        LblEncoderOffset.caption = nActPos - FeedPulsCount
'    End If
    
    IsRunning = False
End Sub

Private Sub CmdFeedFWV2A_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdFeedFWV2_MouseDown Button, Shift, X, Y
End Sub

Private Sub CmdFeedFWV2A_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdFeedFWV2_MouseUp Button, Shift, X, Y
End Sub

Private Sub CmdManualControl_Click()
    If Frame1.Visible = False Then
        Frame1.Move 8, 60
        Frame1.Visible = True
        If languageType = curLanguage Then
            CmdManualControl.caption = "关闭手动控制面板"
            CmdStop.FontName = "黑体"
        Else
            CmdManualControl.caption = "Hide Manual Panel"
        End If
    Else
        Frame1.Visible = False
        
        If languageType = curLanguage Then
            CmdManualControl.caption = "显示手动控制面板"
        Else
            CmdManualControl.caption = "Show Manual Panel"
        End If
    End If
End Sub

Private Sub CmdMotorStart_Click()
    CutOn
End Sub
Sub CutOn()
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 1
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 1
    Else
        WriteIoBit_9030 0, 1, VertMotorPort + 1
    End If
End Sub

Public Sub CmdMotorStop_Click()
    'write_bit 0, VertMotorPort, 0
    CutOff
End Sub
Sub CutOff()
    If CtrlCardType = 0 Then
        write_bit 0, VertMotorPort, 0
    ElseIf CtrlCardType = 4 Then
        WriteOutBit hDmc, VertMotorPort, 0
    Else
        WriteIoBit_9030 0, 0, VertMotorPort + 1
    End If
End Sub

Private Sub CmdMovePoint_Click()
    CurTool = ToolType.MoveElement_Point
    TxtCurTool.Text = "移点"
End Sub

Private Sub CmdMoveToInnerLine_Click()
    'write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
    'Wait 0.1
    'write_bit 0, VertMoveDownPort, 0
    '
    'Do While read_bit(0, VertHighSensor) = 1
    '    DoEvents
    'Loop
    
    StopRunning = False
    IsRunning = True
    VertMoveDown 'Up
    
    PullBack 0, UseMagnetDO
    'VertAngle 0
End Sub

Private Sub CmdMoveToOutLine_Click()
    'write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
    'Wait 0.1
    'write_bit 0, VertMoveDownPort, 0
    '
    'Do While read_bit(0, VertHighSensor) = 1
    '    DoEvents
    'Loop
    
    StopRunning = False
    IsRunning = True
    VertMoveDown 'Up
    
    PushOut Device_CutDepth, UseMagnetDO
    'VertAngle -180
End Sub

Private Sub CmdPause_Click()
    PauseRunning = True
End Sub

Private Sub CmdResetAxisBend_Click()
    set_command_pos 0, BendAxis, 0
    set_actual_pos 0, BendAxis, 0
    TmrBend_Timer
End Sub

Private Sub CmdResetOrg_Click()
    StopRunning = False
    IsRunning = True
    
    'VertMoveFar
    LblReset1.BackColor = RGB(0, 255, 0)
    
    LblReset2.BackColor = RGB(255, 0, 0)
    If Not DebugWithoutSensor Then
        If CtrlCardType = 1 Then
            VertReset_9030
        ElseIf CtrlCardType = 4 Then
            BendReset_GALIL_V8
        End If
        If StopRunning = True Then
            LblReset2.BackColor = RGB(200, 200, 200)
            GoTo Exit_Sub
        End If
    End If
    LblReset2.BackColor = RGB(0, 255, 0)
    
    LblReset3.BackColor = RGB(255, 0, 0)
    If Not DebugWithoutSensor And Device_BenderHome = True Then
        If CtrlCardType = 1 Then
            BendReset_9030_V8
        ElseIf CtrlCardType = 4 Then
            VertReset_GALIL_V8
        End If
        If StopRunning = True Then
            LblReset3.BackColor = RGB(200, 200, 200)
            GoTo Exit_Sub
        End If
    End If
    LblReset3.BackColor = RGB(0, 255, 0)
    IsRunning = False
    ResetOK = True
    Exit Sub
    
Exit_Sub:
    IsRunning = False
    ResetOK = False
End Sub

Private Sub CmdRestAxisFeed_Click()
    set_command_pos 0, FeedAxis, 0
    set_actual_pos 0, FeedAxis, 0
    'ShowFeedPos
End Sub

Private Sub CmdResume_Click()
    PauseRunning = False
End Sub
Sub createfile2printvertpoint()
    fileNum_g = FreeFile
    Open "c:\hd_debug\" + "vertPoint.txt" For Output As #fileNum_g
    Print #fileNum_g, "序号"; Tab(8); "实际切点"; Tab(20); "理论切点"
    
End Sub
Sub closefileofprintvertpoint()
    Close #fileNum_g
End Sub

Function BackSet_func()
Dim pos As Long
Dim state As Integer
Dim Ret As Integer
    'Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
    Ret = SetVel(hDmc, FeedAxis, Device_FeedSpeed / 2)   '20140925
    Ret = SetAcc(hDmc, FeedAxis, Device_FeedAccel)
    Ret = SetDec(hDmc, FeedAxis, Device_FeedAccel)
            
    StopAxis hDmc, FeedAxis
    pos = GetPos(hDmc, FeedAxis)
    PosMoveAbs hDmc, FeedAxis, pos - Device_BackSet * Device_PulsPerMM
    
    CmdRun.Enabled = False
    state = 1
    Do
        Sleep 5
        state = GetStatus(hDmc, FeedAxis)
        If state <> 1 Then
            Exit Do
        End If
        If StopRunning = True Then
            CmdRun.Enabled = True
            Exit Function
        End If
        DoEvents
    Loop
    
    CmdRun.Enabled = True
    
    Wait 0.5
    pos = GetPos(hDmc, FeedAxis)
    PosMoveAbs hDmc, FeedAxis, pos + 10 * Device_PulsPerMM
    
    state = GetStatus(hDmc, FeedAxis)
    Sleep 5
    state = 1
    Do
        Sleep 5
        state = GetStatus(hDmc, FeedAxis)
        If state <> 1 Then
            Exit Do
        End If
        If StopRunning = True Then
            Exit Function
        End If
        DoEvents
    Loop
    
    DefinePos hDmc, FeedAxis, 0
End Function
Private Sub PanButtonBackSet_Click()
Dim pos As Long
    Dim state As Integer
    Dim Ret As Integer
    
    If CtrlCardType = 4 Then
        BackSet_func
        Exit Sub
    End If
    
    Ret = SetAxisStartVel_9030(0, FeedAxis, Device_FeedStartV)
    Ret = SetAxisVel_9030(0, FeedAxis, Device_FeedSpeed / 2) '20140925
    Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
    Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
            
    CeaseAxis_9030 0, FeedAxis
    pos = ReadAxisPos_9030(0, FeedAxis)
    SetAxisPos_9030 0, FeedAxis, pos - Device_BackSet * Device_PulsPerMM
    StartAxis_9030 0, FeedAxis
    CmdRun.Enabled = False
    state = 1
    Do
        Sleep 5
        state = ReadAxisState_9030(0, FeedAxis)
        If state <> 1 Then
            Exit Do
        End If
        If StopRunning = True Then
            CmdRun.Enabled = True
            Exit Sub
        End If
        DoEvents
    Loop
    
    CmdRun.Enabled = True
    
    Wait 0.5
    pos = ReadAxisPos_9030(0, FeedAxis)
    SetAxisPos_9030 0, FeedAxis, pos + 10 * Device_PulsPerMM
    StartAxis_9030 0, FeedAxis
    state = ReadAxisState_9030(0, FeedAxis)
    Sleep 5
    state = 1
    Do
        Sleep 5
        state = ReadAxisState_9030(0, FeedAxis)
        If state <> 1 Then
            Exit Do
        End If
        If StopRunning = True Then
            Exit Sub
        End If
        DoEvents
    Loop
    
    Home_9030 0, FeedAxis
End Sub

Private Sub PanButtonRun_Click()
    '---------------define private variable------------------------
    Dim obj As Object
    Dim ds As Double, ux As Double, uy As Double, feed_puls As Long
    Dim I As Long, i0 As Long
    Dim PosToGo, PosToStop As Long
    Dim status_feedAxis, Ret As Integer
    Dim AngleToAchieve As Double
    Dim PrePosBender As Double  '弯弧器前器状态
    Dim FeedLength_Encoder As Double
    Dim BendDis As Double
    Dim getSetCutPos As Double
    '--------------------------------------------------------------
    
    '----------------------Initialize Globle Variable and Private Variable------------------------------
    StopRunning = False
    AngleToAchieve = 0
    FeedPulsPerMM = Device_PulsPerMM
    
    GetPathXYByFeedPuls -1, ux, uy 'reset
    
    GetPathXYByVertPuls -1, ux, uy  '垂直轴复位
    
    FeedMark_x0 = -9999
    FeedMark_y0 = -9999
    VertMark_x0 = -9999
    VertMark_y0 = -9999
    
    Home_9030 0, 0
    Home_9030 0, 1
    Home_9030 0, 2
    Home_9030 0, 3
    
    HomeFB_9030 0, FeedAxis
    
    CurIndex = 0    '当前运行的段号
    PrePosBender = 0
    
    OldBendangle_g = 0
    PanButtonRun.Enabled = False
    
'    For Each obj In FrmMain
'        obj.Enabled = False
'    Next
    CmdManualControl.Enabled = False
    CmdFeedBkV2A.Enabled = False
    CmdFeedFWV2A.Enabled = False
    PanButton2.Enabled = False
    PanButton3.Enabled = False
    PanButton11.Enabled = False
    PanButton1.Enabled = False
    PanButton5.Enabled = False
    CmdResetOrg.Enabled = False
    
    CmdStop.Enabled = True
    CmdPause.Enabled = True
    CmdResume.Enabled = True
    TxtStatistics.Enabled = True
    TxtStatistics.Text = ""
    
    Timer1.Enabled = True
    TmrGetCurRunState.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text9.Enabled = True
    
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = True
    Label14.Enabled = True
    Label16.Enabled = True
    
    SetAxisStartVel_9030 0, BendAxis, Device_FeedStartV
    SetAxisAcc_9030 0, BendAxis, Device_FeedAccel
    SetAxisDec_9030 0, BendAxis, Device_BendAccel
    SetAxisVel_9030 0, BendAxis, Device_FeedSpeed

    TotalPathOutLength = 0
    PathOutputPointCount = 0
    
    CalculateAllPath Max(val(TxtRunN.Text), 1)  '计算路径，至少重复1次
    PosMove_9030 0
    Sleep 10
    Home_9030 0, FeedAxis
    HomeFB_9030 0, FeedAxis
    createfile2printvertpoint
    '----------------------------------------------------------------------------------------------------
    If TotalPathOutLength > 0 Then
        For I = 1 To PathOutputPointCount
        
'            FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
'            ShowFeedMarkPoint
'            ShowVertMarkPoint
            
            If StopRunning = True Then
                Exit For
            End If
            
            If PathOutputPoint(I).VertType = 2 Or PathOutputPoint(I).VertType = 1 Then
            
                If PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 4 Then '首端点'末端点
                    PushOut Device_CutDepth, UseMagnetDO
                End If
                PosToGo = PathOutputPoint(I).LengthFromStart * FeedPulsPerMM
                
                PosMove_9030 PosToGo
                
                WaitAxisRunComplete FeedAxis
                Sleep (500)
                
                '打印切割点位置，换行
                TxtStatistics.Text = TxtStatistics.Text + "   Len:" + str(Round(PathOutputPoint(I).LengthFromStart, 2)) + vbCrLf
                
                If Device_UseEncoder = True Then
                    'If Device_UseEncoder = False Then
                    '    FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
                    'Else
                        FeedPulsCount = ReadAxisEncodePos_9030(0, FeedAxis)
                    'End If
                    
                    getSetCutPos = PathOutputPoint(I).LengthFromStart '得到指定的切割位置
                    posErr = Round(FeedPulsCount / Device_EncoderPulsPerMM - PathOutputPoint(I).LengthFromStart, 3)
                    If posErr > ErrLmt Or posErr < -1 * ErrLmt Then
                        'TextBendLength.Text = Round(FeedPulsCount / FeedPulsPerMM - 0.4, 3)
                        posErrCompensation posErr
                    End If
                End If
                            
                If Device_UseEncoder = True Then
                    FeedLength_Encoder = ReadAxisEncodePos_9030(0, FeedAxis) / Device_EncoderPulsPerMM '读编码器值长度
                Else
                    FeedLength_Encoder = ReadAxisPos_9030(0, FeedAxis) / Device_PulsPerMM '读脉冲值长度
                End If
                Print #fileNum_g, I; Tab(8); Round(FeedLength_Encoder, 3); Tab(23); _
                                            Round(PathOutputPoint(I).LengthFromStart, 3); Tab(38); _
                                            Round(FeedLength_Encoder - PathOutputPoint(I).LengthFromStart, 3)
            
                '切割操作
                'Wait 2
                VertInnerAngle_done PathOutputPoint(I).AngleToNext + 180 - 22.5, False
                
                SetAxisStartVel_9030 0, FeedAxis, Device_FeedStartV
            ElseIf PathOutputPoint(I).VertType = -9 Then '首端点'末端点
                Exit For
            ElseIf PathOutputPoint(I).VertType = 99999 Then                             '送料总长
                PosToGo = PathOutputPoint(I).LengthFromStart * FeedPulsPerMM
                
                PosMove_9030 PosToGo
                
                WaitAxisRunComplete FeedAxis
                Sleep (10)
                '退出
            ElseIf Abs(PathOutputPoint(I).Radius3P) > 5 And _
                Abs(PathOutputPoint(I).AngleToNext - PrePosBender) > 2 Then             '弯弧连接角度过大必须停止
                PrePosBender = PathOutputPoint(I).AngleToNext
                PosToGo = PathOutputPoint(I).LengthFromStart * FeedPulsPerMM
                
                PosMove_9030 PosToGo
                
                WaitAxisRunComplete FeedAxis
                Sleep (10)
                
                '弯弧
                'BendAngleAbs PathOutputPoint(i).AngleToNext + Device_EmptyDegree, False
                'BendAngleByRadius -VTDir * PathOutputPoint(i).Radius3P, True
                WaitAxisRunComplete BendAxis
            ElseIf Abs(PathOutputPoint(I).Radius3P) > 5 And _
                Abs(PathOutputPoint(I).Radius3P) < Device_BeatMaxRadius Then            '小于最小弯弧半径则进行拍弧
                PrePosBender = PathOutputPoint(I).AngleToNext
                PosToGo = PathOutputPoint(I).LengthFromStart * FeedPulsPerMM
                
                PosMove_9030 PosToGo
                
                WaitAxisRunComplete FeedAxis
                Sleep (10)
                
                '拍弧
                'BendDis = PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart
                'BeatRealAngle -VTDir * PathOutputPoint(I).AngleToNext, BendDis      '*********拍弧********'
                BendAngleByRadius -VTDir * PathOutputPoint(I).Radius3P, True
                WaitAxisRunComplete BendAxis
            ElseIf PathOutputPoint(I).VertType = -9 Then
                CeaseAxis_9030 0, FeedAxis
                CmdStop_Click
                TmrGetCurRunState.Enabled = False
                Exit For
            ElseIf Abs(PathOutputPoint(I).AngleToNext) > 80 Or Abs(PathOutputPoint(I).AngleToNext) < 0.1 Then
                'PrePosBender = PathOutputPoint(i).AngleToNext
                If Abs(PathOutputPoint(I).AngleToNext) > 80 Or Abs(PathOutputPoint(I).AngleToNext) < 0.1 Then
                    PosToGo = PathOutputPoint(I).LengthFromStart * FeedPulsPerMM
                
                    PosMove_9030 PosToGo
                    
                    WaitAxisRunComplete FeedAxis
                    BendAngleAbs 0, False
                    'BendAngleByRadius -VTDir * PathOutputPoint(i).Radius3P, True
                    WaitAxisRunComplete BendAxis
                End If
            Else
                PosToGo = PathOutputPoint(I).LengthFromStart * FeedPulsPerMM
                
                PosMove_9030 PosToGo
            End If
                
            PrePosBender = PathOutputPoint(I).AngleToNext
            
            DoEvents
        Next
    End If
    
    
    '------------运行结束退出处理-----------------------
    
    '------------关闭记录铣刀点位置的文件-----------------------
    closefileofprintvertpoint
    PicPath.DrawWidth = 1

'    For Each obj In FrmMain
'        If Not (TypeOf obj Is Timer) Then
'            obj.Enabled = True
'        End If
'    Next
    TmrGetCurRunState.Enabled = False
    CmdStop_Click
    
    
    
    IsRunning = False
End Sub




Private Sub CmdRun_Click()
'***************************************************************************************************************************
'                                               运行 命令处理函数（）
'***************************************************************************************************************************
    Dim obj As Object
    Dim I As Long, i0 As Long, I1 As Long, lfs0 As Double, lfs00 As Double, blfs0 As Double, BendDis As Double
    Dim job_done As Boolean, paused As Boolean
    Dim vi0 As Long, vi1 As Long, vlfs0 As Double, vblfs0 As Double
    
    'Dim j As Long, start_id As Long
    
    Dim ds As Double, ux As Double, uy As Double, feed_puls As Long
    Dim t As Long, p As Long ', x0 As Single, y0 As Single
    Dim next_piece As Boolean
    Dim next_piece_cur_puls As Long

    Dim nLogPos As Long                   '逻辑位置
    Dim nActPos As Long                   '实际位置
    Dim nSpeed As Long                    '运行速度
    
    Dim puls_per_mm As Double
    Dim Section As Long
    Dim m As Long
    Dim nActPos0 As Long
    Dim k As Long, k0 As Long, d0 As Double, d1 As Double
    
    Dim fast_feed_puls As Long
    Dim fast_feed_puls0 As Long
    
    Dim vert_zone_feed_puls As Long
    Dim vert_zone_feed_puls0 As Long
    
    Dim FeedPulsCount0 As Long
    Dim Status As Integer
    
    Dim CurDis2VertPoint As Double '当前点到切割点距离
    Dim rtnbox As Integer
    
    Dim xiaohulianxu As Boolean
    
    Dim iStart As Integer
    
    Dim getSetCutPos As Double
    
    Dim tempVertMaxInnerAngle As Double '20140726
    
    Dim FirstEndpointofLazhi As Boolean '定义标志，判断是否为拉直距离最后一个端点，亦即弯弧第一个起点。弯弧第一点需停止进料弯弧。
    On Error Resume Next
    
    Dim CurCutIndex As Integer
    Dim NextCutIndex As Integer
    
    Dim Cur99999Index As Integer
    Dim Next99999Index As Integer
    
    CurCutIndex = 0
    NextCutIndex = FindNextCutPointIndex(CurCutIndex)
    
    Cur99999Index = 0
    Next99999Index = FindNext99999PointIndex(CurCutIndex)
    
    xiaohulianxu = IIf(CheckXiaohu.value = 0, False, True)
    
    VertCnt = -1
    
    '20140806拉直距离判断，若小弧连续有效，则提前5mm起效；否则按参数设定值起效
    If xiaohulianxu = True Then
        LazhiDis = Device_MinBendDisMM - 5
    Else
        LazhiDis = Device_MinBendDisMM
    End If
    
    Bendlength = 0
    FirstEndpointofLazhi = True
    
    BeatAngelPointLength = -1 * Device_TurnPointOffsetMM '给BeatAngelPointLength赋一个负值是要使该参数不起作用
    
    '选择使用余料提示
    If FirstRun = False Then
        If curLanguage = 0 Then
            rtnbox = MsgBox("已选择使用余料,是否继续加工？", vbYesNo Or vbQuestion, "系统提示")
        Else
            rtnbox = MsgBox("To use the end material. Continue?", vbYesNo Or vbQuestion, "Warning")
        End If
        If rtnbox = vbNo Then
            For Each obj In FrmMain
                If Not (TypeOf obj Is Timer) Then
                    obj.Enabled = True
                End If
            Next
            CmdStop_Click
            Exit Sub
        End If
    End If
    
    '铣刀角度复位提示
'    If VertResetOK = False Then
'        If curLanguage = 0 Then
'            rtnbox = MsgBox("设备铣刀开机未复位,是否继续加工？", vbYesNo Or vbQuestion, "系统提示")
'        Else
'            rtnbox = MsgBox("Mill Cutter dose not reset to home. Continue？", vbYesNo Or vbQuestion, "Warning")
'        End If
'        If rtnbox = vbNo Then
'            For Each obj In FrmMain
'                If Not (TypeOf obj Is Timer) Then
'                    obj.Enabled = True
'                End If
'            Next
'            CmdStop_Click
'            Exit Sub
'        End If
'    End If
    
    xiaohulianxu = IIf(CheckXiaohu.value = 0, False, True)
    'xiaohulianxu = True
       
    '送料电机基本参数：起始速度，加速度，减速度
    If CtrlCardType = 1 Then
        SetAxisStartVel_9030 0, FeedAxis, Device_FeedStartV
        SetAxisAcc_9030 0, FeedAxis, Device_FeedAccel
        SetAxisDec_9030 0, FeedAxis, Device_FeedAccel
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'FrmRegistration.CheckRegistrationAgain
    
'    If ResetOK = False Then
'        MsgBox "软件启动后未成功进行设备复位。请复位后重新运行", vbExclamation + vbOKOnly, ""
'        Exit Sub
'    End If
    
    'CmdResetOrg_Click
    
     '==========================
    OldBendangle_g = 0
    If CtrlCardType = 1 Then
        SetAxisPos_9030 0, BendAxis, 0
        StartAxis_9030 0, BendAxis
   
        Do
            
            Status = ReadAxisState_9030(0, BendAxis)
            If Status <> 1 Then
                Exit Do
            End If
            
            DoEvents
        Loop
    ElseIf CtrlCardType = 4 Then
        SetVel hDmc, BendAxis, Device_BendSpeed
        SetAcc hDmc, BendAxis, Device_BendAccel
        SetDec hDmc, BendAxis, Device_BendAccel
        PosMoveAbs hDmc, BendAxis, 0
        Do
            
            Status = GetStatus(hDmc, BendAxis)
            If Status = 0 Then
                Exit Do
            End If
            
            DoEvents
        Loop
    End If
    '==========================
    
    
    
    next_piece = False
    IsRunning = True

    For Each obj In FrmMain
        obj.Enabled = False
    Next
    CmdStop.Enabled = True
    CmdPause.Enabled = True
    CmdResume.Enabled = True
    TxtStatistics.Enabled = True
    
    Timer1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text9.Enabled = True
    
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = True
    Label14.Enabled = True
    Label16.Enabled = True

    TotalPathOutLength = 0
    PathOutputPointCount = 0
    
    CalculateAllPath Max(val(TxtRunN.Text), 1)  '计算路径，至少重复1次
        
    'CalculatePath_ 'True
    
    StatusBar1.Panels.Item(2).Text = ""
    StatusBar1.Panels.Item(3).Text = ""
    StatusBar1.Panels.Item(4).Text = ""
    StatusBar1.Panels.Item(6).Text = ""
    StatusBar1.Panels.Item(7).Text = ""
    StatusBar1.Panels.Item(9).Text = ""
    StatusBar1.Panels.Item(11).Text = ""
    
    FraEdit.Visible = False
    TxtStatistics.Visible = True
    TxtStatistics.Text = ""
    
    '最短铣角间距提示
    If searchMinVertDis = True Then
        If curLanguage = 0 Then
            rtnbox = MsgBox("图形中有线段小于<最短铣角间距>,是否继续加工？", vbYesNo Or vbQuestion, "系统提示")
        Else
            rtnbox = MsgBox("There is a milling length shorter than <Milling Min Interval>. Continue?", vbYesNo Or vbQuestion, "Warning!")
        End If
        If rtnbox = vbNo Then
            For Each obj In FrmMain
                If Not (TypeOf obj Is Timer) Then
                    obj.Enabled = True
                End If
            Next
            CmdStop_Click
            Exit Sub
        End If
    End If
    
    '--------------创建记录铣刀点位置的文件--------------------
    createfile2printvertpoint
    
    If TotalPathOutLength > 0 Then
        m = 0
        OldBendangle_g = 0
        TxtRunCount.Text = str(1 + m) '当前加工遍数

        'PicPathCls
        'DrawAll
        
        StopFeed = False
        StopRunning = False
        PauseRunning = False
        paused = False
        
        PortBit(1) = 0
        PortBit(2) = 0
        PortBit(3) = 0
        PortBit(4) = 0
        If CtrlCardType = 0 Then
            write_bit 0, VertMotorPort, 0
            write_bit 0, VertClosePort, 0
            write_bit 0, VertMoveUpPort, 0
            write_bit 0, VertMoveDownPort, 0
        ElseIf CtrlCardType = 1 Then
            WriteIoBit_9030 0, 0, VertMotorPort + 1
            WriteIoBit_9030 0, 0, VertClosePort + 1
            WriteIoBit_9030 0, 0, VertMoveUpPort + 1
            WriteIoBit_9030 0, 0, VertMoveDownPort + 1
            Home_9030 0, BendAxis
        End If
        
        
        '送料轴位置清零
        If Device_UseEncoder = False Then           '不使用编码器
            'set_command_pos 0, FeedAxis, 0 '-Device_HeadDistance * Device_PulsPerMM
            If CtrlCardType = 1 Then
                Home_9030 0, FeedAxis
            ElseIf CtrlCardType = 4 Then
                DefinePos hDmc, FeedAxis, 0
            End If
            FeedPulsPerMM = Device_PulsPerMM
'            Device_FeedOffset = 0
        Else
            If CtrlCardType = 0 Then
                set_actual_pos 0, FeedAxis, 0 '-Device_HeadDistance * Device_EncoderPulsPerMM
            ElseIf CtrlCardType = 1 Then
                HomeFB_9030 0, FeedAxis
            ElseIf CtrlCardType = 4 Then
                DefineEnc hDmc, FeedAxis, 0
            End If
            FeedPulsPerMM = Device_EncoderPulsPerMM
        End If
        
        'ShowFeedPos
        
        
        
        GetPathXYByFeedPuls -1, ux, uy 'reset
        
        GetPathXYByVertPuls -1, ux, uy  '垂直轴复位
        
        FeedMark_x0 = -9999
        FeedMark_y0 = -9999
        VertMark_x0 = -9999
        VertMark_y0 = -9999
        
        'If Device_VertMotorZoneMM = 0 Then
        '    Device_VertMotorZoneMM = Device_FastSpeedMinLenMM * 2
        'End If
        
        i0 = 0
        vi0 = 0
        p = 0
        lfs0 = 0 '-Device_HeadDistance
        blfs0 = 0
        k0 = 0
        
        '获取起始节点序号 iStart
        'PathOutputPoint(iStart).LengthFromStart > Device_HeadDistance
        '将当前位置（脉冲和编码器）设置成Device_HeadDistance
        If FirstRun = False Then
            For I = 1 To PathOutputPointCount
                If PathOutputPoint(I).LengthFromStart > Device_HeadDistance Then
                    iStart = I
                    If CtrlCardType = 1 Then
                        SetAxisOffset_9030 0, FeedAxis, Device_HeadDistance * Device_PulsPerMM
                        SetAxisFBOffset_9030 0, FeedAxis, Device_HeadDistance * Device_EncoderPulsPerMM
                    ElseIf CtrlCardType = 4 Then
                        DefinePos hDmc, FeedAxis, Device_HeadDistance * Device_PulsPerMM
                        DefineEnc hDmc, FeedAxis, Device_HeadDistance * Device_EncoderPulsPerMM

                    End If
                    Exit For
                End If
            Next
        Else
            iStart = 1
            'FirstRun = False   '不使用首端不切割功能
        End If
               
        For I = iStart To PathOutputPointCount
            If StopRunning = True Then
                Exit For
            End If
                            
            If CtrlCardType = 1 Then
                nLogPos = ReadAxisPos_9030(0, FeedAxis)
                nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
            ElseIf CtrlCardType = 4 Then
                nLogPos = GetPos(hDmc, FeedAxis)
                nActPos = GetPosEnc(hDmc, FeedAxis)

            End If
            '决定是否使用编码器脉冲
            If Device_UseEncoder = False Then
                FeedPulsCount0 = nLogPos
            Else
                FeedPulsCount0 = nActPos
            End If

            '====================================
            For k = I + 1 To PathOutputPointCount
                If PathOutputPoint(k).VertType > 0 Or (PathOutputPoint(k).VertType = 0 And Abs(PathOutputPoint(k).Radius3P) > 0 And Abs(PathOutputPoint(k).Radius3P) <= Device_BeatMaxRadius) Then
                    Exit For
                End If
            Next
            If k > PathOutputPointCount Then k = PathOutputPointCount
            
            ds = PathOutputPoint(k).LengthFromStart - PathOutputPoint(I).LengthFromStart
            
            fast_feed_puls = fast_feed_puls0 '用上一段的数据
            
            If ds - Device_FastSpeedMinLenMM > 0 Then
                fast_feed_puls0 = (ds - Device_FastSpeedMinLenMM) * FeedPulsPerMM
            Else
                fast_feed_puls0 = 0
            End If
            
        If I <> 1 Then
'*********************************原来程序算法*******************
            If fast_feed_puls > 0 And Abs(PathOutputPoint(I).Radius3P) < Device_BeatMaxRadius Then '20140927
Debug.Print I; ">>>"
                If CtrlCardType = 1 Then
                    FeedV3_9030
                ElseIf CtrlCardType = 4 Then
                    FeedV3_GALIL
                End If
            Else
Debug.Print I; "<<<"
                'FeedV   '增量运动10000000*dr个脉冲
            End If
            If CtrlCardType = 1 Then
                FeedV3_9030
            ElseIf CtrlCardType = 4 Then
                FeedV3_GALIL
            End If
           
'****************************************************************
        End If
            
'            CurDis2VertPoint = PathOutputPoint(i).LengthFromStart - nActPos / Device_EncoderPulsPerMM
'            If CurDis2VertPoint > Device_FastSpeedMinLenMM Then
'                 FeedV3_9030
'            Else
'                 FeedV_9030
'            End If
            
            '=================================
            For k = I + 1 To PathOutputPointCount
                If PathOutputPoint(k).VertType > 0 Then
                    Exit For
                End If
            Next
            If k > PathOutputPointCount Then k = PathOutputPointCount
            
            ds = PathOutputPoint(k).LengthFromStart - PathOutputPoint(I).LengthFromStart
            
            vert_zone_feed_puls = vert_zone_feed_puls0
            
            If ds - Device_VertMotorZoneMM > 0 Then
                vert_zone_feed_puls0 = (ds - Device_VertMotorZoneMM) * FeedPulsPerMM
            Else
                vert_zone_feed_puls0 = 0
            End If
            
            If vert_zone_feed_puls > 0 Then
                FeedIntoVertMotorZone = False
            Else
                FeedIntoVertMotorZone = True
            End If
            
            '=================================
                        
            'If Not DebugWithoutSensor Then
                For k = I To PathOutputPointCount
                    If PathOutputPoint(k).VertType = VT1 Or _
                        Device_AmericanMaterial = False And _
                            (((ChkStartPointVert90.value = 1 Or Abs(PathOutputPoint(k).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(k).Type = 3) Or _
                            ((ChkEndPointVert90.value = 1 Or Abs(PathOutputPoint(k).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(k).Type = 4)) Or _
                        Device_AmericanMaterial = True And _
                            (PathOutputPoint(k).Type = 3 Or PathOutputPoint(k).Type = 5) Then   '内角 或 线段两端 或 美国型材两端
                        
                        If k > k0 Then
                            
                            'If Device_AmericanMaterial = True And (PathOutputPoint(k).Type = 3 Or PathOutputPoint(k).Type = 5) Then '美国型材末端
                            If PathOutputPoint(k).Type = 4 Then
                                VertInnerAngle_prev Device_TailVertAngle, False        '*********运动函数********'
                                'StopFeedV I
                                'PushOut Device_CutDepth, UseMagnetDO
                                'FeedV
                        
                            ElseIf Device_AmericanMaterial = False And _
                                (((ChkStartPointVert90.value = 1 Or Abs(PathOutputPoint(k).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(k).Type = 3) Or _
                                ((ChkEndPointVert90.value = 1 Or Abs(PathOutputPoint(k).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(k).Type = 4)) Then
                                
                                '经此一句，程序的第一点的angletonext 变为180,铣刀切割角度就不对了。问题出在这里
                                'PathOutputPoint(k).AngleToNext = GetOutputEndPointAngle(k)
                                
                                'PullBack 0, UseMagnetDO
                                
                                If PathOutputPoint(k).Type = 4 Then
                                    VertEndAngle_prev 0, False              '*********运动函数********'
                                Else
                                    VertEndAngle_prev 1, False              '*********运动函数********'
                                End If
                            Else
                                PullBack 0, UseMagnetDO
                                If Device_AmericanMaterial = False Then
                                    'Debug.Print ">>>> inner k="; k, PathOutputPoint(k).AngleToNext
                                    VertInnerAngle_prev PathOutputPoint(k).AngleToNext, False         '*********运动函数********'  step1
                                Else
                                    VertInnerAngle_prev 0, False                                       '*********运动函数********'
                                End If
                            End If
                        End If
                        k0 = k
                        Exit For
                        
                    ElseIf PathOutputPoint(k).VertType = VT2 Then  '外角
                        
                        If k > k0 Then
                            If (PathOutputPoint(k).Type = 3 Or PathOutputPoint(k).Type = 5) Then
                                'VertInnerAngle_prev Device_TailVertAngle, False        '*********运动函数********'
                                'PushOut Device_CutDepth, UseMagnetDO
                                'FeedV
                            ElseIf Device_AmericanMaterial = False And Device_KareanMaterial = False Then
                                PullBack 0, UseMagnetDO
                                'Debug.Print ">>>> outer k="; k, PathOutputPoint(k).AngleToNext
                                VertOuterAngle_prev PathOutputPoint(k).AngleToNext, False              '*********运动函数********'
                                
                            ElseIf Device_KareanMaterial = True Then
                                PullBack 0, UseMagnetDO
                                'VertInnerAngle_prev 0                           '*********运动函数********'
                                tempVertMaxInnerAngle = Device_VertMaxInnerAngle    '暂存铣内角最大角度
                                Device_VertMaxInnerAngle = Device_VertMaxOuterAngle '用铣外加最大角度赋值给铣内角最大角度
                                VertInnerAngle_prev PathOutputPoint(k).AngleToNext + 180, False      '替换上一句
                                Device_VertMaxInnerAngle = tempVertMaxInnerAngle    '回复铣内角最大角度
                                'PushOut 180
                            Else
                                PullBack 0, UseMagnetDO
                                VertOuterAngle_prev 0, False                           '*********运动函数********'
                                
                            End If
                        End If
                        k0 = k
                        Exit For
                    End If
                Next
            'End If
                        
            Do
                'Wait 0.01
                If CtrlCardType = 4 Then
                    'FeedV3_GALIL '本来就不需要，因为单独ST不能实现ceaseaxis功能，才出此下策。
                End If
                If StopRunning = True Then
                    Exit For
                End If
            
                Do
                    If paused = False And PauseRunning = False Then
                        Exit Do
                    ElseIf paused = False And PauseRunning = True Then
                        If CtrlCardType = 0 Then
                            StopFeedV '8940A1 没有 pause/resume 功能
                            
                            write_bit 0, VertMotorPort, 0
                            write_bit 0, VertClosePort, 0
                            write_bit 0, VertMoveUpPort, 0
                            write_bit 0, VertMoveDownPort, 0
                        ElseIf CtrlCardType = 1 Then
                            StopFeedV '8940A1 没有 pause/resume 功能
                            
                            WriteIoBit_9030 0, 0, VertMotorPort + 1
                            WriteIoBit_9030 0, 0, VertClosePort + 1
                            WriteIoBit_9030 0, 0, VertMoveUpPort + 1
                            WriteIoBit_9030 0, 0, VertMoveDownPort + 1
                        ElseIf CtrlCardType = 4 Then
                            StopFeedV '8940A1 没有 pause/resume 功能
                            
                            WriteOutBit hDmc, VertMotorPort, 0
                            
                            
                        End If
                        
                        paused = True
                        
                    ElseIf paused = True And PauseRunning = False Then
                        paused = False
                        If fast_feed_puls > 0 Then
                            'FeedV2
                            If CtrlCardType = 0 Then
                                Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
                            ElseIf CtrlCardType = 4 Then
                                nLogPos = GetPos(hDmc, FeedAxis)
                                nActPos = GetPosEnc(hDmc, FeedAxis)
                            Else
                                nLogPos = ReadAxisPos_9030(0, FeedAxis)
                                nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                            End If
                            
                            
                            If Device_UseEncoder = False Then
                                FeedPulsCount = nLogPos
                            Else
                                FeedPulsCount = nActPos
                            End If
                            
'                            If FeedPulsCount - FeedPulsCount0 < fast_feed_puls Then
'                                FeedV3 fast_feed_puls - (FeedPulsCount - FeedPulsCount0)
'                                Wait 2
'                            Else
                                FeedV3_GALIL
'                            End If
                        Else
                            FeedV3_GALIL
                        End If
                        
                        If CtrlCardType = 0 Then
                            write_bit 0, VertMotorPort, PortBit(1)
                            write_bit 0, VertClosePort, PortBit(2)
                            write_bit 0, VertMoveUpPort, PortBit(3)
                            write_bit 0, VertMoveDownPort, PortBit(4)
                        ElseIf CtrlCardType = 4 Then
                            WriteOutBit hDmc, VertMotorPort, 0
                        Else
                            WriteIoBit_9030 0, PortBit(1), VertMotorPort + 1
                            WriteIoBit_9030 0, PortBit(2), VertClosePort + 1
                            WriteIoBit_9030 0, PortBit(3), VertMoveUpPort + 1
                            WriteIoBit_9030 0, PortBit(4), VertMoveDownPort + 1
                        End If
                        Exit Do
                    End If
                    'StatusBar1.Panels.Item(4).Text = "Speed: 0"
                    DoEvents
                Loop
                
                'Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
                If CtrlCardType = 0 Then
                    Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
                ElseIf CtrlCardType = 4 Then
                    nLogPos = GetPos(hDmc, FeedAxis)
                    nActPos = GetPosEnc(hDmc, FeedAxis)
                Else
                    nLogPos = ReadAxisPos_9030(0, FeedAxis)
                    nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                End If
                
                If Device_UseEncoder = False Then
                    FeedPulsCount = nLogPos
                Else
                    FeedPulsCount = nActPos
                End If
                'Text9.Text = FeedPulsCount / FeedPulsPerMM
                
                'FindNextCutPointIndex(byVal CurIndex)
                'NextCutIndex = FindNextCutPointIndex(CurCutIndex)
                '铣刀点提前减速
                If PathOutputPoint(NextCutIndex).LengthFromStart - FeedPulsCount / FeedPulsPerMM < Device_FastSpeedMinLenMM Then
                    If CtrlCardType = 1 Then
                        FeedV_9030
                    ElseIf CtrlCardType = 4 Then
                        FeedV_GALIL
                        'SetVel hDmc, FeedAxis, Device_FeedStartV
                    End If
                    CurCutIndex = NextCutIndex
                End If
                If PathOutputPoint(NextCutIndex).LengthFromStart - FeedPulsCount / FeedPulsPerMM < 0 Then
                    NextCutIndex = FindNextCutPointIndex(CurCutIndex)
                    FeedV3_GALIL
                    'SetVel hDmc, FeedAxis, Device_FeedSpeed
                End If
                
                '99999前提前暂停
                If PathOutputPoint(Next99999Index).LengthFromStart - FeedPulsCount / FeedPulsPerMM < 10 _
                 And Next99999Index <> Cur99999Index Then
                    StopFeedV
                    Wait Device_DoneWaitingTime
                    Cur99999Index = Next99999Index
                    Next99999Index = FindNext99999PointIndex(Cur99999Index)
                    FeedV3_GALIL
                End If
                
'                '铣刀点提前减速
'                If PathOutputPoint(I).LengthFromStart - FeedPulsCount / FeedPulsPerMM < Device_FastSpeedMinLenMM _
'                    And PathOutputPoint(I).VertType > 0 Then
'                    If CtrlCardType = 1 Then
'                        FeedV_9030
'                    ElseIf CtrlCardType = 4 Then
'                        FeedV_GALIL
'                    End If
'                End If
                
                '铣刀点提前开启铣刀
'                If PathOutputPoint(I).LengthFromStart - FeedPulsCount / FeedPulsPerMM < Device_VertMotorZoneMM _
'                    And PathOutputPoint(I).VertType <> 0 Then
'                    WriteIoBit_9030 0, 1, VertMotorPort + 1
'                End If
                
                
                t = t + 1
                If t Mod 200 = 0 Then
                    ShowFeedMarkPoint
                    ShowVertMarkPoint
                End If
                
                'If fast_feed_puls > 0 Then
                '    If FeedPulsCount - FeedPulsCount0 >= fast_feed_puls Then
                '        fast_feed_puls = 0
                '        FeedV
                '    End If
                'End If
                
                If vert_zone_feed_puls > 0 Then
                    If FeedPulsCount - FeedPulsCount0 >= vert_zone_feed_puls Then
                        vert_zone_feed_puls = 0
                        FeedIntoVertMotorZone = True
                    End If
                End If
                
                'If job_done = False And FeedPulsCount >= PathOutputPoint(I).LengthFromStart * FeedPulsPerMM - Device_FeedOffset Then
                If FeedPulsCount >= PathOutputPoint(I).LengthFromStart * FeedPulsPerMM - Device_FeedOffset Then
                    'job_done = True
                    

                    ShowFeedMarkPoint
                    ShowVertMarkPoint
                    
                    '通过节点前后关系可以分类：铣内角点、铣外角点、线段两端点。
                    '有根据型材的不同（普通型材、美国型材、韩国型材）做出了归类
                    If PathOutputPoint(I).VertType = VT1 _
                        Or Device_AmericanMaterial = False And _
                            (((ChkStartPointVert90.value = 1 Or Abs(PathOutputPoint(I).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(I).Type = 3) Or _
                            ((ChkEndPointVert90.value = 1 Or Abs(PathOutputPoint(I).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(I).Type = 4)) _
                        Or Device_AmericanMaterial = True And _
                            (PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 5) Then  '内角 或 线段两端 或 美国型材两端
                            
                        ds = PathOutputPoint(I).LengthFromStart - lfs0
                        If (I = 1 And ds >= 0) Or (I > 1 And ds > 0.2) Then 'And ds >= Device_VertMinDistance) Then 'Device_VertMinDistance Then
                            
                            Section = Section + 1
                            t = 0
                            
                            If I > 1 Then
                                p = p + 1
                                'TxtStatistics.Text = TxtStatistics.Text + Format(p, "00") + " Len:" + str(Round(ds, 2)) + vbCrLf'分段长度显示
                                TxtStatistics.Text = TxtStatistics.Text + Format(p, "00") + " Len:" + str(Round(PathOutputPoint(I).LengthFromStart, 2)) + vbCrLf '总长显示
                            End If
                                                            
                            StopFeedV I
                            
                            CutOn
                            Wait 1
                            
                            If bUseCompenstion_g = True Then
                                If Device_UseEncoder = False Then
                                    FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
                                Else
                                    FeedPulsCount = ReadAxisEncodePos_9030(0, FeedAxis)
                                End If
                                 
                
                                getSetCutPos = PathOutputPoint(I).LengthFromStart '得到指定的切割位置
                                posErr = Round(FeedPulsCount / FeedPulsPerMM - PathOutputPoint(I).LengthFromStart, 3)
                                If posErr > ErrLmt Or posErr < 0 Then
                                    'TextBendLength.Text = Round(FeedPulsCount / FeedPulsPerMM - 0.4, 3)
                                    posErrCompensation posErr
                                End If
                            End If
                            
                            If Device_UseEncoder = False Then
                                If CtrlCardType = 1 Then
                                    FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
                                ElseIf CtrlCardType = 4 Then
                                    FeedPulsCount = GetPos(hDmc, FeedAxis)
                                End If
                            Else
                                If CtrlCardType = 1 Then
                                    FeedPulsCount = ReadAxisEncodePos_9030(0, FeedAxis)
                                ElseIf CtrlCardType = 4 Then
                                    FeedPulsCount = GetPosEnc(hDmc, FeedAxis)
                                End If
                            End If
                            VertCnt = VertCnt + 1
                            Print #fileNum_g, VertCnt; Tab(8); Round(FeedPulsCount / FeedPulsPerMM, 3); Tab(23); Round(PathOutputPoint(I).LengthFromStart, 3); Tab(38); Round(FeedPulsCount / FeedPulsPerMM - PathOutputPoint(I).LengthFromStart, 3)
                            'Wait 1
                            ShowFeedMarkPoint
                            ShowVertMarkPoint
                                                                
                            lfs0 = PathOutputPoint(I).LengthFromStart
                            blfs0 = lfs0
                                     
                            
                            
                            'If Not DebugWithoutSensor Then
                                'If Device_AmericanMaterial = True And (PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 5) Then '美国型材末端
                                If (PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 4 Or PathOutputPoint(I).Type = 5) Then '美国型材末端
                                    VertUpToMiddleWay = False
'                                    If PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 4 Then
'                                        PushOut Device_CutDepth, UseMagnetDO
'                                        IsCutoff = True
'                                    End If
                                    
                                    
                                    VertInnerAngle_done Device_TailVertAngle, False        '*****铣角运动******'
                                    'Sleep 2000
                                    If PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 4 Then
                                        If bOnlyCutoffInnerAngleSlotPoint = True Then
                                            If PathOutputPoint(I).AngleToNext > Device_VertMinAngle Then
                                                PushOut Device_CutDepth2, UseMagnetDO
                                            Else
                                                PushOut Device_CutDepth, UseMagnetDO
                                            End If
                                        Else
                                            PushOut Device_CutDepth, UseMagnetDO
                                        End If
                                        IsCutoff = True
                                        VertInnerAngle_done Device_TailVertAngle, False        '*****切断运动******'
                                        CutOff
                                    End If
                                    
                                    PullBack 0, UseMagnetDO
                                    IsCutoff = False
                                    
                                    If Device_DoneWaitingTime > 0 And PathOutputPoint(I).Type = 4 Then
                                        Wait Device_DoneWaitingTime
                                        Exit Do
                                    End If
                                    
                                
                                ElseIf Device_AmericanMaterial = False And _
                                    (((ChkStartPointVert90.value = 1 Or Abs(PathOutputPoint(I).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(I).Type = 3) Or _
                                    ((ChkEndPointVert90.value = 1 Or Abs(PathOutputPoint(I).AngleToNext) < Device_VertMinAngle) And PathOutputPoint(I).Type = 4)) Then
                                    
                                    'PathOutputPoint(I).AngleToNext = GetOutputEndPointAngle(I)
                            
                                    If PathOutputPoint(I).Type = 4 Then
                                        VertEndAngle_done 0, False                     '*********运动函数********'
                                        
                                    Else
                                        VertEndAngle_done 1, False                     '*********运动函数********'
                                        
                                    End If
                                Else
                                    If Device_AmericanMaterial = False Then
                                        'Debug.Print "<<<< inner i="; i, PathOutputPoint(i).AngleToNext
                                        
                                        VertInnerAngle_done PathOutputPoint(I).AngleToNext, False      '*********运动函数********'
                                    Else
                                        VertUpToMiddleWay = True
                                        VertInnerAngle_done 0, False                  '*********运动函数********'
                                        
                                        VertUpToMiddleWay = False
                                    End If
                                End If
                            'End If
                        Else
                            PullBack 0, UseMagnetDO
                        End If
                        CutOff
                        
                    ElseIf PathOutputPoint(I).VertType = VT2 Then  '外角
                        ds = PathOutputPoint(I).LengthFromStart - lfs0
                        If (I = 1 And ds >= 0) Or (I > 1 And ds >= Device_VertMinDistance And ds > 0.2) Then
                            
                            Section = Section + 1
                            t = 0
                            
                            If I > 1 Then
                                p = p + 1
                                'TxtStatistics.Text = TxtStatistics.Text + Format(p, "00") + " Len:" + str(Round(ds, 2)) + vbCrLf
                                TxtStatistics.Text = TxtStatistics.Text + Format(p, "00") + " Len:" + str(Round(PathOutputPoint(I).LengthFromStart, 2)) + vbCrLf '总长显示
                            End If
                            
                            StopFeedV I
                            
                            CutOn
                            Wait 1
                            
                            If bUseCompenstion_g = True Then
                                If Device_UseEncoder = False Then
                                    FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
                                Else
                                    FeedPulsCount = ReadAxisEncodePos_9030(0, FeedAxis)
                                End If
                                
                                getSetCutPos = PathOutputPoint(I).LengthFromStart '得到指定的切割位置
                                posErr = Round(FeedPulsCount / FeedPulsPerMM - PathOutputPoint(I).LengthFromStart, 3)
                                If posErr > ErrLmt Or posErr < 0 Then
                                    'TextBendLength.Text = Round(FeedPulsCount / FeedPulsPerMM - 0.4, 3)
                                    posErrCompensation posErr
                                End If
                            End If
                            
                            If Device_UseEncoder = False Then
                                If CtrlCardType = 1 Then
                                    FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
                                ElseIf CtrlCardType = 4 Then
                                    FeedPulsCount = GetPos(hDmc, FeedAxis)
                                End If
                            Else
                                If CtrlCardType = 1 Then
                                    FeedPulsCount = ReadAxisEncodePos_9030(0, FeedAxis)
                                ElseIf CtrlCardType = 4 Then
                                    FeedPulsCount = GetPosEnc(hDmc, FeedAxis)
                                End If
                            End If
                            VertCnt = VertCnt + 1
                            Print #fileNum_g, VertCnt; Tab(8); Round(FeedPulsCount / FeedPulsPerMM, 3); Tab(23); Round(PathOutputPoint(I).LengthFromStart, 3); Tab(38); Round(FeedPulsCount / FeedPulsPerMM - PathOutputPoint(I).LengthFromStart, 3)
                            'Wait 1
                            ShowFeedMarkPoint
                            ShowVertMarkPoint
                            
                            lfs0 = PathOutputPoint(I).LengthFromStart
                            blfs0 = lfs0
                        
                            
                            
                            'If Not DebugWithoutSensor Then
                            '铣外角处理， 美国型材和韩国型材只有铣内角处理，美国型材更只是上升一半位置
                                If Device_AmericanMaterial = False And Device_KareanMaterial = False Then
                                    'Debug.Print "<<<< outer i="; i, PathOutputPoint(i).AngleToNext
                                    VertOuterAngle_done PathOutputPoint(I).AngleToNext, False     '*********运动函数********'
                                    
                                    
                                ElseIf Device_KareanMaterial = True Then     '韩国型材
                                    '不是线段始末端
                                    If PathOutputPoint(I).Type <> 3 And PathOutputPoint(I).Type <> 4 Then
                                        VertUpToMiddleWay = True '0913原来函数没有这句
                                    End If
                                    
                                    VertUpToMiddleWay = True '限制切高
                                    IsCutoff = False
                                    'VertInnerAngle_done 0                       '*********运动函数********'
                                    tempVertMaxInnerAngle = Device_VertMaxInnerAngle    '暂存铣内角最大角度
                                    Device_VertMaxInnerAngle = Device_VertMaxOuterAngle '用铣外加最大角度赋值给铣内角最大角度
'                                    If PathOutputPoint(I).Type = 3 Or PathOutputPoint(I).Type = 4 Then
'                                        PushOut Device_CutDepth, UseMagnetDO
'                                        IsCutoff = True
'                                    End If
                                    
                                    VertInnerAngle_done PathOutputPoint(I).AngleToNext + 180 - 22.5, False '替换上一句，铣槽
                                   ' Sleep 2000
                                    If PathOutputPoint(I).Type = 4 Then
                                        If bOnlyCutoffInnerAngleSlotPoint = False Then
                                            PushOut Device_CutDepth, UseMagnetDO
                                        End If
                                        'IsCutoff = True  '切断高度是最大高度，不能采用
                                       VertInnerAngle_done PathOutputPoint(I).AngleToNext + 180 - 22.5, False '替换上一句，切断
                                       CutOff
                                    End If
                                    
                                    PullBack 0, UseMagnetDO
                                    IsCutoff = False
                                    Device_VertMaxInnerAngle = tempVertMaxInnerAngle    '回复铣内角最大角度
                                    VertUpToMiddleWay = False
                                    
                                    If Device_DoneWaitingTime > 0 And PathOutputPoint(I).Type = 4 Then
                                        Wait Device_DoneWaitingTime
                                        Exit Do
                                    End If
                                    
                                Else                           ' 美国型材
                                    VertUpToMiddleWay = True
                                    VertOuterAngle_done 0, False                      '*********运动函数********'
                                    
                                End If
                                CutOff
                            'End If
                       End If
                       
                    '切点之外如不大于拉直点，不做弯弧或拍弧动作, 角度复原
                    'ElseIf PathOutputPoint(I).LengthFromStart - getSetCutPos < LazhiDis Then
                    '    '
                    '    StopFeedV I
                    '    BendAngleByRadius 0, True '角度复原，不可少
                        '
                        
                    '以下为弯弧处理
                    ElseIf PathOutputPoint(I).VertType <= 0 And Abs(PathOutputPoint(I).Radius3P) > 0 _
                            And PathOutputPoint(I + 1).VertType = 0 And Abs(PathOutputPoint(I + 1).Radius3P) > 0 Then
                                                
                        '计算连续弯弧长度
                        'If PathOutputPoint(I - 1).Radius3P <> 0 Then
                        If PathOutputPoint(I + 1).Radius3P <> 0 _
                            Or (PathOutputPoint(I + 1).VertType = 2) _
                            Or (PathOutputPoint(I + 1).VertType = 1) Then
                                Bendlength = PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart + Bendlength
                            'End If
                        Else
                            FirstEndpointofLazhi = True
                            Bendlength = 0 '不连续则清零
                        End If
                            
                        If PathOutputPoint(I).VertType = -9 Then
                            StopFeedV
                            CmdStop_Click '切断后停止，结束进程
                        ElseIf Abs(PathOutputPoint(I).Radius3P) <= Device_BeatMaxRadius Then
                            StopFeedV I
                            'Print #fileNum_g, i; Tab(8); FeedPulsCount / FeedPulsPerMM; Tab(16)
                            'Wait 0.02
                            
                            BendDis = PathOutputPoint(I).LengthFromStart - blfs0
                            BeatRealAngle -VTDir * PathOutputPoint(I).AngleToNext, BendDis      '*********运动函数********'
                            blfs0 = PathOutputPoint(I).LengthFromStart
                            BeatAngelPointLength = PathOutputPoint(I).LengthFromStart
                            
                        Else
                           
                            k = 0
                            If Abs(PathOutputPoint(I).Radius3P) > 2000 Then
                                k = 1
                            ElseIf I > 1 And I < PathOutputPointCount Then
                                d0 = PathOutputPoint(I).LengthFromStart - PathOutputPoint(I - 1).LengthFromStart
                                d1 = PathOutputPoint(I + 1).LengthFromStart - PathOutputPoint(I).LengthFromStart
                                If d1 > 3 * d0 And d1 > 5 And Abs(PathOutputPoint(I + 1).AngleToNext) < 1 Then
                                    k = 1
                                End If
                            End If
                            
                            
                            
                            If k = 0 Then
                                'If I > 1 And Sgn(PathOutputPoint(I).Radius3P) <> Sgn(PathOutputPoint(I - 1).Radius3P) Then
                                '    StopFeedV
                                '    BendAngleByRadius -VTDir * PathOutputPoint(I).Radius3P, True
                                '    'FeedV
                                'Else
                                '20131105 测试去掉StopFeedV
                                If xiaohulianxu = True Then
                                    'StartAxisVel_9030 0, FeedAxis, Device_FeedStartV
                                    'If Abs(OldBendangle_g - PathOutputPoint(i - 1).AngleToNext) > IncBendangleLmt _
                                    'Or (PathOutputPoint(i).LengthFromStart - PathOutputPoint(i - 1).LengthFromStart < Device_TurnPointOffsetMM) Then
                                    If Abs(OldBendangle_g - PathOutputPoint(I - 1).AngleToNext) > IncBendangleLmt _
                                    Or (PathOutputPoint(I).LengthFromStart - BeatAngelPointLength) < Device_TurnPointOffsetMM _
                                    Or Abs(PathOutputPoint(I + 2).Radius3P) <= Device_BeatMaxRadius _
                                    Or Abs(PathOutputPoint(I + 1).Radius3P) <= Device_BeatMaxRadius _
                                    Or Abs(PathOutputPoint(I).Radius3P) <= Device_BeatMaxRadius Or FirstEndpointofLazhi = True Then
                                        FirstEndpointofLazhi = False
                                        StopFeedV                              '#20150121
                                        '如果停下则更新前次弯弧角，否则不更新
                                        OldBendangle_g = PathOutputPoint(I - 1).AngleToNext
                                        
                                        'If PathOutputPoint(I).LengthFromStart - getSetCutPos > LazhiDis Then
                                        If Bendlength > LazhiDis Then
                                            BendAngleByRadius -VTDir * PathOutputPoint(I).Radius3P, True    '*********弯弧函数********'
                                        End If
                                    Else
                                        'checkdone = false,不等待复位完成
                                        'If PathOutputPoint(I).LengthFromStart - getSetCutPos > LazhiDis Then
                                        If Bendlength > LazhiDis Then
                                            BendAngleByRadius -VTDir * PathOutputPoint(I).Radius3P, False   '*******弯弧指令*******'
                                        End If
                                    End If
                                Else
                                    StopFeedV
                                    '如果停下则更新前次弯弧角，否则不更新
                                    OldBendangle_g = PathOutputPoint(I - 1).AngleToNext
                                    'If PathOutputPoint(I).LengthFromStart - getSetCutPos > LazhiDis Then
                                    If Bendlength > LazhiDis Then
                                        BendAngleByRadius -VTDir * PathOutputPoint(I).Radius3P, True    '*********弯弧函数********'
                                    End If
                                End If
                                '如果停下则更新前次弯弧角，否则不更新
                                'OldBendangle_g = PathOutputPoint(I - 1).AngleToNext
                                    
                                    
                                    
                                'End If
                                
                                '检测是否因拍弧速度过慢，导致进料超过分段节点，产生节点判断误差-------------------------------
                                If CtrlCardType = 0 Then
                                    Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
                                ElseIf CtrlCardType = 4 Then
                                    nLogPos = GetPos(hDmc, FeedAxis)
                                    nActPos = GetPosEnc(hDmc, FeedAxis)
                                Else
                                    nLogPos = ReadAxisPos_9030(0, FeedAxis)
                                    nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                                    nSpeed = ReadAxisVel_9030(0, FeedAxis)
                                End If
                                
                                If Device_UseEncoder = False Then
                                    FeedPulsCount = nLogPos
                                Else
                                    FeedPulsCount = nActPos
                                End If
                    
                                If FeedPulsCount > PathOutputPoint(I + 1).LengthFromStart * FeedPulsPerMM - Device_FeedOffset Then
                                    LblFeedStatus.BackColor = RGB(0, 0, 255)
                                Else
                                    LblFeedStatus.BackColor = RGB(220, 220, 220)
                                End If
                                '---------------------------------------------------------------------------------------------
                            Else
                                StopFeedV
                                BendAngleByRadius 0, True '角度复原，不可少                 '*******运动指令*******'
                                
                            End If

                        End If
                        
                        ShowFeedMarkPoint
                        ShowVertMarkPoint
                        
                    ElseIf PathOutputPoint(I).VertType <= 0 And Abs(PathOutputPoint(I).Radius3P) = 0 Then
                        If PathOutputPoint(I).VertType = -9 Then
                            StopFeedV
                            'CmdStop_Click
                        ElseIf PathOutputPoint(I).VertType = 0 Then
                            
                            If xiaohulianxu = True Then
                                'StartAxisVel_9030 0, FeedAxis, Device_FeedStartV
                                'If Abs(OldBendangle_g - PathOutputPoint(i).AngleToNext) > IncBendangleLmt _
                                'Or (PathOutputPoint(i).LengthFromStart - PathOutputPoint(i - 1).LengthFromStart < Device_TurnPointOffsetMM) Then
                                If Abs(OldBendangle_g - PathOutputPoint(I - 1).AngleToNext) > IncBendangleLmt _
                                Or (PathOutputPoint(I).LengthFromStart - BeatAngelPointLength) < Device_TurnPointOffsetMM _
                                Or Abs(PathOutputPoint(I + 2).Radius3P) <= Device_BeatMaxRadius _
                                Or Abs(PathOutputPoint(I + 1).Radius3P) <= Device_BeatMaxRadius _
                                Or Abs(PathOutputPoint(I).Radius3P) <= Device_BeatMaxRadius Then
                                    StopFeedV                  '#20150121
                                    BendAngleByRadius 0, True
                                Else
                                    'checkdone = false,不等待复位完成
                                    BendAngleByRadius 0, False '角度复原，不可少                       '*******弯弧指令*******'
                                End If
                            Else
                                StopFeedV
                                BendAngleByRadius 0, True
                            End If
                            OldBendangle_g = PathOutputPoint(I - 1).AngleToNext
                            
                            'BendAngleByRadius 0, True '角度复原，不可少                       '*******运动指令*******'
                            
                            'checkdone = false,不等待复位完成
                            'BendAngleByRadius 0, False '角度复原，不可少                       '*******弯弧指令*******'
                            
                            '检测是否因拍弧速度过慢，导致进料超过分段节点，产生节点判断误差-------------------------------
                            'Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
                            If CtrlCardType = 0 Then
                                Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
                            ElseIf CtrlCardType = 4 Then
                                    nLogPos = GetPos(hDmc, FeedAxis)
                                    nActPos = GetPosEnc(hDmc, FeedAxis)
                            Else
                                nLogPos = ReadAxisPos_9030(0, FeedAxis)
                                nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                            End If
                            
                            If Device_UseEncoder = False Then
                                FeedPulsCount = nLogPos
                            Else
                                FeedPulsCount = nActPos
                            End If
                            
                            If FeedPulsCount > PathOutputPoint(I + 1).LengthFromStart * FeedPulsPerMM - Device_FeedOffset Then
                                LblFeedStatus.BackColor = RGB(0, 0, 255)
                            Else
                                LblFeedStatus.BackColor = RGB(220, 220, 220)
                            End If
                            '---------------------------------------------------------------------------------------------
                             
                        ElseIf PathOutputPoint(I).VertType = -1 And Abs(PathOutputPoint(I).AngleToNext) > 0 Then '折角
                            If Device_VertNoTurn = False And PathOutputPoint(I).Type <> 99999 Then '设定不折角为否
                                
                                StopFeedV
                                
                                ShowFeedMarkPoint
                                ShowVertMarkPoint
                                    
                                If CtrlCardType = 1 Then
                                    TurnRealAngle_9030 -VTDir * PathOutputPoint(I).AngleToNext   '*******运动指令*******'
                                ElseIf CtrlCardType = 4 Then
                                    TurnRealAngle_GALIL -VTDir * PathOutputPoint(I).AngleToNext
                                End If
                            Else
                                StopFeedV
                                BendAngleByRadius 0, True '角度复原，不可少                 '*******运动指令*******'
                                
                             End If
                        End If
                        
                        ShowFeedMarkPoint
                        ShowVertMarkPoint
                    End If
                    
                    
                    
                    If PathOutputPoint(I).Type = 88888 Then
                        If next_piece = True Then
                            StopFeedV
                            Wait 1
                            ShowFeedMarkPoint
                            ShowVertMarkPoint
                            
                            BendAngle 0
                                
                            Do
                                If CtrlCardType = 1 Then
                                    Status = ReadAxisState_9030(0, BendAxis)
                                ElseIf CtrlCardType = 4 Then
                                    Status = GetStatus(hDmc, BendAxis)
                                End If
                                If Status <> 1 Then
                                    Exit Do
                                End If
                                
                                DoEvents
                            Loop
                            
                            If Device_DoneWaitingTime > 0 Then
                                Wait Device_DoneWaitingTime
                            Else
                                FrmMsgDlg.LblMessage = "请按任意键或点击 [关闭] 按钮，开始加工下一部件......"
                                FrmMsgDlg.Show
                                Do While FrmMsgDlg.Visible = True
                                    DoEvents
                                Loop
                            End If
                            'PicPathCls
                            'DrawAll
                            
                            m = m + 1
                            TxtRunCount.Text = str(1 + m)

                        End If
                    ElseIf PathOutputPoint(I).Type = 99999 Then
                        next_piece = True
'                    ElseIf PathOutputPoint(I).Type = 4 Then    '切断就停止
'                        If PathOutputPoint(I).VertType = -9 Then
'                            StopFeedV
'                            'CmdStop_Click
'                        End If
'                        next_piece = True
                        
                    End If
                    
                    
                        
                    Exit Do '???
                End If
                
                
                'If FeedPulsCount >= PathOutputPoint(I + 1).LengthFromStart * FeedPulsPerMM - Device_FeedOffset Then
                '    Exit Do
                'End If
                DoEvents
            Loop
        Next
        
        BendAngleByRadius 0, True '角度复原
        '--------------------------
'        SetAxisPos_9030 0, BendAxis, 0
'        StartAxis_9030 0, BendAxis
'        Do
'
'            status = ReadAxisState_9030(0, BendAxis)
'            If status <> 1 Then
'                Exit Do
'            End If
'
'            DoEvents
'        Loop
        
        '--------------------------
        
        StopFeedV
        'sudden_stop 0, FeedAxis
        
        If StopRunning = False Then
            BendAngle 0
            VertAngle 0
            
            If FeedByDCMotor = False Then
                FeedMM Device_DoneDistance + Device_TotalAddDoneDistance / val(TxtRunN.Text), Device_UseEncoder, 0, False
            Else
                FeedMMByDCMotor Device_DoneDistance, 0, False
            End If
            
            'TxtStatistics.Text = TxtStatistics.Text + vbCrLf + " Total:" + Str(Round(TotalPathOutLength, 2)) + vbCrLf + vbCrLf
        Else
'                Exit For
        End If
        
        TotalWorkLength = TotalWorkLength + TotalPathOutLength
        TotalWorkCount = TotalWorkCount + 1
        'TotalWorkTime = TotalWorkTime + TimeDiff(Timer, t0)
        
        'WriteToINI "TotalWorkLength", Str(TotalWorkLength)
        'WriteToINI "TotalWorkBendCount", Str(TotalWorkBendCount)
        'WriteToINI "TotalWorkCount", Str(TotalWorkCount)
        'WriteToINI "TotalWorkTime", Str(TotalWorkTime)
        
        ShowFeedMarkPoint
        ShowVertMarkPoint
            
        PicPathCls
        DrawAll
             
        StopFeed = False
        'End If
    End If
    
    'BendReset_9030_V8
    'BendReset_9030
   
                        
    PicPath.DrawWidth = 1

    For Each obj In FrmMain
        If Not (TypeOf obj Is Timer) Then
            obj.Enabled = True
        End If
    Next
    
    CmdStop_Click
    
    '------------关闭记录铣刀点位置的文件-----------------------
    closefileofprintvertpoint
    
    IsRunning = False
End Sub


Public Sub CmdStop_Click()
'***************************************************************************************************************************
'                               停止命令处理函数
'***************************************************************************************************************************
    Dim obj As Object
    Dim Status As Integer
    
    On Error Resume Next
    
    StopRunning = True
    
    PauseRunning = False
    
    StopFeedV
    
    
    If CtrlCardType = 0 Then
        
        'sudden_stop 0, FeedAxis
        
        sudden_stop 0, BendAxis
        sudden_stop 0, VertAxis
        sudden_stop 0, VertUpDownAxis
        
        PortBit(1) = 0
        PortBit(2) = 0
        PortBit(3) = 0
        PortBit(4) = 0
        
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertClosePort, 0
        write_bit 0, VertMoveUpPort, 0
        write_bit 0, VertMoveDownPort, 0
    ElseIf CtrlCardType = 4 Then
        StopAxis hDmc, FeedAxis
        StopAxis hDmc, BendAxis
        StopAxis hDmc, VertAxis
        StopAxis hDmc, VertUpDownAxis
    Else
        '关闭锯片电机
        WriteIoBit_9030 0, 0, VertMotorPort + 1
        CeaseAxis_9030 0, FeedAxis
        CeaseAxis_9030 0, BendAxis
        CeaseAxis_9030 0, VertAxis
        CeaseAxis_9030 0, VertUpDownAxis
    End If
    'TmrFeed.Enabled = False
    TmrBend.Enabled = False
    
    If CtrlCardType = 1 Then
        If ReadAxisPos_9030(0, BendAxis) <> 0 Then
            SetAxisPos_9030 0, BendAxis, 0
            StartAxis_9030 0, BendAxis
            Do
                Sleep (1)
                Status = ReadAxisState_9030(0, BendAxis)
                If Status = 0 Then     '等待回零完成
                    Exit Do
                End If
                
                
                DoEvents
            Loop
        End If
    End If
    
    
    
    IsRunning = False
    PanButtonRun.Enabled = True
    CmdRun.Enabled = True
    CmdManualControl.Enabled = True
    CmdFeedBkV2A.Enabled = True
    CmdFeedFWV2A.Enabled = True
    PanButton2.Enabled = True
    PanButton3.Enabled = True
    PanButton11.Enabled = True
    PanButton1.Enabled = True
    PanButton5.Enabled = True
    CmdResetOrg.Enabled = True
    
    StopRunning = True
End Sub

Private Sub CmdSoftwareTest_Click()
'***************************************************************************************************************************
'
'***************************************************************************************************************************
    Dim obj As Object
    Dim I As Long, i0 As Long, lfs0 As Double, ds As Double, p As Long
    Dim j As Long, start_id As Long
        
    On Error Resume Next
    
    IsRunning = True
    
    For Each obj In FrmMain
        obj.Enabled = False
    Next
    CmdStop.Enabled = True
    
    'On Error GoTo 0

    For j = 1 To OutputStartPointList.count
        start_id = OutputStartPointList.point_id(j)

        CalculatePath start_id
        
        'Debug.Print "MaxAngle="; Abs(MaxPathOutAngle)
        'Debug.Print "MinAngle="; Abs(MinPathOutAngle)
    
        FraEdit.Visible = False
        TxtStatistics.Visible = True
        TxtStatistics.Text = ""
        
        lfs0 = 0
        p = 0
        
        If TotalPathOutLength > 0 Then
            'If Abs(MaxPathOutAngle) > 90 Then
            '    MsgBox "图形中的最大角度超过 90 度。请调整图形后重新运行"
            'Else
                TmrDevicePortChecking.Enabled = True
                  
                PicPathCls
                DrawAll
                
                StopRunning = False
                i0 = 1
                For I = 1 To PathOutputPointCount
                    If StopRunning = True Then
                        Exit For
                    End If
                    
                    Do '跳过Vert点
                        If i0 > PathOutputPointCount Or PathOutputPoint(i0).VertType <= 0 Then
                            Exit Do
                        Else
                        End If
                        i0 = i0 + 1
                    Loop
                    
                    I = i0 + 1
                    Do '跳过Vert点
                        If I > PathOutputPointCount Or PathOutputPoint(I).VertType <= 0 Then
                            If PathOutputPoint(I).VertType = -1 Then
                                ds = PathOutputPoint(I).LengthFromStart - lfs0
                                If ds > Device_VertMinDistance Then
                                    p = p + 1
                                    'TxtStatistics.Text = TxtStatistics.Text + Format(p, "00") + " Len:" + Str(Round(ds, 2)) + vbCrLf
                                    lfs0 = PathOutputPoint(I).LengthFromStart
                                End If
                            End If
                            Exit Do
                        Else
                        End If
                        I = I + 1
                    Loop
                                
                    If i0 <= PathOutputPointCount And I <= PathOutputPointCount Then
                        ScreenLine PathOutputPoint(i0).ux, PathOutputPoint(i0).uy, 0, PathOutputPoint(I).ux, PathOutputPoint(I).uy, 0
                    End If
                    i0 = I
                Next
                
                PicPathCls
                DrawAll
                
                'TxtStatistics.Text = TxtStatistics.Text + vbCrLf + " Total:" + Str(Round(TotalPathOutLength, 2)) + vbCrLf
                
                
            'End If
        End If
        PicPath.DrawWidth = 1
    
    Next
    
    For Each obj In FrmMain
        If Not (TypeOf obj Is Timer) Then
            obj.Enabled = True
        End If
    Next
    
    CmdStop_Click
    
    IsRunning = False
End Sub

Private Sub CmdStopA_Click()
    CmdStop_Click
End Sub

Private Sub CmdTest_Click()

    'VertThreadStep = 103
   'VertEndAngle_done 1
    'VertInnerAngle_prev 30
    'VertOuterAngle_prev 30
    If CtrlCardType = 4 Then
        DefineEnc hDmc, FeedAxis, 0
        DefinePos hDmc, FeedAxis, 0
        DefinePos hDmc, BendAxis, 0
        DefinePos hDmc, VertAxis, 0
        'DefinePos hDmc, VertUpDownAxis, 0
    Else
        HomeFB_9030 0, FeedAxis
        Home_9030 0, FeedAxis
        Home_9030 0, BendAxis
        Home_9030 0, VertAxis
    End If
End Sub

Private Sub CmdToolBox_Click()
'    If FraToolBox.Visible = False Then
'        ShowToolBox VScroll1.left + 22, HT + 12
'        FraToolBox.Visible = True
'        FraEdit.Visible = False
'        TxtStatistics.Height = PicFrame.Height - FraToolBox.Height - 20
'    Else
        FraToolBox.Visible = False
        FraEdit.Visible = False
        TxtStatistics.Height = PicFrame.Height - 15
'    End If
    TxtStatistics.Move CmdToolBox.left, FrmMain.ScaleHeight - TxtStatistics.Height - 3, CmdToolBox.Width - 3
    TxtStatistics.Visible = True
TxtStatistics.Text = ""
End Sub

Private Sub CmdTurnL_Click()
    TurnLeft
End Sub

Private Sub CmdTurnR_Click()
    TurnRight
End Sub


Private Sub CmdVertDown_Click()
    StopRunning = False
    
    VertMoveDown
'    Dim t As Double, t0 As Double, b As Long
'
'    write_bit 0, VertMoveDownPort, 1
'    Wait 0.1
'    write_bit 0, VertMoveUpPort, 0      '铣刀向上运动
'
'    TmrDevicePortChecking.Enabled = True
'    t0 = Timer
'    Do
'        b = read_bit(0, 17)
'        If b = 0 Then
'            LblVertLowSensor.BackColor = RGB(255, 0, 0)
'            write_bit 0, VertMoveDownPort, 0
'            Exit Do
'        End If
'
'        t = Timer
'        If TimeDiff(t, t0) > 5 Then
'            Exit Do
'        End If
'
'        DoEvents
'    Loop
End Sub

Private Sub CmdVertInnerLine_Click()
    StopRunning = False
    IsRunning = True
    
    VertInnerAngle 0, False
End Sub

Private Sub CmdVertInnerLineA_Click()
    CmdVertInnerLine_Click
End Sub

Private Sub CmdVertLV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long

    IsRunning = True
    StopRunning = False
    If CtrlCardType = 0 Then
        Ret = set_startv(0, VertAxis, Device_VertStartV)
        Ret = set_speed(0, VertAxis, Device_VertSpeed)
        Ret = set_acc(0, VertAxis, Device_VertAccel)
        Ret = continue_move1(0, VertAxis, 1)
    Else
        Ret = SetAxisStartVel_9030(0, VertAxis, Device_VertStartV)
        'Ret = SetAxisVel_9030(0, VertAxis, Device_VertSpeed)
        Ret = SetAxisAcc_9030(0, VertAxis, Device_VertAccel)
        Ret = SetAxisDec_9030(0, VertAxis, Device_VertAccel)
        Ret = StartAxisVel_9030(0, VertAxis, -1 * Device_VertSpeed)
    End If
End Sub

Private Sub CmdVertLV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If CtrlCardType = 0 Then
        Ret = sudden_stop(0, VertAxis)
    Else
        Ret = CeaseAxis_9030(0, VertAxis)
    End If

    IsRunning = False
End Sub

Private Sub CmdVertOutLine_Click()
    
    StopRunning = False
    IsRunning = True
    
    If Device_KareanMaterial = True Then
        IsCutoff = True
        VertInnerAngle -Device_CutDepth, IsCutoff
        IsCutoff = False
    Else
        VertOuterAngle val(TxtVertDeg.Text)
    End If
End Sub

Private Sub CmdVertReset_Click()
    StopRunning = False
    IsRunning = True
    
    If CtrlCardType = 0 Then
        VertReset
    ElseIf CtrlCardType = 1 Then
        VertReset_9030
    Else
        VertReset_GALIL_V8
    End If
    IsRunning = False
End Sub

Private Sub CmdVertRV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long

    IsRunning = True
    StopRunning = False
    If CtrlCardType = 0 Then
        Ret = set_startv(0, VertAxis, Device_VertStartV)
        Ret = set_speed(0, VertAxis, Device_VertSpeed)
        Ret = set_acc(0, VertAxis, Device_VertAccel)
        Ret = continue_move1(0, VertAxis, 0)
    Else
        Ret = SetAxisStartVel_9030(0, VertAxis, Device_VertStartV)
        'Ret = SetAxisVel_9030(0, VertAxis, Device_VertSpeed)
        Ret = SetAxisAcc_9030(0, VertAxis, Device_VertAccel)
        Ret = SetAxisDec_9030(0, VertAxis, Device_VertAccel)
        Ret = StartAxisVel_9030(0, VertAxis, Device_VertSpeed)
    End If
End Sub

Private Sub CmdVertRV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If CtrlCardType = 0 Then
        Ret = sudden_stop(0, VertAxis)
    Else
        Ret = CeaseAxis_9030(0, VertAxis)
    End If

    IsRunning = False
End Sub

Private Sub CmdVertUp_Click()
    StopRunning = False
    If VersionNmb = 81 Then
        VertUpDownReset_V81
    Else
        VertMoveUp
    End If
    
   ' Dim t As Double, t0 As Double, b As Long
   '
   ' If VertUpDownByDCMotor = True Then
   '     write_bit 0, VertMoveUpPort, 1       '铣刀向上运动
   '     Wait 0.1
   '     write_bit 0, VertMoveDownPort, 0
   '
   '     TmrDevicePortChecking.Enabled = True
   '     t0 = Timer
   '     Do
   '         b = read_bit(0, 23)
   '         If b = 0 Then
   '             LblVertHighSensor.BackColor = RGB(255, 0, 0)
   '             write_bit 0, VertMoveUpPort, 0
   '              Exit Do
   '         End If
   '
   '         t = Timer
   '         If TimeDiff(t, t0) > 5 Then
   '             Exit Do
   '         End If
   '     Loop
   '
   ' Else
   '
   '
   ' End If
End Sub
Function searchMinVertDis() As Boolean
Dim rtn As Boolean
Dim vertpoint0 As Double
Dim vertpoint1 As Double
Dim I As Integer
    rtn = False
    vertpoint0 = 0
    vertpoint1 = 0
    For I = 2 To PathOutputPointCount
        If (PathOutputPoint(I).VertType = 1 Or PathOutputPoint(I).VertType = 2) _
        And PathOutputPoint(I).Type <> 88888 _
        And PathOutputPoint(I).Type <> 3 Then
            vertpoint1 = PathOutputPoint(I).LengthFromStart
            If vertpoint1 - vertpoint0 < Device_VertMinDistance Then
                rtn = True
                Exit For
            End If
            vertpoint0 = vertpoint1
        End If
    Next
        
    searchMinVertDis = rtn
    
End Function

Private Sub Command1_Click()
Dim I As Integer
Dim pulse As Double
Dim FeedPulsCount As Long
Dim FeedPulsPerMM As Double
Dim rtnbox As Integer
Dim str As String
Dim str1 As String
Dim str2 As String


FeedPulsCount = 1000
FeedPulsPerMM = 0.21
I = 1
pulse = 555
    MnuImportAI.caption = "Import AI file"
    FormSettings.ChkUseEncoder.caption = "Using Encoder"

'    If searchMinVertDis = True Then
'        rtnbox = MsgBox("图形中有线段小于<最短铣角间距>,是否继续加工？", vbYesNo Or vbQuestion, "系统提示")
'        If rtnbox = vbNo Then
'            Exit Sub
'        End If
'    End If
'
'    createfile2printvertpoint
'    'str = Format(i, "") + "          " + Format(FeedPulsCount / FeedPulsPerMM, "#0.000") + "          " + Format(FeedPulsCount / FeedPulsPerMM, "#0.000")
'    str = Format(I, "")
'    str1 = Format(FeedPulsCount / FeedPulsPerMM, "#.000")
'    str2 = Format(FeedPulsCount / FeedPulsPerMM, "#.000")
'    'Print #fileNum_g, i; Tab(8); FeedPulsCount / FeedPulsPerMM; FeedPulsCount / FeedPulsPerMM
'    Print #fileNum_g, str; Tab(8); str1; Tab(24); str2
'    I = I + 1
'    FeedPulsCount = FeedPulsCount + 320
'    Print #fileNum_g, I; Tab(8); FeedPulsCount / FeedPulsPerMM; FeedPulsCount / FeedPulsPerMM
'    closefileofprintvertpoint
End Sub

Private Sub Form_Initialize()
    
    '若不存在则创建打印调试文件目录
    If dir("c:\hd_debug", vbDirectory) <> "" Then
        'MsgBox "存在"
    Else
        'MsgBox "不存在"
        MkDir "c:\hd_debug"
    End If
    
    If curLanguage = 0 Then
        AppVersion = "HDWZ_CN V80 " & AppVersionData
    Else
        AppVersion = "HDWZ_EN V80 " & AppVersionData
    End If
    
    '===================================================
    If curLanguage = 0 Then
        ChangeFace curLanguage + 1
        ChangeFontByLanguage curLanguage
        languageType = curLanguage
        SelLanguage curLanguage
        If Frame1.Visible = False Then
            If languageType = 0 Then
                CmdManualControl.caption = "显示手动控制面板"
            Else
                CmdManualControl.caption = "Show Manual Panel"
            End If
    
        Else
            If languageType = 0 Then
                CmdManualControl.caption = "关闭手动控制面板"
            Else
                CmdManualControl.caption = "Hide Manual Panel"
            End If
    
    
        End If
    End If
    
    '====================================================
    InitCommonControls
    
    
    
End Sub

Public Sub Form_Load()
    Dim SInfo As SYSTEM_INFO
    
    Dim cnt As Long
    
    
    
If CtrlCardType = 1 Then

    PrePauseSwitchVal = 0
    CurPauseSwitchVal = 0
    'Get the system information
    GetSystemInfo SInfo
    If SInfo.dwNumberOfProcessors > 1 Then
        SetProcessAffinityMask GetCurrentProcess(), 1
    Else
        'MsgBox "CPU是单核CPU"
    End If
ElseIf CtrlCardType = 4 Then
    PrePauseSwitchVal = 1
    CurPauseSwitchVal = 1
End If

    If bUseSplash = True Then
        cnt = 0
        'frmSplash.lblProductName.caption = "HANDUN BENDER"
        frmSplash.Show
        'Timer1.Enabled = True
        Do
            cnt = cnt + 1
            If cnt > 2000000 Then
                Exit Do
            End If
            DoEvents
        Loop
        frmSplash.Hide
    End If

    KeyPreview = True
    
    If App.PrevInstance = True Then
        Unload Me
        End
    End If
    
    NormalColor = RGB(255, 0, 0)
    HighLightColor = RGB(255, 127, 127)
    XORColor = RGB(255, 0, 0)
    
    If IsDemoVersion = True Then
        VersionMark = "(演示版)"
    Else
        VersionMark = ""
    End If
    Me.caption = AppVersion & VersionMark
    
    Device_GetBasicParam
    InitParameter
    FormResize
       
    Me.Toolbar1.Buttons(16).value = tbrPressed
    Me.Toolbar1.Buttons(17).value = tbrUnpressed
    Me.Toolbar1.Buttons(18).value = tbrUnpressed
    ShowDirection = False
    ShowPoints = False
    
    'WritePrivateProfileString "Screen", "Ratio", "1", App.Path & "\Parameters.ini"
    PicPathRatio = GetValueFromINI("Screen", "Ratio", "1", App.Path & "\Parameters.ini")
    
    CurLayer = 1
    DirectionChanged = True
    
    DemoStep = UserMaxX / 200
    
    Me.Show
    
    'PicFace.Move (Me.ScaleWidth - PicFace.Width) / 2, (Me.ScaleHeight - PicFace.Height) / 2
    'PicFace.Visible = True
    'TmrFace.Enabled = True
    
    GetDeviceParameters
    
    FirstRun = True
    ChkUseRemainder.value = 0
    
    If curLanguage = 0 Then
        ChangeFace 2
        SelLanguage 1
        ChangeFontByLanguage curLanguage
    End If
    
    ' 初始化板卡
    Init_Board
    
    'If IsDemoVersion = False Then
    '    CheckRegistration
    'End If
    
    MakeRecentFileMenu ""
    
    AddScrollness PicPath.hWnd
    
'    If ResetDone = False Then
'        Wait 1
'        MsgBox "该程序每次启动后请先执行[复位]功能。 ", vbInformation + vbOKOnly + vbSystemModal, "注意"
'    End If

    If MaterialDirUpset = False Then
        VT1 = 1
        VT2 = 2
        VTDir = 1
    Else
        VT1 = 2
        VT2 = 1
        VTDir = -1
    End If
    
    'If Device_UseEncoder = True Then
    '    LblEncoderOS.Visible = True
    '    LblEncoderOffset.Visible = True
    'End If
    
    PathSmooth = True
    If curLanguage = 0 Then
        ChangeFace 1
        SelLanguage 0
        ChangeFontByLanguage curLanguage
    End If
    
    FeedPulsPerMM = Device_PulsPerMM
End Sub

Private Sub Form_Resize()
    FraToolBox.Visible = True
    FormResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim p As Long
    
    If IsRunning = True Then
        MsgBox "设备正在运行。请设备停止运行后再结束程序。", vbExclamation + vbOKOnly
        Cancel = 1
        Exit Sub
    End If
    
    
    Timer1.Enabled = False
    CutOff
    DMCClose hDmc
    
    'If DataChanged = True Then
    '    If MsgBox("图形数据已改变。退出前是否需要保存改变后的数据？", vbQuestion + vbYesNo, "") = vbYes Then
    '        If CurFileName <> "" Then
    '            p = InStr(UCase(CurFileName), ".DXF")
    '            If p > 0 Then
    '                CurFileName = Mid(CurFileName, 1, p - 1) & ".ITN"
    '            End If
    '            WriteFile CurFileName
    '        Else
    '            MnuSave_Click
    '        End If
    '    End If
    'End If
    
    DeleteAllUndoFiles
    RemoveScrollness PicPath.hWnd
    End
End Sub

Public Sub CmdTool_Click(Index As Integer)
    Dim tv() As Title_Value
    
    CurTool = ToolTask(Index)
    CurToolStep = 0
    
    CurPoint.id = 0
    LastPoint.id = 0
    
    PicPath.SetFocus
    
    If CurTool = ToolType.ZoomIn Or _
       CurTool = ToolType.ZoomOut Or _
       CurTool = ToolType.MoveCanvas Then
        CanShowCursorReferenceLines = False
    Else
        CanShowCursorReferenceLines = True
    End If
    
    TxtCurTool.Text = CmdTool(Index).caption
    MaxBodyID = GetBodyList(BodyList)
    
    FraEdit.Visible = False
    
    If CurTool = ToolType.Unit Then
        CurUnittedGroupID = 0
        
    ElseIf CurTool = ToolType.PieceArray Then
        ReDim tv(5) As Title_Value
        tv(0).t = "X 范围"
        tv(0).v = ViewMaxX
        tv(1).t = "Y 范围"
        tv(1).v = ViewMaxY
        tv(2).t = "X 间距"
        tv(2).v = 0
        tv(3).t = "Y 间距"
        tv(3).v = 0
        tv(4).t = "转角"
        tv(4).v = 0
        
        ChkEdit(0).value = 1
        
        ShowEditData "阵列参数", 5, tv, ToolType.PieceArray
        CurArrayedGroupID = 0
        CmdEdit.Enabled = False
        
    ElseIf CurTool = ToolType.ConvertToSegments Then
        ConvertAllToSegments
        
        PicPathCls
        DrawAll
        
        SaveUndo
    End If
End Sub

Private Sub HScroll1_Change()
    If PicPath.left = -HScroll1.value Then Exit Sub
    PicPath.left = -HScroll1.value
End Sub

Private Sub HScroll1_Scroll()
    If PicPath.left = -HScroll1.value Then Exit Sub
    PicPath.left = -HScroll1.value
End Sub

Private Sub LblEdit_Click(Index As Integer)
    FrmDigiPad.left = (FraEdit.left - 100) * Screen.TwipsPerPixelX
    FrmDigiPad.top = Min((FraEdit.top + 70) * Screen.TwipsPerPixelY + LblEdit(Index).top, Screen.Height - FrmDigiPad.Height - 30 * Screen.TwipsPerPixelY)
    
    FrmDigiPad.Tag = str(Index)
    FrmDigiPad.LblEdit.caption = LblEdit(Index).caption
    FrmDigiPad.TxtEdit.Text = Trim(TxtEdit(Index).Text)
    
    FrmDigiPad.Show
    FraEdit.Tag = "1"
End Sub

Private Sub MnuAbout_Click()
    'PicFace.Visible = True
    If bUseSplash = True Then
        frmSplash.Visible = True
    End If
End Sub

Private Sub MnuBarAddPoint_Click()
    CurTool = ToolType.BreakSegment
    
    If curLanguage = 0 Then
        TxtCurTool.Text = "加点"
    Else
        TxtCurTool.Text = "Add Element Pt."
    End If
End Sub

Private Sub MnuBarDelPoint_Click()
    CurTool = ToolType.DeleteElement_Point
    If curLanguage = 0 Then
        TxtCurTool.Text = "删点"
    Else
        TxtCurTool.Text = "Del Element Pt."
    End If
End Sub

Private Sub MnuBarMeasureAera_Click()
    CurTool = ToolType.MeasureScale
    
    If curLanguage = 0 Then
        TxtCurTool.Text = "测面积"
    Else
        TxtCurTool.Text = "Measure Scale"
    End If
End Sub

Private Sub MnuBarMeasureDis_Click()
    CurTool = ToolType.MeasureDistance
    If curLanguage = 0 Then
        TxtCurTool.Text = "测距离"
    Else
        TxtCurTool.Text = "Measure Distance"
    End If
End Sub

Private Sub MnuBarMovePoint_Click()
    CurTool = ToolType.MoveElement_Point
    
    If curLanguage = 0 Then
        TxtCurTool.Text = "移点"
    Else
        TxtCurTool.Text = "Move Element Pt."
    End If
End Sub

Private Sub MnuBlack_Click()
    ColorMode = 1
    FormResize
    SaveColorMode
    
    MnuWhite.Checked = False
    MnuBlack.Checked = True
End Sub

Public Sub MnuBYDrawingOrder_Click()
    SetDroppingByDrawingOrder
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub


Private Sub mnuCHN_Click()
    ChangeFace 1
    languageType = 0
    SelLanguage 0
    If Frame1.Visible = False Then
        If languageType = 0 Then
            CmdManualControl.caption = "显示手动控制面板"
        Else
            CmdManualControl.caption = "Show Manual Panel"
        End If

    Else
        If languageType = 0 Then
            CmdManualControl.caption = "关闭手动控制面板"
        Else
            CmdManualControl.caption = "Hide Manual Panel"
        End If


    End If
End Sub

Private Sub mnuENG_Click()
    ChangeFace 2
    languageType = 1
    SelLanguage 1
    If Frame1.Visible = False Then
        If languageType = 0 Then
            CmdManualControl.caption = "显示手动控制面板"
        Else
            CmdManualControl.caption = "Show Manual Panel"
        End If
        
    Else
        If languageType = 0 Then
            CmdManualControl.caption = "关闭手动控制面板"
        Else
            CmdManualControl.caption = "Hide Manual Panel"
        End If
        
        
    End If
End Sub

Private Sub MnuEraseAll_Click()
    Dim I As Integer, j As Integer, id As Long
    Dim X As Double, Y As Double
    
    For I = 0 To AuxXLineCount - 1
        For j = 0 To AuxYLineCount - 1
            X = AuxXLine(I)
            Y = AuxYLine(j)
            
            id = CatchPoint(X, Y, 0, CatchPointMode.Normal)
            
            If id > 0 Then
                DeletePoint id
            End If
        Next
    Next
    
    SaveUndo
    PicPathCls
    DrawAll
End Sub

Private Sub MnuEraseDroppingSetting_Click()
    EraseDroppingSetting
    
    SaveUndo
    
    FrmMain.PicPathCls
    DrawAll
End Sub

Private Sub mnuExit_Click()
    Form_Unload 0
End Sub

Private Sub MnuImportAI_Click()
    Dim fn As String, s As String
    
    CmnDlg.Filter = "*.AI|*.AI"
    CmnDlg.ShowOpen
    If CmnDlg.CancelError = False Then
        fn = CmnDlg.FileName
        
        If Trim(fn) <> "" Then
            PicPath.Enabled = False
            
            Me.Refresh
            
            PathOutputPointCount = 0
            
            DeleteAllUndoFiles
            PicPathCls
            DeleteAll
    
            ImportAI fn
                            
            CheckOuterAndInnerLines
            
            FitScreen
            CurFileName = fn
            s = FileName(fn)
            FrmMain.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%) [" & s & "]"
            
            DataChanged = False
            SaveUndo
            
            MakeRecentFileMenu fn
            
            PicPath.Enabled = True
            PicPath.SetFocus
        End If
    End If
    
    FitScreen
    
End Sub

Private Sub MnuImportDXF_Click()
    Dim fn As String, s As String
    
    CmnDlg.Filter = "*.DXF|*.DXF"
    CmnDlg.ShowOpen
    If CmnDlg.CancelError = False Then
        fn = CmnDlg.FileName
        
        If Trim(fn) <> "" Then
            PicPath.Enabled = False
            
            Me.Refresh
            
            PathOutputPointCount = 0
            
            DeleteAllUndoFiles
            PicPathCls
            DeleteAll
    
            ImportDXF fn    '打开指定的文件
                            
            CheckOuterAndInnerLines

            FitScreen
            CurFileName = fn
            s = FileName(fn)
            FrmMain.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%) [" & s & "]"
            
            DataChanged = False
            SaveUndo
            
            MakeRecentFileMenu fn
            
            PicPath.Enabled = True
            PicPath.SetFocus
        End If
    End If
    
    FitScreen
End Sub

Private Sub MnuM45_Click()
    Dim I As Long, t As Double
    
    MnuEraseDroppingSetting_Click
    
    For I = 1 To PointCount
        t = PointList(I).X
        PointList(I).X = PointList(I).Y
        PointList(I).Y = t
    Next
    For I = 1 To ArcCount
        t = ArcList(I).X
        ArcList(I).X = ArcList(I).Y
        ArcList(I).Y = t
        
        t = ArcList(I).a
        ArcList(I).a = ArcList(I).b
        ArcList(I).b = t
        
        ArcList(I).start_angle = Pi / 2 - ArcList(I).start_angle
        ArcList(I).end_angle = Pi / 2 - ArcList(I).end_angle
        
        ArcList(I).ax_angle = -ArcList(I).ax_angle
    Next
    For I = 1 To OutputStartPointList.count
        t = OutputStartPointList.leading_point0(I).X
        OutputStartPointList.leading_point0(I).X = OutputStartPointList.leading_point0(I).Y
        OutputStartPointList.leading_point0(I).Y = t
        
        t = OutputStartPointList.leading_point1(I).X
        OutputStartPointList.leading_point1(I).X = OutputStartPointList.leading_point1(I).Y
        OutputStartPointList.leading_point1(I).Y = t
    Next
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub MnuMatrixByColumn_Click()
    Dim I As Long, j As Long, k As Long, pt() As Path_Point, t As Path_Point, end_pid As Long
    Dim Col_pt(500) As Path_Point, col_x As Double, col_pt_count As Long, m As Long, n As Long
    
    ReDim pt(PointCount)
    
    If PointCount < 1 Then
        Exit Sub
    End If
    
    EraseDroppingSetting

    '-----------------------
    OutputStartPointList.count = 0
    '----------------------------
    
    For I = 1 To PointCount
        pt(I) = PointList(I)
    Next
    
    For I = 1 To PointCount - 1
        For j = I + 1 To PointCount
            If pt(I).X > pt(j).X Then
                t = pt(I)
                pt(I) = pt(j)
                pt(j) = t
            End If
        Next
    Next
    
    I = 0
    j = 0
    col_x = pt(1).X
    Do While I < PointCount
        I = I + 1
        If pt(I).X = col_x Then
            j = j + 1
            Col_pt(j) = pt(I)
        End If
        
        If pt(I).X > col_x Or I = PointCount Then
            k = k + 1
            For m = 1 To j - 1
                For n = m + 1 To j
                    If k Mod 2 = 1 Then
                        If Col_pt(m).Y > Col_pt(n).Y Then
                            t = Col_pt(m)
                            Col_pt(m) = Col_pt(n)
                            Col_pt(n) = t
                        End If
                    Else
                        If Col_pt(m).Y < Col_pt(n).Y Then
                            t = Col_pt(m)
                            Col_pt(m) = Col_pt(n)
                            Col_pt(n) = t
                        End If
                    End If
                Next
            Next
            
            For m = 1 To j
                SetStartDroppingOnChain Col_pt(m).id, 0, end_pid
                
                OutputStartPointList.count = OutputStartPointList.count + 1
                ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.count)
                
                OutputStartPointList.point_id(OutputStartPointList.count) = Col_pt(m).id
                OutputStartPointList.leading_point0(OutputStartPointList.count) = PointList(Col_pt(m).id)
                OutputStartPointList.leading_point1(OutputStartPointList.count) = PointList(end_pid)
            Next
            
            If pt(I).X > col_x Or I < PointCount Then
                col_x = pt(I).X
                I = I - 1
                j = 0
            Else
                Exit Do
            End If
        End If
    Loop
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub MnuMatrixByRow_Click()
    Dim I As Long, j As Long, k As Long, pt() As Path_Point, t As Path_Point, end_pid As Long
    Dim Row_pt(500) As Path_Point, row_y As Double, row_pt_count As Long, m As Long, n As Long
    
    ReDim pt(PointCount)
    
    If PointCount < 1 Then
        Exit Sub
    End If
    
    EraseDroppingSetting

    '-----------------------
    OutputStartPointList.count = 0
    '----------------------------
    
    For I = 1 To PointCount
        pt(I) = PointList(I)
    Next
    
    For I = 1 To PointCount - 1
        For j = I + 1 To PointCount
            If pt(I).Y > pt(j).Y Then
                t = pt(I)
                pt(I) = pt(j)
                pt(j) = t
            End If
        Next
    Next
    
    I = 0
    j = 0
    row_y = pt(1).Y
    Do While I < PointCount
        I = I + 1
        If pt(I).Y = row_y Then
            j = j + 1
            Row_pt(j) = pt(I)
        End If
        
        If pt(I).Y > row_y Or I = PointCount Then
            k = k + 1
            For m = 1 To j - 1
                For n = m + 1 To j
                    If k Mod 2 = 1 Then
                        If Row_pt(m).X > Row_pt(n).X Then
                            t = Row_pt(m)
                            Row_pt(m) = Row_pt(n)
                            Row_pt(n) = t
                        End If
                    Else
                        If Row_pt(m).X < Row_pt(n).X Then
                            t = Row_pt(m)
                            Row_pt(m) = Row_pt(n)
                            Row_pt(n) = t
                        End If
                    End If
                Next
            Next
            
            For m = 1 To j
                SetStartDroppingOnChain Row_pt(m).id, 0, end_pid
                
                OutputStartPointList.count = OutputStartPointList.count + 1
                ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.count)
                
                OutputStartPointList.point_id(OutputStartPointList.count) = Row_pt(m).id
                OutputStartPointList.leading_point0(OutputStartPointList.count) = PointList(Row_pt(m).id)
                OutputStartPointList.leading_point1(OutputStartPointList.count) = PointList(end_pid)
            Next
            
            If pt(I).Y > row_y Or I < PointCount Then
                row_y = pt(I).Y
                I = I - 1
                j = 0
            Else
                Exit Do
            End If
        End If
    Loop
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub MnuMX_Click()
    Dim I As Long, k As Integer
    Dim MinX As Double, MaxX As Double, MeanX As Double
    
    MnuEraseDroppingSetting_Click
    'k = 0
    'For i = 1 To PointCount
    '    If i = 1 Then
    '        MinX = PointList(i).X
    '        MaxX = MinX
    '        k = 1
    '    Else
    '        If MinX > PointList(i).X Then
    '            MinX = PointList(i).X
    '        ElseIf MaxX < PointList(i).X Then
    '            MaxX = PointList(i).X
    '        End If
    '    End If
    'Next
    'For i = 1 To ArcCount
    '    If k = 0 Then
    '        MinX = ArcList(i).X - ArcList(i).a
    '        MaxX = ArcList(i).X + ArcList(i).a
    '        k = 1
    '    Else
    '        If MinX > ArcList(i).X - ArcList(i).a Then
    '            MinX = ArcList(i).X - ArcList(i).a
    '        End If
    '        If MaxX < ArcList(i).X + ArcList(i).a Then
    '            MaxX = ArcList(i).X + ArcList(i).a
    '        End If
    '    End If
    'Next
    
    'MeanX = (MinX + MaxX) / 2
    MeanX = (ViewMinX + ViewMaxX) / 2
    
    For I = 1 To PointCount
        PointList(I).X = 2 * MeanX - PointList(I).X
    Next
    For I = 1 To ArcCount
        ArcList(I).X = 2 * MeanX - ArcList(I).X
        
        ArcList(I).start_angle = Pi - ArcList(I).start_angle
        ArcList(I).end_angle = Pi - ArcList(I).end_angle
        
        ArcList(I).ax_angle = PI2 - ArcList(I).ax_angle
    Next
    
    For I = 1 To OutputStartPointList.count
        OutputStartPointList.leading_point0(I).X = 2 * MeanX - OutputStartPointList.leading_point0(I).X
        
        OutputStartPointList.leading_point1(I).X = 2 * MeanX - OutputStartPointList.leading_point1(I).X
    Next
                
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub MnuMY_Click()
    Dim I As Long, k As Integer
    Dim MinY As Double, MaxY As Double, MeanY As Double
    
    MnuEraseDroppingSetting_Click
    'k = 0
    'For i = 1 To PointCount
    '    If i = 1 Then
    '        MinY = PointList(i).Y
    '        MaxY = MinY
    '        k = 1
    '    Else
    '        If MinY > PointList(i).Y Then
    '            MinY = PointList(i).Y
    '        ElseIf MaxY < PointList(i).Y Then
    '            MaxY = PointList(i).Y
    '        End If
    '    End If
    'Next
    'For i = 1 To ArcCount
    '    If k = 0 Then
    '        MinY = ArcList(i).Y - ArcList(i).B
    '        MaxY = ArcList(i).Y + ArcList(i).B
    '        k = 1
    '    Else
    '        If MinY > ArcList(i).Y - ArcList(i).B Then
    '            MinY = ArcList(i).Y - ArcList(i).B
    '        End If
    '        If MaxY < ArcList(i).Y + ArcList(i).B Then
    '            MaxY = ArcList(i).Y + ArcList(i).B
    '        End If
    '    End If
    'Next
    '
    'MeanY = (MinY + MaxY) / 2
    MeanY = (ViewMinY + ViewMaxY) / 2
    
    For I = 1 To PointCount
        PointList(I).Y = 2 * MeanY - PointList(I).Y
    Next
    For I = 1 To ArcCount
        ArcList(I).Y = 2 * MeanY - ArcList(I).Y
        
        ArcList(I).start_angle = 2 * Pi - ArcList(I).start_angle
        ArcList(I).end_angle = 2 * Pi - ArcList(I).end_angle
        
        ArcList(I).ax_angle = 2 * Pi - ArcList(I).ax_angle
    Next
    For I = 1 To OutputStartPointList.count
        OutputStartPointList.leading_point0(I).Y = 2 * MeanY - OutputStartPointList.leading_point0(I).Y
        
        OutputStartPointList.leading_point1(I).Y = 2 * MeanY - OutputStartPointList.leading_point1(I).Y
    Next
                
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub MnuNew_Click()
    DeleteAllUndoFiles
    PicPathCls
    DeleteAll
    
    ReadUserParameter
    
    DrawGridLines
    
    CurFileName = ""
    FrmMain.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%)"
    
    DataChanged = False
End Sub

Private Sub MnuOpen_Click()
    Dim fn As String, s As String
    
    CmnDlg.Filter = "*.ITN"
    CmnDlg.FileName = ""
    CmnDlg.ShowOpen
    If CmnDlg.FileName <> "" And CmnDlg.CancelError = False Then
        fn = CmnDlg.FileName
        
        PathOutputPointCount = 0
    
        DeleteAllUndoFiles
        PicPathCls
        DeleteAll
        
        ReadFile fn
        
        CheckOuterAndInnerLines

        FitScreen
        
        CurFileName = fn
        s = FileName(fn)
        FrmMain.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%)[" & s & "]"
        
        SaveUndo
        
        MakeRecentFileMenu fn
    
        DataChanged = False
    End If
End Sub

Private Sub MnuOrgLeftDown_Click()
    Device_CoordinateMode = 0
    Device_SetMode
    
    FrmMain.Form_Load
End Sub

Private Sub MnuOrgLeftUp_Click()
    Device_CoordinateMode = 1
    Device_SetMode
    
    FrmMain.Form_Load
End Sub

Private Sub mnuPlatformParameter_Click()
'    FrmUserParameter.Show
End Sub

Private Sub MnuPointAndLine_Click()
    Device_Mode = ModeType.PointsAndLines
    Device_SetMode
    
    FrmMain.Form_Load
End Sub

Private Sub MnuPointOnly_Click()
    Device_Mode = ModeType.PointsOnly
    Device_SetMode
    
    FrmMain.Form_Load
End Sub

Sub FitScreen()
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, xmin As Double, ymin As Double, xmax As Double, ymax As Double, I As Long
    
    MaxBodyID = GetBodyList(BodyList)
    For I = 1 To MaxBodyID
        GetBodyScale I, x0, y0, x1, y1
        If I = 1 Then
            xmin = x0
            ymin = y0
            
            xmax = x1
            ymax = y1
        Else
            If xmin > x0 Then xmin = x0
            If ymin > y0 Then ymin = y0
            
            If xmax < x1 Then xmax = x1
            If ymax < y1 Then ymax = y1
        End If
    Next
    MoveAllBody -xmin + 10, -ymin + 10
            
    Device_UserSize(1) = GetValueFromINI("UserSize", str(1), "1000", App.Path & "\" & App.EXEName & ".ini")
    Device_UserSize(2) = GetValueFromINI("UserSize", str(2), "1000", App.Path & "\" & App.EXEName & ".ini")
        
    If Device_UserSize(1) < xmax - xmin + 20 Then
        Device_UserSize(1) = 50 * Int((xmax - xmin + 20 + 49) / 50)
    End If
    
    If Device_UserSize(2) < ymax - ymin + 20 Then
        Device_UserSize(2) = 50 * Int((ymax - ymin + 20 + 49) / 50)
    End If
    
    ViewMinX = 0
    ViewMaxX = Device_UserSize(1)
    ViewMinY = 0
    ViewMaxY = Device_UserSize(2)
    ViewMargin = 0.03

    DirectionChanged = True
    ShiftX = 0
    ShiftY = 0
    Zoom 0
    
    Me.Refresh
    
    PopAllXORStack
    CloseXORStack
    
    DrawAll
    
    CurTool = ToolType.MoveCanvas
    'TxtCurTool.Text = "平移"
    TxtCurTool.Text = "Move"
End Sub

Private Sub MnuRecentFile_Click(Index As Integer)
    Dim fn As String, s As String
    
    Me.Refresh
    
    fn = MnuRecentFile.Item(Index).Tag
    
    PathOutputPointCount = 0
    
    DeleteAllUndoFiles
    PicPathCls
    DeleteAll
    
    If UCase(Right(fn, 4)) = ".DXF" Then
        ImportDXF fn
        s = FileName(fn) + ".DXF"
    ElseIf UCase(Right(fn, 3)) = ".AI" Then
        ImportAI fn
        s = FileName(fn) + ".DXF"
    Else
        ReadFile fn
        s = FileName(fn)
    End If
                            
    CheckOuterAndInnerLines
    
    FitScreen
    CurFileName = fn
    FrmMain.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%) [" & s & "]"
    
    DataChanged = False
    SaveUndo
    
    MakeRecentFileMenu fn
End Sub

Private Sub MnuRedo_Click()
    Redo
End Sub

Private Sub MnuRegistration_Click()
    'Load FrmRegistration
    'FrmRegistration.Show
    If languageType = 0 Then
    ShellExecute Me.hWnd, "open", "翰顿围字机操作指南.doc", "", "", 1
    ElseIf languageType = 1 Then
        ShellExecute Me.hWnd, "open", "User's Guide for HandunBender.doc", "", "", 1
    End If
End Sub

Private Sub MnuRotate_Click()
    MnuEraseDroppingSetting_Click
    FrmRotate.Show
End Sub

Private Sub MnuSaleUserDefined_Click()
    MnuEraseDroppingSetting_Click
    FrmScale.Show
End Sub

Private Sub MnuSave_Click()
    Dim fn As String
    
    If UCase(Right(CurFileName, 4)) <> ".DXF" And CurFileName <> "" Then
        WriteFile CurFileName
        DataChanged = False
    Else
        If UCase(Right(CurFileName, 4)) = ".DXF" Then
            CurFileName = Mid(CurFileName, 1, Len(CurFileName) - 4) + ".ITN"
        End If
        CmnDlg.Filter = "*.ITN|*.ITN"
        CmnDlg.FileName = CurFileName
        CmnDlg.ShowSave
        If CmnDlg.FileName <> "" And CmnDlg.CancelError = False Then
            fn = CmnDlg.FileName
            
            WriteFile fn
            CurFileName = fn
            
            MakeRecentFileMenu fn
            DataChanged = False
        End If
    End If
End Sub

Private Sub MnuSaveAs_Click()
    Dim fn As String
    
    CmnDlg.Filter = "*.ITN|*.ITN"
    CmnDlg.FileName = ""
    CmnDlg.ShowSave
    If CmnDlg.FileName <> "" And CmnDlg.CancelError = False Then
        fn = CmnDlg.FileName
        WriteFile fn
        CurFileName = fn
        
        MakeRecentFileMenu fn
        DataChanged = False
    End If
End Sub


Private Sub MnuScaleM_Click(Index As Integer)
    Dim I As Long, sx As Double, sy As Double
    
    MnuEraseDroppingSetting_Click
    
    sx = Index
    sy = Index
    
    For I = 1 To PointCount
        PointList(I).X = PointList(I).X * sx
        PointList(I).Y = PointList(I).Y * sy
    Next
    For I = 1 To ArcCount
        If ArcList(I).ax_angle = 0 Or sx = sy Then
            ArcList(I).X = ArcList(I).X * sx
            ArcList(I).Y = ArcList(I).Y * sy
            
            ArcList(I).a = ArcList(I).a * sx
            ArcList(I).b = ArcList(I).b * sy
        Else
            '此处应作更全面的处理
            
            ArcList(I).X = ArcList(I).X * sx
            ArcList(I).Y = ArcList(I).Y * sx
            
            ArcList(I).a = ArcList(I).a * sx
            ArcList(I).b = ArcList(I).b * sx
        End If
    Next
    
    For I = 1 To OutputStartPointList.count
        OutputStartPointList.leading_point0(I).X = OutputStartPointList.leading_point0(I).X * sx
        OutputStartPointList.leading_point0(I).Y = OutputStartPointList.leading_point0(I).Y * sy
        
        OutputStartPointList.leading_point1(I).X = OutputStartPointList.leading_point1(I).X * sx
        OutputStartPointList.leading_point1(I).Y = OutputStartPointList.leading_point1(I).Y * sy
    Next
                
    PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub MnuSetAll_Click()
    Dim I As Integer, j As Integer, k As Integer, c As Long, end_pid As Long
    Dim p As Path_Point
    
    For I = 0 To AuxXLineCount - 1
        For j = 0 To AuxYLineCount - 1
            If I Mod 2 = 0 Then
                k = j
            Else
                k = AuxYLineCount - 1 - j
            End If
            p.X = AuxXLine(I)
            p.Y = AuxYLine(k)
            p.Layer = 1
            
            c = PointCount
            CatchOrAddPoint p
            
            If Device_Mode = 1 And PointCount > c Then
                SetStartDroppingOnChain PointCount, 0, end_pid
                
                OutputStartPointList.count = OutputStartPointList.count + 1
                ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.count)
                
                OutputStartPointList.point_id(OutputStartPointList.count) = PointCount
                OutputStartPointList.leading_point0(OutputStartPointList.count) = PointList(PointCount)
                OutputStartPointList.leading_point1(OutputStartPointList.count) = PointList(end_pid)
            End If
        Next
    Next
    
    SaveUndo
    DrawAllPoints
End Sub

Private Sub MnuSetParam_Click()
'    ChangeFace 2
'    SelLanguage 1
'    ChangeFace 1
'    SelLanguage 0
If curLanguage = 0 Then
    ChangeParamSetLanguage 2
    ChangeParamSetLanguage 1
End If
    FormSettings.Show
End Sub

Private Sub MnuShift_Click()
    MnuEraseDroppingSetting_Click
    FrmShift.Show
End Sub

Private Sub MnuShiftToOrg_Click()
    Dim I As Long
    Dim MinX As Double, MinY As Double
    Dim dX As Double, dy As Double
    
    MnuEraseDroppingSetting_Click
    
    For I = 1 To PointCount
        If I = 1 Then
            MinX = PointList(I).X
            MinY = PointList(I).Y
        Else
            If MinX > PointList(I).X Then
                MinX = PointList(I).X
            End If
            If MinY > PointList(I).Y Then
                MinY = PointList(I).Y
            End If
        End If
    Next
    
    dX = -MinX
    dy = -MinY
    
    For I = 1 To PointCount
        PointList(I).X = PointList(I).X + dX
        PointList(I).Y = PointList(I).Y + dy
    Next
    For I = 1 To ArcCount
        ArcList(I).X = ArcList(I).X + dX
        ArcList(I).Y = ArcList(I).Y + dy
    Next
    
    PicPathCls
    DrawAll
    
    SaveUndo
End Sub

Private Sub mnuShowPointList_Click()
    FrmShowPointList.Show
End Sub

Private Sub MnuShowTest_Click()
'    FrmTest.Show
End Sub

Private Sub MnuUndo_Click()
    Undo
End Sub

Private Sub MnuWhite_Click()
    ColorMode = 0
    FormResize
    SaveColorMode
        
    MnuWhite.Checked = True
    MnuBlack.Checked = False
End Sub

Function Zoom(ByVal Zoom_Index As Integer) As Boolean
'***************************************************************************************************************************
'
'***************************************************************************************************************************
    Dim d As Double, v As Double, s As String, fn As String, z0 As Double, h0 As Double, v0 As Double
    
    On Error Resume Next
    
    h0 = HScroll1.value
    v0 = VScroll1.value
    
    DrawCursorReferenceLines 0, 0, 0
    
    Select Case Zoom_Index
        Case 0: d = 1
        Case 1: d = 1.25
        Case 2: d = 1.5
        Case 3: d = 2
        Case 4: d = 4
        Case 5: d = 8
        Case 6: d = 12
        Case 7: d = 16
        Case 8: d = 20
    End Select
    
    ZoomFactor = d
    
    fn = CurFileName
    If fn <> "" Then
        s = FileName(fn)
        Me.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%) [" & s & "]"
    Else
        Me.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%)"
    End If
    Me.MousePointer = 11
    
    PicPath.Move 0, 0
    HScroll1.value = 0
    VScroll1.value = 0
    PicPathCls

    PicPath.Width = d * PicFrame.ScaleWidth
    PicPath.Height = d * PicFrame.ScaleHeight
    
    PicPath.ScaleMode = 3
    If Device_CoordinateMode = 0 Then
        PicPath.ScaleTop = PicPath.Height * PicPathRatio - 3
        PicPath.ScaleHeight = -PicPath.Height * PicPathRatio + 2
    Else
        PicPath.ScaleTop = 0
        PicPath.ScaleHeight = PicPath.Height * PicPathRatio - 2
    End If
    
    PathMinX = 0
    PathMaxX = PicPath.ScaleWidth - 1
    If Device_CoordinateMode = 0 Then
        PathMinY = 0
        PathMaxY = -(PicPath.ScaleHeight - 1)
    Else
        PathMinY = 0
        PathMaxY = PicPath.ScaleHeight - 1
    End If
    
    PicPath.PSet (0, 0), PicPath.BackColor
    
    v = HScroll1.value
    HScroll1.value = IIf(v < PicPath.Width, v, PicPath.Width - 1)
    PicPath.left = -HScroll1.value
    
    HScroll1.Min = 0
    HScroll1.Max = PicPath.Width - PicFrame.ScaleWidth
    HScroll1.LargeChange = IIf(d < 2, HScroll1.Max + 1, PicFrame.ScaleWidth + 1)
    
    v = VScroll1.value
    VScroll1.value = IIf(v < PicPath.Height, v, PicPath.Height - 1)
    PicPath.top = -VScroll1.value
    
    VScroll1.Min = 0
    VScroll1.Max = PicPath.Height - PicFrame.ScaleHeight
    VScroll1.LargeChange = IIf(d < 2, VScroll1.Max + 1, PicFrame.ScaleHeight + 1)
    
    If PicPath.Visible = True Then
        PicPath.SetFocus
    End If
    
    SetUserScale ViewMaxX, ViewMaxY, ViewMargin
    DrawAll
    
    Me.MousePointer = 0
    
    utw = 3 * GetUserDistance(0, 0, TrapWidth, 0)
'Debug.Print "utw="; utw
    Zoom = True
End Function

Private Sub MnuZoom_Click(Index As Integer)
'***************************************************************************************************************************
'                                   菜单缩放处理函数
'***************************************************************************************************************************
    Dim d As Double
    
    'Zoom Index
    Select Case Index
        Case 0: d = 1
        Case 1: d = 1.25
        Case 2: d = 1.5
        Case 3: d = 2
        Case 4: d = 4
        Case 5: d = 8
        Case 6: d = 12
        Case 7: d = 16
        Case 8: d = 20
    End Select

    ZoomFactor = d
    ShiftX = 0
    ShiftY = 0
    
    PicPathCls
    DrawAll
    
    utw = 3 * GetUserDistance(0, 0, TrapWidth, 0)
'Debug.Print "utw="; utw
End Sub

Sub MouseScrollWheel(ByVal X As Single, ByVal Y As Single, ByVal zoom_mode As Long)
'***************************************************************************************************************************
'                                       鼠标滚轮处理函数
'***************************************************************************************************************************
    Dim s As String, fn As String
    Dim ux As Double, uy As Double, x1 As Single, y1 As Single
    
    ConvertPathToUser X, Y, ux, uy

    
    If zoom_mode = 1 And ZoomFactor < 100 Then
        ZoomFactor = ZoomFactor * 1.2
    ElseIf zoom_mode = -1 And ZoomFactor > 0.5 Then
        ZoomFactor = ZoomFactor / 1.2
    Else
        Exit Sub
    End If
    
    ConvertUserToPath ux, uy, x1, y1
    
    ShiftX = ShiftX - (x1 - X)
    ShiftY = ShiftY - (y1 - Y)
    
    PicPathCls
    DrawAll False
    
    fn = CurFileName
    If fn <> "" Then
        s = FileName(fn)
        Me.caption = AppVersion & VersionMark & " (" & Trim(str(Round(ZoomFactor * 100))) & "%) [" & s & "]"
    Else
        Me.caption = AppVersion & VersionMark & " (" & Trim(str(Round(ZoomFactor * 100))) & "%)"
    End If
    Me.MousePointer = 0
    
    utw = 3 * GetUserDistance(0, 0, TrapWidth, 0)
'Debug.Print "utw="; utw
End Sub

Private Sub PanButton1_Click()
    CurTool = ToolType.StopPoint
    If curLanguage = 0 Then
        TxtCurTool.Text = "设终点"
    Else
        TxtCurTool.Text = "Set End"
    End If
End Sub

Private Sub PanButton11_Click()
    CurTool = ToolType.StartPoint
    If curLanguage = 0 Then
        TxtCurTool.Text = "设起点"
    Else
        TxtCurTool.Text = "Set Start"
    End If
End Sub

Private Sub CmdTurnReset_Click()
    StopRunning = False
    
    BendReset
End Sub

Private Sub PanButton2_Click()
    CurTool = ToolType.Reverse
    If curLanguage = 0 Then
        TxtCurTool.Text = "内轮廓"
    Else
        TxtCurTool.Text = "Set Innerline"
    End If
End Sub

Private Sub PanButton3_Click()
    CurTool = ToolType.Reverse
    If curLanguage = 0 Then
    TxtCurTool.Text = "外轮廓"
    Else
    TxtCurTool.Text = "Set Outline"
    End If
End Sub

Private Sub PanButton4_Click()
    CurTool = ToolType.MeasureDistance
    If curLanguage = 0 Then
        TxtCurTool.Text = "测距离"
    Else
        TxtCurTool.Text = "Measure Distance"
    End If
End Sub

Private Sub PanButton5_Click()
    MnuEraseDroppingSetting_Click
End Sub

Private Sub PanButton6_Click()
    CurTool = ToolType.MeasureScale
    If curLanguage = 0 Then
        TxtCurTool.Text = "测面积"
    Else
        TxtCurTool.Text = "Measure Scale"
    End If
End Sub



Private Sub PicFace_Click()
    PicFace.Visible = False
End Sub

Private Sub PicPath_DblClick()
'***************************************************************************************************************************
'
'***************************************************************************************************************************

    'MnuZoom_Click 0
'Exit Sub

    If (CurTool = ToolType.SetCircle And CurToolStep = 2) Or _
       (CurTool = ToolType.SetCircle_3p And CurToolStep >= 1.5) Or _
       (CurTool = ToolType.SetEllipse And CurToolStep >= 1.5) Then
        PopAllXORStack
        CloseXORStack
        
        'Make a circle/ellipse
        CurArc.end_angle = CurArc.start_angle + PI2
        ArcList(CurArc.id) = CurArc
        PointList(CurArc.point1_id).X = PointList(CurArc.point0_id).X
        PointList(CurArc.point1_id).Y = PointList(CurArc.point0_id).Y
        DrawArc ArcList(CurArc.id)
        
        PicPath_MouseDown 9, 0, 0, 0
        CurToolStep = 0
        
        SaveUndo
        
    End If
    
    
End Sub

Private Sub PicPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************************************************************************************
'
'***************************************************************************************************************************
    Dim ux As Double, uy As Double, ux0 As Double, uy0 As Double, ux1 As Double, uy1 As Double, cx As Double, cy As Double
    Dim id As Long, r As Double, r2 As Double, a As Double, da As Double, sa As Double, ea As Double
    Dim x0 As Single, y0 As Single, dX As Single, dy As Single, w As Double, h As Double, x1 As Single, y1 As Single
    Dim I As Long, j As Long, k As Long, n As Long, id0 As Long, id1 As Long, body_id As Long, group_id As Long
    Dim Ret As Boolean, top As Single
        
    Static ux_min As Double, uy_min As Double, ux_max As Double, uy_max As Double
    
    Dim TempPoint As Path_Point, NullPoint As Path_Point, end_pid As Long
    
    Dim NullArc As Path_Arc
    
    DrawCursorReferenceLines 0, 0, 0
        
    If TxtStatistics.Visible = True And FraToolBox.Visible = True Then
        TxtStatistics.Visible = False
    End If

    'If FraToolBox.Visible = False And catched_pid = 0 Then
    '    CurTool = ToolType.MoveCanvas
    '    TxtCurTool.Text = "平移"
    '    GoTo Exit_Sub
    'End If
    
    If Button = 2 Then
        If CurTool = ToolType.SetSegment Or CurTool = ToolType.ConnectTwoPoints Then
            PopAllXORStack
            CloseXORStack
        
            LastPoint.id = 0
            CurPoint.id = 0
            CurToolStep = 0
            
        ElseIf CurTool = ToolType.SetSPLine Then
            PopAllXORStack
            CloseXORStack
            
            If TempSPline.vertex_count > 2 Then
                AddSPLineByTempSPline
                DrawSPLine SPLineList(SPLineCount)
            End If
            CurToolStep = 0
            
            SaveUndo
        ElseIf CurTool = ToolType.Unit Then
            CurUnittedGroupID = 0
            
        Else
            'CurTool = ToolType.MoveCanvas
            'TxtCurTool.Text = "平移"
        
        End If
        
        GoTo Exit_Sub
    End If
    
    ConvertPathToUser X, Y, ux, uy
    
    If CurTool = ToolType.MoveCanvas Then 'Move
        path_x0 = X
        path_y0 = Y
        
    ElseIf CurTool = ToolType.SetPoint Then 'Add Point
        If ux >= 0 And ux <= ViewMaxX And uy >= 0 And uy <= ViewMaxY Then
            AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
            
'            If Device_Mode = 1 Then
'                ' Add dropping point directly
'                '------------------------------------------------------------------------
'                SetStartDroppingOnChain PointCount, 0, end_pid
'
'                OutputStartPointList.Count = OutputStartPointList.Count + 1
'                ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.Count)
'                ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.Count)
'                ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.Count)
'
'                OutputStartPointList.point_id(OutputStartPointList.Count) = PointCount
'                OutputStartPointList.leading_point0(OutputStartPointList.Count) = PointList(PointCount)
'                OutputStartPointList.leading_point1(OutputStartPointList.Count) = PointList(end_pid)
'                '------------------------------------------------------------------------
'            End If
    
            DrawPoint PointList(PointCount)
            
            ReDim tv(2) As Title_Value
            tv(0).t = "X"
            tv(0).v = ux
            tv(1).t = "Y"
            tv(1).v = uy
            ShowEditData "点编号:" & Trim(str(PointCount)), 2, tv, ToolType.SetPoint
            
            SaveUndo
        Else
            Beep
        End If
        
    ElseIf CurTool = ToolType.SetSegment Then 'Add Segment
        If CurToolStep = 0 Then
            If catched_pid > 0 Then
                CurPoint = PointList(catched_pid)
                LastPoint = CurPoint
                catched_pid = 0
            Else
                If ux >= 0 And ux <= ViewMaxX And uy >= 0 And uy <= ViewMaxY Then
                    AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
        
                    LastPoint = CurPoint
                    DrawPoint PointList(PointCount)
                    
                    ReDim tv(2) As Title_Value
                    tv(0).t = "X"
                    tv(0).v = ux
                    tv(1).t = "Y"
                    tv(1).v = uy

                    ShowEditData "点编号:" & Trim(str(PointCount)), 2, tv, ToolType.SetSegment
                Else
                    Beep
                End If
            End If
            CurToolStep = 1
                
        Else
            PopAllXORStack
            CloseXORStack
            
            k = 0
            
            If catched_pid > 0 Then
                LastPoint = CurPoint
                CurPoint = PointList(catched_pid)
                
                For I = 1 To SegmentCount
                    If SegmentList(I).point1_id = catched_pid Then
                        PopAllXORStack
                        CloseXORStack
                    
                        LastPoint.id = 0
                        CurPoint.id = 0
                        CurToolStep = 0
                        GoTo Exit_Sub
                    End If
                Next
                
                For I = 1 To ArcCount
                    If ArcList(I).point1_id = catched_pid Or _
                       (ArcList(I).point0_id = catched_pid And ArcList(I).Type = ArcType.RoundedCorner) Then
                        PopAllXORStack
                        CloseXORStack
                    
                        LastPoint.id = 0
                        CurPoint.id = 0
                        CurToolStep = 0
                        GoTo Exit_Sub
                    End If
                Next
                
                For I = 1 To SPLineCount
                    For j = 1 To SPLineList(I).vertex_count - 1 'only poin0_id allowed (j = 0)
                        If SPLineList(I).vertex_id(j) = catched_pid Then
                            PopAllXORStack
                            CloseXORStack
                        
                            LastPoint.id = 0
                            CurPoint.id = 0
                            CurToolStep = 0
                            GoTo Exit_Sub
                        End If
                    Next
                Next
                
                For I = 1 To SegmentCount
                    If SegmentList(I).point0_id = catched_pid Then
                        k = 1
                        Exit For
                    End If
                Next
                
                If k = 0 Then
                    For I = 1 To ArcCount
                        If ArcList(I).point0_id = catched_pid Then
                            k = 1
                            Exit For
                        End If
                    Next
                End If
            
                If k = 0 Then
                    For I = 1 To SPLineCount
                        If SPLineList(I).point0_id = catched_pid Then
                            k = 1
                            Exit For
                        End If
                    Next
                End If
            
            Else
                'If ChkCatchHVLine.Value = 1 Then
                If Toolbar1.Buttons(21).value = tbrPressed Then
                    ConvertUserToPath CurPoint.X, CurPoint.Y, x0, y0
            
                    If X <> x0 And Y <> y0 Then
                        dX = Abs(X - x0)
                        dy = Abs(Y - y0)
            
                        If dX <= HVTrapWidth Or dy <= HVTrapWidth Then
                            If dX >= dy Then
                                Y = y0
                            Else
                                X = x0
                            End If
                        End If
                    End If
                End If
                
                If ux >= 0 And ux <= ViewMaxX And uy >= 0 And uy <= ViewMaxY Then
                    AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
                    
                    ReDim tv(2) As Title_Value
                    tv(0).t = "X"
                    tv(0).v = ux
                    tv(1).t = "Y"
                    tv(1).v = uy
                    ShowEditData "点编号:" & Trim(str(PointCount)), 2, tv, ToolType.SetSegment
                Else
                    Beep
                    GoTo Exit_Sub
                End If
        
            End If
            
            AddSegment LastPoint.id, CurPoint.id
            SegmentList(SegmentCount).Type = SegmentType.NormalSegment
            DrawSegment SegmentList(SegmentCount)
            DrawPoint PointList(PointCount)
            
            SaveUndo
            
            If k = 1 Then
                PopAllXORStack
                CloseXORStack
            
                LastPoint.id = 0
                CurPoint.id = 0
                CurToolStep = 0
                GoTo Exit_Sub
            End If
        End If
        OpenXORStack
        
    ElseIf CurTool = ToolType.SetSPLine Then 'Create SPLine
        If CurToolStep = 0 Then
            If catched_pid > 0 Then
                CurPoint = PointList(catched_pid)
                LastPoint = CurPoint
                catched_pid = 0
            Else
                AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.SPLinePoint
        
                LastPoint = CurPoint
                DrawPoint PointList(PointCount)
                
                ReDim tv(2) As Title_Value
                tv(0).t = "X"
                tv(0).v = ux
                tv(1).t = "Y"
                tv(1).v = uy
                ShowEditData "点编号:" & Trim(str(PointCount)), 2, tv, ToolType.SetSPLine
            End If
            CurToolStep = 1
            
            TempSPline.vertex_count = 2
            ReDim TempSPline.vertex_id(TempSPline.vertex_count - 1)
            TempSPline.vertex_id(TempSPline.vertex_count - 2) = CurPoint.id
            TempSPline.point0_id = CurPoint.id
            TempSPline.Layer = CurLayer
            TempSPline.color = LayerColor(CurLayer)
            TempSPline.segment_between_points = SPLine_SegmentBetweenPoints
            
            OpenXORStack
        Else
            PopAllXORStack
            CloseXORStack
            
            j = 0
            
            If catched_pid > 0 Then
                For I = 1 To SegmentCount
                    If SegmentList(I).point1_id = catched_pid Then
                        j = -1
                        Exit For
                    End If
                Next
                
                If j = 0 Then
                    For I = 1 To ArcCount
                        If ArcList(I).point1_id = catched_pid Or _
                           (ArcList(I).point0_id = catched_pid And ArcList(I).Type = ArcType.RoundedCorner) Then
                            j = -1
                            Exit For
                        End If
                    Next
                End If
                
                If j = 0 Then
                    For I = 1 To SegmentCount
                        If SegmentList(I).point0_id = catched_pid Then 'closed to a segment
                            body_id = SegmentList(I).body_id
                            group_id = SegmentList(I).group_id
                            j = 1
                            Exit For
                        End If
                    Next
                End If
                
                If j = 0 Then
                    For I = 1 To ArcCount
                        If ArcList(I).point0_id = catched_pid Then 'closed to an arc
                            body_id = ArcList(I).body_id
                            group_id = ArcList(I).group_id
                            j = 1
                            Exit For
                        End If
                    Next
                End If
                
                If j = 0 Then
                    If TempSPline.point0_id = catched_pid Then 'closed
                        j = 1
                    Else
                        For I = 1 To TempSPline.vertex_count - 2
                            If TempSPline.vertex_id(I) = catched_pid Then
                                j = -1
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                If j = 0 Then
                    For I = 1 To SPLineCount
                        If SPLineList(I).point0_id = catched_pid Then
                            body_id = SPLineList(I).body_id
                            group_id = SPLineList(I).group_id
                            j = 1
                            Exit For
                        End If
                    Next
                End If
                
                If j <> -1 Then
                    LastPoint = CurPoint
                    CurPoint = PointList(catched_pid)
                
                    If j = 1 Then
                        TempSPline.vertex_id(TempSPline.vertex_count - 1) = CurPoint.id
                        
                        TempSPline.vertex_id(TempSPline.vertex_count - 1) = CurPoint.id
                        TempSPline.point1_id = CurPoint.id
            
                        'DrawSPLine TempSPline
                        
                        TempSPline.vertex_count = TempSPline.vertex_count + 1  'add one more point for AddSPLineByTempSPLine
                        ReDim Preserve TempSPline.vertex_id(TempSPline.vertex_count - 1)
                        
                        AddSPLineByTempSPline
                        DrawSPLine SPLineList(SPLineCount)
                        CurToolStep = 0
                        
                        If body_id <> TempSPline.body_id Then
                            ReplaceBodyID body_id, SPLineList(SPLineCount).body_id
                            ReplaceGroupID group_id, SPLineList(SPLineCount).group_id
                        End If
                        
                        SaveUndo
                        GoTo Exit_Sub
                    End If
                Else
                    'LastPoint.id = 0
                    'CurPoint.id = 0
                    'CurToolStep = 0
                    GoTo Exit_Sub
                End If
            Else
                'If ChkCatchHVLine.Value = 1 Then
                If Toolbar1.Buttons(21).value = tbrPressed Then
                    ConvertUserToPath CurPoint.X, CurPoint.Y, x0, y0
            
                    If X <> x0 And Y <> y0 Then
                        dX = Abs(X - x0)
                        dy = Abs(Y - y0)
            
                        If dX <= HVTrapWidth Or dy <= HVTrapWidth Then
                            If dX >= dy Then
                                Y = y0
                            Else
                                X = x0
                            End If
                        End If
                    End If
                End If
                
                AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.SPLinePoint
                catched_pid = PointCount
            End If
            
            TempSPline.vertex_id(TempSPline.vertex_count - 1) = CurPoint.id
            TempSPline.point1_id = CurPoint.id

            OpenXORStack
            DrawPoint PointList(CurPoint.id)
            
            ReDim tv(2) As Title_Value
            tv(0).t = "X"
            tv(0).v = ux
            tv(1).t = "Y"
            tv(1).v = uy
            ShowEditData "点编号:" & Trim(str(PointCount)), 2, tv, ToolType.SetSPLine
            
            If CurToolStep > 1 Then
                DrawSPLine TempSPline, Not LayerColor(CurLayer)
            End If
            
            TempSPline.vertex_count = TempSPline.vertex_count + 1
            ReDim Preserve TempSPline.vertex_id(TempSPline.vertex_count - 1)
            
            CurToolStep = CurToolStep + 1
        End If
        
    ElseIf CurTool = ToolType.MoveElement_Point Then
        CurGroupID = 0
        If catched_pid > 0 Then  'Move Point
            PicPathCls
            DrawGridLines
            DrawGroupExcept PointList(catched_pid).group_id
            
            OpenXORStack
            DrawGroup PointList(catched_pid).group_id, XORColor
            CurGroupID = PointList(catched_pid).group_id
        End If
        
    ElseIf CurTool = ToolType.MoveElement Then
        CurGroupID = 0
        If catched_pid > 0 Then  'Move Point
            PicPathCls
            DrawGridLines
            DrawGroupExcept PointList(catched_pid).group_id
            
            OpenXORStack
            DrawGroup PointList(catched_pid).group_id, XORColor
            CurGroupID = PointList(catched_pid).group_id
            
        ElseIf catched_sid > 0 Then 'Move Segment
            PicPathCls
            DrawGridLines
            DrawGroupExcept SegmentList(catched_sid).group_id
            
            OpenXORStack
            DrawGroup SegmentList(catched_sid).group_id, XORColor
            CurGroupID = SegmentList(catched_sid).group_id
            
        ElseIf (catched_cid > 0 Or catched_aid > 0) Then 'Catch And Move Arc
            If catched_cid > 0 Then
                id = catched_cid
            Else
                id = catched_aid
            End If
            '
            If ArcList(id).Type = ArcType.RoundedCorner Then
                path_R0 = ArcList(id).a
            
                ux0 = PointList(ArcList(id).point_id).X
                uy0 = PointList(ArcList(id).point_id).Y
                ux1 = ArcList(id).X
                uy1 = ArcList(id).Y
            
                path_D0 = Sqr((ux1 - ux0) ^ 2 + (uy1 - uy0) ^ 2)
            End If
            
            PicPathCls
            DrawGridLines
            DrawGroupExcept ArcList(id).group_id
            
            OpenXORStack
            DrawGroup ArcList(id).group_id, XORColor
            CurGroupID = ArcList(id).group_id
            
        ElseIf catched_spid > 0 Then 'Move SPLine
            PicPathCls
            DrawGridLines
            DrawGroupExcept SPLineList(catched_spid).group_id
            
            OpenXORStack
            DrawGroup SPLineList(catched_spid).group_id, XORColor
            CurGroupID = SPLineList(catched_spid).group_id
            
       End If
    
    ElseIf (CurTool = ToolType.DeleteElement Or CurTool = ToolType.DeleteElement_Point) And catched_pid > 0 Then 'Catch And Delete Point
        For I = 1 To ArcCount
            If ArcList(I).point0_id = catched_pid Or ArcList(I).point1_id = catched_pid Then
                Exit For
            End If
        Next
        If I = ArcCount Then '不删除圆弧的端点
            GoTo Exit_Sub
        End If
        
        If PointList(catched_pid).Type = PointType.ArcPoint Then '不删除圆弧的圆心
            GoTo Exit_Sub
        End If
        
        If PointList(catched_pid).Type = PointType.SPLinePoint Then
            For I = 1 To SPLineCount
                For j = 0 To SPLineList(I).vertex_count - 1
                    If SPLineList(I).vertex_id(j) = catched_pid Then
                        If SPLineList(I).vertex_count <= 3 Then
                            MsgBox "不能删除只有三个点的样条曲线上的点。   ", vbOKOnly + vbExclamation, ""
                            catched_pid = 0
                            catched_spid = 0
                            PicPathCls
                            DrawAll
                            GoTo Exit_Sub
                        End If
                
                        If SPLineList(I).vertex_count = 5 Then
                            If SPLineList(I).vertex_id(0) = SPLineList(I).vertex_id(4) Then
                                MsgBox "不能删除只有四个点的封闭样条曲线上的点。   ", vbOKOnly + vbExclamation, ""
                                catched_pid = 0
                                catched_spid = 0
                                PicPathCls
                                DrawAll
                                GoTo Exit_Sub
                            End If
                        End If
                        
                        For k = j To SPLineList(I).vertex_count - 2
                            SPLineList(I).vertex_id(k) = SPLineList(I).vertex_id(k + 1)
                        Next
                        SPLineList(I).vertex_count = SPLineList(I).vertex_count - 1
                        ReDim Preserve SPLineList(I).vertex_id(SPLineList(I).vertex_count - 1)
                        Exit For
                    End If
                Next
            Next
        End If
        
        DeletePoint catched_pid
        catched_pid = 0
        
        PicPathCls
        DrawAll
        SaveUndo
                
    ElseIf CurTool = ToolType.DeleteElement Then
        If catched_cid > 0 Then
            If ArcList(catched_cid).Type = ArcType.RoundedCorner Then
                DeletePoint ArcList(catched_cid).point0_id
                DeletePoint ArcList(catched_cid).point1_id
                DeletePoint ArcList(catched_cid).pointm_id
        
                DeleteArc catched_cid
                catched_cid = 0
                catched_aid = 0
        
                PicPathCls
                DrawAll
        
                SaveUndo
                GoTo Exit_Sub
            End If
        End If
        
        group_id = 0
        If catched_sid > 0 Then
            'group_id = SegmentList(catched_sid).group_id
            
            DeleteSegment catched_sid
        ElseIf catched_cid > 0 Then
            group_id = ArcList(catched_cid).group_id
        ElseIf catched_aid > 0 Then
            group_id = ArcList(catched_aid).group_id
        ElseIf catched_spid > 0 Then
            group_id = SPLineList(catched_spid).group_id
        End If
        If group_id > 0 Then
            DeleteGroup group_id
        End If
        catched_cid = 0
        catched_aid = 0
        catched_sid = 0
        catched_spid = 0
        
        PicPathCls
        DrawAll
        SaveUndo
        
    ElseIf CurTool = ToolType.CopyElement Then
        CurGroupID = 0
        If catched_sid > 0 Then
            CurGroupID = SegmentList(catched_sid).group_id
        ElseIf catched_cid > 0 Then
            CurGroupID = ArcList(catched_cid).group_id
        ElseIf catched_aid > 0 Then
            CurGroupID = ArcList(catched_aid).group_id
        ElseIf catched_spid > 0 Then
            CurGroupID = SPLineList(catched_spid).group_id
        End If
        If CurGroupID > 0 Then
            CurGroupID = CopyGroup(CurGroupID)
        End If
            
        catched_cid = 0
        catched_aid = 0
        catched_sid = 0
        catched_spid = 0
        
        'PicPathCls
        'DrawAll
        'SaveUndo

    ElseIf CurTool = ToolType.MirrorElement Then
        CurGroupID = 0
        If catched_sid > 0 Then
            CurGroupID = SegmentList(catched_sid).group_id
        ElseIf catched_cid > 0 Then
            CurGroupID = ArcList(catched_cid).group_id
        ElseIf catched_aid > 0 Then
            CurGroupID = ArcList(catched_aid).group_id
        ElseIf catched_spid > 0 Then
            CurGroupID = SPLineList(catched_spid).group_id
        End If
        If CurGroupID > 0 Then
            GetGroupCenter CurGroupID, CurGroupCenterX, CurGroupCenterY
            MirrorGroup CurGroupID, CurGroupCenterX
        End If
            
        catched_cid = 0
        catched_aid = 0
        catched_sid = 0
        catched_spid = 0
        
        PicPathCls
        DrawAll
        SaveUndo

    ElseIf CurTool = ToolType.RotateElement Then
        CurGroupID = 0
        If catched_sid > 0 Then
            CurGroupID = SegmentList(catched_sid).group_id
        ElseIf catched_cid > 0 Then
            CurGroupID = ArcList(catched_cid).group_id
        ElseIf catched_aid > 0 Then
            CurGroupID = ArcList(catched_aid).group_id
        ElseIf catched_spid > 0 Then
            CurGroupID = SPLineList(catched_spid).group_id
        End If
        If CurGroupID > 0 Then
            GetGroupCenter CurGroupID, CurGroupCenterX, CurGroupCenterY
            
            CurPoint = NullPoint
            CurPoint.X = CurGroupCenterX
            CurPoint.Y = CurGroupCenterY
            
            ReDim tv(1) As Title_Value
            tv(0).t = "角度"
            tv(0).v = 0
            ShowEditData "旋转物体", 1, tv, ToolType.RotateElement
            
            PicPathCls
            DrawGridLines
            DrawGroupExcept CurGroupID
            
            OpenXORStack
            DrawGroup CurGroupID, XORColor
            DrawPoint CurPoint, RGB(255, 255, 255)
        End If
        
        catched_cid = 0
        catched_aid = 0
        catched_sid = 0
        catched_spid = 0
        path_a0 = -1
        
    ElseIf CurTool = ToolType.MakeHole1 Or _
           CurTool = ToolType.MakeHole2 Or _
           CurTool = ToolType.MakeHole3 Then
           
        CurPointIndex = 0
        If catched_pid > 0 Then
            CurPointIndex = catched_pid
            
            If CurTool = ToolType.MakeHole1 Then
                I = 1
            ElseIf CurTool = ToolType.MakeHole2 Then
                I = 2
            ElseIf CurTool = ToolType.MakeHole3 Then
                I = 3
            End If
            
            k = PointList(CurPointIndex).HoleType
            If I = k Then
                PointList(CurPointIndex).HoleType = 0
            Else
                PointList(CurPointIndex).HoleType = I
            End If
            
            PicPathCls
            DrawAll
            SaveUndo
        End If
            
    ElseIf CurTool = ToolType.EditElement Then
        CurPointIndex = 0
        CurSegmentIndex = 0
        CurArcIndex = 0

        If catched_pid > 0 Then
            CurPointIndex = catched_pid
            ReDim tv(2) As Title_Value
            tv(0).t = "X"
            tv(0).v = Round(PointList(CurPointIndex).X, NumDigitsAfterDecimal)
            tv(1).t = "Y"
            tv(1).v = Round(PointList(CurPointIndex).Y, NumDigitsAfterDecimal)
            ShowEditData "点编号:" & Trim(str(CurPointIndex)), 2, tv, ToolType.EditElement
            
        ElseIf catched_sid > 0 Then
            CurSegmentIndex = catched_sid
            If SegmentOnBox(catched_sid, ux, uy, ux0, uy0, id0, body_id) = True Then
                ReDim tv(4) As Title_Value
                tv(0).t = "X"
                tv(0).v = ux
                tv(1).t = "Y"
                tv(1).v = uy
                tv(2).t = "宽"
                tv(2).v = ux0
                tv(3).t = "高"
                tv(3).v = uy0
                ShowEditData "矩形 (线号:" & Trim(str(body_id)) & ")", 4, tv, ToolType.EditElement
            End If
        
        ElseIf catched_cid > 0 Or catched_aid > 0 Then
            CurArcIndex = Max(catched_cid, catched_aid)
            
            CurArc = ArcList(CurArcIndex)
        
            If ArcList(CurArcIndex).a <> ArcList(CurArcIndex).b Then 'Ellipse
                ReDim tv(6) As Title_Value
                tv(0).t = "X"
                tv(0).v = Round(CurArc.X, NumDigitsAfterDecimal)
                tv(1).t = "Y"
                tv(1).v = Round(CurArc.Y, NumDigitsAfterDecimal)
                tv(2).t = "A"
                tv(2).v = Round(CurArc.a, NumDigitsAfterDecimal)
                tv(3).t = "B"
                tv(3).v = Round(CurArc.b, NumDigitsAfterDecimal)
                tv(4).t = "始角"
                tv(4).v = Round(CurArc.start_angle * 180 / Pi, NumDigitsAfterDecimal)
                tv(5).t = "夹角"
                tv(5).v = Round((CurArc.end_angle - CurArc.start_angle) * 180 / Pi, NumDigitsAfterDecimal)
                ShowEditData "椭园 (线号:" & Trim(str(CurArc.body_id)) & ")", 6, tv, ToolType.SetEllipse
            
            ElseIf CurArc.Type = ArcType.RoundedCorner Then
                ReDim tv(1) As Title_Value
                tv(0).t = "半径"
                tv(0).v = Round(CurArc.a, NumDigitsAfterDecimal)
                ShowEditData "倒圆角 (线号:" & Trim(str(CurArc.body_id)) & ")", 1, tv, ToolType.RoundCornerByPoint
                
            Else 'If PointList(CurArc.pointm_id).X = CurArc.X And PointList(CurArc.pointm_id).Y = CurArc.Y Then 'CR
                
                ReDim tv(5) As Title_Value
                tv(0).t = "X"
                tv(0).v = Round(CurArc.X, NumDigitsAfterDecimal)
                tv(1).t = "Y"
                tv(1).v = Round(CurArc.Y, NumDigitsAfterDecimal)
                tv(2).t = "半径"
                tv(2).v = Round(CurArc.a, NumDigitsAfterDecimal)
                tv(3).t = "始角"
                tv(3).v = Round(CurArc.start_angle * 180 / Pi, NumDigitsAfterDecimal)
                tv(4).t = "夹角"
                tv(4).v = Round((CurArc.end_angle - CurArc.start_angle) * 180 / Pi, NumDigitsAfterDecimal)
                ShowEditData "园 (线号:" & Trim(str(CurArc.body_id)) & ")", 5, tv, ToolType.SetCircle
                
'            Else '3P
'
'                ReDim tv(6) As Title_Value
'                tv(0).t = "X1"
'                tv(0).v = Round(PointList(CurArc.point0_id).X, NumDigitsAfterDecimal)
'                tv(1).t = "Y1"
'                tv(1).v = Round(PointList(CurArc.point0_id).Y, NumDigitsAfterDecimal)
'                tv(2).t = "X2"
'                tv(2).v = Round(PointList(CurArc.pointm_id).X, NumDigitsAfterDecimal)
'                tv(3).t = "Y2"
'                tv(3).v = Round(PointList(CurArc.pointm_id).Y, NumDigitsAfterDecimal)
'                tv(4).t = "X3"
'                tv(4).v = Round(PointList(CurArc.point1_id).X, NumDigitsAfterDecimal)
'                tv(5).t = "Y3"
'                tv(5).v = Round(PointList(CurArc.point1_id).Y, NumDigitsAfterDecimal)
'                ShowEditData "园3P (线号:" & Trim(Str(CurArc.body_id)) & ")", 6, tv, ToolType.SetCircle_3p
            End If
        End If
        
    ElseIf CurTool = ToolType.BreakSegment Then
        If catched_sid > 0 Then 'Catch And Insert Segment
            InsertSegment catched_sid, ux, uy
            catched_sid = 0
            
        ElseIf catched_spid > 0 Then
            'n = SPLineList(catched_spid).segment_between_points
            
            '提高样条曲线的精度（线段数）
            'SPLineList(catched_spid).segment_between_points = 10 * SPLineList(catched_spid).segment_between_points
            
            'PicPathCls
            'DrawAll
        
            GetSPlinePoint catched_spid, ux, uy, cx, cy, j
            
            AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.SPLinePoint
            PointList(PointCount).body_id = SPLineList(catched_spid).body_id
            PointList(PointCount).group_id = SPLineList(catched_spid).group_id
            
            SPLineList(catched_spid).vertex_count = SPLineList(catched_spid).vertex_count + 1
            ReDim Preserve SPLineList(catched_spid).vertex_id(SPLineList(catched_spid).vertex_count)
            
            k = Int((j - 1) / (SPLineList(catched_spid).segment_between_points + 1)) + 1
            For I = SPLineList(catched_spid).vertex_count To k + 1 Step -1
                SPLineList(catched_spid).vertex_id(I) = SPLineList(catched_spid).vertex_id(I - 1)
            Next
            SPLineList(catched_spid).vertex_id(k) = PointList(PointCount).id
            
            '还原样条曲线的精度（线段数）
            'SPLineList(catched_spid).segment_between_points = n
            
            catched_spid = 0
        End If
        
        PicPathCls
        DrawAll
        
        SaveUndo
        
    ElseIf CurTool = ToolType.BreakSegment_2 Then
        If catched_sid > 0 Then
            ux = (PointList(SegmentList(catched_sid).point0_id).X + PointList(SegmentList(catched_sid).point1_id).X) / 2
            uy = (PointList(SegmentList(catched_sid).point0_id).Y + PointList(SegmentList(catched_sid).point1_id).Y) / 2
            InsertSegment catched_sid, ux, uy
            catched_sid = 0
        
        ElseIf catched_spid > 0 Then
            Dim Pts() As PolygonPoint
            Dim pts1 As PolygonPoint, pts2 As PolygonPoint
            
            SplinePoints SPLineList(catched_spid), Pts(), SPLine_SegmentBetweenPoints
        
            GetSPlinePoint catched_spid, ux, uy, cx, cy, j
            
            AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.SPLinePoint
            PointList(PointCount).body_id = SPLineList(catched_spid).body_id
            PointList(PointCount).group_id = SPLineList(catched_spid).group_id
            
            SPLineList(catched_spid).vertex_count = SPLineList(catched_spid).vertex_count + 1
            ReDim Preserve SPLineList(catched_spid).vertex_id(SPLineList(catched_spid).vertex_count)
            
            k = Int((j - 1) / (SPLineList(catched_spid).segment_between_points + 1)) + 1
            For I = SPLineList(catched_spid).vertex_count To k + 1 Step -1
                SPLineList(catched_spid).vertex_id(I) = SPLineList(catched_spid).vertex_id(I - 1)
            Next
            SPLineList(catched_spid).vertex_id(k) = PointList(PointCount).id
            
            pts1 = Pts((k - 1) * (SPLineList(catched_spid).segment_between_points + 1) + SPLineList(catched_spid).segment_between_points / 2)
            pts2 = Pts((k - 1) * (SPLineList(catched_spid).segment_between_points + 1) + SPLineList(catched_spid).segment_between_points / 2 + 1)
                
            ux0 = pts1.X
            uy0 = pts1.Y
            ux = pts2.X
            uy = pts2.Y
                        
            PointList(PointCount).X = (ux + ux0) / 2
            PointList(PointCount).Y = (uy + uy0) / 2
            
            catched_spid = 0
        End If
        PicPathCls
        DrawAll
        
        SaveUndo
        
    ElseIf CurTool = ToolType.SetCircle Or _
           CurTool = ToolType.SetCircle_3p Or _
           CurTool = ToolType.SetEllipse Then
           
        FraEdit.Visible = False
        
        If CurToolStep = 0 Then
            CurArc = NullArc
            
            If CurTool = ToolType.SetCircle Or CurTool = ToolType.SetEllipse Then
                If CurTool = ToolType.SetCircle Then
                    AddArc ux, uy, LayerZValue(CurLayer), 0, 0, 0, PI2, 0, 0, 0, CurLayer, ArcType.CircleCR
                ElseIf CurTool = ToolType.SetEllipse Then
                    AddArc ux, uy, LayerZValue(CurLayer), 0, 0, 0, PI2, 0, 0, 0, CurLayer, ArcType.Ellipse
                End If
                
                AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
                ArcList(ArcCount).point0_id = PointList(PointCount).id
                PointList(PointCount).arc_id = ArcList(ArcCount).id
                PointList(PointCount).body_id = ArcList(ArcCount).body_id
                PointList(PointCount).group_id = ArcList(ArcCount).group_id
                
                AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
                ArcList(ArcCount).pointm_id = PointList(PointCount).id
                PointList(PointCount).arc_id = ArcList(ArcCount).id
                PointList(PointCount).body_id = ArcList(ArcCount).body_id
                PointList(PointCount).group_id = ArcList(ArcCount).group_id
                
                AddPoint ux, uy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
                ArcList(ArcCount).point1_id = PointList(PointCount).id
                PointList(PointCount).arc_id = ArcList(ArcCount).id
                PointList(PointCount).body_id = ArcList(ArcCount).body_id
                PointList(PointCount).group_id = ArcList(ArcCount).group_id
                
                PointList(ArcList(ArcCount).point0_id).Status = PointStatus.temp
                PointList(ArcList(ArcCount).point1_id).Status = PointStatus.temp
                
                CurArc = ArcList(ArcCount)
                CurToolStep = 1
                DrawArc CurArc
                
                OpenXORStack
            Else ' CurTool = ToolType.SetCircle_3P
                If catched_pid = 0 Then
                    LastPoint.X = ux
                    LastPoint.Y = uy
                    LastPoint.id = 0
                    DrawPoint LastPoint, XORColor
                Else
                    LastPoint.X = PointList(catched_pid).X
                    LastPoint.Y = PointList(catched_pid).Y
                    LastPoint.id = catched_pid
                    catched_pid = 0
                End If
                CurToolStep = 0.1
            End If
            
        ElseIf CurToolStep = 0.1 Then ' Only for CurTool = ToolType.SetCircle_3P
            PopAllXORStack
            CloseXORStack
                    
            If catched_pid = 0 Then
                CurPoint.X = ux
                CurPoint.Y = uy
                CurPoint.id = 0
                DrawPoint CurPoint, XORColor
                
                OpenXORStack
            Else
                CurPoint.X = PointList(catched_pid).X
                CurPoint.Y = PointList(catched_pid).Y
                CurPoint.id = catched_pid
                catched_pid = 0
            End If
            CurToolStep = 0.2
            
        ElseIf CurToolStep = 0.2 Then
            PopAllXORStack
            CloseXORStack
            
DrawPoint CurPoint, RGB(1, 0, 0) '擦掉中间点

            OpenXORStack
                
            ux1 = ux
            uy1 = uy
            
            If catched_pid > 0 Then
                ux1 = PointList(catched_pid).X
                uy1 = PointList(catched_pid).Y
            End If
            
            Ret = GetCircleBy3Points(LastPoint.X, LastPoint.Y, CurPoint.X, CurPoint.Y, ux1, uy1, cx, cy, r, sa, ea)
            If Ret = False Then
                CloseXORStack
                PicPathCls
                DrawAll
                CurToolStep = 0
                GoTo Exit_Sub
            End If
            AddArc cx, cy, LayerZValue(CurLayer), r, r, 0, PI2, 0, 0, 0, CurLayer, ArcType.CircleCR '.Circle3P
            
            If LastPoint.id = 0 Then
                AddPoint LastPoint.X, LastPoint.Y, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
                ArcList(ArcCount).point0_id = PointList(PointCount).id
                PointList(PointCount).arc_id = ArcList(ArcCount).id
                PointList(PointCount).body_id = ArcList(ArcCount).body_id
                PointList(PointCount).group_id = ArcList(ArcCount).group_id
            Else
                ArcList(ArcCount).point0_id = LastPoint.id
                PointList(LastPoint.id).Type = PointType.ArcPoint
                PointList(LastPoint.id).arc_id = ArcList(ArcCount).id
                If PointList(LastPoint.id).body_id = 0 Then
                    PointList(LastPoint.id).body_id = ArcList(ArcCount).body_id
                    PointList(LastPoint.id).group_id = ArcList(ArcCount).group_id
                Else
                    ArcList(ArcCount).body_id = PointList(LastPoint.id).body_id
                    ArcList(ArcCount).group_id = PointList(LastPoint.id).group_id
                    BodyCount = BodyCount - 1
                End If
            End If
            
            AddPoint cx, cy, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
            ArcList(ArcCount).pointm_id = PointList(PointCount).id
            PointList(PointCount).arc_id = ArcList(ArcCount).id
            
            If catched_pid = 0 Then
                AddPoint ux1, uy1, LayerZValue(CurLayer), CurLayer, PointType.ArcPoint
                ArcList(ArcCount).point1_id = PointList(PointCount).id
                PointList(PointCount).arc_id = ArcList(ArcCount).id
                'PointList(PointCount).body_id = ArcList(ArcCount).body_id
            Else
                ArcList(ArcCount).point1_id = catched_pid
                PointList(catched_pid).Type = PointType.ArcPoint
                PointList(catched_pid).arc_id = ArcList(ArcCount).id
                'PointList(catched_pid).body_id = ArcList(ArcCount).body_id
                catched_pid = 0
            End If
            
            ArcList(ArcCount).start_angle = sa
            ArcList(ArcCount).end_angle = ea
                
            CurArc = ArcList(ArcCount)
            CurPoint.X = ux1
            CurPoint.Y = uy1
            CurToolStep = 2
                        
            PopAllXORStack
            DrawArc CurArc, XORColor
            DrawArcEndLine CurArc, XORColor
            
        ElseIf CurToolStep = 1 Then
            If CurTool = ToolType.SetCircle Then
                r = Sqr((ux - CurArc.X) ^ 2 + (uy - CurArc.Y) ^ 2)
                If r > 0 Then
                    'DrawArcStartLine CurArc, xorcolor
                    'DrawArc CurArc, xorcolor, False
                    
                    a = GetArcAngle(CurArc.X, CurArc.Y, ux, uy)
                    
                    CurArc.a = r
                    CurArc.b = r
                    CurArc.start_angle = a
                    CurArc.end_angle = a
                    
                    If catched_pid > 0 Then
                        PointList(CurArc.point0_id).body_id = 0
                        PointList(CurArc.point0_id).group_id = 0
                        
                        PointList(CurArc.point0_id).arc_id = 0
                        PointList(CurArc.point0_id).Status = PointStatus.Deleted
                        
                        CurArc.point0_id = catched_pid
                        If PointList(catched_pid).body_id > 0 Then
                            CurArc.body_id = PointList(catched_pid).body_id
                            CurArc.group_id = PointList(catched_pid).group_id
                        End If
                    Else
                        PointList(CurArc.point0_id).X = ux
                        PointList(CurArc.point0_id).Y = uy
                        PointList(CurArc.point0_id).Status = PointStatus.Normal
                    End If
                    
                    PointList(CurArc.point1_id).X = ux
                    PointList(CurArc.point1_id).Y = uy
                    DrawArc CurArc, XORColor
                    
                    CurToolStep = 2
                End If
                
                
            ElseIf CurTool = ToolType.SetEllipse Then
                r = Abs(ux - CurArc.X)
                r2 = Abs(uy - CurArc.Y)
                
                If r > 0 And r2 > 0 Then
                    DrawArc CurArc, XORColor, False
                    
                    a = GetArcAngle(CurArc.X, CurArc.Y, ux, uy)
                    
                    CurArc.a = r
                    CurArc.b = r2
                    CurArc.start_angle = a
                    CurArc.end_angle = a + PI2
                    
                    PointList(CurArc.point0_id).X = Cos(a) * CurArc.a + CurArc.X
                    PointList(CurArc.point0_id).Y = Sin(a) * CurArc.b + CurArc.Y
                    PointList(CurArc.point1_id).X = Cos(a) * CurArc.a + CurArc.X
                    PointList(CurArc.point1_id).Y = Sin(a) * CurArc.b + CurArc.Y
                    
                    DrawArc CurArc, XORColor, False
                    
                    DrawArcStartLine CurArc, XORColor
                    CurToolStep = 1.5
                End If
            End If
            angle_adjust = 0
            
       ElseIf CurToolStep = 1.5 Then
            PopAllXORStack
            
            a = GetArcAngle(CurArc.X, CurArc.Y, ux, uy)
            If a <> -999999 Then
                CurArc.start_angle = a
                CurArc.end_angle = a
            
                ux0 = Cos(a) * CurArc.a + CurArc.X
                uy0 = Sin(a) * CurArc.b + CurArc.Y
                
                id = CatchPoint(ux0, uy0, utw, NotStartPoint)
                If id > 0 Then
                    PointList(CurArc.point0_id).body_id = 0
                    PointList(CurArc.point0_id).group_id = 0
                    
                    PointList(CurArc.point0_id).arc_id = 0
                    PointList(CurArc.point0_id).Status = PointStatus.Deleted

                    CurArc.point0_id = id
                    If PointList(id).body_id > 0 Then
                        CurArc.body_id = PointList(id).body_id
                        CurArc.group_id = PointList(id).group_id
                    End If
                Else
                    PointList(CurArc.point0_id).Status = PointStatus.Normal
                End If
        
                DrawArc CurArc, XORColor
                DrawArcStartLine CurArc, XORColor
                CurToolStep = 2
            End If
            
       ElseIf CurToolStep = 2 Then
            body_id = PointList(CurArc.point0_id).body_id
            group_id = PointList(CurArc.point0_id).group_id
            
            CurArc.body_id = body_id
            CurArc.group_id = group_id
            
            If PointList(CurArc.point1_id).body_id > 0 Then
                ReplaceBodyID PointList(CurArc.point1_id).body_id, body_id
                ReplaceGroupID PointList(CurArc.point1_id).group_id, group_id
            End If
            PointList(CurArc.point1_id).body_id = body_id
            PointList(CurArc.pointm_id).body_id = body_id
            PointList(CurArc.point1_id).group_id = group_id
            PointList(CurArc.pointm_id).group_id = group_id
            
            PointList(CurArc.point1_id).Status = PointStatus.Normal
            
            '-------------------------------------------------------------------------------
            '夹角超过+/-2PI则设为闭合
            k = 0
            If CurArc.end_angle - CurArc.start_angle >= PI2 Then
                CurArc.end_angle = CurArc.start_angle + PI2
                k = 1
            ElseIf CurArc.end_angle - CurArc.start_angle <= -PI2 Then
                CurArc.end_angle = CurArc.start_angle - PI2
                k = 1
            End If

            'PointList(CurArc.point0_id).X = Cos(CurArc.start_angle) * CurArc.a + CurArc.X
            'PointList(CurArc.point0_id).Y = Sin(CurArc.start_angle) * CurArc.B + CurArc.Y
            If k = 1 Then
                PointList(CurArc.point1_id).X = PointList(CurArc.point0_id).X
                PointList(CurArc.point1_id).Y = PointList(CurArc.point0_id).Y
            Else
                PointList(CurArc.point1_id).X = Cos(CurArc.end_angle) * CurArc.a + CurArc.X
                PointList(CurArc.point1_id).Y = Sin(CurArc.end_angle) * CurArc.b + CurArc.Y
            End If
            '--------------------------------------------------------------------------------
            ArcList(CurArc.id) = CurArc
            
            PopAllXORStack
            CloseXORStack
            DrawArc CurArc
            
            CurToolStep = 0
            
            If CurTool = ToolType.SetCircle Or CurTool = ToolType.SetCircle_3p Then
                
                ReDim tv(5) As Title_Value
                tv(0).t = "X"
                tv(0).v = Round(CurArc.X, NumDigitsAfterDecimal)
                tv(1).t = "Y"
                tv(1).v = Round(CurArc.Y, NumDigitsAfterDecimal)
                tv(2).t = "半径"
                tv(2).v = Round(CurArc.a, NumDigitsAfterDecimal)
                tv(3).t = "始角"
                tv(3).v = Round(CurArc.start_angle * 180 / Pi, NumDigitsAfterDecimal)
                tv(4).t = "夹角"
                tv(4).v = Round((CurArc.end_angle - CurArc.start_angle) * 180 / Pi, NumDigitsAfterDecimal)
                ShowEditData "园CR (线号:" & Trim(str(CurArc.body_id)) & ")", 5, tv, ToolType.SetCircle
                
            'ElseIf CurTool = ToolType.SetCircle_3p Then
            '
            '    ReDim tv(6) As Title_Value
            '    tv(0).t = "X1"
            '    tv(0).v = Round(PointList(CurArc.point0_id).X, NumDigitsAfterDecimal)
            '    tv(1).t = "Y1"
            '    tv(1).v = Round(PointList(CurArc.point0_id).Y, NumDigitsAfterDecimal)
            '    tv(2).t = "X2"
            '    tv(2).v = Round(PointList(CurArc.pointm_id).X, NumDigitsAfterDecimal)
            '    tv(3).t = "Y2"
            '    tv(3).v = Round(PointList(CurArc.pointm_id).Y, NumDigitsAfterDecimal)
            '    tv(4).t = "X3"
            '    tv(4).v = Round(CurPoint.X, NumDigitsAfterDecimal)
            '    tv(5).t = "Y3"
            '    tv(5).v = Round(CurPoint.Y, NumDigitsAfterDecimal)
            '    ShowEditData "园3P (线号:" & Trim(Str(CurArc.body_id)) & ")", 6, tv, ToolType.SetCircle_3p
                
            ElseIf CurTool = ToolType.SetEllipse Then
            
                ReDim tv(6) As Title_Value
                tv(0).t = "X"
                tv(0).v = Round(CurArc.X, NumDigitsAfterDecimal)
                tv(1).t = "Y"
                tv(1).v = Round(CurArc.Y, NumDigitsAfterDecimal)
                tv(2).t = "A"
                tv(2).v = Round(CurArc.a, NumDigitsAfterDecimal)
                tv(3).t = "B"
                tv(3).v = Round(CurArc.b, NumDigitsAfterDecimal)
                tv(4).t = "始角"
                tv(4).v = Round(CurArc.start_angle * 180 / Pi, NumDigitsAfterDecimal)
                tv(5).t = "夹角"
                tv(5).v = Round((CurArc.end_angle - CurArc.start_angle) * 180 / Pi, NumDigitsAfterDecimal)
                ShowEditData "椭园 (线号:" & Trim(str(CurArc.body_id)) & ")", 6, tv, ToolType.SetEllipse
                
            End If
            
            For id = PointCount To 1 Step -1
                If PointList(id).Status = PointStatus.Deleted Then
                    DeletePoint id
                End If
            Next
            
            SaveUndo
        End If
        
    ElseIf (CurTool = ToolType.RoundCornerByPoint Or CurTool = ToolType.RoundCornerByPoint_2) And catched_pid > 0 Then 'round corner
        Ret = RoundCorner(catched_pid, CornerR)
        If CurTool = ToolType.RoundCornerByPoint_2 Then
            ArcList(ArcCount).color = -99999
        End If
        
        PicPathCls
        DrawAll
        
        If Ret = True Then
            SaveUndo
        End If
        
    ElseIf CurTool = ToolType.SetBox Then 'Create a box
        If CurToolStep = 0 Then
            If catched_pid = 0 Then
                LastPoint.X = ux
                LastPoint.Y = uy
                LastPoint.id = 0
            Else
                LastPoint.X = PointList(catched_pid).X
                LastPoint.Y = PointList(catched_pid).Y
                LastPoint.id = catched_pid
                catched_pid = 0
            End If
            CurToolStep = 1
            OpenXORStack
        Else
            CloseXORStack
                    
            k = 0
            If MouseInWindow(PicFrame) = True Then
                ux0 = LastPoint.X
                uy0 = LastPoint.Y
                
                If catched_pid = 0 Then
                    CurPoint.X = ux
                    CurPoint.Y = uy
                    CurPoint.id = 0
                    
                    ux1 = ux
                    uy1 = uy
                    
                    If LastPoint.X <> ux And LastPoint.Y <> uy Then
                        AddBox LastPoint, CurPoint
                        
                        SaveUndo
                        k = 1
                    End If
                Else
                    CurPoint.X = PointList(catched_pid).X
                    CurPoint.Y = PointList(catched_pid).Y
                    CurPoint.id = catched_pid
                                    
                    ux1 = CurPoint.X
                    uy1 = CurPoint.Y
                    
                    If LastPoint.X <> CurPoint.X And LastPoint.Y <> CurPoint.Y Then
                        AddBox LastPoint, CurPoint
                        
                        SaveUndo
                        k = 1
                    End If
                End If
            End If
            
            CurToolStep = 0
            PicPathCls
            DrawAll
            
            If k = 1 Then
                ReDim tv(4) As Title_Value
                tv(0).t = "X"
                tv(0).v = ux0
                tv(1).t = "Y"
                tv(1).v = uy0
                tv(2).t = "宽"
                tv(2).v = ux1 - ux0
                tv(3).t = "高"
                tv(3).v = uy1 - uy0
                ShowEditData "矩形 (线号:" & Trim(str(BodyCount)) & ")", 4, tv, ToolType.SetBox
            End If
            
        End If
        
    ElseIf CurTool = ToolType.ConnectTwoPoints Then
        If CurToolStep = 0 And catched_pid > 0 Then
            
            DrawPoint PointList(catched_pid), 0
            LastPoint = PointList(catched_pid)
            
            catched_pid = 0
            CurToolStep = 1
            OpenXORStack
            
        ElseIf CurToolStep = 1 And catched_pid > 0 Then
            PopAllXORStack
            
            CloseXORStack
        
            DrawPoint PointList(catched_pid), 0
            AddSegment LastPoint.id, catched_pid
            SegmentList(SegmentCount).Type = SegmentType.NormalSegment
            DrawSegment SegmentList(SegmentCount)
        
            catched_pid = 0
            CurToolStep = 0
            
            SaveUndo
        End If
    ElseIf CurTool = ToolType.Reverse Then
        
        EraseDroppingSetting '反向时起始、终止点之间会出现问题，因此先取消
            
        If catched_sid > 0 Then
            If TxtCurTool.Text = "内轮廓" Or TxtCurTool.Text = "Set Innerline" Then
                I = -1
            Else
                I = 1
            End If
            
            j = IsSegmentsClockwise(SegmentList(catched_sid).point0_id)
            If j = 1 And I = -1 Or j = -1 And I = 1 Then
                ReverseDirection catched_sid, 0, True, True
                DirectionChanged = True
                
                PicPathCls
                DrawAll
                
                SaveUndo
            End If
            
        ElseIf catched_cid > 0 Or catched_aid > 0 Then
            'catched_cid and catched_aid are both the same id of a ArcList element
            If catched_cid > 0 Then
                id = catched_cid
            Else
                id = catched_aid
            End If
            ReverseDirection id, 1, True, True
            DirectionChanged = True

            PicPathCls
            DrawAll

            SaveUndo
        ElseIf catched_spid > 0 Then
            ReverseDirection catched_spid, 2, True, True
            DirectionChanged = True
            
            PicPathCls
            DrawAll
            
            SaveUndo
        ElseIf catched_pid > 0 Then
            k = SeekSPlineByPointID(catched_pid)
            If k > 0 Then
                ReverseDirection k, 2, True, True
                DirectionChanged = True
                
                PicPathCls
                DrawAll
                
                SaveUndo
            End If
        End If
        
        For I = 1 To OutputStartPointList.count
            If PointList(OutputStartPointList.leading_point1(I).id).action = StartDropping Then
                OutputStartPointList.point_id(I) = OutputStartPointList.leading_point1(I).id
                CurPoint = OutputStartPointList.leading_point0(I)
                OutputStartPointList.leading_point0(I) = OutputStartPointList.leading_point1(I)
                OutputStartPointList.leading_point1(I) = CurPoint
            End If
        Next
        
    ElseIf CurTool = ToolType.ZoomIn Or CurTool = ToolType.ZoomOut Then
        Dim lvl As Integer, zf(8) As Double, d As Integer
        
        lvl = 7 '8
        
        zf(0) = 1
        zf(1) = 1.25
        zf(2) = 1.5
        zf(3) = 2
        zf(4) = 4
        zf(5) = 8
        zf(6) = 12
        zf(7) = 16
        zf(8) = 20
    
        If CurTool = ToolType.ZoomIn Then
            d = 1
        ElseIf CurTool = ToolType.ZoomOut Then
            d = -1
        End If
        
        For I = 0 To lvl
            If d = 1 And I < lvl Or d = -1 And I > 0 Then
                If ZoomFactor = zf(I) Then
                    DrawCursorReferenceLines -9999, 0, 0
                    
                    If Zoom(I + d) = False Then
                        Exit Sub
                    End If
                    
                    If I + d > 0 Then
                        x1 = zf(I + d) * X / zf(I)
                        y1 = zf(I + d) * Y / zf(I)
                    
                        If x1 <= PicFrame.Width / 2 Then
                            PicPath.left = 0
                        ElseIf PicPath.ScaleWidth - x1 <= PicFrame.Width / 2 Then
                            PicPath.left = PicFrame.Width - PicPath.ScaleWidth
                        Else
                            PicPath.left = PicFrame.Width / 2 - x1
                        End If
                        If PicPath.left > 0 Then
                            PicPath.left = 0
                        End If
                        HScroll1.value = -PicPath.left
                        
                        If (-PicPath.ScaleHeight - y1) <= PicFrame.Height / 2 Then
                            PicPath.top = 0
                        ElseIf y1 <= PicFrame.Height / 2 Then
                            PicPath.top = PicFrame.Height + PicPath.ScaleHeight
                        Else
                            PicPath.top = PicFrame.Height / 2 - (-PicPath.ScaleHeight - y1)
                        End If
                        If PicPath.top > 0 Then
                            PicPath.top = 0
                        End If
                        VScroll1.value = -PicPath.top
                    End If
                    
                    d = 0
                    Exit For
                End If
            End If
        Next
        If d <> 0 Then
            Beep
        End If
        
    ElseIf CurTool = ToolType.StartPoint And catched_pid > 0 Then
    
        If PointList(catched_pid).Type = PointType.SPLinePoint Then
            id = SPLineList(SeekSPlineByPointID(catched_pid)).point0_id '起点
        Else
            id = catched_pid
        End If

        k = 0
        If catched_pid = 1 Or catched_pid = PointCount Or Toolbar1.Buttons(18).value = tbrPressed Then
            k = 1
        Else
            ux = PointList(catched_pid).X
            uy = PointList(catched_pid).Y
            
            'ux0 = PointList(catched_pid - 1).X
            'uy0 = PointList(catched_pid - 1).Y
            'ux1 = PointList(catched_pid + 1).X
            'uy1 = PointList(catched_pid + 1).Y
            
            id0 = 0
            id1 = 0
            For I = 1 To SegmentCount
                If SegmentList(I).point1_id = catched_pid Then
                    ux0 = PointList(SegmentList(I).point0_id).X
                    uy0 = PointList(SegmentList(I).point0_id).Y
                    id0 = 1
                End If
                
                If SegmentList(I).point0_id = catched_pid Then
                    ux1 = PointList(SegmentList(I).point1_id).X
                    uy1 = PointList(SegmentList(I).point1_id).Y
                    id1 = 1
                End If
                
                If id0 = 1 And id1 = 1 Then
                    Exit For
                End If
            Next
            
            If I <= SegmentCount Then
                If Abs(GetAngle(ux0, uy0, ux, uy, ux1, uy1)) >= Device_VertMinAngle Then
                    k = 1
                End If
            End If
        End If
        
        If k = 1 Then
            PopAllXORStack
            CloseXORStack
            
'            EraseDroppingSetting '只设一个起点
    
            SetStartDroppingOnChain id, 0, end_pid
            For I = 1 To OutputStartPointList.count
                If OutputStartPointList.point_id(I) = id Then
                    Exit For
                End If
            Next
            If I > OutputStartPointList.count Then
                OutputStartPointList.count = OutputStartPointList.count + 1
                ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.count)
                ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.count)
                I = OutputStartPointList.count
    
                OutputStartPointList.leading_point1(I) = PointList(end_pid)
            End If
            OutputStartPointList.point_id(I) = id
            OutputStartPointList.leading_point0(I) = PointList(id)
            
            'CalculatePath id
            CalculateAllPath
            
            PicPathCls
            DrawAll
            
            ShowCalculation '1
            
                        
            SaveUndo
        End If
        
    ElseIf CurTool = ToolType.StopPoint And catched_pid > 0 Then
    
        k = 0
        If catched_pid = 1 Or catched_pid = PointCount Or Toolbar1.Buttons(18).value = tbrPressed Then
            k = 1
        Else
            ux = PointList(catched_pid).X
            uy = PointList(catched_pid).Y
            
            'ux0 = PointList(catched_pid - 1).X
            'uy0 = PointList(catched_pid - 1).Y
            'ux1 = PointList(catched_pid + 1).X
            'uy1 = PointList(catched_pid + 1).Y
            
            id0 = 0
            id1 = 0
            For I = 1 To SegmentCount
                If SegmentList(I).point1_id = catched_pid Then
                    ux0 = PointList(SegmentList(I).point0_id).X
                    uy0 = PointList(SegmentList(I).point0_id).Y
                    id0 = 1
                End If
                
                If SegmentList(I).point0_id = catched_pid Then
                    ux1 = PointList(SegmentList(I).point1_id).X
                    uy1 = PointList(SegmentList(I).point1_id).Y
                    id1 = 1
                End If
                
                If id0 = 1 And id1 = 1 Then
                    Exit For
                End If
            Next
            
            If I <= SegmentCount Then
                If Abs(GetAngle(ux0, uy0, ux, uy, ux1, uy1)) >= Device_VertMinAngle Then
                    k = 1
                End If
            End If
        End If
        
        If k = 1 Then
            PopAllXORStack
            CloseXORStack
            
            'k = 0
            'If PointList(catched_pid).Type = PointType.NormalPoint Then
            '    For I = 1 To SegmentCount
            '        If SegmentList(I).point0_id = catched_pid Or SegmentList(I).point1_id = catched_pid Then
            '            Exit For
            '        End If
            '    Next
            '    If I > SegmentCount Then
            '        PointList(catched_pid).action = ActionType.No_Action
            '        k = 1
            '    End If
            'End If
            'If k = 0 Then
            '    SetStopDroppingOnChain catched_pid, 0
            'End If
            
            For I = 1 To OutputStartPointList.count
                If OutputStartPointList.point_id(I) = catched_pid Then
                    SetStopDroppingOnChain catched_pid, 0
                    PointList(catched_pid).action = No_Action
                    
                    For j = I + 1 To OutputStartPointList.count
                        OutputStartPointList.point_id(j - 1) = OutputStartPointList.point_id(j)
                        OutputStartPointList.leading_point0(j - 1) = OutputStartPointList.leading_point0(j)
                        OutputStartPointList.leading_point1(j - 1) = OutputStartPointList.leading_point1(j)
                    Next
                    OutputStartPointList.count = OutputStartPointList.count - 1
                    ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.count)
                    ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.count)
                    ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.count)
                    Exit For
                Else
                    SetStartDroppingOnChain OutputStartPointList.point_id(I), catched_pid, end_pid
                    If catched_pid = end_pid Then
                        SetStopDroppingOnChain catched_pid, 0
                        OutputStartPointList.leading_point1(I) = PointList(end_pid)
                        Exit For
                    End If
                End If
            Next
            
            'CalculatePath OutputStartPointList.point_id(OutputStartPointList.Count)
            CalculateAllPath
            
            PicPathCls
            DrawAll
            
            ShowCalculation
            
            
            
            SaveUndo
        End If
                    
    ElseIf CurTool = ToolType.MakeHole1 Or CurTool = ToolType.MakeHole2 Or CurTool = ToolType.MakeHole3 Then
'        PopAllXORStack
'        CloseXORStack
'
'        If CurToolStep = 0 Then
'            If catched_pid > 0 Then
'                CurPoint = PointList(catched_pid)
'                OutputStartPointList.leading_point0(CurOutputPointIndex).X = CurPoint.X
'                OutputStartPointList.leading_point0(CurOutputPointIndex).Y = CurPoint.Y
'                CurToolStep = 1
'            End If
'        Else
'            OutputStartPointList.leading_point0(CurOutputPointIndex).X = ux
'            OutputStartPointList.leading_point0(CurOutputPointIndex).Y = uy
'            CurToolStep = 0
'        End If
'
'        PicPathCls
'        DrawAll
'        SaveUndo
        
    ElseIf CurTool = ToolType.PieceArray Then
        If catched_groupid > 0 And CurArrayedGroupID = 0 Then
            CurArrayedGroupID = catched_groupid
            PieceArray CurArrayedGroupID, val(TxtEdit(0).Text), val(TxtEdit(1).Text), val(TxtEdit(2).Text), val(TxtEdit(3).Text), val(TxtEdit(4).Text)
            
            CmdEdit.Enabled = True
            PicPathCls
            DrawAll
            SaveUndo
        End If
        
    ElseIf CurTool = ToolType.Unit Then
        If Button = 1 Then
            If catched_groupid > 0 Then
                If CurUnittedGroupID = 0 Then
                    CurUnittedGroupID = catched_groupid
                ElseIf catched_groupid <> CurUnittedGroupID Then
                    ReplaceGroupID catched_groupid, CurUnittedGroupID
                End If
                
                GetBodyList BodyList
                PicPathCls
                DrawAll
                
                SaveUndo
            End If
        'ElseIf Button = 2 Then
        '    CurUnittedGroupID = 0
        End If
        
    ElseIf CurTool = ToolType.Seperate Then
        If catched_groupid > 0 Then
            SeperateBodyFromGroup catched_groupid
            GetBodyList BodyList
            
            PicPathCls
            DrawAll
            SaveUndo
        End If
        
    ElseIf CurTool = ToolType.MeasureDistance Then
        If CurToolStep = 0 Then
            If catched_pid > 0 Then
                LastPoint = PointList(catched_pid)
            Else
                LastPoint.X = ux
                LastPoint.Y = uy
            End If
            CurToolStep = 1
                
        Else
            PopAllXORStack
            CloseXORStack
            OpenXORStack
            
            If catched_pid > 0 Then
                ux = PointList(catched_pid).X
                uy = PointList(catched_pid).Y
            End If
            TxtCurData.BackColor = RGB(255, 255, 180)
            TxtCurData.Text = "L= " & Format(Sqr((ux - LastPoint.X) ^ 2 + (uy - LastPoint.Y) ^ 2), "0.0#mm")
            TxtStatistics.Text = TxtStatistics.Text + vbCrLf + "L= " + Format(Sqr((ux - LastPoint.X) ^ 2 + (uy - LastPoint.Y) ^ 2), "0.0# mm") + vbCrLf
            
            ConvertUserToPath LastPoint.X, LastPoint.Y, x0, y0
            ConvertUserToPath ux, uy, x1, y1
            
            LineOut x0, y0, x1, y1, RGB(0, 0, 255)
            CurToolStep = 0
        End If
        
    ElseIf CurTool = ToolType.MeasureScale Then
        If catched_groupid > 0 Then
            If Shift = 0 Then
                ux_min = 0
                uy_min = 0
                ux_max = 0
                uy_max = 0
            End If
            
            GetGroupScale catched_groupid, ux0, uy0, ux, uy
            
            If ux_min = 0 And uy_min = 0 And ux_max = 0 And uy_max = 0 Then
                ux_min = ux0
                uy_min = uy0
                ux_max = ux
                uy_max = uy
            Else
                ux_min = Min(ux_min, ux0)
                uy_min = Min(uy_min, uy0)
                ux_max = Max(ux_max, ux)
                uy_max = Max(uy_max, uy)
            End If
                
            TxtCurData.BackColor = RGB(255, 255, 180)
            TxtCurData.Text = "W= " & Format(ux_max - ux_min, "#,##0.0#mm") & "  H= " & Format(uy_max - uy_min, "#,##0.0#mm") & "  S= " & Format((ux_max - ux_min) * (uy_max - uy_min), "#,##0.0#mm2")
            TxtStatistics.Text = TxtStatistics.Text + vbCrLf + "W= " & Format(ux_max - ux_min, "#,##0.0#mm") + vbCrLf + "H= " + Format(uy_max - uy_min, "#,##0.0#mm") + vbCrLf + "S= " + Format((ux_max - ux_min) * (uy_max - uy_min), "#,##0.0#mm2") + vbCrLf
    
            ConvertUserToPath ux_min, uy_min, x0, y0
            ConvertUserToPath ux_max, uy_max, x1, y1
            
            OpenXORStack
            
            LineOut x0, y0, x1, y0, RGB(0, 0, 255)
            LineOut x1, y0, x1, y1, RGB(0, 0, 255)
            LineOut x1, y1, x0, y1, RGB(0, 0, 255)
            LineOut x0, y1, x0, y0, RGB(0, 0, 255)
        End If
        
    End If

Exit_Sub:
    path_MouseDnX = X
    path_MouseDnY = Y
    path_x0 = X
    path_y0 = Y
    
    MouseDnUX0 = ux
    MouseDnUY0 = uy
    
    path_ShiftX0 = ShiftX
    path_ShiftY0 = ShiftY
    
    DrawCursorReferenceLines X, Y, 1
End Sub

Private Sub PicPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim x0 As Single, y0 As Single, ux0 As Double, uy0 As Double
    Dim dX As Single, dy As Single, dux As Double, duy As Double
    Dim ux As Double, uy As Double, uz As Double, id As Long, id0 As Long, id1 As Long
    Dim left As Long, top As Long, I As Long, j As Long, Ret As Boolean
    Dim r As Double, r2 As Double, a As Double, da As Double, k As Integer, d As Double
    Dim cx As Double, cy As Double, sa As Double, ea As Double
    Dim group_id As Long
    
    Dim scroll_wheel As Long
    
    On Error Resume Next
    
'Debug.Print PicPath.top

    scroll_wheel = GetScrollMovement(PicPath.hWnd)
    If scroll_wheel <> 0 Then
        'PicPath.AutoRedraw = True
        
        MouseScrollWheel path_x0, path_y0, scroll_wheel

        CurTool = ToolType.MoveCanvas
        'TxtCurTool.Text = "平移"
        TxtCurTool.Text = "Move"
        Exit Sub
    End If
    
    DrawCursorReferenceLines 0, 0, 0
    
    If path_x0 = X And path_y0 = Y Then
        GoTo Exit_Sub
    End If
    
   ConvertPathToUser X, Y, ux, uy
    
    If TxtCurData.BackColor <> RGB(255, 255, 255) Then
        PopAllXORStack
        CloseXORStack
    End If
    TxtCurData.BackColor = RGB(255, 255, 255)
    TxtCurData.Text = " X: " & Format(ux, "0.0") & ",   Y: " & Format(uy, "0.0")
    
    uz = LayerZValue(CurLayer)
    'ShowPosition ux, uy, uz, OnlyControlBar
    
    TmrCheckMouse.Enabled = True
    If MouseInWindow(PicFrame) = False Then
        GoTo Exit_Sub
    End If
        
'PicToolTip.Move PicFrame.left + 30 + X, PicFrame.top + PicPath.Height - Y + 30
'PicPath.Refresh
'PicToolTip.Visible = True

    'x,y must be integer
    'Debug.Print "x, y ="; X; Y
    
    If Button = 2 Then
        ShiftX = path_ShiftX0 + (X - path_MouseDnX)
        ShiftY = path_ShiftY0 + (Y - path_MouseDnY)
        
        PicPathCls
        PicPath.AutoRedraw = False
        DrawAll False
        GoTo Exit_Sub
    End If
    
    If CurTool = ToolType.MoveCanvas Then
        If Button > 0 Then
            ShiftX = path_ShiftX0 + (X - path_MouseDnX)
            ShiftY = path_ShiftY0 + (Y - path_MouseDnY)
            
            PicPathCls
            PicPath.AutoRedraw = False
            DrawAll False
            GoTo Exit_Sub 'Don't Keep path_x0, path_y0
           
        End If
        
    ElseIf CurTool = ToolType.SetSegment Then 'Segment
        If CurToolStep = 1 Then
            PopAllXORStack
            CloseXORStack
        End If
        
        If CurToolStep = 0 Then
            k = CatchPointMode.NotStartPoint
        Else
            k = CatchPointMode.Normal
        End If
        id = CatchPoint(ux, uy, utw, k)
        If id > 0 Then
            If catched_pid > 0 And id <> catched_pid Then
                DrawPoint PointList(catched_pid), NormalColor
            End If
            
            DrawPoint PointList(id), HighLightColor
            catched_pid = id
            
        ElseIf catched_pid > 0 Then
            DrawPoint PointList(catched_pid), NormalColor
            catched_pid = 0
        End If
        
        If CurToolStep = 1 Then
            OpenXORStack
            ConvertUserToPath CurPoint.X, CurPoint.Y, x0, y0
        
            If catched_pid = 0 Then
                'If ChkCatchHVLine.Value = 1 Then
                If Toolbar1.Buttons(21).value = tbrPressed Then
                    If X <> x0 And Y <> y0 Then
                        dX = Abs(X - x0)
                        dy = Abs(Y - y0)
            
                        If dX <= HVTrapWidth Or dy <= HVTrapWidth Then
                            If dX >= dy Then
                                Y = y0
                            Else
                                X = x0
                            End If
                        End If
                    End If
                End If
                LineOut x0, y0, X, Y, XORColor
            Else
                For I = 1 To SegmentCount
                    If SegmentList(I).point1_id = catched_pid Then
                        GoTo Exit_Sub
                    End If
                Next
                For I = 1 To ArcCount
                    If ArcList(I).point1_id = catched_pid Or _
                       (ArcList(I).point0_id = catched_pid And ArcList(I).Type = ArcType.RoundedCorner) Then
                        GoTo Exit_Sub
                    End If
                Next
                For I = 1 To SPLineCount
                    For j = 1 To SPLineList(I).vertex_count - 1 'only poin0_id allowed (j = 0)
                        If SPLineList(I).vertex_id(j) = catched_pid Then
                            GoTo Exit_Sub
                        End If
                    Next
                Next
                LineOut x0, y0, X, Y, XORColor
            End If
            
            ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
        End If
    
    ElseIf CurTool = ToolType.SetSPLine Then 'Spline
        If CurToolStep >= 1 Then
           PopAllXORStack
           CloseXORStack
        End If
        
        If CurToolStep = 0 Then
            k = CatchPointMode.NotStartPoint
        Else
            k = CatchPointMode.Normal
        End If
        id = CatchPoint(ux, uy, utw, k)
        If id > 0 Then
            If catched_pid > 0 And id <> catched_pid Then
                DrawPoint PointList(catched_pid), NormalColor
            End If
            
            DrawPoint PointList(id), HighLightColor
            catched_pid = id
            
        ElseIf catched_pid > 0 Then
            DrawPoint PointList(catched_pid), NormalColor
            catched_pid = 0
        End If
        
        If CurToolStep >= 1 Then
            OpenXORStack
            ConvertUserToPath CurPoint.X, CurPoint.Y, x0, y0
        
            If catched_pid = 0 Then
                'If ChkCatchHVLine.Value = 1 Then
                If Toolbar1.Buttons(21).value = tbrPressed Then
                    If X <> x0 And Y <> y0 Then
                        dX = Abs(X - x0)
                        dy = Abs(Y - y0)
            
                        If dX <= HVTrapWidth Or dy <= HVTrapWidth Then
                            If dX >= dy Then
                                Y = y0
                            Else
                                X = x0
                            End If
                        End If
                    End If
                End If
            Else
                For I = 1 To SegmentCount
                    If SegmentList(I).point1_id = catched_pid Then
                        GoTo Exit_Sub
                    End If
                Next
                For I = 1 To ArcCount
                    If ArcList(I).point1_id = catched_pid Or _
                       (ArcList(I).point0_id = catched_pid And ArcList(I).Type = ArcType.RoundedCorner) Then
                        GoTo Exit_Sub
                    End If
                Next
                For I = 1 To SPLineCount
                    If SPLineList(I).point0_id <> catched_pid Then
                        For j = 0 To SPLineList(I).vertex_count - 1
                            If SPLineList(I).vertex_id(j) = catched_pid Then
                                GoTo Exit_Sub
                            End If
                        Next
                    End If
                Next
                For I = 1 To TempSPline.vertex_count - 2 'i=0: start point, closed
                    If TempSPline.vertex_id(I) = catched_pid Then
                        If I < TempSPline.vertex_count - 2 Then
                            GoTo Exit_Sub
                        Else
                            k = TempSPline.vertex_count
                            TempSPline.vertex_count = k - 1
                            DrawSPLine TempSPline, Not LayerColor(CurLayer)
                            TempSPline.vertex_count = k
                            GoTo Exit_Sub
                        End If
                    End If
                Next
            End If
            LineOut x0, y0, X, Y, XORColor
            
            If CurToolStep > 1 Then
                ConvertPathToUser X, Y, ux, uy
                PointList(0).X = ux
                PointList(0).Y = uy
                
                TempSPline.vertex_id(TempSPline.vertex_count - 1) = id '0
                DrawSPLine TempSPline, Not LayerColor(CurLayer)
            End If
            
            ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
        End If
    
    ElseIf CurTool = ToolType.MakeHole1 Or _
           CurTool = ToolType.MakeHole2 Or _
           CurTool = ToolType.MakeHole3 Then
           
            k = CatchPointMode.NotArcPoint
            id = CatchPoint(ux, uy, utw, k)
            If id > 0 Then
                If catched_pid > 0 And id <> catched_pid Then
                    DrawPoint PointList(catched_pid), NormalColor
                End If
                
                If catched_groupid > 0 Then
                    DrawGroup catched_groupid, NormalColor
                End If
                
                DrawPoint PointList(id), HighLightColor
                
                catched_pid = id
                catched_sid = 0
                catched_cid = 0
                catched_aid = 0
                catched_spid = 0
                catched_groupid = 0
                GoTo Exit_Sub
                
            ElseIf catched_pid > 0 Then
                DrawPoint PointList(catched_pid), NormalColor
                catched_pid = 0
                catched_groupid = 0
            End If
            
    ElseIf CurTool = ToolType.EditElement Or _
           CurTool = ToolType.MoveElement Or _
           CurTool = ToolType.MoveElement_Point Or _
           CurTool = ToolType.DeleteElement Or _
           CurTool = ToolType.BreakSegment Or _
           CurTool = ToolType.BreakSegment_2 Or _
           CurTool = ToolType.Reverse Or _
           CurTool = ToolType.CopyElement Or _
           CurTool = ToolType.RotateElement Or _
           CurTool = ToolType.MirrorElement Or _
           CurTool = ToolType.Unit Or _
           CurTool = ToolType.Seperate Or _
           CurTool = ToolType.MeasureScale Or _
           CurTool = ToolType.PieceArray Then
    
        'ConvertPathToUser x, y, ux, uy

        catched_bodyid = 0
        
        If Button = 0 Then
            If CurTool <> ToolType.BreakSegment And CurTool <> ToolType.BreakSegment_2 And CurTool <> ToolType.DeleteElement And Toolbar1.Buttons(18).value = tbrPressed Then
                k = CatchPointMode.NotArcPoint
                id = CatchPoint(ux, uy, utw, k)
                If id > 0 Then
                    If catched_pid > 0 And id <> catched_pid Then
                        DrawPoint PointList(catched_pid), NormalColor
                    End If
                    
                    If catched_groupid > 0 Then
                        DrawGroup catched_groupid, NormalColor
                    End If
                    
                    DrawPoint PointList(id), HighLightColor
                    
                    catched_pid = id
'                    catched_sid = 0
'                    catched_cid = 0
'                    catched_aid = 0
'                    catched_spid = 0
                    catched_groupid = PointList(id).group_id
                    GoTo Exit_Sub
                    
                ElseIf catched_pid > 0 Then
                    DrawPoint PointList(catched_pid), NormalColor
                    catched_pid = 0
                    catched_groupid = 0
                End If
            End If
            
            If CurTool = ToolType.DeleteElement Then 'delete only one segment
                id = CatchSegment(ux, uy, utw)
                If id > 0 Then
                    If catched_sid > 0 And id <> catched_sid Then
                        DrawSegment SegmentList(catched_sid), NormalColor
                    End If
                    
                    DrawSegment SegmentList(id), HighLightColor
                    catched_sid = id
                    GoTo Exit_Sub
                Else
                    If catched_sid > 0 Then
                        DrawSegment SegmentList(catched_sid), NormalColor
                    End If
                    catched_sid = 0
                    
                End If
            Else
                id = CatchSegment(ux, uy, utw)
                If id > 0 Then
                    
                    'if ---- select group ---
                      group_id = SegmentList(id).group_id
                      If catched_groupid > 0 And catched_groupid <> group_id Then
                          DrawGroup catched_groupid, 0 'NormalColor
                      End If
                    
                      catched_pid = 0
                      catched_cid = 0
                      catched_aid = 0
                      catched_spid = 0
                    
                      DrawGroup group_id, HighLightColor
                    
                      catched_sid = id
                      catched_groupid = group_id
                    'else
                    '
                    '   If catched_sid > 0 And id <> catched_sid Then
                    '       DrawSegment SegmentList(catched_sid), NormalColor
                    '   End If
                   '
                   '    DrawSegment SegmentList(id), HighLightColor
                   '    catched_sid = id
                   '
                    'end if
                    
                    GoTo Exit_Sub
                    
                Else
                    If catched_groupid > 0 Then
                        DrawGroup catched_groupid, 0 'NormalColor
                        catched_groupid = 0
                    End If
                End If
            End If
            
            If CurTool <> ToolType.BreakSegment And CurTool <> ToolType.BreakSegment_2 Then
                id = CatchArcCenter(ux, uy, utw)
                If id > 0 Then
                    group_id = ArcList(id).group_id
                    If ArcList(id).Type <> ArcType.RoundedCorner Then
                        If catched_groupid > 0 And catched_groupid <> group_id Then
                            DrawGroup catched_groupid, NormalColor
                        End If
                        DrawGroup group_id, HighLightColor
                    Else
                    '    增强显示圆心
                    End If
                        
                    catched_pid = 0
                    catched_sid = 0
                    catched_aid = 0
                    catched_spid = 0
                    
                    'If catched_cid > 0 And id <> catched_cid Then
                    '    DrawArc ArcList(catched_cid), NormalColor
                    'End If
                    '
                    'DrawArc ArcList(catched_cid), HighLightColor
                    catched_cid = id
                    catched_groupid = group_id
                    GoTo Exit_Sub
                Else
                    If catched_cid > 0 Then
                        DrawArc ArcList(catched_cid), NormalColor
                    End If
                
                    catched_cid = 0
                End If
                
                id = CatchArc(ux, uy, utw)
                If id > 0 Then
                    group_id = ArcList(id).group_id
                    If catched_groupid > 0 And catched_groupid <> group_id Then
                        DrawGroup catched_groupid, NormalColor
                    End If
                    
                    catched_pid = 0
                    catched_sid = 0
                    catched_cid = 0
                    catched_spid = 0
                    
                    'If catched_cid > 0 And id <> catched_cid Then
                    '    DrawArc ArcList(catched_cid), NormalColor
                    'End If
                    DrawGroup group_id, HighLightColor
                    
                    'DrawArc ArcList(catched_cid), HighLightColor
                    catched_aid = id
                    catched_groupid = group_id
                    GoTo Exit_Sub
                    
                Else
                    If catched_cid > 0 Then
                        DrawArc ArcList(catched_cid), NormalColor
                    End If
                
                    catched_cid = 0
                End If
            End If
            
            id = CatchSPline(ux, uy, utw)
            If id > 0 Then
                group_id = SPLineList(id).group_id
                If catched_groupid > 0 And catched_groupid <> group_id Then
                    DrawGroup catched_groupid, NormalColor
                End If
                
                catched_pid = 0
                catched_sid = 0
                catched_cid = 0
                catched_aid = 0
                
                DrawGroup SPLineList(id).group_id, HighLightColor
                
                catched_spid = id
                catched_groupid = group_id
                GoTo Exit_Sub
                
            Else
                If catched_spid > 0 Then
                    DrawGroup SPLineList(catched_spid).group_id, NormalColor
                End If
                catched_spid = 0
            End If
                        
            '--------------------------------------------------------------
            
        ElseIf CurTool = ToolType.CopyElement Then
            If CurGroupID > 0 Then
                ConvertPathToUser path_x0, path_y0, ux0, uy0
                
                dX = ux - ux0
                dy = uy - uy0
                
                If dX <> 0 Or dy <> 0 Then
                    MoveGroup CurGroupID, dX, dy
                    
                    PopAllXORStack
                    DrawGroup CurGroupID
                End If
                
                ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
            End If
        
        ElseIf CurTool = ToolType.RotateElement Then

            If CurGroupID > 0 Then
                a = GetArcAngle(CurGroupCenterX, CurGroupCenterY, ux, uy)
                If a < 0 Then
                    a = a + PI2
                End If
                
                If path_a0 >= 0 Then
                    RotateGroup CurGroupID, CurGroupCenterX, CurGroupCenterY, a - path_a0
                End If
                path_a0 = a
                
                PopAllXORStack
                DrawGroup CurGroupID
                DrawPoint CurPoint, RGB(255, 255, 255)
                
                TxtCurData.Text = TxtCurData.Text & ",   A:" & str(Round(a * 180 / Pi, NumDigitsAfterDecimal)) & "°"
            End If
            
        ElseIf catched_pid > 0 Then
            If CurTool = ToolType.MoveElement Or CurTool = ToolType.MoveElement_Point Then 'Move
                PopAllXORStack
                If MovePoint(catched_pid, ux, uy) = True Then
                    DrawGroup PointList(catched_pid).group_id, XORColor, True
                Else
                    CloseXORStack
                    PicPathCls
                    DrawAll
                    GoTo Exit_Sub
                End If
                
                ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
            End If
            
        ElseIf catched_sid > 0 Then
            If CurTool = ToolType.MoveElement Then 'Move
                PopAllXORStack
                
                ConvertPathToUser path_x0, path_y0, ux0, uy0
                
                MoveGroup SegmentList(catched_sid).group_id, ux - ux0, uy - uy0
                DrawGroup SegmentList(catched_sid).group_id, XORColor, True
                
                ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
            End If
            
        ElseIf catched_cid > 0 Or catched_aid > 0 Then
            If CurTool = ToolType.MoveElement Then 'Move
            
                PopAllXORStack
                '
                ''catched_cid and catched_aid are both the same id of a ArcList element
                If catched_cid > 0 Then
                    id = catched_cid
                Else
                    id = catched_aid
                End If

                '
                If catched_cid > 0 And ArcList(id).Type = ArcType.RoundedCorner Then
                    id1 = ArcList(id).point_id
                    ux0 = PointList(id1).X
                    uy0 = PointList(id1).Y
                
                    d = Sqr((ux - ux0 - dux) * (ux - ux0 - dux) + (uy - uy0 - duy) * (uy - uy0 - duy))
                
                    If catched_cid > 0 Then
                        RoundCorner id1, path_R0 * (d / path_D0)
                        TxtCurData.Text = TxtCurData.Text & ",   R:" & str(Round(path_R0 * (d / path_D0), NumDigitsAfterDecimal))
                    Else
                        RoundCorner id1, path_R0 * (d / (path_D0 - path_R0))
                        TxtCurData.Text = TxtCurData.Text & ",   R:" & str(Round(path_R0 * (d / (path_D0 - path_R0)), NumDigitsAfterDecimal))
                    End If
                
                Else
                    ConvertPathToUser path_x0, path_y0, ux0, uy0
                    
                    If catched_cid > 0 Then
                        id = catched_cid
                    Else
                        id = catched_aid
                    End If
                    
                    MoveGroup ArcList(id).group_id, ux - ux0, uy - uy0
                    ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
                End If
                
                DrawGroup ArcList(id).group_id, XORColor, True
            End If
        ElseIf catched_spid > 0 Then
            If CurTool = ToolType.MoveElement Then 'Move
            
                PopAllXORStack
                
                ConvertPathToUser path_x0, path_y0, ux0, uy0
                
                MoveGroup SPLineList(catched_spid).group_id, ux - ux0, uy - uy0
                DrawGroup SPLineList(catched_spid).group_id, XORColor, True
                ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0
            End If
        End If
        
'    ElseIf CurTool = ToolType.BreakSegment Or CurTool = ToolType.BreakSegment_2 Then
'
'        'ConvertPathToUser x, y, ux, uy
'
'        If Button = 0 Then
'            id = CatchSegment(ux, uy, (UserMaxX - UserMinX) / 100 / ZoomFactor)
'            If id > 0 Then
'                If catched_sid > 0 And catched_sid <> id Then
'                    SegmentList(catched_sid).color = 0
'                    DrawSegment SegmentList(catched_sid)
'                End If
'
'                SegmentList(id).color = HighLightColor
'                DrawSegment SegmentList(id)
'                catched_sid = id
'
'            ElseIf catched_sid > 0 Then
'                SegmentList(catched_sid).color = 0
'                DrawSegment SegmentList(catched_sid)
'                catched_sid = 0
'
'            End If
'        End If
        
    ElseIf CurTool = ToolType.SetCircle Or _
           CurTool = ToolType.SetCircle_3p Or _
           CurTool = ToolType.SetEllipse Then
           
        If CurToolStep >= 1 Then
            PopAllXORStack
            CloseXORStack
        End If
        
        If CurToolStep < 1 Then
            If CurTool = ToolType.SetCircle_3p Then
                If CurToolStep = 0 Then
                    k = CatchPointMode.NotStartPoint
                ElseIf CurToolStep = 0.1 Then
                    k = CatchPointMode.Alone
                Else
                    k = CatchPointMode.NotEndPoint
                End If
                
                id = CatchPoint(ux, uy, utw, k)
                If id > 0 Then
                    If catched_pid > 0 And id <> catched_pid Then
                        DrawPoint PointList(catched_pid), NormalColor
                    End If
                    
                    DrawPoint PointList(id), HighLightColor
                    catched_pid = id
                    
                ElseIf catched_pid > 0 Then
                    DrawPoint PointList(catched_pid), NormalColor
                    catched_pid = 0
                End If
                
                If CurToolStep = 0.2 Then
                    Ret = GetCircleBy3Points(LastPoint.X, LastPoint.Y, CurPoint.X, CurPoint.Y, ux, uy, cx, cy, r, sa, ea)
                    If Ret = True Then
                        PopAllXORStack
                        CloseXORStack
                        
                        OpenXORStack
                        
                        CurArc.X = cx
                        CurArc.Y = cy
                        CurArc.a = r
                        CurArc.b = r
                        CurArc.ax_angle = 0
                        CurArc.start_angle = sa
                        CurArc.end_angle = ea
                        
                        PointList(CurArc.point0_id).X = Cos(CurArc.start_angle) * r + CurArc.X
                        PointList(CurArc.point0_id).Y = Sin(CurArc.start_angle) * r + CurArc.Y
                        PointList(CurArc.point1_id).X = Cos(CurArc.end_angle) * r + CurArc.X
                        PointList(CurArc.point1_id).Y = Sin(CurArc.end_angle) * r + CurArc.Y
                        
                        DrawArc CurArc, XORColor, False
    
                        TxtCurData.Text = TxtCurData.Text & ",   R:" & str(Round(r, NumDigitsAfterDecimal))
                    End If
                End If
            End If
        
        ElseIf CurToolStep = 1 Then
            OpenXORStack
            
            k = CatchPointMode.NotStartPoint
            id = CatchPoint(ux, uy, utw, k)
            If id > 0 Then
                If catched_pid > 0 And id <> catched_pid Then
                    DrawPoint PointList(catched_pid), NormalColor
                End If
                
                DrawPoint PointList(id), HighLightColor
                catched_pid = id
                
            ElseIf catched_pid > 0 Then
                DrawPoint PointList(catched_pid), NormalColor
                catched_pid = 0
            End If
                
            If CurTool = ToolType.SetCircle Then
                r = Sqr((ux - CurArc.X) ^ 2 + (uy - CurArc.Y) ^ 2)
                
                CurArc.a = r
                CurArc.b = r
                
                a = GetArcAngle(CurArc.X, CurArc.Y, ux, uy)
                If a <> -999999 Then
                    CurArc.ax_angle = 0
                    CurArc.start_angle = a
                    CurArc.end_angle = a + PI2
                End If
                
                PointList(CurArc.point0_id).X = Cos(CurArc.start_angle) * r + CurArc.X
                PointList(CurArc.point0_id).Y = Sin(CurArc.start_angle) * r + CurArc.Y
                PointList(CurArc.point1_id).X = Cos(CurArc.end_angle) * r + CurArc.X
                PointList(CurArc.point1_id).Y = Sin(CurArc.end_angle) * r + CurArc.Y
                
                ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0, 1

           ElseIf CurTool = ToolType.SetEllipse Then
                r = Abs(ux - CurArc.X)
                r2 = Abs(uy - CurArc.Y)
                
                CurArc.a = r
                CurArc.b = r2
                
                PointList(CurArc.point0_id).X = CurArc.a + CurArc.X
                PointList(CurArc.point0_id).Y = CurArc.Y
                PointList(CurArc.point1_id).X = CurArc.a + CurArc.X
                PointList(CurArc.point1_id).Y = CurArc.Y
                
                ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0, 2
            End If
            
            DrawArc CurArc, XORColor, False
            If CurTool = ToolType.SetCircle Then
                DrawArcStartLine CurArc, XORColor
            End If
            
        ElseIf CurToolStep = 1.5 Then
            OpenXORStack
            
            'ConvertPathToUser x, y, ux, uy

            a = GetArcAngle(CurArc.X, CurArc.Y, ux, uy)
            If a <> -999999 Then
                CurArc.ax_angle = 0
                CurArc.start_angle = a
                CurArc.end_angle = a + PI2
                
                PointList(CurArc.point0_id).X = Cos(a) * CurArc.a + CurArc.X
                PointList(CurArc.point0_id).Y = Sin(a) * CurArc.b + CurArc.Y
                PointList(CurArc.point1_id).X = Cos(a) * CurArc.a + CurArc.X
                PointList(CurArc.point1_id).Y = Sin(a) * CurArc.b + CurArc.Y
                
                TxtCurData.Text = "A:" & str(Round(a * 180 / Pi, NumDigitsAfterDecimal)) & "°"
            End If
            DrawArc CurArc, XORColor, False
            DrawArcStartLine CurArc, XORColor
            
        ElseIf CurToolStep = 2 Then
            OpenXORStack
            
            'ConvertPathToUser x, y, ux, uy

            a = GetArcAngle(CurArc.X, CurArc.Y, ux, uy)
            If a <> -999999 Then
                da = a - CurArc.end_angle + angle_adjust
                If da > Pi Then
                    angle_adjust = angle_adjust - PI2
                ElseIf da < -Pi Then
                    angle_adjust = angle_adjust + PI2
                End If

                CurArc.end_angle = a + angle_adjust
            End If
            
            PointList(CurArc.point1_id).X = Cos(CurArc.end_angle) * CurArc.a + CurArc.X
            PointList(CurArc.point1_id).Y = Sin(CurArc.end_angle) * CurArc.b + CurArc.Y
            DrawArc CurArc, XORColor
            DrawArcEndLine CurArc, XORColor
            
            da = (CurArc.end_angle - CurArc.start_angle) * 180 / Pi
            If da > 360 Then
                da = 360
            ElseIf da < -360 Then
                da = -360
            End If
            TxtCurData.Text = "A:" & str(Round(da, NumDigitsAfterDecimal)) & "°"
        End If
        
    ElseIf CurTool = ToolType.RoundCornerByPoint Or CurTool = ToolType.RoundCornerByPoint_2 Then 'Round Corner
        'ConvertPathToUser x, y, ux, uy

        If Button = 0 Then
            id = CatchPoint(ux, uy, utw, CatchPointMode.Normal)
            If id > 0 Then
                If catched_pid > 0 Then
                    PointList(catched_pid).color = 0
                End If
                
                PointList(id).color = HighLightColor
                catched_pid = id
                
                DrawAll
            ElseIf catched_pid > 0 Then
                PointList(catched_pid).color = 0
                catched_pid = 0
                
                DrawAll
            End If
        End If
        
    ElseIf CurTool = ToolType.SetBox Then
        If CurToolStep = 1 Then
            PopAllXORStack
            CloseXORStack
        End If
        
        id = CatchPoint(ux, uy, utw, Alone)
        If id > 0 Then
            If catched_pid > 0 And id <> catched_pid Then
                DrawPoint PointList(catched_pid), NormalColor
            End If
            
            DrawPoint PointList(id), HighLightColor
            catched_pid = id
            
        ElseIf catched_pid > 0 Then
            DrawPoint PointList(catched_pid), NormalColor
            catched_pid = 0
        End If
        
        If CurToolStep = 1 Then
            OpenXORStack
            
            ConvertUserToPath LastPoint.X, LastPoint.Y, x0, y0
            
            LineOut x0, y0, x0, Y, XORColor
            LineOut x0, Y, X, Y, XORColor
            LineOut X, Y, X, y0, XORColor
            LineOut X, y0, x0, y0, XORColor
            
            ShowOperationData ux, uy, MouseDnUX0, MouseDnUY0, 3
        End If
        
    ElseIf CurTool = ToolType.ConnectTwoPoints Then 'Connect two points

        If CurToolStep = 0 Then
            id = CatchPoint(ux, uy, utw, NotStartPoint)
            If id > 0 Then
                If catched_pid > 0 And id <> catched_pid Then
                    DrawPoint PointList(catched_pid), NormalColor
                End If
                
                DrawPoint PointList(id), HighLightColor
                catched_pid = id
                
            ElseIf catched_pid > 0 Then
                DrawPoint PointList(catched_pid), NormalColor
                catched_pid = 0
            End If
        
        ElseIf CurToolStep = 1 Then
            PopAllXORStack
            
            ConvertUserToPath LastPoint.X, LastPoint.Y, x0, y0
            LineOut x0, y0, X, Y, XORColor
            
            id = CatchPoint(ux, uy, utw, NotEndPoint)
            If id > 0 Then
                
                XORStack.Enabled = False
                
                If catched_pid > 0 And id <> catched_pid Then
                    DrawPoint PointList(catched_pid), NormalColor
                End If
                
                DrawPoint PointList(id), HighLightColor
                catched_pid = id
                
                XORStack.Enabled = True
                
            ElseIf catched_pid > 0 Then
                
                XORStack.Enabled = False
                
                DrawPoint PointList(catched_pid), NormalColor
                catched_pid = 0
                
                XORStack.Enabled = True
            End If
        End If
            
    ElseIf CurTool = ToolType.StartPoint Or CurTool = ToolType.StopPoint Or CurTool = ToolType.DeleteElement_Point Then
        'ConvertPathToUser x, y, ux, uy

        If Button = 0 Then
            id = CatchPoint(ux, uy, utw, CatchPointMode.NotArcCenter)
            If id > 0 Then
                If catched_pid > 0 And id <> catched_pid Then
                    DrawPoint PointList(catched_pid), NormalColor
                End If
                
                If catched_sid > 0 Then
                    DrawSegment SegmentList(catched_sid), NormalColor
                    catched_sid = 0
                End If
                
                If catched_cid > 0 Then
                    DrawArc ArcList(catched_cid), NormalColor
                    catched_cid = 0
                End If
                
                DrawPoint PointList(id), HighLightColor
                catched_pid = id
                GoTo Exit_Sub
                
            ElseIf catched_pid > 0 Then
                DrawPoint PointList(catched_pid), NormalColor
                catched_pid = 0
            End If
        End If
            
    ElseIf CurTool = ToolType.MakeHole1 Or CurTool = ToolType.MakeHole2 Or CurTool = ToolType.MakeHole3 Then
        
'        If CurToolStep = 1 Then
'            PopAllXORStack
'            CloseXORStack
'        End If
'
'        If CurToolStep = 0 Then
'            If CurTool = ToolType.SetWayIn Then
'                id = CatchPoint(ux, uy, utw, CatchPointMode.OutputStartPoint)
'            Else
'                id = CatchPoint(ux, uy, utw, CatchPointMode.OutputEndPoint)
'            End If
'
'            If id > 0 Then
'                If catched_pid > 0 And id <> catched_pid Then
'                    DrawPoint PointList(catched_pid), NormalColor
'                End If
'
'                DrawPoint PointList(id), HighLightColor
'                catched_pid = id
'
'            ElseIf catched_pid > 0 Then
'                DrawPoint PointList(catched_pid), NormalColor
'                catched_pid = 0
'            End If
'        End If
'
'        If CurToolStep = 1 Then
'            id = CatchPoint(ux, uy, utw, CatchPointMode.Normal)
'            If id > 0 Then
'                GoTo Exit_Sub
'            End If
'
'            OpenXORStack
'
'            ConvertUserToPath CurPoint.X, CurPoint.Y, x0, y0
'
'            'If ChkCatchHVLine.Value = 1 Then
'            If Toolbar1.Buttons(21).Value = tbrPressed Then
'                If X <> x0 And Y <> y0 Then
'                    dX = Abs(X - x0)
'                    dY = Abs(Y - y0)
'
'                    If dX <= HVTrapWidth Or dY <= HVTrapWidth Then
'                        If dX >= dY Then
'                            Y = y0
'                        Else
'                            X = x0
'                        End If
'                    End If
'                End If
'            End If
'            LineOut x0, y0, X, Y, XORColor
'        End If
'
    ElseIf CurTool = ToolType.MeasureDistance Then
        If CurToolStep = 1 Then
            PopAllXORStack
            CloseXORStack
        End If
        
        k = CatchPointMode.Normal
        id = CatchPoint(ux, uy, utw, k)
        If id > 0 Then
            If catched_pid > 0 And id <> catched_pid Then
                DrawPoint PointList(catched_pid), NormalColor
            End If
            
            DrawPoint PointList(id), HighLightColor
            catched_pid = id
            
        ElseIf catched_pid > 0 Then
            DrawPoint PointList(catched_pid), NormalColor
            catched_pid = 0
        End If
        
        If CurToolStep = 1 Then
            OpenXORStack
            ConvertUserToPath LastPoint.X, LastPoint.Y, x0, y0
        
            If catched_pid = 0 Then
                If Toolbar1.Buttons(21).value = tbrPressed Then
                    If X <> x0 And Y <> y0 Then
                        dX = Abs(X - x0)
                        dy = Abs(Y - y0)
            
                        If dX <= HVTrapWidth Or dy <= HVTrapWidth Then
                            If dX >= dy Then
                                Y = y0
                            Else
                                X = x0
                            End If
                        End If
                    End If
                End If
            End If
            LineOut x0, y0, X, Y, RGB(0, 0, 255)
        End If
    
    End If

    path_x0 = X
    path_y0 = Y

Exit_Sub:
TxtCurData.Refresh

    DrawCursorReferenceLines X, Y, 1
End Sub

Private Sub PicPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ux As Double, uy As Double, ux0 As Double, uy0 As Double
    Dim x0 As Single, y0 As Single
    Dim id0 As Long, id1 As Long
    
    DrawCursorReferenceLines 0, 0, 0
    
    If CurTool = ToolType.MoveElement Or CurTool = ToolType.MoveElement_Point Then
        CloseXORStack
                
        PicPathCls
        DrawAll
            
        If CurGroupID > 0 Then
            SaveUndo
        End If
        
    ElseIf CurTool = ToolType.CopyElement Then
        CloseXORStack
        
        If X = path_MouseDnX And Y = path_MouseDnY And CurGroupID > 0 Then
            MoveGroup CurGroupID, 50, 20
        End If
        
        PicPathCls
        DrawAll
        
        If CurGroupID > 0 Then
            SaveUndo
        End If
        
    ElseIf CurTool = ToolType.RotateElement Then
        CloseXORStack
                    
        If X = path_MouseDnX And Y = path_MouseDnY Then
            FraEdit.Visible = True
        Else
            PicPathCls
            DrawAll
    
            If CurGroupID > 0 Then
                SaveUndo
            End If
        End If
        
    End If
    
    DrawCursorReferenceLines X, Y, 1
End Sub

Private Sub PicPath_Paint()
    DrawAll
End Sub

Public Sub Timer1_Timer()
Dim outval_vertmotor As Integer
Dim inval As Long
Dim Length As Double
If MotionCardOK = True Then
    If CtrlCardType = 4 Then
        
        
        If Device_UseEncoder = True Then
            FeedPulsCount = GetPosEnc(hDmc, FeedAxis)
            Length = FeedPulsCount / Device_EncoderPulsPerMM
        Else
            FeedPulsCount = GetPos(hDmc, FeedAxis)
            Length = FeedPulsCount / Device_PulsPerMM
        End If
        
        Text1.Text = Round(Length, 3)
        Text2.Text = Round(GetPos(hDmc, BendAxis) / Device_PulsPerDegree, 3)
        Text3.Text = Round(GetPos(hDmc, VertAxis) / Device_VertPulsPerDegree, 3)
        Text4.Text = Round(GetPos(hDmc, VertUpDownAxis) / Device_VertUpDownPulsPerMM, 2)
        
        'Text6.Text = Round(GetPos(hDmc, FeedAxis) / Device_PulsPerMM, 2)
        Text5.Text = Round(GetVel(hDmc, BendAxis) / Device_PulsPerDegree, 2)
        Text6.Text = Round(GetVel(hDmc, FeedAxis) / Device_PulsPerMM, 2)
        Text7.Text = Round(GetVel(hDmc, VertAxis) / Device_VertPulsPerDegree, 2)
        Text8.Text = Round(GetVel(0, VertUpDownAxis) / Device_VertUpDownPulsPerMM, 2)
        'Text9.Text = ReadAxisEncodePos_9030(0, FeedAxis) / Device_EncoderPulsPerMM
        
'绘制进料点和切刀点
        
        outval_vertmotor = ReadOutBit(hDmc, VertMotorPort)
        If outval_vertmotor = 1 Then
            LabelVertMotor.BackColor = RGB(0, 255, 0)
        Else
            LabelVertMotor.BackColor = RGB(255, 255, 255)
        End If

        inval = GetHMStatus(hDmc, BendAxis)
        If inval = 1 Then
            IN5.BackColor = RGB(0, 255, 0)
        Else
            IN5.BackColor = RGB(255, 255, 255)
        End If

        inval = GetHMStatus(hDmc, VertAxis)
        If inval = 1 Then
            IN6.BackColor = RGB(0, 255, 0)
        Else
            IN6.BackColor = RGB(255, 255, 255)
        End If

        inval = GetHMStatus(hDmc, VertUpDownAxis)
        If inval = 1 Then
            IN9.BackColor = RGB(0, 255, 0)
        Else
            IN9.BackColor = RGB(255, 255, 255)
        End If
        'Text9.Text = ReadAxisEncodePos_9030(0, FeedAxis)
        
        TextBendLength.Text = CurIndex
        'TextBendLength.Text = VertThreadStep
        CurPauseSwitchVal = ReadInBit(hDmc, PauseSwitch_GALIL)  'Input is 1 when init
        If CurPauseSwitchVal <> PrePauseSwitchVal Then
            If CurPauseSwitchVal = 0 Then
                If PauseRunning = False Then
                    CmdPause_Click
                Else
                    CmdResume_Click
                End If
            End If
            PrePauseSwitchVal = CurPauseSwitchVal
        End If
    
    
    Else
        FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
        Length = FeedPulsCount / Device_PulsPerMM
        If Device_UseEncoder = True Then
            Text1.Text = Round(ReadAxisEncodePos_9030(0, FeedAxis) / Device_EncoderPulsPerMM, 3)
        Else
            Text1.Text = Round(Length, 3)
        End If
        Text2.Text = Round(ReadAxisPos_9030(0, BendAxis) / Device_PulsPerDegree, 3)
        Text3.Text = Round(ReadAxisPos_9030(0, VertAxis) / Device_VertPulsPerDegree, 3)
        Text4.Text = Round(ReadAxisPos_9030(0, VertUpDownAxis) / Device_VertUpDownPulsPerMM, 2)
        Text5.Text = Round(ReadAxisVel_9030(0, BendAxis) / Device_PulsPerDegree, 2)
        Text6.Text = Round(ReadAxisVel_9030(0, FeedAxis) / Device_PulsPerMM, 2)
        Text7.Text = Round(ReadAxisVel_9030(0, VertAxis) / Device_VertPulsPerDegree, 2)
        Text8.Text = Round(ReadAxisVel_9030(0, VertUpDownAxis) / Device_VertUpDownPulsPerMM, 2)
        'Text9.Text = ReadAxisEncodePos_9030(0, FeedAxis) / Device_EncoderPulsPerMM
        
'绘制进料点和切刀点
        
        outval_vertmotor = ReadOsBit_9030(0, VertMotorPort + 1)
        If outval_vertmotor = 1 Then
            LabelVertMotor.BackColor = RGB(0, 255, 0)
        Else
            LabelVertMotor.BackColor = RGB(255, 255, 255)
        End If
        
        inval = ReadIOBit_9030(0, 5)
        If inval <> 0 Then
            IN5.BackColor = RGB(0, 255, 0)
        Else
            IN5.BackColor = RGB(255, 255, 255)
        End If
        
        inval = ReadIOBit_9030(0, 6)
        If inval <> 0 Then
            IN6.BackColor = RGB(0, 255, 0)
        Else
            IN6.BackColor = RGB(255, 255, 255)
        End If
        
        inval = ReadIOBit_9030(0, 9)
        If inval <> 0 Then
            IN9.BackColor = RGB(0, 255, 0)
        Else
            IN9.BackColor = RGB(255, 255, 255)
        End If
        
        CurPauseSwitchVal = ReadIOBit_9030(0, PauseSwitch)
        If CurPauseSwitchVal <> PrePauseSwitchVal Then
            If CurPauseSwitchVal = 1 Then
                If PauseRunning = False Then
                    CmdPause_Click
                Else
                    CmdResume_Click
                End If
            End If
            PrePauseSwitchVal = CurPauseSwitchVal
        End If
        
        'Text9.Text = ReadAxisEncodePos_9030(0, FeedAxis)
        TextBendLength.Text = CurIndex
        'TextBendLength.Text = VertThreadStep
    End If
End If
End Sub

Public Sub TmrBend_Timer()
    Dim nLogPos As Long                   '逻辑位置
    Dim nActPos As Long                   '实际位置
    Dim nSpeed As Long                    '运行速度
        
    Get_CurrentInf BendAxis, nLogPos, nActPos, nSpeed
    'StatusBar1.Panels.Item(5).Text = "Bend"
    StatusBar1.Panels.Item(7).Text = "Pos:" + str(nLogPos) + " /" + str(Round(nLogPos / Device_PulsPerDegree, 2)) + " deg"
    'StatusBar1.Panels.Item(7).Text = "Speed:" + Str(nSpeed)
    
    'ProgressBar1.value = Round(nLogPos / Device_PulsPerDegree, 0) + 180
End Sub

'Private Sub ReadData_Click()
'    Dim end_pid As Long
'
'    Device_ReadDeviceData
'    ConvertPulsToUser Device_CurPulsPos(1), Device_CurPulsPos(2), Device_CurPulsPos(3), CurX, CurY, CurZ
'
'    ShowCurHeadPos 2
'
'    AddPoint CurX, CurY, LayerZValue(CurLayer), CurLayer, PointType.NormalPoint
'    If Device_Mode = 1 Then
'        ' Add dropping point directly
'        '-------------------------------------------------------------------------------------
'        SetStartDroppingOnChain PointCount, 0, end_pid
'
'        OutputStartPointList.Count = OutputStartPointList.Count + 1
'        ReDim Preserve OutputStartPointList.point_id(OutputStartPointList.Count)
'        ReDim Preserve OutputStartPointList.leading_point0(OutputStartPointList.Count)
'        ReDim Preserve OutputStartPointList.leading_point1(OutputStartPointList.Count)
'
'        OutputStartPointList.point_id(OutputStartPointList.Count) = PointCount
'        OutputStartPointList.leading_point0(OutputStartPointList.Count) = PointList(PointCount)
'        OutputStartPointList.leading_point1(OutputStartPointList.Count) = PointList(end_pid)
'        '---------------------------------------------------------------------------------------
'    End If
'
'    DrawPoint PointList(PointCount)
'
'    SaveUndo
'    ShowCurHeadPos
'End Sub


'Public Sub CmdReset_Click()
'    Dim obj As Object
'    Dim I As Integer, PulsX As Long, PulsY As Long, PulsZ As Long
'
'    On Error Resume Next
'
'    DeviceMotion = True
'
'    For Each obj In FrmMain
'        obj.Enabled = False
'    Next
'    CmdStop.Enabled = True
'
'    TmrShowCurHeadPos.Enabled = True
'    TmrDevicePortChecking.Enabled = True
'
'    Device_ReadDeviceData '确保读取了当前位置数据，因为 TmrShowCurHeadPos 可能尚未触发'
'
'    'Device_FastMoveToPosXYZ 0, 0, 0
'
'    ConvertUserToPuls UserOrgX, UserOrgY, UserOrgZ, PulsX, PulsY, PulsZ
'    Device_FastMoveToPosXYZ UserOrgX, UserOrgY, UserOrgZ
'    TmrShowCurHeadPos.Enabled = False
'
'    For I = 1 To 100
'        Sleep 50
'        TmrShowCurHeadPos_Timer
'        'If Device_CurPulsPos(1) = 0 And Device_CurPulsPos(2) = 0 And Device_CurPulsPos(3) = 0 Then
'        If Device_CurPulsPos(1) = PulsX And Device_CurPulsPos(2) = PulsY And Device_CurPulsPos(3) = PulsZ Then
'            Exit For
'        End If
'        DoEvents
'    Next
'
'    ShowCurHeadPos 0
 '
'    For Each obj In FrmMain
'        If Not TypeOf obj Is Timer Then
'            obj.Enabled = True
'        End If
'    Next
'
'    DeviceMotion = False
'End Sub

Private Sub TmrCheckMouse_Timer()
    Dim x0 As Single, y0 As Single
    
    If MouseInWindow(PicFrame) = False Then
        TmrCheckMouse.Enabled = False
        
        DrawCursorReferenceLines 0, 0, 0

        PopAllXORStack
        CloseXORStack
        
        DrawCursorReferenceLines -1, -1, 1
        
        PicToolTip.Visible = False
    End If
End Sub

Private Sub TmrDevicePortChecking_Timer()
    Dim Sp As Long, pos As Long, I As Integer, b As Long
    
'    On Error Resume Next
'
'    get_speed 0, 1, Sp
'    StatusBar1.Panels(6).Text = "1 轴速度 p/s:" & Str(Sp * Ratio(1)) & ",  mm/s:" & Format(100 * Sp * Ratio(1) / PP100MM, "0.0")
'
'    get_command_pos 0, 1, pos
'    StatusBar1.Panels(7).Text = "1 轴位置 p:" & Str(pos) & ",  mm:" & Format(100 * pos / PP100MM, "0.0#")
'
'    get_speed 0, 2, Sp
'    StatusBar1.Panels(8).Text = "2 轴速度 p/s:" & Str(Sp * Ratio(2)) & ",  d/s:" & Format(Sp * Ratio(2) / PPD, "0.0")
'
'    get_command_pos 0, 2, pos
'    StatusBar1.Panels(9).Text = "2 轴位置 p:" & Str(pos) & ",  d:" & Format(pos / PPD, "0.0##")
'
'    get_speed 0, 3, Sp
'    StatusBar1.Panels(10).Text = "3 轴速度 p/s:" & Str(Sp * Ratio(3))
'
'    get_command_pos 0, 3, pos
'    StatusBar1.Panels(11).Text = "3 轴位置 p:" & Str(pos)
'
'    StatusBar1.Panels(12).Text = "进刀长度:" & Format(PathOutLength, " 0.0")
'    StatusBar1.Panels(13).Text = "总长度:" & Format(TotalPathOutLength, " 0.0")
'
'    If OutputStatus(2) = 1 Then
'        LblHole1.BackColor = RGB(255, 0, 0)
'    Else
'        LblHole1.BackColor = FrmMain.BackColor
'    End If
'
'    If OutputStatus(3) = 1 Then
'        LblHole2.BackColor = RGB(255, 0, 0)
'    Else
'        LblHole2.BackColor = FrmMain.BackColor
'    End If
'
'    If OutputStatus(4) = 1 Then
'        LblHole3.BackColor = RGB(255, 0, 0)
'    Else
'        LblHole3.BackColor = FrmMain.BackColor
'    End If
    
    For I = 0 To 31
        'b = read_bit(0, i)
        If I = 0 Then
            b = read_bit(0, 0)
        ElseIf I = 1 Then
            b = read_bit(0, 1)
        ElseIf I = 2 Then
            b = read_bit(0, 2)
        ElseIf I = 3 Then
            b = read_bit(0, 3)
        ElseIf I = 4 Then
            b = read_bit(0, 4)
        ElseIf I = 5 Then
            b = read_bit(0, 5)
        ElseIf I = 6 Then
            b = read_bit(0, 6)
        ElseIf I = 7 Then
            b = read_bit(0, 7)
        ElseIf I = 8 Then
            b = read_bit(0, 8)
        ElseIf I = 9 Then
            b = read_bit(0, 9)
        ElseIf I = 10 Then
            b = read_bit(0, 10)
        ElseIf I = 11 Then
            b = read_bit(0, 11)
        ElseIf I = 12 Then
            b = read_bit(0, 12)
        ElseIf I = 13 Then
            b = read_bit(0, 13)
        ElseIf I = 14 Then
            b = read_bit(0, 14)
        ElseIf I = 15 Then
            b = read_bit(0, 15)
        ElseIf I = 16 Then
            b = read_bit(0, 16)
        ElseIf I = 17 Then
            b = read_bit(0, 17)
        ElseIf I = 18 Then
            b = read_bit(0, 18)
        ElseIf I = 19 Then
            b = read_bit(0, 19)
        ElseIf I = 20 Then
            b = read_bit(0, 20)
        Else
            b = read_bit(0, I)
        End If
        'If B = 0 Then
        '    LblIn(I).BackColor = RGB(255, 0, 0)
        'Else
        '    LblIn(I).BackColor = RGB(0, 255, 0)
        'End If
    
        If I = 23 Then
            If b = 0 Then
                LblVertHighSensor.BackColor = RGB(255, 0, 0)
            Else
                LblVertHighSensor.BackColor = RGB(0, 255, 0)
            End If

        ElseIf I = 17 Then
            If b = 0 Then
                LblVertLowSensor.BackColor = RGB(255, 0, 0)
            Else
                LblVertLowSensor.BackColor = RGB(0, 255, 0)
            End If

'        ElseIf i = 22 Then
'            If b = 0 Then
'                LblVertOrgSensor.BackColor = RGB(255, 0, 0)
'            Else
'                LblVertOrgSensor.BackColor = RGB(0, 255, 0)
'            End If

'        ElseIf I = 10 Then
'            If B = 0 Then
'                LblElevatorHighSensor.BackColor = RGB(255, 0, 0)
'            Else
'                LblElevatorHighSensor.BackColor = RGB(0, 255, 0)
'            End If

'        ElseIf I = 11 Then
'            If B = 0 Then
'                LblElevatorLowSensor.BackColor = RGB(255, 0, 0)
'            Else
'                LblElevatorLowSensor.BackColor = RGB(0, 255, 0)
'            End If

        End If
    
'        If I = 0 Then 'XLMT+
'            If b = 0 Then
'                LblXLMT_P.BackColor = RGB(255, 0, 0)
'            Else
'                LblXLMT_P.BackColor = RGB(0, 255, 0)
'            End If
'
'        ElseIf I = 1 Then 'XLMT-
'            If b = 0 Then
'                LblXLMT_M.BackColor = RGB(255, 0, 0)
'            Else
'                LblXLMT_M.BackColor = RGB(0, 255, 0)
'            End If
'
'        ElseIf I = 8 Then 'YLMT+
'            If b = 0 Then
'                LblYLMT_P.BackColor = RGB(255, 0, 0)
'            Else
'                LblYLMT_P.BackColor = RGB(0, 255, 0)
'            End If
'
'        ElseIf I = 9 Then 'YLMT-
'            If b = 0 Then
'                LblYLMT_M.BackColor = RGB(255, 0, 0)
'            Else
'                LblYLMT_M.BackColor = RGB(0, 255, 0)
'            End If
'
'        ElseIf I = 18 Then 'ZSTOP0
'            If b = 0 Then
'                LblZLMT_P.BackColor = RGB(255, 0, 0)
'            Else
'                LblZLMT_P.BackColor = RGB(0, 255, 0)
'            End If
'
'        ElseIf I = KnifeTestSensorPort1 Then '
'            If b = 0 Then
'                LblKnifeTest1.BackColor = RGB(255, 0, 0)
'            Else
'                LblKnifeTest1.BackColor = RGB(0, 255, 0)
'            End If
'
'        ElseIf I = KnifeTestSensorPort2 Then '
'            If b = 0 Then
'                LblKnifeTest2.BackColor = RGB(255, 0, 0)
'            Else
'                LblKnifeTest2.BackColor = RGB(0, 255, 0)
'            End If
'        End If
    Next
    
    '--------------------------------------------------------------------------------------------------------
'    Dim h As Long, m As Long, s As Long, t As Double
'
'    If DeviceMotion = True Then
'        t = TimeDiff(Timer, RunningStartTime)
'
'        h = Int(t / 3600)
'        m = Int((t - h * 3600) / 60)
'        s = Int(t - h * 3600 - m * 60)
'
'        StatusBar1.Panels(15).Text = "运行时间:" & Format(h, "00") & ":" & Format(m, "00") & ":" & Format(s, "00")
'    End If
End Sub

 Sub ShowFeedPos()
    Dim nLogPos As Long                   '逻辑位置
    Dim nActPos As Long                   '实际位置
    Dim nSpeed As Long                    '运行速度
    
    If CtrlCardType = 0 Then
        Get_CurrentInf FeedAxis, nLogPos, nActPos, nSpeed
    Else
        nLogPos = ReadAxisPos_9030(0, FeedAxis)
        nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
        nSpeed = ReadAxisVel_9030(0, FeedAxis)
    End If
    
    StatusBar1.Panels.Item(2).Text = "Pos:" + str(nLogPos) + " /" + str(Round(nLogPos / Device_PulsPerMM, 2)) + " mm"
    StatusBar1.Panels.Item(3).Text = "EncPos:" + str(nActPos) + " /" + str(Round(nActPos / Device_EncoderPulsPerMM, 2)) + " mm"
    StatusBar1.Panels.Item(4).Text = "Speed:" + str(nSpeed)
End Sub

Sub ShowFeedMarkPoint(Optional Add_Offset As Boolean = True)
    Dim feed_puls As Long, ux As Double, uy As Double, X As Single, Y As Single
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint FeedMark_x0, FeedMark_y0
    FrmMain.PicPath.DrawMode = 13

    GetPathXYByFeedPuls FeedPulsCount + IIf(Add_Offset, Device_FeedOffset, 0), ux, uy
    
    ConvertUserToPath ux, uy, X, Y
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint X, Y
    FeedMark_x0 = X
    FeedMark_y0 = Y
    FrmMain.PicPath.DrawMode = 13
End Sub

Sub ShowVertMarkPoint(Optional Add_Offset As Boolean = True)
    Dim feed_puls As Long, ux As Double, uy As Double, X As Single, Y As Single
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint VertMark_x0, VertMark_y0, RGB(255, 0, 255)
    FrmMain.PicPath.DrawMode = 13

    GetPathXYByVertPuls FeedPulsCount + Device_HeadDistance * FeedPulsPerMM + IIf(Add_Offset, Device_FeedOffset, 0), ux, uy
    ConvertUserToPath ux, uy, X, Y
    
    FrmMain.PicPath.DrawMode = 7
    MarkPoint X, Y, RGB(255, 0, 255)
    VertMark_x0 = X
    VertMark_y0 = Y
    FrmMain.PicPath.DrawMode = 13
End Sub

Private Sub TmrFace_Timer()
    PicFace.Visible = False
    TmrFace.Enabled = False
End Sub

Private Sub TmrFeedV3Thread_Timer()
    Dim Status As Long
    
    get_status 0, FeedAxis, Status
    If Status = 0 Then
        TmrFeedV3Thread.Enabled = False
        FeedV
    End If
End Sub

Private Sub TmrGetCurRunState_Timer()

    Dim Length As Double
    Dim AngleToGo As Double
    If TotalPathOutLength > 0 Then
        'FeedPulsCount = ReadAxisPos_9030(0, FeedAxis)
        Length = FeedPulsCount / Device_PulsPerMM
        ShowFeedMarkPoint
        ShowVertMarkPoint
        If CurIndex >= PathOutputPointCount - 1 Then
            Exit Sub
        End If
        If Length > PathOutputPoint(CurIndex).LengthFromStart Then
    
            CurIndex = CurIndex + 1
        End If
        If Abs(PathOutputPoint(CurIndex).Radius3P) > Device_BeatMaxRadius And _
            Abs(PathOutputPoint(CurIndex + 1).Radius3P) > Device_BeatMaxRadius And _
            Abs(PathOutputPoint(CurIndex).AngleToNext) > 0.5 Then
            SetAxisVel_9030 0, FeedAxis, Device_FeedSpeed * 0.5
            StartAxis_9030 0, FeedAxis
            
'            If PathOutputPoint(CurIndex).AngleToNext > 0 Then
'                AngleToGo = PathOutputPoint(CurIndex).AngleToNext + Device_EmptyDegree
'            Else
'                AngleToGo = PathOutputPoint(CurIndex).AngleToNext - Device_EmptyDegree2
'            End If
'            SetAxisVel_9030 0, BendAxis, Device_BendSpeed
'            SetAxisPos_9030 0, BendAxis, AngleToGo * Device_PulsPerDegree
'            StartAxis_9030 0, BendAxis

'            If CurIndex = 40 Then
'                Wait 0.2
'            End If

            BendAngleByRadius -VTDir * PathOutputPoint(CurIndex).Radius3P, False
        Else
            SetAxisVel_9030 0, FeedAxis, Device_FeedSpeed
            StartAxis_9030 0, FeedAxis
        End If
    
    End If
End Sub

'Private Sub TmrKey_Timer()
'    Dim k As Integer, t As Variant
'    Static dX As Double, dY As Double, d As Double, t0 As Variant
'
'    If t0 = 0 Then
'        t0 = Timer
'    End If
'
'    t = Timer
'    If t - t0 < 2 Then
'        d = 0.001
'    ElseIf t - t0 < 4 Then
'        d = 0.005
'    ElseIf t - t0 < 6 Then
'        d = 0.01
'    Else
'        d = 0.05
'    End If
'
'    If Device_CoordinateMode = 0 Then
'        If KeyDown = 1 Then
'            dX = 0
'            dY = dY + d
'        ElseIf KeyDown = 2 Then
'            dX = 0
'            dY = dY - d
'        ElseIf KeyDown = 3 Then
'            dX = dX - d
'            dY = 0
'        ElseIf KeyDown = 4 Then
'            dX = dX + d
'            dY = 0
'        Else
'            dX = 0
'            dY = 0
'            t0 = 0
'            TmrKey.Enabled = False
'            Exit Sub
'        End If
'    Else
'        If KeyDown = 4 Then
'            dX = 0
'            dY = dY + d
'        ElseIf KeyDown = 3 Then
'            dX = 0
'            dY = dY - d
'        ElseIf KeyDown = 1 Then
'            dX = dX - d
'            dY = 0
'        ElseIf KeyDown = 2 Then
'            dX = dX + d
'            dY = 0
'        Else
'            dX = 0
'            dY = 0
'            t0 = 0
'            TmrKey.Enabled = False
'            Exit Sub
'        End If
'    End If
'
'    Device_ReadDeviceData
'    ConvertPulsToUser Device_CurPulsPos(1), Device_CurPulsPos(2), Device_CurPulsPos(3), CurX, CurY, CurZ
'
'    CurX = CurX + dX
'    CurY = CurY + dY
'
'    If CurX < 0 Then
'        CurX = 0
'    ElseIf CurX > ViewMaxX Then
'        CurX = ViewMaxX
'    End If
'
'    If CurY < 0 Then
'        CurY = 0
'    ElseIf CurY > ViewMaxY Then
'        CurY = ViewMaxY
'    End If
'
'    Device_ConMoveToPosXYZ CurX, CurY, CurZ, False, False
'
'    Device_ReadDeviceData
'    ConvertPulsToUser Device_CurPulsPos(1), Device_CurPulsPos(2), Device_CurPulsPos(3), CurX, CurY, CurZ
'
'    '对脉冲与毫米之间的相互转换误差引起的出界进行校正
'    k = 0
'    If CurX < 0 Then
'        CurX = 0
'        k = 1
'    ElseIf CurX > ViewMaxX Then
'        CurX = ViewMaxX
'        k = 1
'    End If
'
'    If CurY < 0 Then
'        CurY = 0
'        k = 1
'    ElseIf CurY > ViewMaxY Then
'        CurY = ViewMaxY
'        k = 1
'    End If
'
'    If k = 1 Then
'        Device_ConMoveToPosXYZ CurX, CurY, CurZ, False, False
'    End If
'
'    ShowPosition CurX, CurY, CurZ, BothStatusAndControlBar
'    ShowCurHeadPos
'End Sub

'Private Sub TmrPortActionThread_Timer()
'    Select Case TmrPortActionThread.Tag
'        Case "1"
'            CmdRun_Click
'        'Case "2"
'        '    CmdResetOrg_Click
'        'Case "3"
'        '    CmdPause_Click
'    End Select
'
'    TmrPortActionThread.Enabled = False
'    TmrPortActionThread.Tag = ""
'End Sub

Public Sub TmrReadDevicePos_Timer()
    Dim X As Single, Y As Single, x0 As Single, y0 As Single, dw As Integer, m As Long
    Dim ux As Double, uy As Double, uz As Double
    
    Static x00 As Single, y00 As Single
        
    x0 = PicPath.CurrentX
    y0 = PicPath.CurrentY
    
    'Device_ReadDeviceData
    'ConvertPulsToUser Device_CurPulsPos(1), Device_CurPulsPos(2), Device_CurPulsPos(3), ux, uy, uz
    
    'ShowPosition ux, uy, uz, OnlyControlBar
    
    'ConvertPulsToPath Device_CurPulsPos(1), Device_CurPulsPos(2), X, Y
    
    PicPath.DrawMode = 13 'Copy Pen
    m = ColorMode1 Mod 2
    If ColorMode1 > 0 Then
        dw = PicPath.DrawWidth
        PicPath.DrawWidth = dw + 2
        PicPath.Line -(X, Y), IIf(ColorMode = 0, RGB(255, 255, 255), RGB(0, 0, 0)) 'erase background
        
        PicPath.DrawWidth = dw
        PicPath.Line (x00, y00)-(x0, y0), IIf(ColorMode = 0, IIf(m = 0, RGB(0, 0, 0), RGB(0, 128, 255)), IIf(m = 0, RGB(255, 255, 255), RGB(255, 128, 0)))
    End If
    PicPath.Line (x0, y0)-(X, Y), IIf(ColorMode = 0, IIf(m = 0, RGB(0, 0, 0), RGB(0, 128, 255)), IIf(m = 0, RGB(255, 255, 255), RGB(255, 128, 0)))
    
    x00 = x0
    y00 = y0
End Sub



Private Sub TmrVertThread_Timer()
    Dim Status As Long
    
    If VertThreadStep = 100 Then
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertMotorPort, 0
        write_bit 0, VertMotorPort, 0
        
        LblVertMotorMode.BackColor = RGB(0, 255, 0)
        'If CtrlCardType = 0 Then
            get_status 0, VertUpDownAxis, Status
        'Else
        '    status = ReadAxisState_9030(0, VertUpDownAxis)
        'End If
        If Status = 0 Then
            'VertAngle VertThreadAngle, False
            VertThreadStep = 101
        End If
           
    ElseIf VertThreadStep = 101 Then
        'If CtrlCardType = 0 Then
            get_status 0, VertUpDownAxis, Status
        'Else
        '    status = ReadAxisState_9030(0, VertUpDownAxis)
        'End If
        If Status = 0 Then
            If FeedIntoVertMotorZone = True Then
                'If CtrlCardType = 0 Then
                    write_bit 0, VertMotorPort, 1
                    write_bit 0, VertMotorPort, 1
                    write_bit 0, VertMotorPort, 1
                'Else
                '     WriteIoBit_9030 0, 1, VertMotorPort + 1
                'End If
                
                LblVertMotorMode.BackColor = RGB(255, 0, 0)
                
                VertThreadStep = 102
                VertThreadTime = Timer
            'Else
            '    write_bit 0, VertMotorPort, 0
            '    write_bit 0, VertMotorPort, 0
            '    write_bit 0, VertMotorPort, 0
            '
            '    LblVertMotorMode.BackColor = RGB(0, 255, 0)
            End If
        End If
    
    ElseIf VertThreadStep = 102 Then
        If Timer - VertThreadTime >= 0.02 Then
            'get_status 0, VertAxis, status
            'If CtrlCardType = 0 Then
                get_status 0, VertAxis, Status
            'Else
            '    status = ReadAxisState_9030(0, VertAxis)
            'End If
            If Status = 0 Then
                TmrVertThread.Enabled = False
                VertThreadStep = 103
            End If
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            MnuNew_Click
        Case "Open"
            'MnuOpen_Click
            MnuImportAI_Click
        Case "ImportDXF"
            MnuImportDXF_Click
        Case "Save"
            MnuSave_Click
        Case "SaveAs"
            MnuSaveAs_Click
        Case "Undo"
            MnuUndo_Click
        Case "Redo"
            MnuRedo_Click
        Case "SetAll"
            MnuBYDrawingOrder_Click
        Case "Unset"
            MnuEraseDroppingSetting_Click
        Case "ShowGridLines"
            PicPathCls
            DrawAll
        Case "ShowDirection"
            If Button.value = tbrPressed Then
                ShowDirection = True
            Else
                ShowDirection = False
            End If
            PicPathCls
            DrawAll
        Case "ShowPoints"
            If Button.value = tbrPressed Then
                ShowPoints = True
                CmdAddPoint.Enabled = True
                CmdDeletePoint.Enabled = True
                CmdMovePoint.Enabled = True
            Else
                ShowPoints = False
                CmdAddPoint.Enabled = False
                CmdDeletePoint.Enabled = False
                CmdMovePoint.Enabled = False
           End If
                                       
            PointSize = 5

            PicPathCls
            DrawAll
'        Case "Narrow"
'            If Device_UserSize(1) > 300 Then
'                Device_UserSize(1) = Device_UserSize(1) - 50
'
'                ViewMinX = 0
'                ViewMaxX = Device_UserSize(1)
'                ViewMinY = 0
'                ViewMaxY = Device_UserSize(2)
'                ViewMargin = 0.03
'
'                Zoom 0
'                Me.Refresh
'
'                PicPathCls
'                DrawAll
'            End If
'        Case "Wide"
'            If Device_UserSize(1) < 3000 Then
'                Device_UserSize(1) = Device_UserSize(1) + 50
'
'                ViewMinX = 0
'                ViewMaxX = Device_UserSize(1)
'                ViewMinY = 0
'                ViewMaxY = Device_UserSize(2)
'                ViewMargin = 0.03
'
'                Zoom 0
'                Me.Refresh
'
'                PicPathCls
'                DrawAll
'            End If
'        Case "High"
'            If Device_UserSize(2) < 1000 Then
'                Device_UserSize(2) = Device_UserSize(2) + 50
'
'                ViewMinX = 0
'                ViewMaxX = Device_UserSize(1)
'                ViewMinY = 0
'                ViewMaxY = Device_UserSize(2)
'                ViewMargin = 0.03
'
'                Zoom 0
'                Me.Refresh
'
'                PicPathCls
'                DrawAll
'            End If
'        Case "Low"
'            If Device_UserSize(2) > 300 Then
'                Device_UserSize(2) = Device_UserSize(2) - 50
'
'                ViewMinX = 0
'                ViewMaxX = Device_UserSize(1)
'                ViewMinY = 0
'                ViewMaxY = Device_UserSize(2)
'                ViewMargin = 0.03
'
'                Zoom 0
'                Me.Refresh
'
'                PicPathCls
'                DrawAll
'            End If
    End Select
End Sub

Private Sub TxtBendDeg_DblClick()
    SetDigiPad "FrmMain", "TxtBendDeg"
End Sub

Private Sub TxtEdit_Change(Index As Integer)
    FraEdit.Tag = "1"
End Sub

Private Sub TxtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then 'key return
        If Index < TxtEdit.count - 1 Then
            TxtEdit(Index + 1).SetFocus
            TxtEdit(Index + 1).SelStart = Len(TxtEdit(Index + 1).Text)
        Else
            CmdEdit_Click
        End If
    ElseIf KeyCode = 38 Then 'key up
        If Index > 0 Then
            TxtEdit(Index - 1).SetFocus
            TxtEdit(Index - 1).SelStart = Len(TxtEdit(Index - 1).Text)
        End If
    ElseIf KeyCode = 40 Then 'key down
        If Index < TxtEdit.count - 1 Then
            TxtEdit(Index + 1).SetFocus
            TxtEdit(Index + 1).SelStart = Len(TxtEdit(Index + 1).Text)
        End If
    End If
End Sub

Private Sub TxtEndPointAdjustMM_Change()
    On Error Resume Next
    Device_EndPointAdjustMM = val(TxtEndPointAdjustMM.Text)
    'Device_EndPointAdjustMM = Device_EndComp
End Sub

Private Sub TxtEndPointAdjustMM_DblClick()
    SetDigiPad "FrmMain", "TxtEndPointAdjustMM"
End Sub

Private Sub TxtFeedMM_DblClick()
    SetDigiPad "FrmMain", "TxtFeedMM"
End Sub


Private Sub TxtRunN_DblClick()
    SetDigiPad "FrmMain", "TxtRunN"
End Sub

Private Sub TxtStartPointAdjustMM_Change()
    On Error Resume Next
    Device_StartPointAdjustMM = val(TxtStartPointAdjustMM.Text)
    'Device_StartPointAdjustMM = Device_StartComp
End Sub

Private Sub TxtStartPointAdjustMM_DblClick()
    SetDigiPad "FrmMain", "TxtStartPointAdjustMM"
End Sub

Private Sub TxtVertDeg_DblClick()
    SetDigiPad "FrmMain", "TxtVertDeg"
End Sub

Private Sub VScroll1_Change()
    If PicPath.top = -VScroll1.value Then Exit Sub
    'PicPath.Visible = False
    PicPath.top = -VScroll1.value
    'DrawAll
    'PicPath.Visible = True
    PicPath.Refresh
End Sub

Private Sub VScroll1_Scroll()
    If PicPath.top = -VScroll1.value Then Exit Sub
    'PicPath.Visible = False
    PicPath.top = -VScroll1.value
    'DrawAll
    'PicPath.Visible = True
End Sub

Public Sub FormResize()
    Dim er As Boolean
    
    On Error GoTo ErrorHandler
    
    If Me.WindowState = 1 Then
        Exit Sub
    End If
    
    If Me.Width < 1024 * Screen.TwipsPerPixelX Then
        Me.Width = 1024 * Screen.TwipsPerPixelX
    End If
    
    If Me.Height < 580 * Screen.TwipsPerPixelY Then
        Me.Height = 580 * Screen.TwipsPerPixelY
    End If
    
    If Me.WindowState = 0 Then
        Me.Move 0, 0
    End If
    
    TxtCurTool.Move Toolbar1.Buttons(29).left * Screen.TwipsPerPixelX, 0, 50 * Screen.TwipsPerPixelX, Toolbar1.Buttons(29).Height * Screen.TwipsPerPixelY
    TxtCurData.Move TxtCurTool.left + 50 * Screen.TwipsPerPixelX, 0, (Toolbar1.Buttons(29).Width - 50) * Screen.TwipsPerPixelX, Toolbar1.Buttons(29).Height * Screen.TwipsPerPixelY
    
    PicFrame.left = WL
    PicFrame.Width = Me.Width / Screen.TwipsPerPixelX - (WL + WR)
    PicFrame.top = HT
    PicFrame.Height = Me.Height / Screen.TwipsPerPixelY - (HT + HB) - 10 - StatusBar1.Height
    PicFrame.ScaleMode = 3
    
    HScroll1.left = WL
    HScroll1.top = PicFrame.top + PicFrame.Height
    HScroll1.Width = PicFrame.Width
    
    VScroll1.left = PicFrame.left + PicFrame.Width
    VScroll1.top = HT
    VScroll1.Height = PicFrame.Height
    
    CmdToolBox.left = VScroll1.left + 2
    CmdToolBox.top = HT + 1
    CmdToolBox.Width = WR - 11
    
    CmdToolBox_Click
    
    If ColorMode = 0 Then
        PicPath.BackColor = RGB(255, 255, 255)
    Else
        PicPath.BackColor = RGB(0, 0, 0)
    End If
    
    PicPath.left = 0
    PicPath.top = 0
    PicPath.Width = PicFrame.ScaleWidth
    PicPath.Height = PicFrame.ScaleHeight
    
    ImgLogo.top = Me.ScaleHeight - 80
    
    ShiftX = 0
    ShiftY = 0
    Zoom 0
    Me.Refresh
    
    If er = True Then
        'MsgBox "请将屏幕分辨率调整到 1024*768 以上，否则软件界面不完整。", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    er = True
    Resume Next
End Sub

Public Sub ShowCurHeadPos(Optional mode As Integer = 1)
    '0-reset, 1-normal, 2-erase, 3-start
    Dim X As Single, Y As Single, d As Integer, r As Integer
    Static x0 As Single, y0 As Single
    
    d = 10
    r = 5
    
    If mode = 1 Or mode = 0 Or mode = 2 Then
        If ShowHead = True Then
            PicPath.DrawMode = 7 'XOR pen
            PicPath.Line (x0 - d, y0)-(x0 + d + 1, y0), RGB(0, 0, 255)
            PicPath.Line (x0, y0 - d)-(x0, y0 + d + 1), RGB(0, 0, 255)
            PicPath.Circle (x0, y0), r, RGB(0, 0, 255)
        End If
        
        If mode = 0 Or mode = 2 Then
            PicPath.DrawMode = 13 'Copy pen
            ShowHead = False
        End If
        
        If mode = 0 Then 'reset
            x0 = 0
            y0 = 0
         End If
    End If
    
    If mode = 1 Or mode = 3 Then
        PicPath.DrawMode = 7
        
        ConvertUserToPath CurX, CurY, X, Y
        PicPath.Line (X - d, Y)-(X + d + 1, Y), RGB(0, 0, 255)
        PicPath.Line (X, Y - d)-(X, Y + d + 1), RGB(0, 0, 255)
        PicPath.Circle (X, Y), r, RGB(0, 0, 255)
        x0 = X
        y0 = Y
        
        ShowHead = True
    End If
End Sub

Public Sub ShowHeadPos(Optional mode As Integer = 1)
    '1-show, 2-erase
    Dim X As Single, Y As Single, d As Integer, r As Integer, clr As Long
    Static x0 As Single, y0 As Single, Shown As Boolean
    
    
    d = 10
    r = 5
    
    clr = RGB(128, 128, 128)
    If mode = 1 Or mode = 2 Then
        If Shown = True Then
            PicPath.DrawMode = 7 'XOR pen
            PicPath.Line (x0 - d, y0)-(x0 + d + 1, y0), clr
            PicPath.Line (x0, y0 - d)-(x0, y0 + d + 1), clr
            PicPath.Circle (x0, y0), r, clr
        End If
        
        If mode = 2 Then
            PicPath.DrawMode = 13 'Copy pen
            Shown = False
        End If
    End If
    
    If mode = 1 Then
        PicPath.DrawMode = 7
        
        ConvertUserToPath CurX, CurY, X, Y
        PicPath.Line (X - d, Y)-(X + d + 1, Y), clr
        PicPath.Line (X, Y - d)-(X, Y + d + 1), clr
        PicPath.Circle (X, Y), r, clr
        x0 = X
        y0 = Y
        
        Shown = True
    End If
End Sub

Public Sub PicPathCls()
    PicPath.Cls
    ShowHead = False
End Sub


Private Sub VScroll2_Change()
    TxtTotalRun.Text = VScroll2.value
End Sub

Public Sub MakeRecentFileMenu(ByVal FilePathAndName As String)
    Dim I As Long, s As String, p As Long, RecentFile() As String
    
    On Error Resume Next
    
    If FilePathAndName = "" Then
        ReDim RecentFile(RecentFileCount)

        p = 0
        For I = 0 To RecentFileCount - 1
            RecentFile(I) = Trim(GetStringFromINI("RecentFile", str(I + 1), "", App.Path & "\" & App.EXEName & "_RF.ini"))
            If dir(RecentFile(I)) = "" Then
                RecentFile(I) = ""
                p = 1
            End If
        Next
        If p = 1 Then
            DeleteFile App.Path & "\" & App.EXEName & "_RF.ini"
            
            p = 0
            For I = 0 To RecentFileCount - 1
                If RecentFile(I) <> "" Then
                    WritePrivateProfileString "RecentFile", str(p + 1), RecentFile(I), App.Path & "\" & App.EXEName & "_RF.ini"
                    p = p + 1
                End If
            Next
        End If
        
        For I = 0 To RecentFileCount - 1
            If I > 0 Then
                Load MnuRecentFile(I)
            End If
            
            s = Trim(GetStringFromINI("RecentFile", str(I + 1), "", App.Path & "\" & App.EXEName & "_RF.ini"))
            If s = "" Then
                MnuRecentFile.Item(I).Tag = ""
                MnuRecentFile.Item(I).Visible = False
            Else
                MnuRecentFile.Item(I).Tag = s
                If UCase(Right(s, 4)) = ".DXF" Then
                    s = FileName(s) + ".DXF"
                Else
                    s = FileName(s)
                End If
                
                MnuRecentFile.Item(I).caption = "&" & Format(I + 1, "0.") & s
                MnuRecentFile.Item(I).Visible = True
            End If
        Next
    Else
        For I = 0 To RecentFileCount - 2
            s = Trim(MnuRecentFile.Item(I).Tag)
            If s = "" Or UCase(s) = UCase(FilePathAndName) Then
                MnuRecentFile.Item(I).Tag = MnuRecentFile.Item(I + 1).Tag
                MnuRecentFile.Item(I + 1).Tag = ""
            End If
        Next
        
        For I = RecentFileCount - 1 To 1 Step -1
            s = Trim(MnuRecentFile.Item(I - 1).Tag)
            MnuRecentFile.Item(I).Tag = s
            
            If s = "" Then
                MnuRecentFile.Item(I).Visible = False
            Else
                If UCase(Right(s, 4)) = ".DXF" Then
                    s = FileName(s) + ".DXF"
                Else
                    s = FileName(s)
                End If
                
                MnuRecentFile.Item(I).caption = "&" & Format(I + 1, "0.") & s
                MnuRecentFile.Item(I).Visible = True
            End If
        Next
        MnuRecentFile.Item(0).Tag = FilePathAndName
        
        s = FileName(FilePathAndName)
        If UCase(Right(FilePathAndName, 4)) = ".DXF" Then
            s = s + ".DXF"
        End If
        MnuRecentFile.Item(0).caption = "&" & Format(1, "0.") & s
        MnuRecentFile.Item(0).Visible = True
        
        Me.caption = AppVersion & VersionMark & " (" & Trim(str(ZoomFactor * 100)) & "%) [" & s & "]"
        
        For I = 0 To RecentFileCount - 1
            If Trim(MnuRecentFile.Item(I).Tag) <> "" Then
                WritePrivateProfileString "RecentFile", str(I + 1), MnuRecentFile.Item(I).Tag, App.Path & "\" & App.EXEName & "_RF.ini"
            End If
        Next
    End If
End Sub

Function MovePoint(ByVal pid As Long, ByVal ux As Double, uy As Double) As Boolean
    Dim ux0 As Double, uy0 As Double, id0 As Long, id1 As Long, I As Long, Ret As Boolean
    
    MovePoint = True
    ux0 = PointList(pid).X
    uy0 = PointList(pid).Y
    
    If PointList(pid).Type = PointType.BoxPoint Then
        For I = 1 To PointCount
            If PointList(I).id <> pid And PointList(I).body_id = PointList(pid).body_id Then
                If PointList(I).X = PointList(pid).X Then
                    PointList(I).X = ux
                ElseIf PointList(I).Y = PointList(pid).Y Then
                    PointList(I).Y = uy
                End If
            End If
        Next
    End If
    
    PointList(pid).X = ux
    PointList(pid).Y = uy
    
    If PointList(pid).method = PointMethod.RoundedCorner Then
        Ret = RoundCorner(pid, ArcList(PointList(pid).arc_id).a)
        If Ret = False Then
            PointList(pid).X = ux0
            PointList(pid).Y = uy0
            MovePoint = False
            Exit Function
        End If
    End If
    
    For I = 1 To SegmentCount
        id0 = SegmentList(I).point0_id
        id1 = SegmentList(I).point1_id
        
        If id0 = pid Then
            If PointList(id1).method = PointMethod.RoundedCorner Then
                Ret = RoundCorner(id1, ArcList(PointList(id1).arc_id).a)
                If Ret = False Then
                    PointList(pid).X = ux0
                    PointList(pid).Y = uy0
                    MovePoint = False
                    Exit Function
                End If
            End If
            
        ElseIf id1 = catched_pid Then
            If PointList(id0).method = PointMethod.RoundedCorner Then
                Ret = RoundCorner(id0, ArcList(PointList(id0).arc_id).a)
                If Ret = False Then
                    PointList(pid).X = ux0
                    PointList(pid).Y = uy0
                    MovePoint = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    For I = 1 To OutputStartPointList.count
        If OutputStartPointList.leading_point0(I).id = pid Then
            OutputStartPointList.leading_point0(I).X = OutputStartPointList.leading_point0(I).X + (ux - ux0)
            OutputStartPointList.leading_point0(I).Y = OutputStartPointList.leading_point0(I).Y + (uy - uy0)
        End If
        If OutputStartPointList.leading_point1(I).id = pid Then
            OutputStartPointList.leading_point1(I).X = OutputStartPointList.leading_point1(I).X + (ux - ux0)
            OutputStartPointList.leading_point1(I).Y = OutputStartPointList.leading_point1(I).Y + (uy - uy0)
        End If
    Next
End Function

Sub ShowOperationData(ByVal ux As Double, ByVal uy As Double, ux0 As Double, uy0 As Double, Optional mode As Long = 0)
    Dim dX As Double, dy As Double, Length As Double, angle As Double
    
    dX = Round(ux - ux0, NumDigitsAfterDecimal)
    dy = Round(uy - uy0, NumDigitsAfterDecimal)
    Length = Round(Sqr(dX * dX + dy * dy), NumDigitsAfterDecimal)
    If Length > 0 Then
        angle = Round(GetArcAngle(ux0, uy0, ux, uy) * 180 / Pi, NumDigitsAfterDecimal)
    Else
        angle = 0
    End If
    If mode = 0 Then
        TxtCurData.Text = TxtCurData.Text & ",   DX:" & str(dX) & ",   DY:" & str(dy) & ",   L:" & str(Length) & ",   A:" & str(angle) & "°"
    ElseIf mode = 1 Then
        TxtCurData.Text = TxtCurData.Text & ",   DX:" & str(dX) & ",   DY:" & str(dy) & ",   R:" & str(Length) & ",   A:" & str(angle) & "°"
    ElseIf mode = 2 Then
        TxtCurData.Text = TxtCurData.Text & ",   RA:" & str(Abs(dX)) & ",   RB:" & str(Abs(dy))
    ElseIf mode = 3 Then
        TxtCurData.Text = TxtCurData.Text & ",   W:" & str(dX) & ",   H:" & str(dy)
    End If
End Sub

Sub MnuFile_Click()
    FraEdit.Visible = False
End Sub

Sub MnuEdit_Click()
    FraEdit.Visible = False
End Sub

Sub MnuView_Click()
    FraEdit.Visible = False
End Sub

Sub MnuTool_Click()
    FraEdit.Visible = False
End Sub

Sub MnuMachine_Click()
    FraEdit.Visible = False
End Sub

Sub MnuHelp_Click()
    FraEdit.Visible = False
End Sub

Private Sub Init_Board()
    Dim count As Integer
    
    count = Init_Card
    
    If CtrlCardType = 0 Then
        If count < 1 Then
            MsgBox "8940A1卡未正确安装"
            MotionCardOK = False
        Else
            MotionCardOK = True
        End If
    ElseIf CtrlCardType = 4 Then
        If count <> 0 Then
            MsgBox "未注册控制卡！Controller is not registered!", vbOKOnly, AppVersion
            MotionCardOK = False
        Else
            MotionCardOK = True
            Timer1.Enabled = True
        End If
    Else
        If count <> 0 Then
            MsgBox "9030卡未正确安装", vbOKOnly, AppVersion
            MotionCardOK = False
        Else
            MotionCardOK = True
            Timer1.Enabled = True
        End If
    End If
    
End Sub

Private Sub CmdVertLow_Click()
'    If ChkAutoReset.value = 0 Then
'        VertUp 0, ChkVertMotor.value
'    Else
'        Vert 0, ChkVertMotor.value
'    End If
End Sub

Private Sub CmdVertHigh_Click()
'    If ChkAutoReset.value = 0 Then
'        VertUp 1, ChkVertMotor.value
'    Else
'        Vert 1, ChkVertMotor.value
'    End If
End Sub

Public Sub ShowCalculation(Optional show_text As Boolean = True)
    'mode=1:set start point, mode=-1:set stop point, mode=0:show only
    
    Dim I As Long, lfs0 As Double, ds As Double, p As Long
    Dim i0 As Long, ux As Double, uy As Double, X As Single, Y As Single
    
    If show_text Then
        TxtStatistics.Text = ""
    End If
    
    lfs0 = Device_HeadDistance '0
    p = 0
    'SumCount = 0
    For I = 1 To PathOutputPointCount
        If PathOutputPoint(I).VertType = -1 Then
            ds = PathOutputPoint(I).LengthFromStart - lfs0
            If ds > 0 Then 'Device_VertMinDistance Then
                lfs0 = PathOutputPoint(I).LengthFromStart
                
                If PathOutputPoint(i0).Type <> 99999 Then
                    p = p + 1
                    If show_text Then
                        TxtStatistics.Text = TxtStatistics.Text + Format(p, "00") + " Len:" + str(Round(ds, 2)) + vbCrLf
                    End If
                    
                    ux = PathOutputPoint(i0).ux '+ (PathOutputPoint(I).ux - PathOutputPoint(i0).ux) / 2
                    uy = PathOutputPoint(i0).uy '+ (PathOutputPoint(I).uy - PathOutputPoint(i0).uy) / 2
                    
                    ConvertUserToPath ux, uy, X, Y
                    '这里是显示段号的地方   201502
                    PicPath.CurrentX = X - 0.3
                    PicPath.CurrentY = Y + 0.3
                    'PicPath.CurrentX = X - 30
                    'PicPath.CurrentY = Y + 30
                    PicPath.ForeColor = RGB(255, 255, 255)
                    PicPath.Print str(p)
                Else
                    If show_text Then
                        TxtStatistics.Text = TxtStatistics.Text + "   Len:" + str(Round(ds, 2)) + vbCrLf
                    End If
                End If
            End If
            
            i0 = I
        End If
    Next
    If p > 0 Then
        If show_text Then
            TxtStatistics.Text = TxtStatistics.Text + vbCrLf + "Total:" + str(Round(TotalPathOutLength, 2)) + vbCrLf
        End If
        
        'If mode = 1 Then
        '    SumCount = SumCount + 1
        '    SumTotalPathOutLength0 = SumTotalPathOutLength
        '    SumTotalPathOutLength = TotalPathOutLength + SumTotalPathOutLength + IIf(SumCount > 1, Device_DoneDistance, 0)
        '
        'ElseIf mode = -1 Then
        '    SumTotalPathOutLength = TotalPathOutLength + SumTotalPathOutLength0 + IIf(SumCount > 1, Device_DoneDistance, 0)
        '
        'End If
        
        'If SumCount > 1 Then
        '    TxtStatistics.Text = TxtStatistics.Text + "Sum(" + Trim(str(SumCount)) + "):" + str(Round(SumTotalPathOutLength, 2)) + vbCrLf
        'End If
    End If
    
    FraEdit.Visible = False
    TxtStatistics.Visible = True
End Sub

Sub ElevatorUp()
'Exit Sub
'
'    PortBit(5) = 1
'    PortBit(6) = 0
'    write_bit 0, ElevatorUpPort, 1
'    write_bit 0, ElevatorDownPort, 0
'    Do
'        If read_bit(0, ElevatorUpSensor) = 0 Then
'            Exit Do
'        End If
'
'        If StopRunning = True Then
'            Exit Do
'        End If
'
'        DoEvents
'    Loop
'    PortBit(5) = 0
'    write_bit 0, ElevatorUpPort, 0
End Sub

Sub ElevatorDown()
Exit Sub
'
'    PortBit(5) = 0
'    PortBit(6) = 1
'    write_bit 0, ElevatorUpPort, 0
'    write_bit 0, ElevatorDownPort, 1
'    Do
'        If read_bit(0, ElevatorDownSensor) = 0 Then
'            Exit Do
'        End If
'
'        If StopRunning = True Then
'            Exit Do
'        End If
'
'        DoEvents
'    Loop
'    PortBit(6) = 0
'    write_bit 0, ElevatorDownPort, 0
End Sub

Sub TurnLeft()
'    If ChkVertBeforeTurn.value = 1 Then
'        If Not OptVertLow.value = False Or Not OptVertHigh.value = False Then
'            Vert IIf(OptVertLow.value = True, 0, 1), 1
'        End If
'
'        FeedMM Device_HeadDistance, Device_UseEncoder,  0.5
'    End If
    
    TurnAngle -val(TxtTurnDeg.Text) - IIf(ChkAddEmptyDegree2.value = 0, 0, Device_EmptyDegree)
End Sub

Sub TurnRight()
'    If ChkVertBeforeTurn.value = 1 Then
'        If Not OptVertLow.value = False Or Not OptVertHigh.value = False Then
'            Vert IIf(OptVertLow.value = True, 0, 1), 1
'        End If
'
'        FeedMM Device_HeadDistance, Device_UseEncoder, 0.5
'    End If
    
    TurnAngle val(TxtTurnDeg.Text) + IIf(ChkAddEmptyDegree2.value = 0, 0, Device_EmptyDegree2)
End Sub
Sub ChangeFontByLanguage(ByVal curLanguage As Integer)
    If curLanguage = 0 Then
        CmdFeedBkV2A.FontName = "黑体"
        CmdFeedFWV2A.FontName = "黑体"
        PanButton3.FontName = "黑体"
        PanButton2.FontName = "黑体"
        PanButton11.FontName = "黑体"
        PanButton1.FontName = "黑体"
        PanButton5.FontName = "黑体"
        CmdResetOrg.FontName = "黑体"
        CmdRun.FontName = "黑体"
        CmdStop.FontName = "黑体"
        CmdResume.FontName = "黑体"
        CmdPause.FontName = "黑体"
    Else
        CmdFeedBkV2A.FontName = "Arial"
        CmdFeedFWV2A.FontName = "Arial"
        PanButton3.FontName = "Arial"
        PanButton2.FontName = "Arial"
        PanButton11.FontName = "Arial"
        PanButton1.FontName = "Arial"
        PanButton5.FontName = "Arial"
        CmdResetOrg.FontName = "Arial"
        CmdRun.FontName = "Arial"
        CmdStop.FontName = "Arial"
        CmdResume.FontName = "Arial"
        CmdPause.FontName = "Arial"

    End If
End Sub
Sub ChangeFace(ByVal lm As Integer)
    Dim obj As Object
    Dim ff As Long
    Dim CaptionUsed(500) As Boolean, Caption1(500) As String, Caption2(500) As String, CaptionCount As Long
    Dim ln As String, p As Long, I As Long
    
    Static lm0 As Long
    
    If lm = lm0 Or (lm0 = 0 And lm = 1) Then
        Exit Sub
    End If
        
    On Error GoTo ErrorHandler
    
    CaptionCount = 0
    ff = FreeFile
    Open App.Path & "\" & "English.txt" For Input As ff
    Do While Not EOF(ff)
        Line Input #ff, ln
        If Trim(ln) <> "" Then
            CaptionCount = CaptionCount + 1
            
            p = InStr(ln, ",")
            CaptionUsed(CaptionCount) = False
            Caption1(CaptionCount) = Trim(Mid(ln, 1, p - 1))
            Caption2(CaptionCount) = Trim(Mid(ln, p + 1))
            
'Debug.Print CaptionCount; Caption0(CaptionCount), Caption1(CaptionCount)
        End If
    Loop
    Close #ff
        
    On Error Resume Next
    
    For Each obj In FrmMain
        p = 0
        p = IIf(obj.caption = "", 1, 2)
        If p > 0 Then
            For I = 1 To CaptionCount
                'If CaptionUsed(I) = False Then
                    If lm = 2 Then
                        If obj.caption = Caption1(I) Then
                            obj.caption = Caption2(I)
                            CaptionUsed(I) = True
                            Exit For
                        End If
                    Else
                        If obj.caption = Caption2(I) Then
                            obj.caption = Caption1(I)
                            CaptionUsed(I) = True
                            Exit For
                        End If
                    End If
                'End If
            Next
        End If
    Next
    
    For Each obj In FormSettings
        p = 0
        p = IIf(obj.caption = "", 1, 2)
        If p > 0 Then
            For I = 1 To CaptionCount
                'If CaptionUsed(I) = False Then
                    If lm = 2 Then
                        If obj.caption = Caption1(I) Then
                            obj.caption = Caption2(I)
                            CaptionUsed(I) = True
                            Exit For
                        End If
                    Else
                        If obj.caption = Caption2(I) Then
                            obj.caption = Caption1(I)
                            CaptionUsed(I) = True
                            Exit For
                        End If
                    End If
                'End If
            Next
        End If
    Next
    
    lm0 = lm
    
ErrorHandler:

End Sub

Sub ChangeParamSetLanguage(ByVal lm As Integer)
    Dim obj As Object
    Dim ff As Long
    Dim CaptionUsed(500) As Boolean, Caption1(500) As String, Caption2(500) As String, CaptionCount As Long
    Dim ln As String, p As Long, I As Long
    
    Static lm0 As Long
    
    If lm = lm0 Or (lm0 = 0 And lm = 1) Then
        Exit Sub
    End If
        
    On Error GoTo ErrorHandler
    
    CaptionCount = 0
    ff = FreeFile
    Open App.Path & "\" & "English.txt" For Input As ff '绝对路劲+文件名
    Do While Not EOF(ff)
        Line Input #ff, ln
        If Trim(ln) <> "" Then
            CaptionCount = CaptionCount + 1
            
            p = InStr(ln, ",")
            CaptionUsed(CaptionCount) = False
            Caption1(CaptionCount) = Trim(Mid(ln, 1, p - 1))
            Caption2(CaptionCount) = Trim(Mid(ln, p + 1))
            
'Debug.Print CaptionCount; Caption0(CaptionCount), Caption1(CaptionCount)
        End If
    Loop
    Close #ff
        
    On Error Resume Next
    
'    For Each obj In FrmMain
'        p = 0
'        p = IIf(obj.caption = "", 1, 2)
'        If p > 0 Then
'            For I = 1 To CaptionCount
'                'If CaptionUsed(I) = False Then
'                    If lm = 2 Then
'                        If obj.caption = Caption1(I) Then
'                            obj.caption = Caption2(I)
'                            CaptionUsed(I) = True
'                            Exit For
'                        End If
'                    Else
'                        If obj.caption = Caption2(I) Then
'                            obj.caption = Caption1(I)
'                            CaptionUsed(I) = True
'                            Exit For
'                        End If
'                    End If
'                'End If
'            Next
'        End If
'    Next
    
    For Each obj In FormSettings
        p = 0
        p = IIf(obj.caption = "", 1, 2)
        If p > 0 Then
            For I = 1 To CaptionCount
                'If CaptionUsed(I) = False Then
                    If lm = 2 Then
                        If obj.caption = Caption1(I) Then
                            obj.caption = Caption2(I)
                            CaptionUsed(I) = True
                            Exit For
                        End If
                    Else
                        If obj.caption = Caption2(I) Then
                            obj.caption = Caption1(I)
                            CaptionUsed(I) = True
                            Exit For
                        End If
                    End If
                'End If
            Next
        End If
    Next
    
    lm0 = lm
    
ErrorHandler:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Ret As Integer
    If IsRunning = False Then
        If KeyCode = vbKeyLeft Then
            Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
            Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
            Ret = StartAxisVel_9030(0, FeedAxis, -1 * Device_FeedSpeed)
            
        ElseIf KeyCode = vbKeyRight Then
            Ret = SetAxisAcc_9030(0, FeedAxis, Device_FeedAccel)
            Ret = SetAxisDec_9030(0, FeedAxis, Device_FeedAccel)
            Ret = StartAxisVel_9030(0, FeedAxis, Device_FeedSpeed)
        
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsRunning = False Then
        If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
            
                
                SetAxisStopDec_9030 0, FeedAxis, 100000
                StopAxis_9030 0, FeedAxis
                
                'LblEncoderOffset.caption = nActPos - FeedPulsCount
            
        End If
    End If
End Sub


Sub SelLanguage(mode As Integer)
    If mode = 0 Then
        Label16.caption = "铣刀"
        Label13.caption = "弯弧角度"
        Label14.caption = "切割高度"
        CmdVertLV.caption = "向左"
        CmdVertRV.caption = "向右"
        CmdFeedBkV2.caption = "连续快退"
        CmdFeedFWV2.caption = "连续快进"
        CheckBendAbs.caption = "绝对角度"
        CmdTest.caption = "清零"
        Label7.caption = "当前次数"
        PanButton3.caption = "外轮廓"
        PanButton2.caption = "内轮廓"
        CmdFeedBkV2A.caption = "快退"
        CmdFeedFWV2A.caption = "快进"
    ElseIf mode = 1 Then
        Label16.caption = "SPINDLE"
        Label13.caption = "BendAng"
        Label14.caption = "CutHigh"
        CmdVertLV.caption = "L"
        CmdVertRV.caption = "R"
        CmdFeedBkV2.caption = "BWFast"
        CmdFeedFWV2.caption = "FWFast"
        CheckBendAbs.caption = "Abs Angel"
        CmdTest.caption = "CLR"
        Label7.caption = "LoopCnt"
        PanButton3.caption = "SetOuter"
        PanButton2.caption = "SetInner"
        CmdFeedBkV2A.caption = "BWFast"
        CmdFeedFWV2A.caption = "FWFast"
    End If
End Sub

Sub posErrCompensation(errpos As Double)
    Dim Curpos As Long
    Dim Status As Integer
    If errpos > 2 Or errpos < -2 Then
        'TextBendLength.Text = "err"
        'Exit Sub
    End If
    'Wait 0.5
    Curpos = ReadAxisPos_9030(0, FeedAxis)
    PostoComp = Curpos - (errpos) * Device_PulsPerMM - 0 * Device_PulsPerMM
    SetAxisPos_9030 0, FeedAxis, PostoComp
    'TextBendLength.Text = Round(PostoComp / Device_PulsPerMM, 3)
    StartAxis_9030 0, FeedAxis
    Sleep (100)
    Do
        Sleep (10)
        Status = ReadAxisState_9030(0, FeedAxis)
        If Status = 0 Then     '等待FeedAxis停止运动
            Exit Do
        End If
        
        If StopRunning = True Then
            VertResetOK = False
            Exit Sub
        End If
        DoEvents
    Loop
End Sub

Function FindNextCutPointIndex(ByVal CurIndex As Integer) As Integer
    Dim I As Long
    Dim j As Long
    Dim length_arr As Long
    
    I = CurIndex + 1    '数组PathOutputPoint从1开始，CurIndex最小取值0
    
    length_arr = UBound(PathOutputPoint)
    
    Do While I < length_arr
        If PathOutputPoint(I).VertType > 0 Then
            FindNextCutPointIndex = I
            Exit Function
        End If
        I = I + 1
    Loop
    FindNextCutPointIndex = length_arr   '返回最大值
    
End Function

Function FindNext99999PointIndex(ByVal CurIndex As Integer) As Integer
    Dim I As Long
    Dim j As Long
    Dim length_arr As Long
    
    I = CurIndex + 1    '数组PathOutputPoint从1开始，CurIndex最小取值0
    
    length_arr = UBound(PathOutputPoint)
    
    Do While I < length_arr
        If PathOutputPoint(I).Type = 99999 Then
            FindNext99999PointIndex = I
            Exit Function
        End If
        I = I + 1
    Loop
    FindNext99999PointIndex = length_arr   '返回最大值
    
End Function
