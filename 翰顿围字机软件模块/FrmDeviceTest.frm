VERSION 5.00
Begin VB.Form FrmDeviceTest 
   BackColor       =   &H8000000B&
   Caption         =   "弯刀机控制试验程序"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   688
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   869
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000A&
      Caption         =   "IO 控制"
      Height          =   2655
      Left            =   6600
      TabIndex        =   122
      Top             =   7440
      Width           =   6255
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 4"
         Height          =   375
         Index           =   4
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 5"
         Height          =   375
         Index           =   5
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 6"
         Height          =   375
         Index           =   6
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 7"
         Height          =   375
         Index           =   7
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 8"
         Height          =   375
         Index           =   8
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 9"
         Height          =   375
         Index           =   9
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 10"
         Height          =   375
         Index           =   10
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 11"
         Height          =   375
         Index           =   11
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 12"
         Height          =   375
         Index           =   12
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 13"
         Height          =   375
         Index           =   13
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 14"
         Height          =   375
         Index           =   14
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 15"
         Height          =   375
         Index           =   15
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 0"
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 1"
         Height          =   375
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 2"
         Height          =   375
         Index           =   2
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOutput 
         Caption         =   "Out 3"
         Height          =   375
         Index           =   3
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   202
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   201
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   200
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   199
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   198
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   197
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   196
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   195
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   194
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   3480
         TabIndex        =   193
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   192
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   191
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   190
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   189
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   5280
         TabIndex        =   188
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   5640
         TabIndex        =   187
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   186
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   600
         TabIndex        =   185
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   960
         TabIndex        =   184
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   183
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   1680
         TabIndex        =   182
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   2040
         TabIndex        =   181
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   2400
         TabIndex        =   180
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   2760
         TabIndex        =   179
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   3120
         TabIndex        =   178
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   3480
         TabIndex        =   177
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   26
         Left            =   3840
         TabIndex        =   176
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   27
         Left            =   4200
         TabIndex        =   175
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   28
         Left            =   4560
         TabIndex        =   174
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   29
         Left            =   4920
         TabIndex        =   173
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   30
         Left            =   5280
         TabIndex        =   172
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label LblIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   31
         Left            =   5640
         TabIndex        =   171
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   170
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label83 
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         Height          =   255
         Left            =   5640
         TabIndex        =   169
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         Height          =   255
         Left            =   240
         TabIndex        =   168
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label85 
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         Height          =   255
         Left            =   5640
         TabIndex        =   167
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label86 
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         Height          =   255
         Left            =   3120
         TabIndex        =   166
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label87 
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         Height          =   255
         Left            =   3120
         TabIndex        =   165
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label88 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   255
         Left            =   1680
         TabIndex        =   164
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label89 
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   255
         Left            =   4560
         TabIndex        =   163
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label90 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Left            =   600
         TabIndex        =   162
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label91 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   255
         Left            =   960
         TabIndex        =   161
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label92 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   255
         Left            =   1320
         TabIndex        =   160
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label93 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         Height          =   255
         Left            =   2040
         TabIndex        =   159
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label94 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         Height          =   255
         Left            =   2400
         TabIndex        =   158
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label95 
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         Height          =   255
         Left            =   2760
         TabIndex        =   157
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label96 
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         Height          =   255
         Left            =   3480
         TabIndex        =   156
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label97 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   3840
         TabIndex        =   155
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label98 
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         Height          =   255
         Left            =   4200
         TabIndex        =   154
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label99 
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         Height          =   255
         Left            =   4920
         TabIndex        =   153
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label100 
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         Height          =   255
         Left            =   5280
         TabIndex        =   152
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label101 
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         Height          =   255
         Left            =   600
         TabIndex        =   151
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label102 
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         Height          =   255
         Left            =   960
         TabIndex        =   150
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label103 
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         Height          =   255
         Left            =   1320
         TabIndex        =   149
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label104 
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         Height          =   255
         Left            =   1680
         TabIndex        =   148
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label105 
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         Height          =   255
         Left            =   2040
         TabIndex        =   147
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label106 
         BackStyle       =   0  'Transparent
         Caption         =   "22"
         Height          =   255
         Left            =   2400
         TabIndex        =   146
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label107 
         BackStyle       =   0  'Transparent
         Caption         =   "23"
         Height          =   255
         Left            =   2760
         TabIndex        =   145
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label108 
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         Height          =   255
         Left            =   3480
         TabIndex        =   144
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label109 
         BackStyle       =   0  'Transparent
         Caption         =   "26"
         Height          =   255
         Left            =   3840
         TabIndex        =   143
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label110 
         BackStyle       =   0  'Transparent
         Caption         =   "27"
         Height          =   255
         Left            =   4200
         TabIndex        =   142
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label111 
         BackStyle       =   0  'Transparent
         Caption         =   "28"
         Height          =   255
         Left            =   4560
         TabIndex        =   141
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label112 
         BackStyle       =   0  'Transparent
         Caption         =   "29"
         Height          =   255
         Left            =   4920
         TabIndex        =   140
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label113 
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         Height          =   255
         Left            =   5280
         TabIndex        =   139
         Top             =   1920
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      Caption         =   "开孔控制"
      Height          =   2655
      Left            =   240
      TabIndex        =   85
      Top             =   7440
      Width           =   6255
      Begin VB.CommandButton CmdReset3 
         Caption         =   "凸轮复位"
         Height          =   495
         Left            =   2880
         TabIndex        =   114
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdZStop0 
         Height          =   255
         Left            =   5760
         TabIndex        =   105
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox TxtReadPos3P 
         Height          =   285
         Left            =   3480
         TabIndex        =   103
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtReadSpeed3P 
         Height          =   285
         Left            =   3480
         TabIndex        =   101
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdMakeHole3 
         Caption         =   "开鹰嘴孔"
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CmdMakeHole2 
         Caption         =   "开直切孔"
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdMakeHole1 
         Caption         =   "开桥位孔"
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtAccl3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   91
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TxtSpeed3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   90
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtStartV3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   89
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "(第三轴控制开孔电机)"
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "当前位置(p)"
         Height          =   255
         Left            =   2520
         TabIndex        =   102
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "当前速度(p/s)"
         Height          =   255
         Left            =   2400
         TabIndex        =   100
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "(out 4/7)"
         Height          =   255
         Left            =   5520
         TabIndex        =   99
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "(out 3/6)"
         Height          =   255
         Left            =   5520
         TabIndex        =   97
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "(out 2/5)"
         Height          =   255
         Left            =   5520
         TabIndex        =   95
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LblZLMT_P 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5400
         TabIndex        =   93
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "原点(ZStop0)"
         Height          =   255
         Left            =   4200
         TabIndex        =   92
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(p/ss)"
         Height          =   255
         Left            =   360
         TabIndex        =   88
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(p/s)"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(p/s)"
         Height          =   255
         Left            =   360
         TabIndex        =   86
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "连续加工控制"
      Height          =   855
      Left            =   240
      TabIndex        =   76
      Top             =   6480
      Width           =   12615
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   5160
         TabIndex        =   229
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   5160
         TabIndex        =   228
         Top             =   210
         Width           =   255
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000A&
         Height          =   615
         Left            =   7560
         TabIndex        =   225
         Top             =   120
         Width           =   2055
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000B&
            Caption         =   "弯正角"
            Height          =   255
            Left            =   240
            TabIndex        =   227
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000B&
            Caption         =   "弯负角"
            Height          =   255
            Left            =   1080
            TabIndex        =   226
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TxtBendingMarginMM 
         Height          =   285
         Left            =   1320
         TabIndex        =   223
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxtBendingCount 
         Height          =   285
         Left            =   4200
         TabIndex        =   212
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton CmdStopRunning 
         Caption         =   "停止"
         Height          =   495
         Left            =   11160
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdRun 
         Caption         =   "运行"
         Height          =   495
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtBendingD 
         Height          =   285
         Left            =   6360
         TabIndex        =   80
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtBendingMM 
         Height          =   285
         Left            =   3120
         TabIndex        =   78
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "运动角度(°)"
         Height          =   255
         Left            =   5400
         TabIndex        =   224
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "端距(mm)"
         Height          =   255
         Left            =   480
         TabIndex        =   222
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "步数"
         Height          =   255
         Left            =   3720
         TabIndex        =   211
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "实际角度(°)"
         Height          =   255
         Left            =   5400
         TabIndex        =   79
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "步距(mm)"
         Height          =   255
         Left            =   2280
         TabIndex        =   77
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "弯刀控制"
      Height          =   6255
      Left            =   6600
      TabIndex        =   39
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdMoveRight 
         Caption         =   "负转"
         Height          =   495
         Left            =   3240
         TabIndex        =   221
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton CmdMoveLeft 
         Caption         =   "正转"
         Height          =   495
         Left            =   2520
         TabIndex        =   220
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton CmdStop2 
         Caption         =   "停止"
         Height          =   495
         Left            =   1200
         TabIndex        =   219
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox TxtBendHeadOffsetD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   218
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtBreakingAngle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   214
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton CmdClip 
         Caption         =   "剪刀动作"
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmdClipperUp 
         Caption         =   "剪刀架升起"
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmdTrackUp 
         Caption         =   "导正器升起"
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TxtLeftGapD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   119
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtRightGapD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   118
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton CmdKnifeBreak 
         Caption         =   "挠断刀片"
         Height          =   495
         Left            =   4920
         TabIndex        =   115
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdYLMTM 
         Height          =   195
         Left            =   5760
         TabIndex        =   108
         Top             =   5520
         Width           =   255
      End
      Begin VB.CommandButton CmdYLMTP 
         Height          =   195
         Left            =   5760
         TabIndex        =   107
         Top             =   5280
         Width           =   255
      End
      Begin VB.TextBox TxtMaxSpeedDPS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   71
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtReadBendingPosD 
         Height          =   285
         Left            =   2640
         TabIndex        =   65
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox TxtReadBendingSpeedD 
         Height          =   285
         Left            =   2640
         TabIndex        =   64
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox TxtReadBendingPosP 
         Height          =   285
         Left            =   1440
         TabIndex        =   63
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox TxtReadBendingSpeedP 
         Height          =   315
         Left            =   1440
         TabIndex        =   62
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox TxtAccl2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   53
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox TxtSpeed2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   52
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TxtStartV2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   51
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TxtBackStartV2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   50
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TxtBackSpeed2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   49
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TxtBackAccl2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   48
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton CmdBendPositive 
         Caption         =   "弯正角"
         Height          =   495
         Left            =   2520
         TabIndex        =   44
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdBendNegative 
         Caption         =   "弯负角"
         Height          =   495
         Left            =   3720
         TabIndex        =   43
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdReset2 
         Caption         =   "弯刀头复位"
         Height          =   495
         Left            =   1200
         TabIndex        =   42
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TxtPPD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   41
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtBendDegrees 
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "中心偏移量(°)"
         Height          =   255
         Left            =   3840
         TabIndex        =   217
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "挠断往复角度(°)"
         Height          =   255
         Left            =   4080
         TabIndex        =   213
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "(Out 10/11)"
         Height          =   255
         Left            =   5040
         TabIndex        =   210
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "(Out 9)"
         Height          =   255
         Left            =   3960
         TabIndex        =   209
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "(Out 8)"
         Height          =   255
         Left            =   2760
         TabIndex        =   208
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "正向间隙角度(°)"
         Height          =   255
         Left            =   3720
         TabIndex        =   121
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "负向间隙角度(°)"
         Height          =   255
         Left            =   3720
         TabIndex        =   120
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "负限位(YLMT-)"
         Height          =   255
         Left            =   4080
         TabIndex        =   75
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "正限位(YLMT+)"
         Height          =   255
         Left            =   4080
         TabIndex        =   74
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label LblYLMT_M 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5400
         TabIndex        =   73
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label LblYLMT_P 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5400
         TabIndex        =   72
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "最大速度(度/秒)(°/s)"
         Height          =   255
         Left            =   720
         TabIndex        =   70
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "(°)"
         Height          =   255
         Left            =   2280
         TabIndex        =   69
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "当前位置(p)"
         Height          =   255
         Left            =   360
         TabIndex        =   68
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "(°/s)"
         Height          =   255
         Left            =   2280
         TabIndex        =   67
         Top             =   5160
         Width           =   375
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "当前速度(p/s)"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(°/ss)"
         Height          =   255
         Left            =   1440
         TabIndex        =   61
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(°/s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   60
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(°/s)"
         Height          =   255
         Left            =   1440
         TabIndex        =   59
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(°/s)"
         Height          =   255
         Left            =   4080
         TabIndex        =   58
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(°/s)"
         Height          =   255
         Left            =   3840
         TabIndex        =   57
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(°/ss)"
         Height          =   255
         Left            =   3960
         TabIndex        =   56
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "弯刀头速度"
         Height          =   255
         Left            =   2520
         TabIndex        =   55
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "空回速度"
         Height          =   255
         Left            =   5040
         TabIndex        =   54
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "(第二轴控制弯刀电机)"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "每度脉冲数(p/°)"
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "运动角度(°)"
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   3000
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "刀片输送控制"
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdHolderFeed 
         Caption         =   "刀夹前进"
         Height          =   495
         Left            =   2520
         TabIndex        =   216
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton CmdHolderBack 
         Caption         =   "刀夹后退"
         Height          =   495
         Left            =   3720
         TabIndex        =   215
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton CmdKnifeLoad 
         Caption         =   "刀片上料"
         Height          =   495
         Left            =   2520
         TabIndex        =   204
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmdCut 
         Caption         =   "剪断刀片"
         Height          =   495
         Left            =   4920
         TabIndex        =   203
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox TxtFeedMaxMM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   117
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdXLMTM 
         Height          =   195
         Left            =   5640
         TabIndex        =   113
         Top             =   5520
         Width           =   255
      End
      Begin VB.CommandButton CmdKnifeFeedHold 
         Caption         =   "输送夹紧"
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmdXLMTP 
         Height          =   195
         Left            =   5640
         TabIndex        =   106
         Top             =   5280
         Width           =   255
      End
      Begin VB.TextBox TxtPP100MM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtMaxSpeedMMPS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtHoldDelay 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "停止"
         Height          =   495
         Left            =   1320
         TabIndex        =   31
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton CmdKnifeBackHold 
         Caption         =   "返回夹紧"
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TxtBackAccl 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   24
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox TxtBackSpeed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   23
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TxtBackStartV 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton CmdContinueFeed 
         Caption         =   "连续输送"
         Height          =   495
         Left            =   4920
         TabIndex        =   21
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdKnifeFeed 
         Caption         =   "刀片前进"
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdKnifeBack 
         Caption         =   "刀片后退"
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdReset1 
         Caption         =   "复位"
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TxtMoveMM 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox TxtReadKnifeSpeedP 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox TxtReadKnifePosP 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox TxtStartV 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TxtSpeed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TxtAccl 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox TxtReadKnifeSpeedMM 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox TxtReadKnifePosMM 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label Label118 
         BackStyle       =   0  'Transparent
         Caption         =   "输送行程(mm)"
         Height          =   255
         Left            =   3840
         TabIndex        =   116
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblXLMT_M 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   112
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label115 
         BackStyle       =   0  'Transparent
         Caption         =   "负限位(XLMT-)"
         Height          =   255
         Left            =   3960
         TabIndex        =   111
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label114 
         BackStyle       =   0  'Transparent
         Caption         =   "(Out 1)"
         Height          =   255
         Left            =   5160
         TabIndex        =   110
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label LblXLMT_P 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   84
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label Label68 
         BackStyle       =   0  'Transparent
         Caption         =   "正限位(XLMT+)"
         Height          =   255
         Left            =   3960
         TabIndex        =   83
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "每100毫米脉冲数(p/100mm)"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "最大速度(毫米/秒)(mm/s)"
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "(Out 0)"
         Height          =   255
         Left            =   3960
         TabIndex        =   34
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "夹紧延时(ms)"
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "空回速度"
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "刀片速度"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(mm/ss)"
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(mm/s)"
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(mm/s)"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "(第一轴控制输送电机)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "初速度(mm/s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "驱动速度(mm/s)"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "加速度(mm/ss)"
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "移动距离(mm)"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "当前速度(p/s)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "(mm/s)"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "当前位置(p)"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "(mm)"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   5520
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   6240
   End
End
Attribute VB_Name = "FrmDeviceTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub CmdBendNegative_Click()
    Dim d As Double
    
    If ResetDone = False Then
        Reset2
        KnifeFeedMM SensorLenMM1 + 20
        KnifeBreak
    End If
                    
    CmdBendNegative.Enabled = False
    d = -Val(TxtBendDegrees.Text)
    KnifeBendD d, 0, False
    CmdBendNegative.Enabled = True
End Sub

Private Sub CmdBendPositive_Click()
    Dim d As Double
    
    If ResetDone = False Then
        Reset2
        KnifeFeedMM SensorLenMM1 + 20
        KnifeBreak
    End If
                    
    CmdBendPositive.Enabled = False
    d = Val(TxtBendDegrees.Text)
    KnifeBendD d, 0, False
    CmdBendPositive.Enabled = True
End Sub

Private Sub CmdClip_Click()
    CmdClip.BackColor = RGB(255, 0, 0)
    Clip
    CmdClip.BackColor = FrmDeviceTest.CmdRun.BackColor
End Sub

Private Sub CmdClipperUp_Click()
    If OutputStatus(9) = 0 Then
        CmdClipperUp.BackColor = RGB(255, 0, 0)
        ClipperUp
    Else
        CmdClipperUp.BackColor = FrmDeviceTest.CmdRun.BackColor
        ClipperDown
    End If
End Sub

Private Sub CmdContinueFeed_Click()
    StopRun = False
    StopFeed = False
    
    KnifeFeedMM 1000000
End Sub

Private Sub CmdCut_Click()
    KnifeMakeHoleAndCut
End Sub

Private Sub CmdHolderBack_Click()
    KnifeBackHolderOff
    
    KnifeFeedHolderOff
    KnifeBack Val(TxtMoveMM.Text)
End Sub

Private Sub CmdHolderFeed_Click()
    KnifeBackHolderOff
    
    KnifeFeedHolderOff
    KnifeFeed Val(TxtMoveMM.Text)
End Sub

Private Sub CmdKnifeBack_Click()
    StopRun = False
    StopFeed = False
    
    KnifeBackMM Val(TxtMoveMM.Text)
End Sub

Private Sub CmdKnifeBreak_Click()
    If ResetDone = False Then
        Reset2
    End If
                
    KnifeBreak
End Sub

Private Sub CmdKnifeFeed_Click()
    StopRun = False
    StopFeed = False
    
    KnifeFeedMM Val(TxtMoveMM.Text)
End Sub

Private Sub CmdKnifeBackHold_Click()
    If OutputStatus(1) = 0 Then
        CmdKnifeBackHold.BackColor = RGB(255, 0, 0)
        KnifeBackHolderOn
    Else
        CmdKnifeBackHold.BackColor = FrmDeviceTest.CmdRun.BackColor
        KnifeBackHolderOff
    End If
End Sub

Private Sub CmdKnifeFeedHold_Click()
    If OutputStatus(0) = 0 Then
        CmdKnifeFeedHold.BackColor = RGB(255, 0, 0)
        KnifeFeedHolderOn
    Else
        CmdKnifeFeedHold.BackColor = FrmDeviceTest.CmdRun.BackColor
        KnifeFeedHolderOff
    End If
End Sub

Private Sub CmdKnifeLoad_Click()
    If ResetDone = False Then
        Reset2
    End If
                
    KnifeLoad
End Sub

Private Sub CmdMakeHole1_Click()
    CmdMakeHole1.BackColor = RGB(255, 0, 0)
    MakeHole 1
    CmdMakeHole1.BackColor = FrmDeviceTest.CmdRun.BackColor
End Sub

Private Sub CmdMakeHole2_Click()
    CmdMakeHole2.BackColor = RGB(255, 0, 0)
    MakeHole 2
    CmdMakeHole2.BackColor = FrmDeviceTest.CmdRun.BackColor
End Sub

Private Sub CmdMakeHole3_Click()
    CmdMakeHole3.BackColor = RGB(255, 0, 0)
    MakeHole 3
    CmdMakeHole3.BackColor = FrmDeviceTest.CmdRun.BackColor
End Sub

Private Sub CmdMoveLeft_Click()
    Dim sv As Long, Sp As Long, Accl As Long
    
    sv = startv2 * PPD
    Sp = speed2 * PPD
    Accl = Accl2 * PPD
    SetSpeed 2, sv, Sp, Accl, 0, 0

    pmove 0, 2, 2000
End Sub

Private Sub CmdMoveRight_Click()
    Dim sv As Long, Sp As Long, Accl As Long
    
    sv = startv2 * PPD
    Sp = speed2 * PPD
    Accl = Accl2 * PPD
    SetSpeed 2, sv, Sp, Accl, 0, 0

    pmove 0, 2, -2000
End Sub

Private Sub CmdOutput_Click(Index As Integer)
'    Static status(15) As Byte
    
    If OutputStatus(Index) = 0 Then
        OutputStatus(Index) = 1
    Else
        OutputStatus(Index) = 0
    End If
    
    If Index = 0 Then
        CmdKnifeFeedHold.BackColor = CmdOutput(Index).BackColor
    ElseIf Index = 1 Then
        CmdKnifeBackHold.BackColor = CmdOutput(Index).BackColor
    End If
    
    'write_bit 0, Index, OutputStatus(Index)
'    Select Case Index
'        Case 0
'            write_bit 0, 0, OutputStatus(Index)
'        Case 1
'            write_bit 0, 1, OutputStatus(Index)
'        Case 2
'            write_bit 0, 2, OutputStatus(Index)
'        Case 3
'            write_bit 0, 3, OutputStatus(Index)
'        Case 4
'            write_bit 0, 4, OutputStatus(Index)
'        Case 5
'            write_bit 0, 5, OutputStatus(Index)
'        Case 6
'            write_bit 0, 6, OutputStatus(Index)
'        Case 7
'            write_bit 0, 7, OutputStatus(Index)
'        Case 8
'            write_bit 0, 8, OutputStatus(Index)
'        Case 9
'            write_bit 0, 9, OutputStatus(Index)
'        Case 10
'            write_bit 0, 10, OutputStatus(Index)
'        Case 11
'            write_bit 0, 11, OutputStatus(Index)
'        Case 12
'            write_bit 0, 12, OutputStatus(Index)
'        Case 13
'            write_bit 0, 13, OutputStatus(Index)
'        Case 14
'            write_bit 0, 14, OutputStatus(Index)
'        Case 15
'            write_bit 0, 15, OutputStatus(Index)
'    End Select

If OutputStatus(Index) = 0 Then
    Select Case Index
        Case 0
            write_bit 0, 0, 0
        Case 1
            write_bit 0, 1, 0
        Case 2
            write_bit 0, 2, 0
        Case 3
            write_bit 0, 3, 0
        Case 4
            write_bit 0, 4, 0
        Case 5
            write_bit 0, 5, 0
        Case 6
            write_bit 0, 6, 0
        Case 7
            write_bit 0, 7, 0
        Case 8
            write_bit 0, 8, 0
        Case 9
            write_bit 0, 9, 0
        Case 10
            write_bit 0, 10, 0
        Case 11
            write_bit 0, 11, 0
        Case 12
            write_bit 0, 12, 0
        Case 13
            write_bit 0, 13, 0
        Case 14
            write_bit 0, 14, 0
        Case 15
            write_bit 0, 15, 0
    End Select
Else
    Select Case Index
        Case 0
            write_bit 0, 0, 1
        Case 1
            write_bit 0, 1, 1
        Case 2
            write_bit 0, 2, 1
        Case 3
            write_bit 0, 3, 1
        Case 4
            write_bit 0, 4, 1
        Case 5
            write_bit 0, 5, 1
        Case 6
            write_bit 0, 6, 1
        Case 7
            write_bit 0, 7, 1
        Case 8
            write_bit 0, 8, 1
        Case 9
            write_bit 0, 9, 1
        Case 10
            write_bit 0, 10, 1
        Case 11
            write_bit 0, 11, 1
        Case 12
            write_bit 0, 12, 1
        Case 13
            write_bit 0, 13, 1
        Case 14
            write_bit 0, 14, 1
        Case 15
            write_bit 0, 15, 1
    End Select
End If

    Wait 0.02
End Sub

Private Sub CmdReset1_Click()
    Reset1
End Sub

Private Sub CmdReset2_Click()
    Reset2
End Sub

Private Sub CmdReset3_Click()
    Reset3
End Sub

Private Sub CmdRun_Click()
    Dim mm As Double, d As Double, t As Boolean, i As Long
    
    If ResetDone = False Then
        Reset2
    
        KnifeFeedMM 100000, True
        KnifeFeedMM SensorLenMM1 + 10
        KnifeBreak
    End If
                
    KnifeBackHolderOff
    KnifeFeedHolderOn
    
    StopRun = False
    StopFeed = False
    
    
    mm = Val(TxtBendingMarginMM.Text)
    KnifeFeedMM mm
    
    For i = 1 To Val(TxtBendingCount)
        If StopRun = True Then
            Exit For
        End If
        
        mm = 0
        If i > 1 Then
            mm = Val(TxtBendingMM.Text)
            KnifeFeedMM mm
        End If
        If StopRun = True Then
            Exit For
        End If
        
        d = IIf(Option1.value = True, Val(TxtBendingD.Text), -Val(TxtBendingD.Text))
        t = Option4.value
        
        KnifeBendD d, mm, t
    Next
    If StopRun = False Then
        mm = Val(TxtBendingMarginMM.Text)
        KnifeFeedMM mm
        KnifeBreak
        KnifeFeedHolderOff
    End If
End Sub

Private Sub CmdStop_Click()
    StopFeed = True
    sudden_stop 0, 1
End Sub

Private Sub Command1_Click()
    sudden_stop 0, 2
End Sub

Private Sub CmdStop2_Click()
    sudden_stop 0, 2
    Wait 1
    sudden_stop 0, 2
End Sub

Private Sub CmdStopRunning_Click()
    StopRun = True
    StopFeed = True
    
    sudden_stop 0, 1
    sudden_stop 0, 2
End Sub

Private Sub CmdTrackUp_Click()
    If OutputStatus(8) = 0 Then
        CmdTrackUp.BackColor = RGB(255, 0, 0)
        TrackUp
    Else
        CmdTrackUp.BackColor = FrmDeviceTest.CmdRun.BackColor
        TrackDown
    End If
End Sub

Private Sub CmdXLMTM_Click()
    sudden_stop 0, 1
End Sub

Private Sub CmdXLMTP_Click()
    sudden_stop 0, 1
End Sub

Private Sub CmdYLMTM_Click()
    sudden_stop 0, 2
End Sub

Private Sub CmdYLMTP_Click()
    sudden_stop 0, 2
End Sub

Private Sub CmdZStop0_Click()
    dec_stop 0, 3
End Sub

Private Sub Form_Load()
    Dim rtn As Integer, i As Integer, max_speed As Long
        
'    rtn = adt850_initial
'    If rtn <= 0 Then
'        Exit Sub
'    End If
'
'    '第一轴，用于进退刀
'    rtn = set_limit_mode(0, 1, 0, 0)    '设置限位
'
'    '第二轴，用于弯刀
'    rtn = set_limit_mode(0, 2, 0, 0)    '设置限位
'
'    SetMaxSpeed 1, 400000
'    SetMaxSpeed 2, 200000
'    SetMaxSpeed 3, 800000
'
'    set_command_pos 0, 1, 0
'    set_command_pos 0, 2, 0
'    set_command_pos 0, 3, 0

    LoadParameters
    
    TxtPP100MM.Text = Str(PP100MM)
    
    TxtMaxSpeedMMPS.Text = Str(MaxSpeedMMPS)
    max_speed = MaxSpeedMMPS * PP100MM / 100
    SetMaxSpeed 1, max_speed
    
    TxtStartV.Text = Str(startv)
    TxtSpeed.Text = Str(speed)
    TxtAccl.Text = Str(Accl)
    TxtBackStartV.Text = Str(BackStartV)
    TxtBackSpeed.Text = Str(BackSpeed)
    TxtBackAccl.Text = Str(BackAccl)
    
    TxtMoveMM.Text = GetFromINI("MoveMM")
    
    'TxtFeedMaxMM.Text = Str(FeedMaxMM)
    TxtHoldDelay.Text = Str(HoldDelay)
    
    '------------------------------------------------
    TxtPPD.Text = Str(PPD)
    TxtMaxSpeedDPS.Text = Str(MaxSpeedDPS)
    max_speed = MaxSpeedDPS * PPD
    SetMaxSpeed 2, max_speed

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
    
    TxtBendDegrees.Text = GetFromINI("BendDegrees")
    
    TxtStartV3.Text = Str(startv3)
    TxtSpeed3.Text = Str(speed3)
    TxtAccl3.Text = Str(Accl3)
    
    Option1.value = IIf(GetFromINI("Option1") = "1", True, False)
    Option2.value = Not Option1.value
    Option4.value = IIf(GetFromINI("Option4") = "1", True, False)
    Option3.value = Not Option4.value
    
    TxtBendingMarginMM.Text = GetFromINI("BendingMarginMM")
    TxtBendingMM.Text = GetFromINI("BendingMM")
    TxtBendingD.Text = GetFromINI("BendingD")
    TxtBendingCount.Text = GetFromINI("BendingCount")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Option1_Click()
    WriteToINI "Option1", IIf(Option1.value = True, "1", "0")
End Sub

Private Sub Option2_Click()
    WriteToINI "Option1", IIf(Option2.value = True, "0", "1")
End Sub

Private Sub Option3_Click()
    WriteToINI "Option4", IIf(Option3.value = True, "0", "1")
End Sub

Private Sub Option4_Click()
    WriteToINI "Option4", IIf(Option4.value = True, "1", "0")
End Sub

Private Sub Timer1_Timer()
    Dim Sp As Long, pos As Long, i As Integer, b As Long
    
    On Error Resume Next
    
    get_speed 0, 1, Sp
    TxtReadKnifeSpeedP.Text = Str(Sp * Ratio(1))
    TxtReadKnifeSpeedMM.Text = Str(100 * Sp * Ratio(1) / PP100MM)
    
    get_command_pos 0, 1, pos
    TxtReadKnifePosP.Text = Str(pos)
    TxtReadKnifePosMM.Text = Str(100 * pos / PP100MM)
    
    get_speed 0, 2, Sp
    TxtReadBendingSpeedP.Text = Str(Sp * Ratio(2))
    TxtReadBendingSpeedD.Text = Str(Sp * Ratio(2) / PPD)
    
    get_command_pos 0, 2, pos
    TxtReadBendingPosP.Text = Str(pos)
    TxtReadBendingPosD.Text = Str(pos / PPD)
    
    get_speed 0, 3, Sp
    TxtReadSpeed3P.Text = Str(Sp * Ratio(3))
    
    get_command_pos 0, 3, pos
    TxtReadPos3P.Text = Str(pos)
    
    For i = 0 To 15
        If OutputStatus(i) = 1 Then
            CmdOutput(i).BackColor = RGB(255, 0, 0)
        Else
            CmdOutput(i).BackColor = FrmDeviceTest.CmdRun.BackColor
        End If
    Next
    
    For i = 0 To 31
        'b = read_bit(0, i)
        If i = 0 Then
            b = read_bit(0, 0)
        ElseIf i = 1 Then
            b = read_bit(0, 1)
        ElseIf i = 2 Then
            b = read_bit(0, 2)
        ElseIf i = 3 Then
            b = read_bit(0, 3)
        ElseIf i = 4 Then
            b = read_bit(0, 4)
        ElseIf i = 5 Then
            b = read_bit(0, 5)
        ElseIf i = 6 Then
            b = read_bit(0, 6)
        ElseIf i = 7 Then
            b = read_bit(0, 7)
        ElseIf i = 8 Then
            b = read_bit(0, 8)
        ElseIf i = 9 Then
            b = read_bit(0, 9)
        ElseIf i = 10 Then
            b = read_bit(0, 10)
        ElseIf i = 11 Then
            b = read_bit(0, 11)
        ElseIf i = 12 Then
            b = read_bit(0, 12)
        ElseIf i = 13 Then
            b = read_bit(0, 13)
        ElseIf i = 14 Then
            b = read_bit(0, 14)
        ElseIf i = 15 Then
            b = read_bit(0, 15)
        ElseIf i = 16 Then
            b = read_bit(0, 16)
        ElseIf i = 17 Then
            b = read_bit(0, 17)
        ElseIf i = 18 Then
            b = read_bit(0, 18)
        ElseIf i = 19 Then
            b = read_bit(0, 19)
        ElseIf i = 20 Then
            b = read_bit(0, 20)
        Else
            b = read_bit(0, i)
        End If
        If b = 0 Then
            LblIn(i).BackColor = RGB(255, 0, 0)
        Else
            LblIn(i).BackColor = RGB(0, 255, 0)
        End If
    
        If i = 0 Then 'XLMT+
            If b = 0 Then
                LblXLMT_P.BackColor = RGB(255, 0, 0)
            Else
                LblXLMT_P.BackColor = RGB(0, 255, 0)
            End If
        
        ElseIf i = 1 Then 'XLMT-
            If b = 0 Then
                LblXLMT_M.BackColor = RGB(255, 0, 0)
            Else
                LblXLMT_M.BackColor = RGB(0, 255, 0)
            End If
        
        ElseIf i = 8 Then 'YLMT+
            If b = 0 Then
                LblYLMT_P.BackColor = RGB(255, 0, 0)
            Else
                LblYLMT_P.BackColor = RGB(0, 255, 0)
            End If
        
        ElseIf i = 9 Then 'YLMT-
            If b = 0 Then
                LblYLMT_M.BackColor = RGB(255, 0, 0)
            Else
                LblYLMT_M.BackColor = RGB(0, 255, 0)
            End If
    
        ElseIf i = 18 Then 'ZSTOP0
            If b = 0 Then
                LblZLMT_P.BackColor = RGB(255, 0, 0)
            Else
                LblZLMT_P.BackColor = RGB(0, 255, 0)
            End If
        End If
    Next
End Sub

Sub Wait(ByVal s As Double)
    Dim tm0 As Double, tm As Double
    
    tm0 = Timer
    Do
        tm = Timer
        If tm - tm0 >= s Then
            Exit Do
        ElseIf tm < tm0 Then
            If tm + 86400 - tm0 >= s Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
End Sub

Private Sub TxtBendDegrees_Change()
    WriteToINI "BendDegrees", TxtBendDegrees.Text
End Sub

Private Sub TxtBendingCount_Change()
    WriteToINI "BendingCount", TxtBendingCount.Text
End Sub

Private Sub TxtBendingD_Change()
    WriteToINI "BendingD", TxtBendingD.Text
End Sub

Private Sub TxtBendingMarginMM_Change()
    WriteToINI "BendingMarginMM", TxtBendingMarginMM.Text
End Sub

Private Sub TxtBendingMM_Change()
    WriteToINI "BendingMM", TxtBendingMM.Text
End Sub

Private Sub TxtMoveMM_Change()
    WriteToINI "MoveMM", TxtMoveMM.Text
End Sub
