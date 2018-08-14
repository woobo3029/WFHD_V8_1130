VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FormSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ParameterSetting"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
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
   ScaleHeight     =   646
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtStartComp2 
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
      Left            =   2805
      TabIndex        =   189
      Top             =   4590
      Width           =   975
   End
   Begin VB.TextBox TxtStartComp 
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
      Left            =   2805
      TabIndex        =   188
      Top             =   4290
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   9000
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   8895
      Begin VB.TextBox TxtEndPtAdjustMM 
         Height          =   285
         Left            =   7965
         TabIndex        =   203
         Top             =   4470
         Width           =   870
      End
      Begin VB.TextBox TxtStartPtAdjustMM 
         Height          =   285
         Left            =   7965
         TabIndex        =   202
         Top             =   4140
         Width           =   870
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   105
         TabIndex        =   197
         Top             =   210
         Width           =   420
      End
      Begin VB.TextBox TxtEndComp2 
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
         TabIndex        =   191
         Top             =   4455
         Width           =   975
      End
      Begin VB.TextBox TxtEndComp 
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
         Left            =   5880
         TabIndex        =   190
         Top             =   4140
         Width           =   975
      End
      Begin VB.TextBox TxtSearchDegree 
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
         TabIndex        =   185
         Top             =   630
         Width           =   480
      End
      Begin VB.TextBox TxtInnerCompRatio 
         Height          =   285
         Left            =   7965
         TabIndex        =   179
         Top             =   3765
         Width           =   870
      End
      Begin VB.TextBox TxtFeedOffset 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7440
         TabIndex        =   175
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox ChkVertNoTurn 
         Alignment       =   1  'Right Justify
         Caption         =   "No Turn Angle"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7170
         TabIndex        =   36
         Top             =   6600
         Width           =   1545
      End
      Begin VB.CommandButton CmdCalculateAhead 
         Caption         =   "Calc. Encoder Ahead(Pulse)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7200
         TabIndex        =   173
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtMaterialThickMM 
         Height          =   270
         Left            =   5880
         TabIndex        =   171
         Top             =   3795
         Width           =   975
      End
      Begin VB.CheckBox ChkBenderHome 
         Alignment       =   1  'Right Justify
         Caption         =   "Bender Home Enable"
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
         Left            =   4860
         TabIndex        =   170
         Top             =   1365
         Width           =   1995
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
         Height          =   4245
         Left            =   105
         TabIndex        =   97
         Top             =   4710
         Width           =   8730
         Begin VB.TextBox TxtLinearization 
            Height          =   285
            Left            =   7845
            TabIndex        =   207
            Text            =   "Text1"
            Top             =   3870
            Width           =   795
         End
         Begin VB.TextBox TxtCutDepth2 
            Height          =   285
            Left            =   7845
            TabIndex        =   205
            Text            =   "Text1"
            Top             =   3555
            Width           =   795
         End
         Begin VB.TextBox TxtBackSet 
            Height          =   315
            Left            =   7875
            TabIndex        =   198
            Text            =   "Text1"
            Top             =   1590
            Width           =   765
         End
         Begin VB.TextBox TxtBeatPtOffset 
            Height          =   270
            Left            =   2295
            TabIndex        =   187
            Top             =   3885
            Width           =   855
         End
         Begin VB.TextBox TxtBeatAngModify 
            Height          =   270
            Left            =   2295
            TabIndex        =   186
            Top             =   3525
            Width           =   855
         End
         Begin VB.TextBox TxtCutDepth 
            Height          =   285
            Left            =   7830
            TabIndex        =   183
            Text            =   "Text1"
            Top             =   3210
            Width           =   795
         End
         Begin VB.TextBox TxtCutoffHeight 
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
            Left            =   5445
            TabIndex        =   181
            Top             =   3900
            Width           =   855
         End
         Begin VB.TextBox TxtMinContinuousMM 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7830
            TabIndex        =   176
            Text            =   "Text1"
            Top             =   2565
            Width           =   795
         End
         Begin VB.TextBox TxtTurnAngleDeg 
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
            Left            =   7830
            TabIndex        =   174
            Text            =   "Text1"
            Top             =   2250
            Width           =   795
         End
         Begin VB.CheckBox ChkVertUpDownHome1 
            Caption         =   "VertUpdown HomeEnable"
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
            Left            =   11400
            TabIndex        =   137
            Top             =   3360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox ChkBenderHome1 
            Caption         =   "Bender HomeEnable"
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
            Left            =   10320
            TabIndex        =   136
            Top             =   3360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton CmdCalculateAhead1 
            BackColor       =   &H8000000D&
            Caption         =   "Calc. Encoder Ahead(Puls)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9480
            MaskColor       =   &H00FF0000&
            TabIndex        =   135
            Top             =   1800
            UseMaskColor    =   -1  'True
            Width           =   1935
         End
         Begin VB.TextBox TxtFeedOffset1 
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
            Left            =   11880
            TabIndex        =   134
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
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
            Height          =   270
            Left            =   5445
            TabIndex        =   133
            Top             =   2940
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
            Height          =   270
            Left            =   7845
            TabIndex        =   132
            Top             =   1230
            Width           =   795
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
            Left            =   4800
            TabIndex        =   131
            Top             =   1560
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
            Left            =   2880
            TabIndex        =   130
            Top             =   1560
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
            Height          =   315
            Left            =   1920
            TabIndex        =   129
            Top             =   1560
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
            Height          =   270
            Left            =   7845
            TabIndex        =   128
            Top             =   480
            Width           =   795
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
            Left            =   -255
            TabIndex        =   127
            Top             =   3900
            Visible         =   0   'False
            Width           =   855
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
            Left            =   -255
            TabIndex        =   126
            Top             =   3585
            Visible         =   0   'False
            Width           =   855
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
            Left            =   5445
            TabIndex        =   125
            Top             =   3240
            Width           =   855
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
            Height          =   315
            Left            =   4800
            TabIndex        =   124
            Top             =   1200
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
            Height          =   315
            Left            =   2880
            TabIndex        =   123
            Top             =   1200
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
            Height          =   315
            Left            =   1920
            TabIndex        =   122
            Top             =   1200
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
            Height          =   270
            Left            =   7845
            TabIndex        =   121
            Top             =   855
            Width           =   795
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
            Left            =   7830
            TabIndex        =   120
            Top             =   2910
            Width           =   795
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
            Height          =   315
            Left            =   4800
            TabIndex        =   119
            Top             =   840
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
            Height          =   315
            Left            =   2880
            TabIndex        =   118
            Top             =   840
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
            Height          =   315
            Left            =   1920
            TabIndex        =   117
            Top             =   840
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
            Height          =   315
            Left            =   0
            TabIndex        =   116
            Top             =   2520
            Visible         =   0   'False
            Width           =   615
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
            Left            =   2295
            TabIndex        =   115
            Top             =   3240
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
            Left            =   5445
            TabIndex        =   114
            Top             =   3585
            Width           =   855
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
            Left            =   2295
            TabIndex        =   113
            Top             =   2940
            Width           =   855
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
            Height          =   315
            Left            =   4800
            TabIndex        =   112
            Top             =   2520
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
            Height          =   315
            Left            =   2880
            TabIndex        =   111
            Top             =   2520
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
            Height          =   315
            Left            =   -480
            TabIndex        =   110
            Top             =   1920
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
            Height          =   315
            Left            =   -360
            TabIndex        =   109
            Top             =   495
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
            Height          =   315
            Left            =   -480
            TabIndex        =   108
            Top             =   945
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
            Height          =   315
            Left            =   1920
            TabIndex        =   107
            Top             =   2520
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
            Height          =   315
            Left            =   4800
            TabIndex        =   106
            Top             =   2220
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
            Height          =   315
            Left            =   2880
            TabIndex        =   105
            Top             =   2220
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
            Height          =   315
            Left            =   1920
            TabIndex        =   104
            Top             =   2220
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
            Height          =   315
            Left            =   4800
            TabIndex        =   103
            Top             =   1920
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
            Height          =   315
            Left            =   2880
            TabIndex        =   102
            Top             =   1920
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
            Height          =   315
            Left            =   1920
            TabIndex        =   101
            Top             =   1920
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
            Height          =   315
            Left            =   4800
            TabIndex        =   100
            Top             =   540
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
            Height          =   315
            Left            =   2880
            TabIndex        =   99
            Top             =   540
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
            Height          =   315
            Left            =   1920
            TabIndex        =   98
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label67 
            Caption         =   "Linearization"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6555
            TabIndex        =   206
            Top             =   3915
            Width           =   1200
         End
         Begin VB.Label Label66 
            Caption         =   "Cut Depth2"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6840
            TabIndex        =   204
            Top             =   3585
            Width           =   915
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            Caption         =   "BackSet(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6180
            TabIndex        =   199
            Top             =   1605
            Width           =   1605
         End
         Begin VB.Label Label58 
            Caption         =   "Cut Depth"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6915
            TabIndex        =   184
            Top             =   3285
            Width           =   825
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            Caption         =   "Cutoff Height(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3750
            TabIndex        =   182
            Top             =   3945
            Width           =   1605
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "PointOffset(mm)"
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
            Left            =   6345
            TabIndex        =   178
            Top             =   2610
            Width           =   1365
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            Caption         =   "TurnAngle(deg)"
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
            Left            =   6435
            TabIndex        =   177
            Top             =   2295
            Width           =   1275
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            Caption         =   "Mill pre-on Stroke(mm)"
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
            Left            =   3150
            TabIndex        =   169
            Top             =   3000
            Width           =   2190
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            Caption         =   "Division Waiting(s)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6045
            TabIndex        =   168
            Top             =   1290
            Width           =   1815
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            Caption         =   "Mill Home"
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
            Left            =   390
            TabIndex        =   167
            Top             =   1605
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
            Left            =   3840
            TabIndex        =   166
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            Caption         =   "Feeding Break(mm)"
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
            Left            =   6135
            TabIndex        =   165
            Top             =   525
            Width           =   1710
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Mill Vertical"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   450
            TabIndex        =   164
            Top             =   930
            Width           =   1455
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "BeatPointOffset"
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
            Left            =   765
            TabIndex        =   163
            Top             =   3915
            Width           =   1470
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "BeatAngModify"
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
            Left            =   750
            TabIndex        =   162
            Top             =   3600
            Width           =   1485
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Blade Degree(Deg)"
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
            Left            =   3285
            TabIndex        =   161
            Top             =   3300
            Width           =   2055
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Mill Rotating"
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
            Left            =   570
            TabIndex        =   160
            Top             =   1245
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
            Height          =   315
            Left            =   3840
            TabIndex        =   159
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Space Interval(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6135
            TabIndex        =   158
            Top             =   915
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "Cutter Radius(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6375
            TabIndex        =   157
            Top             =   2970
            Width           =   1560
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
            Height          =   315
            Left            =   3840
            TabIndex        =   156
            Top             =   840
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
            Left            =   0
            TabIndex        =   155
            Top             =   5040
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
            Left            =   -480
            TabIndex        =   154
            Top             =   4920
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Mill Min Interval(mm)"
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
            Left            =   210
            TabIndex        =   153
            Top             =   3300
            Width           =   2055
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
            Height          =   315
            Left            =   3840
            TabIndex        =   152
            Top             =   2520
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
            Height          =   315
            Left            =   3840
            TabIndex        =   151
            Top             =   2220
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
            Height          =   315
            Left            =   3840
            TabIndex        =   150
            Top             =   1920
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
            Height          =   315
            Left            =   -480
            TabIndex        =   149
            Top             =   1440
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
            Height          =   315
            Left            =   3840
            TabIndex        =   148
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Bending Min Radius(mm)"
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
            Left            =   3240
            TabIndex        =   147
            Top             =   3645
            Width           =   2115
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Milling Start Angle(Deg)"
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
            Left            =   75
            TabIndex        =   146
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Bender Home"
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
            Left            =   570
            TabIndex        =   145
            Top             =   2580
            Width           =   1335
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ManualBending"
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
            Left            =   570
            TabIndex        =   144
            Top             =   2235
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
            Left            =   -360
            TabIndex        =   143
            Top             =   2880
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Acc(p/s^2)"
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
            Left            =   4860
            TabIndex        =   142
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Speed(p/s)"
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
            Left            =   3000
            TabIndex        =   141
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "StartSpd(p/s)"
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
            Left            =   1770
            TabIndex        =   140
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Bending"
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
            Left            =   570
            TabIndex        =   139
            Top             =   1935
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Feed Speed"
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
            Left            =   810
            TabIndex        =   138
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CheckBox ChkVertUpDownHome 
         Alignment       =   1  'Right Justify
         Caption         =   "VertUpdown Home Enable"
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
         Left            =   1215
         TabIndex        =   96
         Top             =   2385
         Width           =   2400
      End
      Begin VB.TextBox TxtMinBendDisMM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   94
         Top             =   1680
         Width           =   975
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
         Height          =   270
         Left            =   2610
         TabIndex        =   84
         Top             =   3450
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
         Height          =   270
         Left            =   2610
         TabIndex        =   82
         Top             =   3090
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
         Height          =   285
         Left            =   2655
         TabIndex        =   80
         Top             =   1350
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
         Left            =   13080
         TabIndex        =   70
         Top             =   1680
         Visible         =   0   'False
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
         Height          =   270
         Left            =   5880
         TabIndex        =   68
         Top             =   3090
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
         Height          =   270
         Left            =   5880
         TabIndex        =   66
         Top             =   3450
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
         Left            =   6375
         TabIndex        =   52
         Top             =   960
         Width           =   480
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
         Height          =   270
         Left            =   2640
         TabIndex        =   48
         Top             =   2025
         Width           =   975
      End
      Begin VB.TextBox TxtVertUpDownAdjustmentMM 
         Enabled         =   0   'False
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
         Left            =   2610
         TabIndex        =   45
         Top             =   2745
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
         Height          =   270
         Left            =   2640
         TabIndex        =   43
         Top             =   1680
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
         Left            =   2610
         TabIndex        =   42
         Top             =   3795
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
         TabIndex        =   39
         Top             =   2400
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
         TabIndex        =   37
         Top             =   2025
         Width           =   975
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
         Left            =   5880
         TabIndex        =   15
         Top             =   960
         Width           =   480
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
         TabIndex        =   13
         Top             =   2745
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
         Left            =   6375
         TabIndex        =   10
         Top             =   630
         Width           =   480
      End
      Begin VB.CheckBox ChkUseEncoder 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Encoder"
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
         Left            =   1890
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
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
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   975
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
         Top             =   285
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
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Height          =   2415
         Left            =   7200
         TabIndex        =   85
         Top             =   1305
         Width           =   1635
         Begin VB.TextBox TxtExtendMM 
            Enabled         =   0   'False
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
            Left            =   240
            TabIndex        =   90
            Top             =   660
            Width           =   795
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
            Left            =   240
            TabIndex        =   89
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox TxtTailVertAngle 
            Enabled         =   0   'False
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
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   735
         End
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
            Left            =   1125
            TabIndex        =   87
            Top             =   1395
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox ChkKareanMaterial 
            Caption         =   "Single Cut"
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
            TabIndex        =   86
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label45 
            Caption         =   "Splice Len(mm)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   93
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label52 
            Caption         =   "M.V.Stroke(mm)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   1800
            Width           =   1275
         End
         Begin VB.Label Label53 
            Caption         =   "End-Slot Degree"
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
            Left            =   90
            TabIndex        =   91
            Top             =   1035
            Width           =   1500
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   105
         TabIndex        =   196
         Top             =   585
         Width           =   420
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         Caption         =   "EndComp(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6855
         TabIndex        =   201
         Top             =   4515
         Width           =   1095
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         Caption         =   "PreComp(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6855
         TabIndex        =   200
         Top             =   4185
         Width           =   1095
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         Caption         =   "Outer EndComp(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4140
         TabIndex        =   195
         Top             =   4500
         Width           =   1770
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         Caption         =   "Inner EndComp(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4215
         TabIndex        =   194
         Top             =   4185
         Width           =   1695
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "Outer PreComp(mm)"
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
         Left            =   570
         TabIndex        =   193
         Top             =   4500
         Width           =   2055
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         Caption         =   "Inner PreComp(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   192
         Top             =   4185
         Width           =   2115
      End
      Begin VB.Label Label56 
         Caption         =   "Left Bend Comp.Ratio"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6975
         TabIndex        =   180
         Top             =   3720
         Width           =   960
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Comp Rad(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3975
         TabIndex        =   172
         Top             =   3855
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Mill Vertical  Adj(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   95
         Top             =   2760
         Width           =   2490
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "InnerLine Compensation(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   83
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "OuterLine Compensation(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   81
         Top             =   3120
         Width           =   2520
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "BenderSpringback"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   79
         Top             =   1365
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
         Left            =   10800
         TabIndex        =   69
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Outer Angle(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3735
         TabIndex        =   67
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Inner Angle(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3855
         TabIndex        =   65
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Mill Vertical Strock(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   47
         Top             =   2070
         Width           =   2460
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Min Len for Bender(mm)"
         Enabled         =   0   'False
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
         Left            =   3855
         TabIndex        =   46
         Top             =   1740
         Width           =   2055
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Mill Vertical Pulse/mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   570
         TabIndex        =   44
         Top             =   1725
         Width           =   2055
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Min continuous Lines Len(mm)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   41
         Top             =   3855
         Width           =   2580
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Mill R.Search Stroke(Deg)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3135
         TabIndex        =   40
         Top             =   2445
         Width           =   2775
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Mill Rotating Pulse/Deg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3720
         TabIndex        =   38
         Top             =   2070
         Width           =   2145
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Bender L/R Idle Stroke(Deg)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   4365
         TabIndex        =   14
         Top             =   945
         Width           =   1440
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Mill-Bender Diatance(mm)"
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
         Left            =   3615
         TabIndex        =   12
         Top             =   2790
         Width           =   2295
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Bender Search/Adj(Deg)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3735
         TabIndex        =   11
         Top             =   660
         Width           =   2085
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Encoder Pulse/mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Bending Pulse/Deg"
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
         Left            =   3960
         TabIndex        =   5
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Feeding Pulse/mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   585
         TabIndex        =   4
         Top             =   645
         Width           =   1950
      End
   End
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
      Left            =   15165
      TabIndex        =   57
      Top             =   5220
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox ChkLocked 
      Caption         =   "Data Locked"
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
      Left            =   210
      TabIndex        =   56
      Top             =   9345
      Value           =   1  'Checked
      Width           =   2175
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
      Left            =   15345
      TabIndex        =   53
      Top             =   7470
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   15435
      TabIndex        =   51
      Top             =   6705
      Visible         =   0   'False
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
      Left            =   15390
      TabIndex        =   25
      Top             =   7830
      Visible         =   0   'False
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
      Left            =   15480
      TabIndex        =   24
      Top             =   5850
      Visible         =   0   'False
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
      Left            =   16620
      TabIndex        =   23
      Top             =   4665
      Visible         =   0   'False
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
      Left            =   15480
      TabIndex        =   22
      Top             =   7065
      Visible         =   0   'False
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
      Left            =   15120
      TabIndex        =   21
      Top             =   8520
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
      Left            =   15480
      TabIndex        =   20
      Top             =   6300
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
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
      Left            =   11040
      TabIndex        =   2
      Top             =   9270
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
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
      Left            =   7455
      TabIndex        =   1
      Top             =   9270
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000A&
      Caption         =   "Angle(deg)/Bend Radius(mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8880
      Left            =   9120
      TabIndex        =   16
      Top             =   240
      Width           =   5775
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   78
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   77
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   76
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   75
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtBendDis 
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   73
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtMaterialThickMM1 
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
         Left            =   3720
         TabIndex        =   72
         Top             =   840
         Visible         =   0   'False
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
         Left            =   1560
         TabIndex        =   54
         Top             =   840
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
         ItemData        =   "FormSettingsV7.frx":0000
         Left            =   1560
         List            =   "FormSettingsV7.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   360
         Width           =   1455
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
         Left            =   2310
         TabIndex        =   28
         Top             =   8340
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
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
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
            TabIndex        =   33
            Top             =   0
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
            TabIndex        =   32
            Top             =   120
            Visible         =   0   'False
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
            TabIndex        =   30
            Top             =   120
            Visible         =   0   'False
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
            TabIndex        =   35
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
            TabIndex        =   31
            Top             =   180
            Visible         =   0   'False
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
            Left            =   135
            TabIndex        =   29
            Top             =   150
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton CmdSortAngleTable 
         Caption         =   "Sort"
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
         Left            =   150
         TabIndex        =   18
         Top             =   8445
         Width           =   675
      End
      Begin VB.CommandButton CmdShowCurve 
         Caption         =   "Show Curve"
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
         Left            =   900
         TabIndex        =   17
         Top             =   8445
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid GrdAngleTable 
         Height          =   7095
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   12515
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
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Section"
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
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Name"
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
         Left            =   480
         TabIndex        =   55
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Caption         =   "Section Thickness(mm)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4560
         TabIndex        =   71
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
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
      Left            =   15540
      TabIndex        =   64
      Top             =   3960
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
      Left            =   15540
      TabIndex        =   63
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
      TabIndex        =   62
      Top             =   2880
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
      Left            =   15540
      TabIndex        =   61
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
      Left            =   15540
      TabIndex        =   60
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
      Left            =   15600
      TabIndex        =   59
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
      Left            =   15600
      TabIndex        =   58
      Top             =   1260
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
      Left            =   17460
      TabIndex        =   27
      Top             =   4665
      Visible         =   0   'False
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
      Left            =   15180
      TabIndex        =   26
      Top             =   4665
      Visible         =   0   'False
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
        If curLanguage = 0 Then
            If MsgBox("错误的参数设置将导致设备运行异常。请确定是否放弃该操作？ ", vbQuestion + vbYesNo + vbSystemModal, "") = vbNo Then
                For Each obj In Me
                    obj.Enabled = True
                Next
                'ShowZAxisMode
                'ShowHeadMode
                TxtTailVertAngle.Enabled = False
                TxtExtendMM.Enabled = False
            Else
                ChkLocked.value = 1
            End If
        Else
            If MsgBox("Wrong Parameters  would let the device run abnormally.Abandon this operation？ ", vbQuestion + vbYesNo + vbSystemModal, "") = vbNo Then
                For Each obj In Me
                    obj.Enabled = True
                Next
                'ShowZAxisMode
                'ShowHeadMode
                TxtTailVertAngle.Enabled = False
                TxtExtendMM.Enabled = False
            Else
                ChkLocked.value = 1
            End If
        End If
    End If
End Sub

Private Sub ChkUseEncoder_Click()
    If ChkUseEncoder.value = 1 Then
        TxtEncoderPulsPerMM.Visible = True
        Label4.Visible = True
        'CmdCalculateAhead.Visible = True
        TxtFeedOffset.Enabled = True
    Else
        TxtEncoderPulsPerMM.Visible = False
        Label4.Visible = False
       ' CmdCalculateAhead.Visible = False
        TxtFeedOffset.Enabled = False
    End If
End Sub

Private Sub CmbMaterial_Click()
    Dim I As Long, t As Long
    
    Device_CurMaterial = "Material" + Format(CmbMaterial.ListIndex, "00")
    TxtMaterialName.Text = CmbMaterial.List(CmbMaterial.ListIndex)
    
    WritePrivateProfileString "Device", "CurMaterial", Device_CurMaterial, App.Path & "\Parameters.ini"
    
    For I = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(I, 0) = ""
        GrdAngleTable.TextMatrix(I, 1) = ""
        For t = 1 To MaxBendDisNo
            GrdAngleTable.TextMatrix(I, t + 1) = ""
        Next
    Next
    
    LoadParameters
    For I = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(I, 0) = str(I)
        GrdAngleTable.TextMatrix(I, 1) = Format(KeyAngle(I), " 0.0###")
        For t = 1 To MaxBendDisNo
            'GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0 And BendDis(t) = 0, "", Format(RealAngle(t, I), " 0.0###"))
            GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0, "", Format(RealAngle(t, I), " 0.0###"))
        Next
    Next
    
    Device_MaterialThickMM = GetValueFromINI("MaterialThickMM", Device_CurMaterial, "0.8", App.Path & "\Parameters.ini")
    TxtMaterialThickMM.Text = Format(Device_MaterialThickMM, " 0.0###")
End Sub

Private Sub CmdCalculateAhead_Click()
    Dim I As Long, ahead As Double, ahead_sum As Double
    Dim Ret As Long, nActPos As Long, Pos0 As Long, FeedPulsCount As Long, n As Long
    If curLanguage = 1 Then
        FrmMsgDlg.LblMessage.caption = "Load the material. Machine will feed automatically。"
        FrmMsgDlg.CmdClose.caption = "OK"
    Else
        FrmMsgDlg.LblMessage.caption = "请上好型材。本功能将自动进料 10 次。"
        FrmMsgDlg.CmdClose.caption = "确定"
    End If
    
    FrmMsgDlg.Show 1
    If CtrlCardType = 1 Then
    
        '送料电机基本参数：起始速度，加速度，减速度
        SetAxisStartVel_9030 0, FeedAxis, Device_FeedStartV
        SetAxisAcc_9030 0, FeedAxis, Device_FeedAccel
        SetAxisDec_9030 0, FeedAxis, Device_FeedAccel
        ''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf CtrlCardType = 4 Then ''''''''''''''
        SetAcc hDmc, FeedAxis, Device_FeedAccel
        SetDec hDmc, FeedAxis, Device_FeedAccel
    End If
    
    n = 10
    
    Do While FrmMsgDlg.Visible = True
        DoEvents
    Loop
    
    If Device_FastSpeedMinLenMM <= 50 Then
        FeedPulsCount = Device_FastSpeedMinLenMM * Device_EncoderPulsPerMM
    Else
        FeedPulsCount = 50 * Device_EncoderPulsPerMM
    End If
    
    FrmMain.TxtStatistics.Text = "Pulses to stop：" + vbCrLf + vbCrLf
    ahead_sum = 0
    For I = 1 To n
        If CtrlCardType = 0 Then
            get_actual_pos 0, FeedAxis, nActPos
        ElseIf CtrlCardType = 4 Then
            nActPos = GetPosEnc(hDmc, FeedAxis)
            FeedV_GALIL
        Else
            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
            FeedV_9030
            
        End If
        Pos0 = nActPos
        
        'DCMotorFeedFWOn
        
        
        Do
            If CtrlCardType = 0 Then
                get_actual_pos 0, FeedAxis, nActPos
            ElseIf CtrlCardType = 4 Then
                nActPos = GetPosEnc(hDmc, FeedAxis)
            
            Else
                nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
                
            End If
            
            If nActPos - Pos0 >= FeedPulsCount Then
                Pos0 = nActPos
                Exit Do
            End If
            DoEvents
        Loop
        
        'DCMotorFeedFWOff
        StopFeedV
        
        Wait 2
        
        If CtrlCardType = 0 Then
            get_actual_pos 0, FeedAxis, nActPos
        ElseIf CtrlCardType = 4 Then
            nActPos = GetPosEnc(hDmc, FeedAxis)
        Else
            nActPos = ReadAxisEncodePos_9030(0, FeedAxis)
            
        End If
        ahead = nActPos - Pos0
        FrmMain.TxtStatistics.Text = FrmMain.TxtStatistics.Text + str(ahead) + vbCrLf

        ahead_sum = ahead_sum + ahead
    Next
    
    TxtFeedOffset.Text = str(Round(ahead_sum / n, 1))
    
    IsRunning = False
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdMake_Click()
    Dim AngleDEG As Double, FeedMM As Double, n As Long, I As Long, s As Double, dw As Double
           
    AngleDEG = val(TxtAngleDeg.Text)
    'AngleDEG = AngleDEG + IIf(AngleDEG > 0, Device_EmptyDegree, IIf(AngleDEG = 0, 0, -Device_EmptyDegree))
    FeedMM = val(TxtFeedMM.Text)
    'n = Val(TxtN.Text)
    
    StopRunning = False
    IsRunning = True
    
    '复位
    BendReset
    ''铣槽-高位
    'FrmMain.Vert 1, 1
    VertInnerAngle 0, False
    
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
            VertInnerAngle 0, False
            '进料
            FeedMMByDCMotor Device_HeadDistance, 0, False
            '弯弧
            BendAngle 0
        Else
            '进料
            FeedMMByDCMotor FeedMM, 0, False
            ''铣槽-高位
            'FrmMain.Vert 1, 1
            VertInnerAngle 0, False
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
        VertInnerAngle 0, False
        '进料
        FeedMMByDCMotor -s - dw + Device_HeadDistance - 10, 0, False
        Wait 1
        '拍弧
        For I = 1 To n
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
        VertInnerAngle 0, False
        '进料
        FeedMMByDCMotor Device_HeadDistance / 4, 0, False
        ''铣槽-高位
        'FrmMain.Vert 1, 1
        VertInnerAngle 0, False
        '进料
        FeedMMByDCMotor Device_HeadDistance - Device_HeadDistance / 4, 0, False
        BeatAngle AngleDEG, True
        '进料
        FeedMMByDCMotor Device_HeadDistance / 4, 0, False
        
    ElseIf OptSymmetryTest.value = True Then
        '进料
        FeedMMByDCMotor Device_HeadDistance, 0, False
        '铣槽
        VertInnerAngle 0, False
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
    
    a = val(TxtA.Text)
    l = val(TxtL.Text)
    
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
    Dim t As Long, I As Long
    
    Device_PulsPerMM = val(TxtPulsPerMM.Text)
    Device_EncoderPulsPerMM = val(TxtEncoderPulsPerMM.Text)
    Device_UseEncoder = IIf(ChkUseEncoder.value = 1, True, False)
    
    Device_BenderHome = IIf(ChkBenderHome.value = 1, True, False)
    Device_VertUpdownHome = IIf(ChkVertUpDownHome.value = 1, True, False)
    
    Device_PulsPerDegree = val(TxtPulsPerDegree.Text)
    Device_AdjustmentDegree = val(TxtAdjustmentDegree.Text)
    Device_SearchDegree = val(TxtSearchDegree.Text)
    
    Device_EmptyDegree = val(TxtEmptyDegree.Text)
    
    'Device_AdjustmentDegree2 = Val(TxtAdjustmentDegree2.Text)
    Device_EmptyDegree2 = val(TxtEmptyDegree2.Text)
    
    'Device_VertMotorDrive = IIf(ChkVertMotorDrive.value = 1, True, False)
    'Device_VertAllHigh = IIf(ChkVertAllHigh.value = 1, True, False)
    Device_VertNoTurn = IIf(ChkVertNoTurn.value = 1, True, False)
    
    Device_VertUpDownPulsPerMM = val(TxtVertUpDownPulsPerMM.Text)
    Device_VertUpDownAdjustmentMM = val(TxtVertUpDownAdjustmentMM.Text)
    Device_MinBendDisMM = val(TxtMinBendDisMM.Text)
    
    Device_VertUpDownMM = val(TxtVertUpDownMM.Text)
    Device_InnerCompRatio = val(TxtInnerCompRatio.Text)
    
    Device_StartComp = val(TxtStartComp.Text)
    Device_EndComp = val(TxtEndComp.Text)
    Device_StartComp2 = val(TxtStartComp2.Text)
    Device_EndComp2 = val(TxtEndComp2.Text)
    
    Device_StartPointAdjustMM = val(TxtStartPtAdjustMM.Text)
    Device_EndPointAdjustMM = val(TxtEndPtAdjustMM.Text)
    
    Device_VertPulsPerDegree = val(TxtVertPulsPerDegree.Text)
    Device_VertAdjustmentDegree = val(TxtVertAdjustmentDegree.Text)
    
    Device_HeadDistance = val(TxtHeadDistance.Text)
    Device_DoneDistance = val(TxtDoneDistance.Text)
    Device_BackSet = val(TxtBackSet.Text)
    Device_DoneWaitingTime = val(TxtDoneWaitingTime.Text)
    Device_ExtendMM = val(TxtExtendMM.Text)
    
    'Device_WaitUpTime = Val(TxtWaitUpTime.Text)
    'Device_WaitDownTime = Val(TxtWaitDownTime.Text)
    
    Device_FeedStartV = val(TxtFeedStartV.Text)
    Device_FeedSpeed = val(TxtFeedSpeed.Text)
    Device_FeedAccel = val(TxtFeedAccel.Text)
    Device_FeedOffset = val(TxtFeedOffset.Text)
    
    Device_BeatAngModify = val(TxtBeatAngModify.Text)
    Device_BeatPtOffset = val(TxtBeatPtOffset.Text)
    
    'Device_ManualFeedStartV = Val(TxtManualFeedStartV.Text)
    'Device_ManualFeedSpeed = Val(TxtManualFeedSpeed.Text)
    'Device_ManualFeedAccel = Val(TxtManualFeedAccel.Text)
    'Device_ManualFeedOffset = Val(TxtManualFeedOffset.Text)
    
    Device_BendStartV = val(TxtBendStartV.Text)
    Device_BendSpeed = val(TxtBendSpeed.Text)
    Device_BendAccel = val(TxtBendAccel.Text)
    
    Device_ManualBendStartV = val(TxtManualBendStartV.Text)
    Device_ManualBendSpeed = val(TxtManualBendSpeed.Text)
    Device_ManualBendAccel = val(TxtManualBendAccel.Text)
    
    Device_ResetBendStartV = val(TxtResetBendStartV.Text)
    Device_ResetBendSpeed = val(TxtResetBendSpeed.Text)
    Device_ResetBendAccel = val(TxtResetBendAccel.Text)
    
    'Device_TurnFeedStartV = Val(TxtTurnFeedStartV.Text)
    'Device_TurnFeedSpeed = Val(TxtTurnFeedSpeed.Text)
    'Device_TurnFeedAccel = Val(TxtTurnFeedAccel.Text)
    Device_VertUpDownStartV = val(TxtTurnFeedStartV.Text)
    Device_VertUpDownSpeed = val(TxtTurnFeedSpeed.Text)
    Device_VertUpDownAccel = val(TxtTurnFeedAccel.Text)
    
    Device_TurnFeedStartV = val(TxtVertStartV.Text)
    Device_TurnFeedSpeed = val(TxtTurnFeedSpeed.Text)
    Device_TurnFeedAccel = val(TxtTurnFeedAccel.Text)
    
    Device_VertStartV = val(TxtVertStartV.Text)
    Device_VertSpeed = val(TxtVertSpeed.Text)
    Device_VertAccel = val(TxtVertAccel.Text)
    
    Device_ResetVertStartV = val(TxtResetVertStartV.Text)
    Device_ResetVertSpeed = val(TxtResetVertSpeed.Text)
    Device_ResetVertAccel = val(TxtResetVertAccel.Text)
    
    Device_VertMinAngle = val(TxtVertMinAngle.Text)
    Device_VertMinDistance = val(TxtVertMinDistance.Text)
    Device_BeatMaxRadius = val(TxtBeatMaxRadius.Text)
    
    Device_TurnFeedMM = val(TxtTurnFeedMM.Text)
    Device_CutRadiusMM = val(TxtCutRadiusMM.Text)
    Device_CutDepth = val(TxtCutDepth.Text)
    Device_CutDepth2 = val(TxtCutDepth2.Text)
    Device_Linearization = val(TxtLinearization.Text)
    
    Device_CutoffHeight = val(TxtCutoffHeight.Text)
    
    Device_TurnPointOffsetMM = val(TxtTurnPointOffsetMM.Text)
    Device_MinContinuousMM = val(TxtMinContinuousMM.Text)
    
    Device_VertKnifeDegree = val(TxtVertKnifeDegree.Text)

    Device_VertMaxOuterAngle = val(TxtVertMaxOuterAngle.Text)
    Device_VertMaxInnerAngle = val(TxtVertMaxInnerAngle.Text)
    
    Device_OuterAngleAdjustMM = val(TxtOuterAngleAdjustMM.Text)
    Device_InnerAngleAdjustMM = val(TxtInnerAngleAdjustMM.Text)
    
    Device_OuterLineTerminalAdjustMM = val(TxtOuterLineTerminalAdjustMM.Text)
    Device_InnerLineTerminalAdjustMM = val(TxtInnerLineTerminalAdjustMM.Text)
    
    Device_BenderBacklash = val(TxtBenderBacklash.Text)
    Device_BenderSpringback = val(TxtBenderSpringback.Text)
    Device_TurnAngleDeg = val(TxtTurnAngleDeg.Text)
    
    Device_FastSpeedMinLenMM = val(TxtFastSpeedMinLenMM.Text)
    Device_VertMotorZoneMM = val(TxtVertMotorZoneMM.Text)
    
    Device_AmericanMaterial = IIf(ChkAmericanMaterial.value = 1, True, False)
    Device_TailVertAngle = val(TxtTailVertAngle.Text)
    Device_VertUpDownMM_A = val(TxtVertUpDownMM_A.Text)
    Device_KareanMaterial = IIf(ChkKareanMaterial.value = 1, True, False)
    
    SetDeviceParameters
        
    'FrmMain.ChkStartPointVert90.Visible = Not Device_AmericanMaterial
    'FrmMain.ChkEndPointVert90.Visible = Not Device_AmericanMaterial
    
    '------------------------------------------------------------------
    
    CmdSortAngleTable_Click
    
    For t = 1 To MaxBendDisNo
        BendDis(t) = val(TxtBendDis(t).Text)
    Next
    
    SupplementKeyCount = 0
    For I = 1 To GrdAngleTable.Rows - 1
        If val(GrdAngleTable.TextMatrix(I, 1)) <> 0 Then
            SupplementKeyCount = I
            KeyAngle(I) = val(GrdAngleTable.TextMatrix(I, 1))
            For t = 1 To MaxBendDisNo
                RealAngle(t, I) = val(GrdAngleTable.TextMatrix(I, t + 1))
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
    For I = 1 To SupplementKeyCount
        WriteToINI_A "Key" & Trim(str(I)), str(KeyAngle(I))
        For t = 1 To MaxBendDisNo
            WriteToINI_A "Real" & Trim(str(I)) & "_" & Trim(str(t)), str(RealAngle(t, I))
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

Private Sub Command1_Click()
    FormGetPulse.Text1.Text = FormSettings.TxtPulsPerMM.Text
    FormGetPulse.caption = "Calculate Motor Pulse per mm"
    FormGetPulse.Show
End Sub

Private Sub Command2_Click()
    FormGetPulse.Text1.Text = FormSettings.TxtEncoderPulsPerMM.Text
    FormGetPulse.caption = "Calculate Encoder Pulse per mm"
    FormGetPulse.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Result = 0 Then
        If CtrlCardType = 1 Then
            CeaseAxis_9030 0, FeedAxis
            StopFeedV
        ElseIf CtrlCardType = 4 Then
            StopAxis hDmc, 0
        End If
    End If
    IsRunning = False
End Sub
Public Sub Form_Load()
    Dim t As Long, I As Long
    Dim obj As Object
    
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, SWP_Flags
    
    TxtPulsPerMM.Text = str(Device_PulsPerMM)
    TxtEncoderPulsPerMM.Text = str(Device_EncoderPulsPerMM)
    ChkUseEncoder.value = IIf(Device_UseEncoder = True, 1, 0)
    If (ChkUseEncoder.value = True) Then
        'CmdCalculateAhead.Visible = True
    Else
        'CmdCalculateAhead.Visible = False
    End If
    
    ChkBenderHome.value = IIf(Device_BenderHome = True, 1, 0)
    ChkVertUpDownHome.value = IIf(Device_VertUpdownHome = True, 1, 0)
    
    TxtPulsPerDegree.Text = str(Device_PulsPerDegree)
    
    TxtAdjustmentDegree.Text = Format(Device_AdjustmentDegree, " 0.0#######")
    TxtSearchDegree.Text = Format(Device_SearchDegree, " 0.0#######")
    
    TxtEmptyDegree.Text = Format(Device_EmptyDegree, " 0.0#######")
    
    'TxtAdjustmentDegree2.Text = Format(Device_AdjustmentDegree2, " 0.0#######")
    TxtEmptyDegree2.Text = Format(Device_EmptyDegree2, " 0.0#######")
    
    'ChkVertMotorDrive.value = IIf(Device_VertMotorDrive = True, 1, 0)
    'ChkVertAllHigh.value = IIf(Device_VertAllHigh = True, 1, 0)
    ChkVertNoTurn.value = IIf(Device_VertNoTurn = True, 1, 0)
    
    TxtVertUpDownPulsPerMM.Text = Format(Device_VertUpDownPulsPerMM, " 0.0#######")
    TxtVertUpDownAdjustmentMM.Text = Format(Device_VertUpDownAdjustmentMM, " 0.0#######")
    TxtMinBendDisMM.Text = Format(Device_MinBendDisMM, " 0.0#######")
    
    TxtVertUpDownMM.Text = Format(Device_VertUpDownMM, " 0.0#######")
    TxtInnerCompRatio.Text = Format(Device_InnerCompRatio, " 0.0#######")
    
    TxtStartComp.Text = Format(Device_StartComp, " 0.0#######")
    TxtEndComp.Text = Format(Device_EndComp, " 0.0#######")
    TxtStartComp2.Text = Format(Device_StartComp2, " 0.0#######")
    TxtEndComp2.Text = Format(Device_EndComp2, " 0.0#######")
    
    TxtStartPtAdjustMM.Text = Format(Device_StartPointAdjustMM, " 0.0#######")
    TxtEndPtAdjustMM.Text = Format(Device_EndPointAdjustMM, " 0.0#######")
    
    TxtVertPulsPerDegree.Text = Format(Device_VertPulsPerDegree, " 0.0#######")
    TxtVertAdjustmentDegree.Text = Format(Device_VertAdjustmentDegree, " 0.0#######")
    
    TxtHeadDistance.Text = Format(Device_HeadDistance, " 0.0###")
    TxtDoneDistance.Text = Format(Device_DoneDistance, " 0.0###")
    TxtBackSet.Text = Format(Device_BackSet, " 0.0###")
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
    
    TxtCutDepth.Text = Format(Device_CutDepth, " 0.0###")
    TxtCutDepth2.Text = Format(Device_CutDepth2, " 0.0###")
    TxtLinearization.Text = Format(Device_Linearization, " 0.0###")
    
    TxtCutoffHeight.Text = Format(Device_CutoffHeight, "0.0###")
    
    TxtTurnPointOffsetMM.Text = Format(Device_TurnPointOffsetMM, " 0.0###")
    TxtMinContinuousMM.Text = Format(Device_MinContinuousMM, " 0.0###")
    
    TxtVertKnifeDegree.Text = Format(Device_VertKnifeDegree, " 0.0###")
    TxtVertMaxOuterAngle.Text = Format(Device_VertMaxOuterAngle, " 0.0###")
    TxtVertMaxInnerAngle.Text = Format(Device_VertMaxInnerAngle, " 0.0###")
        
    TxtInnerAngleAdjustMM.Text = Format(Device_InnerAngleAdjustMM, " 0.0###")
    TxtOuterAngleAdjustMM.Text = Format(Device_OuterAngleAdjustMM, " 0.0###")
    
    TxtInnerLineTerminalAdjustMM.Text = Format(Device_InnerLineTerminalAdjustMM, " 0.0###")
    TxtOuterLineTerminalAdjustMM.Text = Format(Device_OuterLineTerminalAdjustMM, " 0.0###")
    
    TxtBenderBacklash.Text = Format(Device_BenderBacklash, " 0.0###")
    TxtBenderSpringback.Text = Format(Device_BenderSpringback, " 0.0###")
    TxtTurnAngleDeg.Text = Format(Device_TurnAngleDeg, " 0.0###")
    
    TxtFastSpeedMinLenMM.Text = Format(Device_FastSpeedMinLenMM, " 0.0###")
    TxtVertMotorZoneMM.Text = Format(Device_VertMotorZoneMM, " 0.0###")
    
    ChkAmericanMaterial.value = IIf(Device_AmericanMaterial = True, 1, 0)
    TxtTailVertAngle.Text = Format(Device_TailVertAngle, " 0.0###")
    TxtVertUpDownMM_A.Text = Format(Device_VertUpDownMM_A, " 0.0###")
    ChkKareanMaterial.value = IIf(Device_KareanMaterial = True, 1, 0)
    
    TxtBeatAngModify.Text = Format(Device_BeatAngModify, " 0.0###")
    TxtBeatPtOffset.Text = Format(Device_BeatPtOffset, " 0.0###")
    
    ChkUseEncoder_Click
    
'-------------------------------------------------------------------
    
    CmbMaterial.Clear
    For t = 1 To 10
        CmbMaterial.AddItem Device_MaterialName(t)
    Next
    CmbMaterial.ListIndex = val(Right(Device_CurMaterial, 2))
    
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
    If curLanguage = 0 Then
        LblString1.caption = "运动角度"
        LblString2.caption = "左弯弧半径"
        LblString3.caption = "右弯弧半径"
        LblString4.caption = "左拍弧角度"
        LblString5.caption = "右拍弧角度"
    Else
        LblString1.caption = "M.Angle"
        LblString2.caption = "L.Radius"
        LblString3.caption = "R.Radius"
        LblString4.caption = "Pat L.Deg"
        LblString5.caption = "Pat R.Deg"
    End If
    
    GrdAngleTable.TextMatrix(0, 1) = LblString1.caption
    
    GrdAngleTable.TextMatrix(0, 2) = LblString2.caption
    
    GrdAngleTable.TextMatrix(0, 3) = LblString3.caption
    
    GrdAngleTable.TextMatrix(0, 4) = LblString4.caption
    
    GrdAngleTable.TextMatrix(0, 5) = LblString5.caption
    'GrdAngleTable.TextMatrix(0, 6) = "拍弧角度"
    
    For I = 1 To SupplementKeyCount
        GrdAngleTable.TextMatrix(I, 0) = str(I)
        GrdAngleTable.TextMatrix(I, 1) = Format(KeyAngle(I), " 0.0###")
        For t = 1 To MaxBendDisNo
            'GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0 And BendDis(t) = 0, "", Format(RealAngle(t, I), " 0.0###"))
            GrdAngleTable.TextMatrix(I, t + 1) = IIf(RealAngle(t, I) = 0, "", Format(RealAngle(t, I), " 0.0###"))
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

Private Sub TxtBackSet_DblClick()
    SetDigiPad "FormSettings", "TxtBackSet"
End Sub

Private Sub TxtBeatAngModify_DblClick()
    SetDigiPad "FormSettings", "TxtBeatAngModify"
End Sub

Private Sub TxtBeatMaxRadius_DblClick()
    SetDigiPad "FormSettings", "TxtBeatMaxRadius"
End Sub

Private Sub TxtBeatPtOffset_DblClick()
    SetDigiPad "FormSettings", "TxtBeatPtOffset"
End Sub

Private Sub TxtBendAccel_DblClick()
    SetDigiPad "FormSettings", "TxtBendAccel"
End Sub

Private Sub TxtBenderBacklash_DblClick()
    SetDigiPad "FormSettings", "TxtBenderBacklash"
End Sub

Private Sub TxtBenderSpringback_DblClick()
    SetDigiPad "FormSettings", "TxtBenderSpringback"
End Sub

Private Sub TxtBendSpeed_Change()
    LblBendSpeed.caption = Format(Round(val(TxtBendSpeed.Text) / Device_PulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtBendSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtBendSpeed"
End Sub

Private Sub TxtBendStartV_DblClick()
    SetDigiPad "FormSettings", "TxtBendStartV"
End Sub

Private Sub TxtCutDepth_DblClick()
    SetDigiPad "FormSettings", "TxtCutDepth"
End Sub
Private Sub TxtCutDepth2_DblClick()
    SetDigiPad "FormSettings", "TxtCutDepth2"
End Sub
'Linearization
Private Sub TxtLinearization_DblClick()
    SetDigiPad "FormSettings", "TxtLinearization"
End Sub

Private Sub TxtCutoffHeight_DblClick()
    SetDigiPad "FormSettings", "TxtCutoffHeight"
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

Private Sub TxtEndComp_DblClick()
    SetDigiPad "FormSettings", "TxtEndComp"
End Sub


Private Sub TxtEndComp2_DblClick()
    SetDigiPad "FormSettings", "TxtEndComp2"
End Sub

Private Sub TxtEndPtAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtEndPtAdjustMM"
End Sub

Private Sub TxtExtendMM_DblClick()
    SetDigiPad "FormSettings", "TxtExtendMM"
End Sub

Private Sub TxtFastSpeedMinLenMM_DblClick()
    SetDigiPad "FormSettings", "TxtFastSpeedMinLenMM"
End Sub

Private Sub TxtFeedAccel_DblClick()
    SetDigiPad "FormSettings", "TxtFeedAccel"
End Sub

Private Sub TxtFeedMM_DblClick()
    SetDigiPad "FormSettings", "TxtFeedMM"
End Sub

Private Sub TxtFeedOffset_DblClick()
    SetDigiPad "FormSettings", "TxtFeedOffset"
End Sub

Private Sub TxtFeedSpeed_Change()
    'LblFeedSpeed.caption = Format(Round(60 * Val(TxtFeedSpeed.Text) / Device_PulsPerMM / 1000, 2), " 0.0## m/min")
    LblFeedSpeed.caption = Format(Round(val(TxtFeedSpeed.Text) / Device_PulsPerMM, 2), " 0.0## mm/s")
End Sub

Private Sub TxtFeedSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtFeedSpeed"
End Sub

Private Sub TxtFeedStartV_DblClick()
    SetDigiPad "FormSettings", "TxtFeedStartV"
End Sub

Private Sub TxtHeadDistance_DblClick()
    SetDigiPad "FormSettings", "TxtHeadDistance"
End Sub

Private Sub TxtInnerAngleAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtInnerAngleAdjustMM"
End Sub

Private Sub TxtInnerCompRatio_DblClick()
    SetDigiPad "FormSettings", "TxtInnerCompRatio"
End Sub

Private Sub TxtInnerLineTerminalAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtInnerLineTerminalAdjustMM"
End Sub

Private Sub TxtManualBendAccel_DblClick()
    SetDigiPad "FormSettings", "TxtManualBendAccel"
End Sub

Private Sub TxtManualBendSpeed_Change()
    LblManualBendSpeed.caption = Format(Round(val(TxtManualBendSpeed.Text) / Device_PulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtManualBendSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtManualBendSpeed"
End Sub

Private Sub TxtManualBendStartV_DblClick()
    SetDigiPad "FormSettings", "TxtManualBendStartV"
End Sub

Private Sub TxtManualFeedSpeed_Change()
    LblManualFeedSpeed.caption = Format(Round(60 * val(TxtManualFeedSpeed.Text) / Device_PulsPerMM / 1000, 2), " 0.0## m/m")
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
    CmbMaterial.ListIndex = val(Right(Device_CurMaterial, 2))
    
End Sub

Private Sub TxtMaterialThickMM_Change()
    Device_MaterialThickMM = val(TxtMaterialThickMM.Text)
    WritePrivateProfileString "MaterialThickMM", Device_CurMaterial, str(Device_MaterialThickMM), App.Path & "\Parameters.ini"
End Sub

Private Sub TxtMaterialThickMM_DblClick()
    SetDigiPad "FormSettings", "TxtMaterialThickMM"
End Sub

Private Sub TxtMinBendDisMM_DblClick()
    SetDigiPad "FormSettings", "TxtMinBendDisMM"
End Sub

Private Sub TxtMinContinuousMM_DblClick()
    SetDigiPad "FormSettings", "TxtMinContinuousMM"

End Sub

Private Sub TxtOuterAngleAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtOuterAngleAdjustMM"
End Sub

Private Sub TxtOuterLineTerminalAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtOuterLineTerminalAdjustMM"
End Sub

Private Sub TxtPulsPerDegree_DblClick()
    SetDigiPad "FormSettings", "TxtPulsPerDegree"
End Sub

Private Sub TxtPulsPerMM_DblClick()
    SetDigiPad "FormSettings", "TxtPulsPerMM"
End Sub

Private Sub TxtResetBendAccel_DblClick()
    SetDigiPad "FormSettings", "TxtResetBendAccel"
End Sub

Private Sub TxtResetBendSpeed_Change()
    LblResetBendSpeed.caption = Format(Round(val(TxtResetBendSpeed.Text) / Device_PulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub CmdSortAngleTable_Click()
    Dim r As Long, r2 As Long, c As Long, c2 As Long, a As Double, b As Double, s As String
    
    For c = 1 To MaxBendDisNo
        If val(TxtBendDis(c).Text) = 0 Then
            TxtBendDis(c).Text = str(10000 + c)
        End If
    Next
    
    For c = 1 To MaxBendDisNo - 1
        For c2 = c + 1 To MaxBendDisNo
            a = val(TxtBendDis(c).Text)
            b = val(TxtBendDis(c2).Text)
                
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
        If val(TxtBendDis(c).Text) > 10000 Then
            TxtBendDis(c).Text = ""
        End If
    Next
    
    
    For r = 1 To GrdAngleTable.Rows - 2
        For r2 = r + 1 To GrdAngleTable.Rows - 1
            a = val(GrdAngleTable.TextMatrix(r, 1))
            b = val(GrdAngleTable.TextMatrix(r2, 1))
            
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
        
        a = val(GrdAngleTable.TextMatrix(r, 1))
        GrdAngleTable.TextMatrix(r, 1) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = val(GrdAngleTable.TextMatrix(r, 2))
        GrdAngleTable.TextMatrix(r, 2) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = val(GrdAngleTable.TextMatrix(r, 3))
        GrdAngleTable.TextMatrix(r, 3) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = val(GrdAngleTable.TextMatrix(r, 4))
        GrdAngleTable.TextMatrix(r, 4) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = val(GrdAngleTable.TextMatrix(r, 5))
        GrdAngleTable.TextMatrix(r, 5) = IIf(a = 0, "", Format(a, " 0.0###"))
        
        a = val(GrdAngleTable.TextMatrix(r, 6))
        GrdAngleTable.TextMatrix(r, 6) = IIf(a = 0, "", Format(a, " 0.0###"))
    Next
    
    For r = GrdAngleTable.Rows - 1 To 1 Step -1
        If val(GrdAngleTable.TextMatrix(r, 1)) = 0 Then
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
        If val(GrdAngleTable.TextMatrix(r, 1)) = 0 Then
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
    
    If val(GrdAngleTable.TextMatrix(CurRow, 1)) < 2 And Trim(GrdAngleTable.TextMatrix(CurRow, CurCol)) <> "" Then
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
    LblResetVertSpeed.caption = Format(Round(val(TxtResetVertSpeed.Text) / Device_VertPulsPerDegree, 2), " 0.0## d/s")
End Sub

Private Sub TxtResetVertSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtResetVertSpeed"
End Sub

Private Sub TxtResetVertStartV_DblClick()
    SetDigiPad "FormSettings", "TxtResetVertStartV"
End Sub

Private Sub TxtSearchDegree_DblClick()
    SetDigiPad "FormSettings", "TxtSearchDegree"
End Sub

Private Sub TxtStartComp_DblClick()
    SetDigiPad "FormSettings", "TxtStartComp"
End Sub

Private Sub TxtStartComp2_DblClick()
    SetDigiPad "FormSettings", "TxtStartComp2"
End Sub

Private Sub TxtStartPtAdjustMM_DblClick()
    SetDigiPad "FormSettings", "TxtStartPtAdjustMM"
End Sub

Private Sub TxtTailVertAngle_DblClick()
    SetDigiPad "FormSettings", "TxtTailVertAngle"
End Sub

Private Sub TxtTurnAngleDeg_DblClick()
    SetDigiPad "FormSettings", "TxtTurnAngleDeg"
End Sub

Private Sub TxtTurnFeedAccel_DblClick()
    SetDigiPad "FormSettings", "TxtTurnFeedAccel"
End Sub

Private Sub TxtTurnFeedSpeed_Change()
    LblTurnFeedSpeed.caption = Format(Round(val(TxtTurnFeedSpeed.Text) / Device_VertUpDownPulsPerMM, 2), " 0.0## mm/s")
End Sub

Private Sub TxtTurnFeedSpeed_DblClick()
    SetDigiPad "FormSettings", "TxtTurnFeedSpeed"
End Sub

Private Sub TxtTurnFeedStartV_DblClick()
    SetDigiPad "FormSettings", "TxtTurnFeedStartV"
End Sub

Private Sub TxtTurnPointOffsetMM_DblClick()
    SetDigiPad "FormSettings", "TxtTurnPointOffsetMM"
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
    LblVertSpeed.caption = Format(Round(val(TxtVertSpeed.Text) / Device_VertPulsPerDegree, 2), " 0.0## d/s")
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
