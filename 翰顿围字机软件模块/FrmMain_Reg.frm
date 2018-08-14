VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "HD_WZ 注册管理程序"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6030
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "修改密码"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "产生注册码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   5535
      Begin VB.TextBox TxtCS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TxtRC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   4935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "产生注册码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox TxtMC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox TxtSN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label5 
         Caption         =   "输入校验码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "注册码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "输入特征值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "输入序列号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "产生序列号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "产生序列号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtNSN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox TxtN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "输入顺序号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    TxtNSN.Text = CalcSN(Val(TxtN.Text))
End Sub

Private Sub Command2_Click()
    Dim n As Currency
    Dim cs As String
    
    n = CheckSN(TxtSN.Text)
    If n < 0 Then
        MsgBox "序列号错误 ！", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    cs = GetCRC(SN & APPSerial & mc)
    If Trim(TxtCS.Text) <> cs Then
        MsgBox "特征值或校验码错误 ！", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    TxtRC.Text = CalcRN(TxtSN.Text, TxtMC.Text)
End Sub

Private Sub Command3_Click()
    FrmEditPassword.Show
    FrmEditPassword.SetFocus
End Sub

Private Sub Form_Load()
    Me.Enabled = False
    
    FrmPassword.Show
    FrmPassword.SetFocus
End Sub

Private Sub TxtCS_Change()
    TxtRC.Text = ""
End Sub

Private Sub TxtMC_Change()
    mc = TxtMC.Text
    mc = Replace(mc, vbCr, "")
    mc = Replace(mc, vbLf, "")
    TxtRC.Text = ""
End Sub

Private Sub TxtSN_Change()
    SN = TxtSN.Text
    SN = Replace(SN, vbCr, "")
    SN = Replace(SN, vbLf, "")
    TxtRC.Text = ""
End Sub
