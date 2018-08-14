VERSION 5.00
Begin VB.Form FrmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入密码"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox TxtPW 
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
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "输入密码"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim pw As String
    
    pw = GetRegString(HKEY_CURRENT_USER, "Software\" & APPSerial & "\" & DeviceModel & "\Registry", "PW")
    
    If pw = "" Then
        pw = "wfhdwz2011"
    End If
    
    If LCase(Trim(TxtPW.Text)) = LCase(pw) Then
        FrmMain.Enabled = True
        Unload Me
    Else
        MsgBox "密码输入错误 ！", vbCritical + vbOKOnly + vbSystemModal, ""
        End
    End If
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_Flags
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FrmMain.Enabled = False Then
        End
    End If
End Sub
