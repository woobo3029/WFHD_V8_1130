VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "内部测试"
   ClientHeight    =   7110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox TxtTest 
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    TxtTest.Text = ""
End Sub

Private Sub Form_Load()
    FrmMain.FrmTestVisible = True
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, SWP_Flags
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.FrmTestVisible = False
End Sub
