VERSION 5.00
Begin VB.Form FormGetPulse 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandOK 
      Caption         =   "Apply"
      Height          =   570
      Left            =   2805
      TabIndex        =   9
      Top             =   1905
      Width           =   1530
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2085
      Width           =   1170
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Text            =   "499"
      Top             =   1490
      Width           =   1170
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Text            =   "500"
      Top             =   895
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   300
      Width           =   1170
   End
   Begin VB.CommandButton CommandGetPulse 
      Caption         =   "Calculate"
      Height          =   570
      Left            =   2805
      TabIndex        =   3
      Top             =   945
      Width           =   1530
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Calc. Pulse"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   2130
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Calc. Length"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   1545
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Set Length"
      Height          =   345
      Left            =   210
      TabIndex        =   1
      Top             =   975
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Set Pulse"
      Height          =   300
      Left            =   225
      TabIndex        =   0
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "FormGetPulse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置

Private Sub CommandGetPulse_Click()
    Dim d1 As Double
    Dim d2 As Double
    Dim d3 As Double
        d1 = val(Text1.Text)
        d2 = val(Text2.Text)
        d3 = val(Text3.Text)
        Text4.Text = str(Round(d1 * d2 / d3, 3))
End Sub

Private Sub CommandOK_Click()
   If FormGetPulse.caption = "Calculate Motor Pulse per mm" Then
        FormSettings.TxtPulsPerMM.Text = Text4.Text
    Else
        FormSettings.TxtEncoderPulsPerMM.Text = Text4.Text
    End If
    Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为总在最前
End Sub

Private Sub Text1_DblClick()
    SetDigiPad "FormGetPulse", "Text1"

End Sub

Private Sub Text2_DblClick()
    SetDigiPad "FormGetPulse", "Text2"

End Sub

Private Sub Text3_DblClick()
    SetDigiPad "FormGetPulse", "Text3"
End Sub
