VERSION 5.00
Begin VB.Form FrmDigiPad 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HD_WZ_V8.PanButton PanButton13 
      Height          =   825
      Left            =   3240
      TabIndex        =   14
      Top             =   1320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1455
      Caption         =   "Bksp"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton1 
      Height          =   840
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "1"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin VB.TextBox TxtEdit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin HD_WZ_V8.PanButton PanButton2 
      Height          =   840
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "2"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton12 
      Height          =   840
      Left            =   2295
      TabIndex        =   4
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "."
      FontName        =   "黑体"
      FontSize        =   40
   End
   Begin HD_WZ_V8.PanButton PanButton3 
      Height          =   840
      Left            =   2295
      TabIndex        =   5
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "3"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton4 
      Height          =   840
      Left            =   345
      TabIndex        =   6
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "4"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton5 
      Height          =   840
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "5"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton6 
      Height          =   840
      Left            =   2295
      TabIndex        =   8
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "6"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton7 
      Height          =   840
      Left            =   345
      TabIndex        =   9
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "7"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton8 
      Height          =   840
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "8"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton9 
      Height          =   840
      Left            =   2295
      TabIndex        =   11
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "9"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton10 
      Height          =   840
      Left            =   1320
      TabIndex        =   12
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "0"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton11 
      Height          =   840
      Left            =   345
      TabIndex        =   13
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1482
      Caption         =   "-"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton16 
      Height          =   825
      Left            =   3240
      TabIndex        =   15
      Top             =   4215
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1455
      Caption         =   "OK"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton14 
      Height          =   825
      Left            =   3240
      TabIndex        =   16
      Top             =   2295
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1455
      Caption         =   "Del"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin HD_WZ_V8.PanButton PanButton15 
      Height          =   825
      Left            =   3240
      TabIndex        =   17
      Top             =   3270
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1455
      Caption         =   "Clr"
      FontName        =   "黑体"
      FontSize        =   24
   End
   Begin VB.Label LblEdit 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FrmDigiPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, SWP_Flags
    
    'Me.ZOrder 0
End Sub

Private Sub Form_Load()
    If curLanguage = 0 Then
        PanButton13.caption = "回格"
        PanButton14.caption = "删除"
        PanButton15.caption = "清空"
        PanButton16.caption = "确认"
    End If
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, SWP_Flags

End Sub

Private Sub PanButton11_Click()
    TxtEdit.SelText = "-"
End Sub

Private Sub PanButton12_Click()
    TxtEdit.SelText = "."
End Sub

Private Sub PanButton13_Click()
    If TxtEdit.SelStart > 0 Then
        TxtEdit.SelStart = TxtEdit.SelStart - 1
        TxtEdit.SelLength = 1
        TxtEdit.SelText = ""
    End If
    TxtEdit.SetFocus
End Sub

Private Sub PanButton14_Click()
    TxtEdit.SelLength = 1
    TxtEdit.SelText = ""
    TxtEdit.SetFocus
End Sub

Private Sub PanButton15_Click()
    TxtEdit.Text = ""
End Sub

Private Sub PanButton1_Click()
    TxtEdit.SelText = "1"
End Sub

Private Sub PanButton16_Click()
'    Dim idx As Long
'
'    idx = Val(Me.Tag)
'
'    FrmMain.TxtEdit(idx).Text = str(Val(TxtEdit.Text))
'    FrmMain.TxtEdit(idx).SelStart = Len(FrmMain.TxtEdit(idx).Text)
'    FrmMain.TxtEdit(idx).SetFocus
'    Unload Me

    Dim obj As Object
    
    On Error Resume Next
    
    If Me.Tag = "FrmMain" Then
        For Each obj In FrmMain
            If UCase(obj.Name) = UCase(TxtEdit.Tag) Then
                If TypeOf obj Is TextBox Then
                    obj.Text = TxtEdit.Text
                End If
                Exit For
            End If
        Next
        
    ElseIf Me.Tag = "FormSettings" Then
        For Each obj In FormSettings
            If UCase(obj.Name) = UCase(TxtEdit.Tag) Then
                If TypeOf obj Is TextBox Then
                    obj.Text = TxtEdit.Text
                ElseIf TypeOf obj Is MSFlexGrid Then
                    obj.TextMatrix(val(PanButton1.Tag), val(PanButton2.Tag)) = TxtEdit.Text
                End If
                Exit For
            End If
        Next
    ElseIf Me.Tag = "FormGetPulse" Then
        For Each obj In FormGetPulse
            If UCase(obj.Name) = UCase(TxtEdit.Tag) Then
                If TypeOf obj Is TextBox Then
                    obj.Text = TxtEdit.Text
                ElseIf TypeOf obj Is MSFlexGrid Then
                    obj.TextMatrix(val(PanButton1.Tag), val(PanButton2.Tag)) = TxtEdit.Text
                End If
                Exit For
            End If
        Next
        
    End If
    Unload Me
End Sub

Private Sub PanButton17_Click()

End Sub

Private Sub PanButton2_Click()
    TxtEdit.SelText = "2"
End Sub

Private Sub PanButton3_Click()
    TxtEdit.SelText = "3"
End Sub

Private Sub PanButton4_Click()
    TxtEdit.SelText = "4"
End Sub

Private Sub PanButton5_Click()
    TxtEdit.SelText = "5"
End Sub

Private Sub PanButton6_Click()
    TxtEdit.SelText = "6"
End Sub

Private Sub PanButton7_Click()
    TxtEdit.SelText = "7"
End Sub

Private Sub PanButton8_Click()
    TxtEdit.SelText = "8"
End Sub

Private Sub PanButton9_Click()
    TxtEdit.SelText = "9"
End Sub

Private Sub PanButton10_Click()
    TxtEdit.SelText = "0"
End Sub


