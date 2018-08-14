VERSION 5.00
Begin VB.Form FrmShowCurve 
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   958
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   8280
      Left            =   120
      ScaleHeight     =   8220
      ScaleWidth      =   14100
      TabIndex        =   0
      Top             =   120
      Width           =   14160
   End
   Begin VB.Label Label5 
      Caption         =   "ÅÄ»¡3"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "ÓÒÅÄ»¡"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "×óÅÄ»¡"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "ÓÒÍä»¡"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "×óÍä»¡"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   8520
      Width           =   615
   End
End
Attribute VB_Name = "FrmShowCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim I As Long, i0 As Long, t As Long, clr(10) As Long, max_angle As Double, max_1 As Double, d As Double, cm As Long
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_Flags
    
    For I = 1 To SupplementKeyCount
        If I = 1 Then
            max_angle = KeyAngle(1)
        Else
            max_angle = Max(max_angle, KeyAngle(I))
        End If
    Next
    
    cm = Int(max_angle + 1)
    
    Me.Picture1.ScaleMode = 0
    Me.Picture1.ScaleLeft = -1
    Me.Picture1.ScaleWidth = cm + 2
    Me.Picture1.ScaleTop = 100
    Me.Picture1.ScaleHeight = -110
    
    Me.Picture1.ForeColor = RGB(200, 200, 200)
    Me.Picture1.Line (0, 0)-(cm, 0), RGB(200, 200, 200)
    Me.Picture1.Line (cm, 0)-(cm, 90), RGB(200, 200, 200)
    Me.Picture1.Line (cm, 90)-(0, 90), RGB(200, 200, 200)
    Me.Picture1.Line (0, 90)-(0, 0), RGB(200, 200, 200)
    
    clr(1) = RGB(255, 0, 0)
    clr(2) = RGB(0, 255, 0)
    clr(3) = RGB(0, 0, 255)
    clr(4) = RGB(255, 0, 255)
    clr(5) = RGB(0, 255, 255)
    clr(6) = RGB(255, 255, 0)
    
    Label1.ForeColor = clr(1)
    Label2.ForeColor = clr(2)
    Label3.ForeColor = clr(3)
    Label4.ForeColor = clr(4)
    Label5.ForeColor = clr(5)
    
    For I = 10 To 80 Step 10
        Me.Picture1.Line (0, I)-(cm, I), RGB(100, 100, 100)
    Next
    For I = 1 To cm - 1
        Me.Picture1.Line (I, 0)-(I, 90), RGB(100, 100, 100)
        
        Me.Picture1.CurrentX = I - 0.3
        Me.Picture1.CurrentY = -0.3
        Me.Picture1.Print Str(I)
    Next
    
    For I = 1 To SupplementKeyCount
        If I = 1 Then
            max_1 = RealAngle(1, I)
        Else
            max_1 = Max(max_1, RealAngle(1, I))
        End If
    Next
    max_1 = Max(max_1, 1)
    
    For t = 1 To MaxBendDisNo
        If t = 1 Or t = 2 Then
            d = 90 / max_1
        Else
            d = 1
        End If
        
        i0 = 0
        For I = 1 To SupplementKeyCount
            If RealAngle(t, I) > 0 Then
                Me.Picture1.Line (KeyAngle(I), 0)-(KeyAngle(I), d * RealAngle(t, I)), RGB(128, 128, 128) 'y×ø±êÏß
                Me.Picture1.Circle (KeyAngle(I), d * RealAngle(t, I)), 0.05, clr(t)
                If I > 1 And RealAngle(t, i0) > 0 Then
                    Me.Picture1.Line (KeyAngle(i0), d * RealAngle(t, i0))-(KeyAngle(I), d * RealAngle(t, I)), clr(t)
                End If
                i0 = I
            End If
        Next
    Next
End Sub
