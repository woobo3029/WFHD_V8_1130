VERSION 5.00
Begin VB.Form FrmShowPointList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   Icon            =   "FrmShowPointList.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPointList 
      Height          =   7335
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   10695
   End
End
Attribute VB_Name = "FrmShowPointList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim ff As Long
    Dim s As String, txt As String
    
    CalculateAllPath
    
    TxtPointList.Visible = False
    ff = FreeFile
    txt = ""
    Open "c:\hd_debug\" + "vertPoint.txt" For Input As ff
    Do While Not EOF(ff)
        Line Input #ff, s
        txt = txt & s & vbCrLf
        DoEvents
    Loop
    Close #ff
    TxtPointList.Text = txt
    TxtPointList.Visible = True
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_Flags
End Sub

Private Sub TxtPointList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ff As Long
    Dim s As String, txt As String
        
    If Shift > 0 Then
        TxtPointList.Visible = False
        txt = ""
        ff = FreeFile
        Open "c:\hd_debug\" + "OutputPointlist.txt" For Input As ff
        Do While Not EOF(ff)
            Line Input #ff, s
            txt = txt & s & vbCrLf
            DoEvents
        Loop
        Close #ff
        TxtPointList.Text = txt
        TxtPointList.Visible = True
        
        Me.Refresh
    End If
End Sub
