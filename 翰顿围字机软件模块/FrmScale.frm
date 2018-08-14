VERSION 5.00
Begin VB.Form FrmScale 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3300
   Icon            =   "FrmScale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtScaleY 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "100"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox TxtScaleX 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "100"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Y 方向缩放比例(%)"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "X 方向缩放比例(%)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim I As Long, sx As Double, sy As Double
    
    sx = Val(TxtScaleX.Text)
    sy = Val(TxtScaleY.Text)
    
    For I = 1 To PointCount
        PointList(I).X = PointList(I).X * sx / 100
        PointList(I).Y = PointList(I).Y * sy / 100
    Next
    For I = 1 To ArcCount
        If ArcList(I).ax_angle = 0 Or sx = sy Then
            ArcList(I).X = ArcList(I).X * sx / 100
            ArcList(I).Y = ArcList(I).Y * sy / 100
            
            ArcList(I).a = ArcList(I).a * sx / 100
            ArcList(I).B = ArcList(I).B * sy / 100
        Else
            '此处应作更全面的处理
            
            ArcList(I).X = ArcList(I).X * sx / 100
            ArcList(I).Y = ArcList(I).Y * sx / 100
            
            ArcList(I).a = ArcList(I).a * sx / 100
            ArcList(I).B = ArcList(I).B * sx / 100
        End If
    Next
    
    For I = 1 To OutputStartPointList.Count
        OutputStartPointList.leading_point0(I).X = OutputStartPointList.leading_point0(I).X * sx / 100
        OutputStartPointList.leading_point0(I).Y = OutputStartPointList.leading_point0(I).Y * sy / 100
        
        OutputStartPointList.leading_point1(I).X = OutputStartPointList.leading_point1(I).X * sx / 100
        OutputStartPointList.leading_point1(I).Y = OutputStartPointList.leading_point1(I).Y * sy / 100
    Next
                
    Unload Me
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

