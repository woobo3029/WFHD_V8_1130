VERSION 5.00
Begin VB.Form FrmRotate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "旋转"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3330
   Icon            =   "FrmRotate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtRotateA 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "0.0"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "旋转角度(逆时针方向)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "FrmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    
    Dim i As Long, k As Integer
    Dim MinX As Double, MaxX As Double, MeanX As Double
    Dim MinY As Double, MaxY As Double, MeanY As Double
    Dim angle As Double, CS As Double, SN As Double, x0 As Double, y0 As Double
    
    angle = Val(TxtRotateA.Text) * PI_180
    
    CS = Cos(angle)
    SN = Sin(angle)
    
    MeanX = (ViewMinX + ViewMaxX) / 2
    MeanY = (ViewMinY + ViewMaxY) / 2
    
    'shift to org
    '-------------------------------------------------
    For i = 1 To PointCount
        PointList(i).X = PointList(i).X - MeanX
        PointList(i).Y = PointList(i).Y - MeanY
    Next
    For i = 1 To ArcCount
        ArcList(i).X = ArcList(i).X - MeanX
        ArcList(i).Y = ArcList(i).Y - MeanY
    Next
    For i = 1 To OutputStartPointList.Count
        OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X - MeanX
        OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y - MeanY
        
        OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X - MeanX
        OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y - MeanY
    Next
    '-------------------------------------------------
    
    'rotate
    '-------------------------------------------------
    For i = 1 To PointCount
        x0 = PointList(i).X
        y0 = PointList(i).Y
        PointList(i).X = (CS * x0) - (SN * y0)
        PointList(i).Y = (SN * x0) + (CS * y0)
    Next
    For i = 1 To ArcCount
        x0 = ArcList(i).X
        y0 = ArcList(i).Y
        ArcList(i).X = (CS * x0) - (SN * y0)
        ArcList(i).Y = (SN * x0) + (CS * y0)
        
        ArcList(i).ax_angle = ArcList(i).ax_angle + angle
    Next
    For i = 1 To OutputStartPointList.Count
        x0 = OutputStartPointList.leading_point0(i).X
        y0 = OutputStartPointList.leading_point0(i).Y
        OutputStartPointList.leading_point0(i).X = (CS * x0) - (SN * y0)
        OutputStartPointList.leading_point0(i).Y = (SN * x0) + (CS * y0)
        
        x0 = OutputStartPointList.leading_point1(i).X
        y0 = OutputStartPointList.leading_point1(i).Y
        OutputStartPointList.leading_point1(i).X = (CS * x0) - (SN * y0)
        OutputStartPointList.leading_point1(i).Y = (SN * x0) + (CS * y0)
    Next
    '-------------------------------------------------
    
    'shift back
    '-------------------------------------------------
    For i = 1 To PointCount
        PointList(i).X = PointList(i).X + MeanX
        PointList(i).Y = PointList(i).Y + MeanY
    Next
    For i = 1 To ArcCount
        ArcList(i).X = ArcList(i).X + MeanX
        ArcList(i).Y = ArcList(i).Y + MeanY
    Next
    For i = 1 To OutputStartPointList.Count
        OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X + MeanX
        OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y + MeanY
        
        OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X + MeanX
        OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y + MeanY
    Next
    '-------------------------------------------------
    
    Unload Me
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

