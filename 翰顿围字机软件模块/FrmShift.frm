VERSION 5.00
Begin VB.Form FrmShift 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "平移"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3405
   Icon            =   "FrmShift.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox TxtShiftDY 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtShiftDX 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Y 方向平移量(mm)"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "X 方向平移量(mm)"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim i As Long
    
    For i = 1 To PointCount
        PointList(i).X = PointList(i).X + Val(TxtShiftDX)
        PointList(i).Y = PointList(i).Y + Val(TxtShiftDY)
    Next
    For i = 1 To ArcCount
        ArcList(i).X = ArcList(i).X + Val(TxtShiftDX)
        ArcList(i).Y = ArcList(i).Y + Val(TxtShiftDY)
    Next
    For i = 1 To OutputStartPointList.Count
        OutputStartPointList.leading_point0(i).X = OutputStartPointList.leading_point0(i).X + Val(TxtShiftDX)
        OutputStartPointList.leading_point0(i).Y = OutputStartPointList.leading_point0(i).Y + Val(TxtShiftDY)
        
        OutputStartPointList.leading_point1(i).X = OutputStartPointList.leading_point1(i).X + Val(TxtShiftDX)
        OutputStartPointList.leading_point1(i).Y = OutputStartPointList.leading_point1(i).Y + Val(TxtShiftDY)
    Next
    
    Unload Me
    
    FrmMain.PicPathCls
    DrawAll
    
    SaveUndo
End Sub

