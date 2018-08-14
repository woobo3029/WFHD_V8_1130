VERSION 5.00
Begin VB.Form FrmShapeLib 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " 图库"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ShowInTaskbar   =   0   'False
   Begin HD_WZ_V500.PanButton CmdCreateParallelRectangle 
      Height          =   480
      Left            =   1635
      TabIndex        =   13
      Top             =   2115
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":0000
      Caption         =   "平行四边形"
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateMultiPointStar 
      Height          =   480
      Left            =   1620
      TabIndex        =   12
      Top             =   3630
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":56BB
      Caption         =   " 多角星  "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreate5PointStar 
      Height          =   480
      Left            =   90
      TabIndex        =   11
      Top             =   3615
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":AE81
      Caption         =   " 五角星  "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateTriangle2 
      Height          =   480
      Left            =   105
      TabIndex        =   10
      Top             =   1605
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":1064E
      Caption         =   "直角三角形"
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateTriangle1 
      Height          =   480
      Left            =   1635
      TabIndex        =   9
      Top             =   1095
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":15CC7
      Caption         =   "等腰三角形"
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateTriangle 
      Height          =   480
      Left            =   105
      TabIndex        =   8
      Top             =   1110
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":1B34A
      Caption         =   "等边三角形"
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateCircle 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":209BF
      Caption         =   " 圆       "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateArc 
      Height          =   480
      Left            =   1635
      TabIndex        =   1
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":26020
      Caption         =   " 圆弧     "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateSector 
      Height          =   480
      Left            =   1635
      TabIndex        =   2
      Top             =   615
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":2B711
      Caption         =   " 扇形     "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateEllipse 
      Height          =   480
      Left            =   105
      TabIndex        =   3
      Top             =   615
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":30D84
      Caption         =   " 椭圆    "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateSquare 
      Height          =   480
      Left            =   105
      TabIndex        =   4
      Top             =   2100
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":3640F
      Caption         =   " 正方形  "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreatePolygon 
      Height          =   480
      Left            =   105
      TabIndex        =   5
      Top             =   3105
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":3BA86
      Caption         =   " 正多边形"
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateTrapezoid 
      Height          =   480
      Left            =   105
      TabIndex        =   6
      Top             =   2595
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":41271
      Caption         =   " 等腰梯形"
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateRectangle 
      Height          =   480
      Left            =   1635
      TabIndex        =   7
      Top             =   1620
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":469AE
      Caption         =   " 矩形    "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateCurvePolygon 
      Height          =   480
      Left            =   1635
      TabIndex        =   14
      Top             =   3120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":4C00D
      Caption         =   " 正曲边形 "
      FontSize        =   0
   End
   Begin HD_WZ_V500.PanButton CmdCreateTrapezoid1 
      Height          =   480
      Left            =   1635
      TabIndex        =   15
      Top             =   2610
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Picture         =   "FrmShapeLib.frx":517EA
      Caption         =   " 直角梯形"
      FontSize        =   0
   End
End
Attribute VB_Name = "FrmShapeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdCreate5PointStar_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "长径"
    tv(2).v = 0
    tv(3).t = "短径(0)"
    tv(3).v = 0
    
    ShowEditData "创建五角星", 4, tv, ToolType.Create5PointStar
End Sub

Private Sub CmdCreateCurvePolygon_Click()
    ReDim tv(5) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "边数"
    tv(2).v = 0
    tv(3).t = "半径"
    tv(3).v = 0
    tv(4).t = "径距"
    tv(4).v = 0
    
    ShowEditData "创建正曲多形", 5, tv, ToolType.CreateCurvePolygon
End Sub

Private Sub CmdCreateMultiPointStar_Click()
    ReDim tv(5) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "边数"
    tv(2).v = 0
    tv(3).t = "长径"
    tv(3).v = 0
    tv(4).t = "短径(0)"
    tv(4).v = 0
    
    ShowEditData "创建多角星", 5, tv, ToolType.CreateMultiPointStar
End Sub

Private Sub CmdCreateTriangle_Click()
    ReDim tv(3) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "边长"
    tv(2).v = 0
    
    ShowEditData "等边三角形", 3, tv, ToolType.CreateTriangle
End Sub

Private Sub CmdCreateTriangle1_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "底边"
    tv(2).v = 0
    tv(3).t = "高度"
    tv(3).v = 0
    
    ShowEditData "等腰三角形", 4, tv, ToolType.CreateTriangle1
End Sub

Private Sub CmdCreateTriangle2_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "底边"
    tv(2).v = 0
    tv(3).t = "高度"
    tv(3).v = 0
    
    ShowEditData "直角三角形", 4, tv, ToolType.CreateTriangle2
End Sub

Private Sub CmdCreateParallelRectangle_Click()
    ReDim tv(5) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "宽度"
    tv(2).v = 0
    tv(3).t = "高度"
    tv(3).v = 0
    tv(4).t = "错位"
    tv(4).v = 0
    
    ShowEditData "平行四边形", 5, tv, ToolType.CreateParallelRectangle
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, SWP_Flags
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.FraEdit.Visible = False
End Sub

Public Sub CmdCreateCircle_Click()
    ReDim tv(3) As Title_Value
    tv(0).t = "圆心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "圆心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "半径R"
    tv(2).v = 0
    
    ShowEditData "创建圆", 3, tv, ToolType.CreateCircle
End Sub

Public Sub CmdCreateArc_Click()
    ReDim tv(5) As Title_Value
    tv(0).t = "圆心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "圆心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "外径"
    tv(2).v = 0
    tv(3).t = "内径"
    tv(3).v = 0
    tv(4).t = "夹角"
    tv(4).v = 0
    
    ShowEditData "创建圆弧", 5, tv, ToolType.CreateArc
End Sub

Public Sub CmdCreateSector_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "圆心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "圆心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "半径"
    tv(2).v = 0
    tv(3).t = "夹角"
    tv(3).v = 0
    
    ShowEditData "创建扇形", 4, tv, ToolType.CreateSector
End Sub

Public Sub CmdCreateTrapezoid_Click()
    ReDim tv(5) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "上底"
    tv(2).v = 0
    tv(3).t = "下底"
    tv(3).v = 0
    tv(4).t = "高度"
    tv(4).v = 0
    
    ShowEditData "创建等腰梯形", 5, tv, ToolType.CreateTrapezoid
End Sub

Public Sub CmdCreateTrapezoid1_Click()
    ReDim tv(5) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "上底"
    tv(2).v = 0
    tv(3).t = "下底"
    tv(3).v = 0
    tv(4).t = "高度"
    tv(4).v = 0
    
    ShowEditData "创建直角腰梯形", 5, tv, ToolType.CreateTrapezoid1
End Sub

Public Sub CmdCreateEllipse_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "圆心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "圆心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "半径A"
    tv(2).v = 0
    tv(3).t = "半径B"
    tv(3).v = 0
    
    ShowEditData "创建椭圆", 4, tv, ToolType.CreateEllipse
End Sub

Public Sub CmdCreateSquare_Click()
    ReDim tv(3) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "边长"
    tv(2).v = 0
    
    ShowEditData "创建正方形", 3, tv, ToolType.CreateSquare
End Sub

Public Sub CmdCreatePolygon_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "边数"
    tv(2).v = 0
    tv(3).t = "半径"
    tv(3).v = 0
    
    ShowEditData "创建正多边形", 4, tv, ToolType.CreatePolygon
End Sub

Public Sub CmdCreateRectangle_Click()
    ReDim tv(4) As Title_Value
    tv(0).t = "中心X"
    tv(0).v = Int(ViewMaxX / 20) * 10
    tv(1).t = "中心Y"
    tv(1).v = Int(ViewMaxY / 20) * 10
    tv(2).t = "宽度"
    tv(2).v = 0
    tv(3).t = "高度"
    tv(3).v = 0
    
    ShowEditData "创建矩形", 4, tv, ToolType.CreateRectangle
End Sub



