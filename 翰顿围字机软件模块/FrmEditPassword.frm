VERSION 5.00
Begin VB.Form FrmEditPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�޸�����"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TxtPW2 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox TxtPW1 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "ȷ��������"
      BeginProperty Font 
         Name            =   "����"
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
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmEditPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If TxtPW1.Text = TxtPW2.Text Then
        If Trim(TxtPW1.Text) = "" Then
            MsgBox "����ʹ�ÿ����롣�����޸����롣", vbExclamation + vbOKOnly + vbSystemModal, ""
        Else
            SaveRegString HKEY_CURRENT_USER, "Software\" & APPSerial & "\" & DeviceModel & "\Registry", "PW", Trim(TxtPW1.Text)
            MsgBox "�����޸���ϡ��´�������ʹ�������룡", vbExclamation + vbOKOnly + vbSystemModal, ""
        End If
    Else
        MsgBox "������������벻һ�¡������޸����롣", vbExclamation + vbOKOnly + vbSystemModal, ""
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_Flags
End Sub
