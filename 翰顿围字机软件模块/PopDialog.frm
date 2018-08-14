VERSION 5.00
Begin VB.Form PopupDlg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   1770
   Icon            =   "PopDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   118
End
Attribute VB_Name = "PopupDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    BringWindowToTop Me.hWnd
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, SWP_Flags
End Sub
