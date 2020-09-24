VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin SelfControlDemo.SelfControl SelfControl1 
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1931
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "This should not appear when SelfControl is active and you press Enter or Esc.", vbInformation
End Sub
