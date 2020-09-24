VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public WndProc As Long

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    WndProc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetWindowLongW Me.hWnd, GWL_WNDPROC, AddOf(AddressOf Module1.WndProc)
    SendMessageW Me.hWnd, WM_SETTEXT, 0&, ByVal StrPtr(ChrW$(&H3041))
    SetWindowLongW Me.hWnd, GWL_WNDPROC, WndProc
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Form2
End Sub
