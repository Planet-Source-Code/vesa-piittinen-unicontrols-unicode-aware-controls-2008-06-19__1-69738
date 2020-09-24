VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "SelfTimer demo"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Timer fires even if you move or resize the form :)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Timer As SelfTimer
Attribute Timer.VB_VarHelpID = -1

Dim Quit As Boolean

Private Sub Form_Load()
    Set Timer = New SelfTimer
    Timer.Interval = 1
End Sub

Private Sub Form_Terminate()
    Set Timer = Nothing
End Sub

Private Sub Timer_Timer(ByVal Seconds As Currency)
    Me.Caption = Format$(Seconds, "0.000") & " seconds has passed"
    Quit = Seconds > 5
End Sub
