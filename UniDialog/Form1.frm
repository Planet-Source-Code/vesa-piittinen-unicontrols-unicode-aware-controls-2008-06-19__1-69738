VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show dialog"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin Project1.UniDialog UniDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621956
      FileCustomFilter=   "Form1.frx":0000
      FileDefaultExtension=   "Form1.frx":0020
      FileFilter      =   "Form1.frx":004A
      FileOpenTitle   =   "Form1.frx":0092
      FileSaveTitle   =   "Form1.frx":00CA
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    UniDialog1.ShowOpen
End Sub

Private Sub UniDialog1_OpenCancel(ByVal CancelType As UniDialogFileCancel)
    List1.Clear
    If CancelType = [No error] Then
        List1.AddItem "Cancel was pressed"
    Else
        List1.AddItem "There was an error!"
    End If
End Sub

Private Sub UniDialog1_OpenFile(ByVal Filename As String)
    List1.AddItem Filename
    ' remember this path
    UniDialog1.FileInitialDirectory = Filename
End Sub
