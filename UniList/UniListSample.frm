VERSION 5.00
Begin VB.Form UniListSample 
   Caption         =   "UniList sample"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Focus test"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin UniListDemo.UniList UniList1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3836
      BackColor       =   -2147483643
      BorderStyle     =   1
      CaptureEnter    =   -1  'True
      CaptureEsc      =   -1  'True
      Columns         =   0
      DisableSelect   =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      IntegralHeight  =   0   'False
      MouseIcon       =   "UniListSample.frx":0000
      MousePointer    =   0
      MultiSelect     =   0
      RightToLeft     =   0   'False
      ScrollBars      =   3
      ScrollWidth     =   1000
      Sort            =   -1  'True
      StorageItems    =   500
      StorageMB       =   1
      Style           =   0
      UseEvents       =   -1  'True
      UseTabStops     =   -1  'True
   End
End
Attribute VB_Name = "UniListSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "This message should not appear when you press Enter or Esc in the listbox.", vbInformation
End Sub

Private Sub Form_Load()
    Dim lngA As Long
    UniList1.SortLocale = [Locale Japanese]
    For lngA = &H3041 To &H3093
        UniList1.AddItem "Test " & ChrW$(lngA) & vbTab & "(character: " & Hex$(lngA) & ")"
    Next lngA
    UniList1.AddItem "Point of interest: a Japanese sort order AIUEO!", 0
    UniList1.FontName = "MS Gothic"
    UniList1.FontSize = 12
    UniList1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        UniList1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub

Private Sub UniList1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And UniList1.ListIndex > -1& Then MsgBox UniList1.List(UniList1.ListIndex)
    If KeyAscii = vbKeyEscape Then MsgBox "Escape to the ships!"
End Sub
