VERSION 5.00
Begin VB.Form UniTextSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UniText sample"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11775
   Icon            =   "UniTextSample.frx":0000
   LinkTopic       =   "UniTextSample"
   MaxButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Things to check out:"
      Height          =   2655
      Left            =   6960
      TabIndex        =   8
      Top             =   2040
      Width           =   4695
      Begin VB.Label Label2 
         Caption         =   "3) In design time, edit text by right click > Edit"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "- LineHeight, MouseOver, ScrollToCaret & UseEvents"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "- FirstVisibleLine and LastVisibleLine for their line indexes"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "- GetSel & SetSel for selection"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "- GetLine for line string, Line for line index, Lines for count"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "4) Other programming easing features:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "2) Hover the mouse over the textbox"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "1) Ctrl + mouse wheel - increase and decrease font size"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Simple'n'dirty speed comparison"
      Height          =   1815
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      Begin UniTextDemo.UniText UniText2 
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Alignment       =   0
         Appearance      =   2
         BackColor       =   -2147483643
         BorderStyle     =   2
         Enabled         =   -1  'True
         FileCodepage    =   0
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
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   -1
         MouseIcon       =   "UniTextSample.frx":000C
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "UniTextSample.frx":0028
         UseEvents       =   -1  'True
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update each 10000 times"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "TextBox"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "I decided to take PropertyChanged ""Text"" out to double the access speed. I don't really see a need for it either."
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   4455
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Focus test"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   6735
   End
   Begin UniTextDemo.UniText UniText1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      Alignment       =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   2
      Enabled         =   -1  'True
      FileCodepage    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HideSelection   =   0   'False
      Locked          =   0   'False
      MaxLength       =   -1
      MouseIcon       =   "UniTextSample.frx":0056
      MousePointer    =   0
      MultiLine       =   -1  'True
      PasswordChar    =   ""
      RightToLeft     =   0   'False
      ScrollBars      =   2
      Text            =   "UniTextSample.frx":0072
      UseEvents       =   -1  'True
   End
End
Attribute VB_Name = "UniTextSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "You should not see this message box when pressing enter in UniText", vbInformation
End Sub

Private Sub Command2_Click()
    Dim lngA As Long, sngStart As Single, strText As String
    sngStart = Timer
    For lngA = 1 To 10000
        strText = CStr(lngA) & " - " & Format$(Timer - sngStart, "0.0000")
        Text2.Text = strText
    Next lngA
    sngStart = Timer
    For lngA = 1 To 10000
        strText = CStr(lngA) & " - " & Format$(Timer - sngStart, "0.0000")
        UniText2.Text = strText
    Next lngA
End Sub

Private Sub Form_Load()
    UniText1.SetSel 0, 4
    UniText1.SelText = "TESTING"
End Sub

Private Sub UniText1_Change()
    Text1.Text = UniText1.Text
    Text1.SelStart = UniText1.SelStart
    Text1.SelLength = UniText1.SelLength
End Sub

Private Sub UniText1_FontChanged()
    Debug.Print "fontchange"
End Sub

Private Sub UniText1_KeyDown(KeyCode As Integer, ByVal Shift As UniTextShift)
    If KeyCode = 48 And (Shift And vbCtrlMask) = vbCtrlMask Then Set UniText1.Font = Text1.Font
End Sub

Private Sub UniText1_KeyUp(KeyCode As Integer, ByVal Shift As UniTextShift)
    Text1.SelStart = UniText1.SelStart
    Text1.SelLength = UniText1.SelLength
End Sub

Private Sub UniText1_MouseEnter()
    UniText1.ForeColor = Text1.BackColor
    UniText1.BackColor = Text1.ForeColor
End Sub

Private Sub UniText1_MouseLeave()
    UniText1.BackColor = Text1.BackColor
    UniText1.ForeColor = Text1.ForeColor
End Sub

Private Sub UniText1_MouseUp(Button As UniTextMouseButton, ByVal Shift As UniTextShift, X As Single, Y As Single)
    Text1.SelStart = UniText1.SelStart
    Text1.SelLength = UniText1.SelLength
End Sub

Private Sub UniText1_MouseWheel(ByVal Wheel As UniTextMouseWheel, ByVal Shift As UniTextShift)
    Dim sngFontSize As Single
    If Shift = [No Mask] Then
        Text1.SelStart = UniText1.SelStart
        Text1.SelLength = UniText1.SelLength
    ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
        If Wheel = [Wheel Up] Then
            sngFontSize = UniText1.FontSize + 1
        Else
            sngFontSize = UniText1.FontSize - 1
        End If
        If sngFontSize < 6 Then sngFontSize = 6
        If sngFontSize > 20 Then sngFontSize = 20
        UniText1.FontSize = sngFontSize
    End If
End Sub

Private Sub UniText1_Scroll(ByVal Direction As UniTextScrollDirection)
    Text1.SelStart = UniText1.SelStart
    Text1.SelLength = UniText1.SelLength
End Sub
