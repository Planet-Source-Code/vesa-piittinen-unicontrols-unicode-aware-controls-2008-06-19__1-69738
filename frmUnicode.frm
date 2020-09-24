VERSION 5.00
Begin VB.Form frmUnicode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unicode controls test"
   ClientHeight    =   4695
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8415
   Icon            =   "frmUnicode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel1 
      Height          =   585
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1032
      Alignment       =   2
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmUnicode.frx":000C
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmUnicode.frx":003E
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   -1  'True
      WordWrap        =   0   'False
   End
   Begin UniControls.UniList UniList1 
      Height          =   3615
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6376
      BackColor       =   -2147483643
      BorderStyle     =   2
      CaptureEnter    =   -1  'True
      CaptureEsc      =   0   'False
      CaptureTab      =   0   'False
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
      IntegralHeight  =   -1  'True
      MouseIcon       =   "frmUnicode.frx":005A
      MousePointer    =   0
      MultiSelect     =   0
      RightToLeft     =   0   'False
      ScrollBars      =   2
      ScrollBarVisibility=   0
      ScrollWidth     =   0
      Sort            =   0   'False
      StorageItems    =   500
      StorageMB       =   1
      Style           =   0
      UseEvents       =   0   'False
      UseTabStops     =   -1  'True
   End
   Begin UniControls.UniMenu UniMenu1 
      Left            =   0
      Top             =   0
      _ExtentX        =   10186
      _ExtentY        =   3201
      BorderBar       =   0   'False
      BorderBarHot    =   -1  'True
      BorderBarSelected=   -1  'True
      BorderCheck     =   -1  'True
      BorderImage     =   0   'False
      BorderItem      =   0   'False
      BorderSelected  =   -1  'True
      In3D            =   0   'False
      Panel           =   -1  'True
      ImageSize       =   16
      ColorArrow      =   -2147483635
      ColorArrowSelected=   -2147483641
      ColorBar        =   -2147483633
      ColorBarBorder  =   -2147483633
      ColorBarHot     =   -2147483635
      ColorBarHotBorder=   -2147483635
      ColorBarHotText =   -2147483641
      ColorBarSelected=   -2147483633
      ColorBarSelectedBorder=   -2147483627
      ColorBarSelectedText=   -2147483641
      ColorBarText    =   -2147483641
      ColorBorderBack =   -2147483633
      ColorBorderInner=   -2147483633
      ColorBorderOuter=   -2147483627
      ColorCheck      =   -2147483635
      ColorCheckBack  =   -2147483633
      ColorCheckBorder=   -2147483635
      ColorDisabled   =   -2147483633
      ColorDisabledText=   -2147483632
      ColorImageBack  =   -2147483633
      ColorImageBorder=   -2147483633
      ColorImageShadow=   -2147483635
      ColorItem       =   -2147483633
      ColorItemBorder =   -2147483633
      ColorItemText   =   -2147483641
      ColorSelected   =   -2147483635
      ColorSelectedBorder=   -2147483635
      ColorSelectedText=   -2147483641
      ColorSeparator  =   -2147483632
      ColorSeparatorBack=   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SmoothArrow     =   1
      SmoothArrowSelected=   1
      SmoothBar       =   0
      SmoothBarBorder =   1
      SmoothBarHot    =   2
      SmoothBarHotBorder=   1
      SmoothBarHotText=   0
      SmoothBarSelected=   1
      SmoothBarSelectedBorder=   1
      SmoothBarSelectedText=   0
      SmoothBarText   =   0
      SmoothBorderBack=   3
      SmoothBorderInner=   1
      SmoothBorderOuter=   1
      SmoothCheck     =   1
      SmoothCheckBack =   3
      SmoothCheckBorder=   0
      SmoothDisabled  =   3
      SmoothDisabledText=   1
      SmoothImageBack =   1
      SmoothImageBorder=   1
      SmoothImageShadow=   1
      SmoothItem      =   3
      SmoothItemBorder=   3
      SmoothItemText  =   0
      SmoothSelected  =   2
      SmoothSelectedBorder=   1
      SmoothSelectedText=   0
      SmoothSeparator =   1
      SmoothSeparatorBack=   3
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   435
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   767
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   0
      Caption         =   "frmUnicode.frx":0076
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmUnicode.frx":00A8
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin UniControls.UniText UniText2 
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Alignment       =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   2
      CaptureEnter    =   0   'False
      CaptureEsc      =   0   'False
      CaptureTab      =   -1  'True
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
      MouseIcon       =   "frmUnicode.frx":00C4
      MousePointer    =   0
      MultiLine       =   0   'False
      PasswordChar    =   ""
      RightToLeft     =   0   'False
      ScrollBars      =   0
      Text            =   "frmUnicode.frx":00E0
      UseEvents       =   -1  'True
   End
   Begin UniControls.UniCommand UniCommand1 
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      CaptionB        =   "frmUnicode.frx":0100
      CaptionLen      =   4
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
   End
   Begin UniControls.UniText UniText1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6165
      Alignment       =   0
      Appearance      =   0
      BackColor       =   -2147483624
      BorderStyle     =   2
      CaptureEnter    =   -1  'True
      CaptureEsc      =   0   'False
      CaptureTab      =   0   'False
      Enabled         =   -1  'True
      FileCodepage    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483625
      HideSelection   =   -1  'True
      Locked          =   0   'False
      MaxLength       =   -1
      MouseIcon       =   "frmUnicode.frx":0128
      MousePointer    =   0
      MultiLine       =   -1  'True
      PasswordChar    =   ""
      RightToLeft     =   0   'False
      ScrollBars      =   2
      Text            =   "frmUnicode.frx":0144
      UseEvents       =   -1  'True
   End
   Begin UniControls.UniDialog UniDialog1 
      Left            =   1920
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   323
      FileCustomFilter=   "frmUnicode.frx":03F8
      FileDefaultExtension=   "frmUnicode.frx":0418
      FileFilter      =   "frmUnicode.frx":0438
      FileOpenTitle   =   "frmUnicode.frx":0480
      FileSaveTitle   =   "frmUnicode.frx":04B8
      FolderMessage   =   "frmUnicode.frx":04F0
   End
   Begin VB.Menu mnuDialog 
      Caption         =   "Uni&Dialog"
      Begin VB.Menu mnuFolder 
         Caption         =   "&Folder..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Uni&Menu"
      Begin VB.Menu mnuStyleSet 
         Caption         =   "Classic"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuStyleSet 
         Caption         =   "Office 97"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuStyleSet 
         Caption         =   "Office XP"
         Index           =   2
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Char "
      Begin VB.Menu mnuText 
         Caption         =   "Add text above"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuText2 
         Caption         =   "Add text here: "
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmUnicode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long

Private Const GWL_WNDPROC = -4

Private m_Caption As String

' for SetIcon: http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Const IMAGE_ICON = 1

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Private Const WM_SETICON = &H80

Private Const GW_OWNER = 4

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Property Get CaptionW() As String
    CaptionW = m_Caption
End Property
Public Property Let CaptionW(ByRef NewValue As String)
    Static WndProc As Long, VBWndProc As Long
    m_Caption = NewValue
    ' get window procedures if we don't have them
    If WndProc = 0 Then
        ' the default Unicode window procedure
        WndProc = GetProcAddress(GetModuleHandleW(StrPtr("user32")), "DefWindowProcW")
        ' window procedure of this form
        VBWndProc = GetWindowLongA(hWnd, GWL_WNDPROC)
    End If
    ' ensure we got them
    If WndProc <> 0 Then
        ' replace form's window procedure with the default Unicode one
        SetWindowLongW hWnd, GWL_WNDPROC, WndProc
        ' change form's caption
        SetWindowTextW hWnd, StrPtr(m_Caption)
        ' restore the original window procedure
        SetWindowLongA hWnd, GWL_WNDPROC, VBWndProc
    Else
        ' no Unicode for us
        Caption = m_Caption
    End If
End Property
' Customized from: http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp
Private Sub SetIcon(ByVal sIconResName As String)
    Dim lhWndTop As Long, lhWnd As Long
    Dim cX As Long, cY As Long
    Dim hIconLarge As Long, hIconSmall As Long
    
    cX = GetSystemMetrics(SM_CXICON)
    cY = GetSystemMetrics(SM_CYICON)
    
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cX, cY, LR_SHARED)
    
    SendMessage hWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge
    
    cX = GetSystemMetrics(SM_CXSMICON)
    cY = GetSystemMetrics(SM_CYSMICON)
    
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cX, cY, LR_SHARED)
    
    SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall
End Sub
Private Sub Form_Load()
    SetIcon "AAAAAAAA"
    UniLabel1(0).Caption = ChrW$(&H3042) & ChrW$(&H3044) & ChrW$(&H3046) & ChrW$(&H3048) & ChrW$(&H304A) & " = A I U E O"
    UniLabel1(1).Caption = UniLabel1(0).Caption
    Me.CaptionW = "UniControls sample: " & ChrW$(&H3042) & ChrW$(&H3044) & ChrW$(&H3046) & ChrW$(&H3048) & ChrW$(&H304A)
End Sub

Private Sub mnuFolder_Click()
    UniDialog1.ShowFolder
End Sub

Private Sub mnuOpen_Click()
    UniDialog1.ShowOpen
End Sub

Private Sub mnuStyleSet_Click(Index As Integer)
    UniMenu1.SetStyle Index + 1, True, UniMenu1.In3D, UniMenu1.Panel
End Sub

Private Sub mnuText_Click()
    UniMenu1.Caption(mnuTest) = UniMenu1.Caption(mnuTest) & ChrW$(&H3042 + (Rnd * 20 \ 1))
End Sub

Private Sub mnuText2_Click()
    UniMenu1.Caption(mnuText2) = UniMenu1.Caption(mnuText2) & ChrW$(&H3042 + (Rnd * 20 \ 1))
End Sub

Private Sub UniCommand1_Click()
    UniList1.AddItem UniText2.Text
    UniList1.ListIndex = UniList1.NewIndex
End Sub

Private Sub UniDialog1_FolderSelect(ByVal Path As String)
    MsgBox "You selected " & Path, vbInformation
End Sub

Private Sub UniDialog1_OpenFile(ByVal Filename As String)
    UniText1.LoadFile Filename
    UniText1.SetFocus
End Sub

Private Sub UniLabel1_MouseEnter(Index As Integer)
    With UniLabel1(Index)
        .BackColor = vbHighlight
        .ForeColor = vbHighlightText
    End With
End Sub

Private Sub UniLabel1_MouseLeave(Index As Integer)
    With UniLabel1(Index)
        '.BorderColor = vbButtonFace
        .BackColor = Me.BackColor
        .ForeColor = Me.ForeColor
        '.ShadowColor = vbButtonFace
    End With
End Sub

Private Sub UniMenu1_StyleChange()
    Dim lngA As Long
    For lngA = 0 To 2
        mnuStyleSet(lngA).Checked = (UniMenu1.Style - 1) = lngA
    Next lngA
End Sub

