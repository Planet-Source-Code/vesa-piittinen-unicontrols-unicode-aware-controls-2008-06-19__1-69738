VERSION 5.00
Begin VB.UserControl UniMenu 
   Alignable       =   -1  'True
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   FillColor       =   &H00C0C0C0&
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   ToolboxBitmap   =   "UniMenu.ctx":0000
   Begin VB.Frame fraMenu 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Visible         =   0   'False
      Begin VB.CommandButton cmdCaption 
         Caption         =   "&Set"
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
      Begin UniControls.UniText txtCaption 
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         Alignment       =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         CaptureEnter    =   -1  'True
         CaptureEsc      =   0   'False
         CaptureTab      =   0   'False
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
         MouseIcon       =   "UniMenu.ctx":00FA
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "UniMenu.ctx":0116
         UseEvents       =   -1  'True
      End
      Begin VB.CheckBox chkChecked 
         Caption         =   "Checked"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkWindowList 
         Caption         =   "Window list of MDI childs"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cmbMenu 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblInfo 
         Caption         =   "Caption:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Timer tmrDesignTime 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape shpTitle 
      BackColor       =   &H8000000E&
      BorderColor     =   &H8000000D&
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UniMenu"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "UniMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniMenu 1.9.2 - Unicode menu user control
'* -----------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'* Unicode on 2000/XP/Vista, ANSI fallback on 95/98/ME
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) May not be used as a part of commercial (unicode) controls suite.
'*        4) Free for any other commercial and non-commercial usage.
'*        5) Use at your own risk. No support guaranteed.
'*
'* SUPPORT FOR UNICONTROLS
'* -----------------------
'* http://www.vbforums.com/showthread.php?t=500026
'* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69738&lngWId=1
'*
'* REQUIREMENTS
'* ------------
'* Note: TLBs are compiled to your program so you don't need to distribute the files
'* - OleGuids3.tlb      = Ole Guid and interface definitions 3.0
'* - UniTextModule.bas
'* - UniText.ctl
'*
'* NOTES
'* -----
'* I've been looking for various approaches to menus. Finding a balance with a good solution is hard.
'* This current version works with the existing menu items. The problem is that it is hard to keep
'* things integrated. With this I mean mostly the problem of adding menu icon support. Currently it seems
'* to be impossible to integrate existing VB objects with their respective menu items. I could make some
'* workaround code, but it wouldn't really ever meet my own standards for quality. When you link an image
'* to a menu item, that image should stay there no matter how much changes happen.
'*
'* Thus this has led me to the conclusion that I probably have to go for the uglier world and take full
'* control of the menu. Thus no support for existing menu systems. But, as creating menu stuff is very
'* time consuming and I still have so much to do with other controls as well, this old base code is going
'* to remain here for a while. For the least at the moment this works better than it has ever before.
'*
'* HOW TO ADD TO YOUR PROGRAM
'* --------------------------
'* 1) OPTIONAL: Copy OleGuids3.tlb to Windows system folder.
'* 2) Copy UniTextModule.bas, UniText.ctl, UniText.ctx, UniMenu.ctl and UniMenu.ctx to your project folder.
'* 3) In your project, add a reference to OleGuids3.tlb (Project > References...)
'* 4) Add UniTextModule.bas
'* 5) Add UniText.ctl
'* 6) Add UniMenu.ctl
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.9.2 BETA (2008-06-19)
'* - New: Caption property, takes a Menu object as a parameter
'* - Fixes: font improvements, IDE crash reduction, Refresh works now as intended
'*
'* Version 1.9.1 BETA (2007-12-15)
'* - Tons of code reordering, all private properties are now in the end of the source
'* - Some code logic fixes to remove very random crashing issues under IDE
'*
'* Version 1.9 BETA (2007-12-08)
'* - goodbye WinSubHook, hello SelfSub and SelfHook!
'*
'* Version 1.0
'* - fixed color and settings bugs and added a few more
'* - made it a user control
'* - custom style
'* - property page
'* - Unicode support for 9X/ME via Unicows.dll
'* - three styles: Classic, Office97 and OfficeXP
'* - uses WinSubHook by Paul Caton
'*
'* ----------------------------------------------------------------------
'*
'* TODO:
'*
'* - Easier handling of captions, not everyone knows how to work with UTF-8... may involve trying to subclass VB6 IDE
'* - Menu images?

Option Explicit

Private Const HC_ACTION = 0

Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
Private Const WM_DRAWITEM = &H2B
Private Const WM_ERASEBKGND = &H14
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUSELECT = &H11F
Private Const WM_NCPAINT = &H85
Private Const WM_SHOWWINDOW = &H18

' Selfsub (by Paul Caton, LaVolpe's version)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' Local variables/constants: must declare these regardless if using subclassing, hooking, callbacks
    Private z_scFunk            As Collection   'hWnd/thunk-address collection; initialized as needed
    Private z_hkFunk            As Collection   'hook/thunk-address collection; initialized as needed
    Private z_cbFunk            As Collection   'callback/thunk-address collection; initialized as needed
    Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
    Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
    Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
    Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
    Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

  ' Declarations:
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Enum eThunkType
        SubclassThunk = 0
        HookThunk = 1
        CallbackThunk = 2
    End Enum

    '-Selfsub specific declarations----------------------------------------------------------------------------
    Private Enum eMsgWhen                                                   'When to callback
      MSG_BEFORE = 1                                                        'Callback before the original WndProc
      MSG_AFTER = 2                                                         'Callback after the original WndProc
      MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
    End Enum
    
    ' see ssc_Subclass for complete listing of indexes and what they relate to
    Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
    Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
    Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
    Private Const ALL_MESSAGES  As Long = -1    'All messages will callback
    
    Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    '-SelfHook specific declarations----------------------------------------------------------------------------
    Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
    
    Private Enum eHookType  ' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
      WH_MSGFILTER = -1
      WH_JOURNALRECORD = 0
      WH_JOURNALPLAYBACK = 1
      WH_KEYBOARD = 2
      WH_GETMESSAGE = 3
      WH_CALLWNDPROC = 4
      WH_CBT = 5
      WH_SYSMSGFILTER = 6
      WH_MOUSE = 7
      WH_DEBUG = 9
      WH_SHELL = 10
      WH_FOREGROUNDIDLE = 11
      WH_CALLWNDPROCRET = 12
      WH_KEYBOARD_LL = 13       ' NT/2000/XP+ only, Global hook only
      WH_MOUSE_LL = 14          ' NT/2000/XP+ only, Global hook only
    End Enum
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' API constants
Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const DSS_NORMAL As Long = &H0&
Private Const DST_PREFIXTEXT As Long = &H2&
Private Const DST_BITMAP As Long = &H4&
Private Const DSS_DISABLED As Long = &H20&

Private Const DT_BOTTOM As Long = &H8&
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_CENTER As Long = &H1&
Private Const DT_LEFT As Long = &H0&
Private Const DT_NOCLIP As Long = &H100&
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_RIGHT As Long = &H2&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_TOP As Long = &H0&
Private Const DT_VCENTER As Long = &H4&
Private Const DT_WORDBREAK As Long = &H10&

Private Const GWL_EXSTYLE As Long = -20&
Private Const GWL_STYLE As Long = -16&

Private Const MF_BYCOMMAND As Long = &H0&
Private Const MF_DISABLED As Long = &H2&
Private Const MF_BITMAP As Long = &H4&
Private Const MF_CHECKED As Long = &H8&
Private Const MF_MENUBARBREAK As Long = &H20&
Private Const MF_MENUBREAK As Long = &H40&
Private Const MF_OWNERDRAW As Long = &H100&
Private Const MF_RADIOCHECK As Long = &H200&
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_RIGHTORDER As Long = &H2000&
Private Const MF_RIGHTJUSTIFY As Long = &H4000&

Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MIIM_SUBMENU As Long = &H4&
Private Const MIIM_CHECKMARKS As Long = &H8&
Private Const MIIM_TYPE As Long = &H10&
Private Const MIIM_DATA As Long = &H20&
Private Const MIIM_STRING As Long = &H40&

Private Const OBM_CHECK As Long = 32760&
Private Const OBM_MNARROW As Long = 32739&

Private Const ODS_SELECTED As Long = &H1&
Private Const ODS_GRAYED As Long = &H2&
Private Const ODS_DISABLED As Long = &H4&
Private Const ODS_CHECKED As Long = &H8&
Private Const ODS_FOCUS As Long = &H10&
Private Const ODS_DEFAULT As Long = &H20&
Private Const ODS_HOTTRACK As Long = &H40&

Private Const ODT_MENU As Long = 1&

Private Const SM_CXBORDER As Long = 5&
Private Const SM_CYMENUSIZE As Long = 55&
Private Const SPI_GETNONCLIENTMETRICS As Long = 41&

' API font constants
Private Const FW_BOLD = 700&
Private Const FW_NORMAL = 400&
Private Const LF_FACESIZE = 32&
Private Const LOGPIXELSX = 88&
Private Const LOGPIXELSY = 90&

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    ItemState As Long
    hwndItem As Long
    hDC As Long
    rcItem As RECT
    ItemData As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(31) As Byte
End Type

Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

' these are the API subs and functions we need just to hook, subclass, get menu data, modify menu settings and to draw...
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal Flags As Long) As Long
Private Declare Function DrawStateANSI Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal Flags As Long) As Long
Private Declare Function DrawStateUnicows Lib "unicows" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal Flags As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStrPtr As Long, ByVal nCount As Long, lpRECT As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextANSI Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStrPtr As Long, ByVal nCount As Long, lpRECT As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextUnicows Lib "unicows" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStrPtr As Long, ByVal nCount As Long, lpRECT As RECT, ByVal wFormat As Long) As Long

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRECT As RECT, ByVal hBrush As Long) As Long
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClipBox Lib "gdi32" (ByVal hDC As Long, lpRECT As RECT) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpmii As MENUITEMINFO) As Long
'Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpStrPtr As Long, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, lpBitmapName As Any) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long

' API: character set conversion
Private Declare Function GetACP Lib "kernel32" () As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Public Event StyleChange()

' default property values
Private Const m_def_ImageSize As Long = 16&     ' default image size (13 and 15 are the default sizes, we use 16)

' other custom constants
Private Const MENU_TOP_ID As Long = 666&        ' top menu id always starts from 666...

Public Enum MENU_STYLE
    [Style Custom]
    [Style Classic]
    [Style Office97]
    [Style OfficeXP]
End Enum

Public Enum MENU_UNICODE
    [Detect Windows]
    [Always Unicode]
    [Always ANSI]
End Enum

Private Type MENU_CAPTION
    Index As Long
    Caption As String
End Type

Private Type MENU_COLORS
    Arrow As Long
    ArrowSelected As Long
    Bar As Long
    BarBorder As Long
    BarHot As Long
    BarHotBorder As Long
    BarHotText As Long
    BarSelected As Long
    BarSelectedBorder As Long
    BarSelectedText As Long
    barText As Long
    BorderBack As Long
    BorderInner As Long
    BorderOuter As Long
    Check As Long
    CheckBack As Long
    CheckBorder As Long
    Disabled As Long
    DisabledText As Long
    ImageBack As Long
    ImageBorder As Long
    ImageShadow As Long
    Item As Long
    ItemBorder As Long
    ItemText As Long
    Selected As Long
    SelectedBorder As Long
    SelectedText As Long
    Separator As Long
    SeparatorBack As Long
End Type

Private Type MENU_HOOK
    hWnd As Long
    ' cSubclass
End Type

Private Type MENU_OWNERDRAW
    hWnd As Long
    Index() As Long
End Type

Private Type MENU_STYLES
    BarBorder As Boolean
    BarHotBorder As Boolean
    BarSelectedBorder As Boolean
    CheckBorder As Boolean
    ImageBorder As Boolean
    ItemBorder As Boolean
    SelectedBorder As Boolean
    In3D As Boolean
    InANSI As Boolean
    Panel As Boolean
    ImageSize As Byte
End Type

' helper arrays
Dim Captions() As MENU_CAPTION
Dim SmoothColors(255) As Long

' styling and properties
Dim Colors As MENU_COLORS
Dim Styles As MENU_STYLES

' hooking and subclassing
Dim HookSub() As MENU_HOOK
Dim m_ShowWindow As Boolean
Dim OwnerDraw() As MENU_OWNERDRAW

' properties
Dim m_ColorArrow As Long
Dim m_ColorArrowSelected As Long
Dim m_ColorBar As Long
Dim m_ColorBarBorder As Long
Dim m_ColorBarHot As Long
Dim m_ColorBarHotBorder As Long
Dim m_ColorBarHotText As Long
Dim m_ColorBarSelected As Long
Dim m_ColorBarSelectedBorder As Long
Dim m_ColorBarSelectedText As Long
Dim m_ColorBarText As Long
Dim m_ColorBorderBack As Long
Dim m_ColorBorderInner As Long
Dim m_ColorBorderOuter As Long
Dim m_ColorCheck As Long
Dim m_ColorCheckBack As Long
Dim m_ColorCheckBorder As Long
Dim m_ColorDisabled As Long
Dim m_ColorDisabledText As Long
Dim m_ColorImageBack As Long
Dim m_ColorImageBorder As Long
Dim m_ColorImageShadow As Long
Dim m_ColorItem As Long
Dim m_ColorItemBorder As Long
Dim m_ColorItemText As Long
Dim m_ColorSelected As Long
Dim m_ColorSelectedBorder As Long
Dim m_ColorSelectedText As Long
Dim m_ColorSeparator As Long
Dim m_ColorSeparatorBack As Long
Dim WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Dim m_SmoothArrow As Byte
Dim m_SmoothArrowSelected As Byte
Dim m_SmoothBar As Byte
Dim m_SmoothBarBorder As Byte
Dim m_SmoothBarHot As Byte
Dim m_SmoothBarHotBorder As Byte
Dim m_SmoothBarHotText As Byte
Dim m_SmoothBarSelected As Byte
Dim m_SmoothBarSelectedBorder As Byte
Dim m_SmoothBarSelectedText As Byte
Dim m_SmoothBarText As Byte
Dim m_SmoothBorderBack As Byte
Dim m_SmoothBorderInner As Byte
Dim m_SmoothBorderOuter As Byte
Dim m_SmoothCheck As Byte
Dim m_SmoothCheckBack As Byte
Dim m_SmoothCheckBorder As Byte
Dim m_SmoothDisabled As Byte
Dim m_SmoothDisabledText As Byte
Dim m_SmoothImageBack As Byte
Dim m_SmoothImageBorder As Byte
Dim m_SmoothImageShadow As Byte
Dim m_SmoothItem As Byte
Dim m_SmoothItemBorder As Byte
Dim m_SmoothItemText As Byte
Dim m_SmoothSelected As Byte
Dim m_SmoothSelectedBorder As Byte
Dim m_SmoothSelectedText As Byte
Dim m_SmoothSeparator As Byte
Dim m_SmoothSeparatorBack As Byte
Dim m_StartingStyle As MENU_STYLE
Dim m_Style As MENU_STYLE
Dim m_Unicode As MENU_UNICODE
Dim m_WINNT As Boolean

Dim m_MenuCurrent As Menu

' helper variables
Dim lngOwnerhWnd As Long, lngSyshWnd As Long, lngMenuhWnd As Long
Dim intMenuTopEndID As Integer, lngMenuTopEndID As Long, lngLastMenuParent As Long
Dim blnDesignTime As Boolean, blnEditMode As Boolean, blnInIDE As Boolean, blnParentIsForm As Boolean
Public Property Get BorderBar() As Boolean
    BorderBar = Styles.BarBorder
End Property
Public Property Let BorderBar(ByVal NewValue As Boolean)
    Styles.BarBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get BorderBarHot() As Boolean
    BorderBarHot = Styles.BarHotBorder
End Property
Public Property Let BorderBarHot(ByVal NewValue As Boolean)
    Styles.BarHotBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get BorderBarSelected() As Boolean
    BorderBarSelected = Styles.BarSelectedBorder
End Property
Public Property Let BorderBarSelected(ByVal NewValue As Boolean)
    Styles.BarSelectedBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get BorderCheck() As Boolean
    BorderCheck = Styles.CheckBorder
End Property
Public Property Let BorderCheck(ByVal NewValue As Boolean)
    Styles.CheckBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get BorderImage() As Boolean
    BorderImage = Styles.ImageBorder
End Property
Public Property Let BorderImage(ByVal NewValue As Boolean)
    Styles.ImageBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get BorderItem() As Boolean
    BorderItem = Styles.ItemBorder
End Property
Public Property Let BorderItem(ByVal NewValue As Boolean)
    Styles.ItemBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get BorderSelected() As Boolean
    BorderSelected = Styles.SelectedBorder
End Property
Public Property Let BorderSelected(ByVal NewValue As Boolean)
    Styles.SelectedBorder = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get Caption(ByRef MenuItem As Menu) As String
    If Not MenuItem Is Nothing Then
        Caption = UTF8toStr(MenuItem.Caption)
    End If
End Property
Public Property Let Caption(ByRef MenuItem As Menu, ByRef NewValue As String)
    If Not MenuItem Is Nothing Then
        MenuItem.Caption = StrToUTF8(NewValue)
        Refresh
    End If
End Property
Public Property Get ColorArrow() As OLE_COLOR
    ColorArrow = m_ColorArrow
End Property
Public Property Let ColorArrow(ByVal NewColor As OLE_COLOR)
    m_ColorArrow = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorArrowSelected() As OLE_COLOR
    ColorArrowSelected = m_ColorArrowSelected
End Property
Public Property Let ColorArrowSelected(ByVal NewColor As OLE_COLOR)
    m_ColorArrowSelected = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBar() As OLE_COLOR
    ColorBar = m_ColorBar
End Property
Public Property Let ColorBar(ByVal NewColor As OLE_COLOR)
    m_ColorBar = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarBorder() As OLE_COLOR
    ColorBarBorder = m_ColorBarBorder
End Property
Public Property Let ColorBarBorder(ByVal NewColor As OLE_COLOR)
    m_ColorBarBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarHot() As OLE_COLOR
    ColorBarHot = m_ColorBarHot
End Property
Public Property Let ColorBarHot(ByVal NewColor As OLE_COLOR)
    m_ColorBarHot = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarHotBorder() As OLE_COLOR
    ColorBarHotBorder = m_ColorBarHotBorder
End Property
Public Property Let ColorBarHotBorder(ByVal NewColor As OLE_COLOR)
    m_ColorBarHotBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarHotText() As OLE_COLOR
    ColorBarHotText = m_ColorBarHotText
End Property
Public Property Let ColorBarHotText(ByVal NewColor As OLE_COLOR)
    m_ColorBarHotText = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarSelected() As OLE_COLOR
    ColorBarSelected = m_ColorBarSelected
End Property
Public Property Let ColorBarSelected(ByVal NewColor As OLE_COLOR)
    m_ColorBarSelected = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarSelectedBorder() As OLE_COLOR
    ColorBarSelectedBorder = m_ColorBarSelectedBorder
End Property
Public Property Let ColorBarSelectedBorder(ByVal NewColor As OLE_COLOR)
    m_ColorBarSelectedBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarSelectedText() As OLE_COLOR
    ColorBarSelectedText = m_ColorBarSelectedText
End Property
Public Property Let ColorBarSelectedText(ByVal NewColor As OLE_COLOR)
    m_ColorBarSelectedText = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBarText() As OLE_COLOR
    ColorBarText = m_ColorBarText
End Property
Public Property Let ColorBarText(ByVal NewColor As OLE_COLOR)
    m_ColorBarText = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBorderBack() As OLE_COLOR
    ColorBorderBack = m_ColorBorderBack
End Property
Public Property Let ColorBorderBack(ByVal NewColor As OLE_COLOR)
    m_ColorBorderBack = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBorderinner() As OLE_COLOR
    ColorBorderinner = m_ColorBorderInner
End Property
Public Property Let ColorBorderinner(ByVal NewColor As OLE_COLOR)
    m_ColorBorderInner = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorBorderOuter() As OLE_COLOR
    ColorBorderOuter = m_ColorBorderOuter
End Property
Public Property Let ColorBorderOuter(ByVal NewColor As OLE_COLOR)
    m_ColorBorderOuter = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorCheck() As OLE_COLOR
    ColorCheck = m_ColorCheck
End Property
Public Property Let ColorCheck(ByVal NewColor As OLE_COLOR)
    m_ColorCheck = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorCheckBack() As OLE_COLOR
    ColorCheckBack = m_ColorCheckBack
End Property
Public Property Let ColorCheckBack(ByVal NewColor As OLE_COLOR)
    m_ColorCheckBack = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorCheckBorder() As OLE_COLOR
    ColorCheckBorder = m_ColorCheckBorder
End Property
Public Property Let ColorCheckBorder(ByVal NewColor As OLE_COLOR)
    m_ColorCheckBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorDisabled() As OLE_COLOR
    ColorDisabled = m_ColorDisabled
End Property
Public Property Let ColorDisabled(ByVal NewColor As OLE_COLOR)
    m_ColorDisabled = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorDisabledText() As OLE_COLOR
    ColorDisabledText = m_ColorDisabledText
End Property
Public Property Let ColorDisabledText(ByVal NewColor As OLE_COLOR)
    m_ColorDisabledText = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorImageBack() As OLE_COLOR
    ColorImageBack = m_ColorImageBack
End Property
Public Property Let ColorImageBack(ByVal NewColor As OLE_COLOR)
    m_ColorImageBack = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorImageBorder() As OLE_COLOR
    ColorImageBorder = m_ColorImageBorder
End Property
Public Property Let ColorImageBorder(ByVal NewColor As OLE_COLOR)
    m_ColorImageBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorImageShadow() As OLE_COLOR
    ColorImageShadow = m_ColorImageShadow
End Property
Public Property Let ColorImageShadow(ByVal NewColor As OLE_COLOR)
    m_ColorImageShadow = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorItem() As OLE_COLOR
    ColorItem = m_ColorItem
End Property
Public Property Let ColorItem(ByVal NewColor As OLE_COLOR)
    m_ColorItem = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorItemBorder() As OLE_COLOR
    ColorItemBorder = m_ColorItemBorder
End Property
Public Property Let ColorItemBorder(ByVal NewColor As OLE_COLOR)
    m_ColorItemBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorItemText() As OLE_COLOR
    ColorItemText = m_ColorItemText
End Property
Public Property Let ColorItemText(ByVal NewColor As OLE_COLOR)
    m_ColorItemText = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorSelected() As OLE_COLOR
    ColorSelected = m_ColorSelected
End Property
Public Property Let ColorSelected(ByVal NewColor As OLE_COLOR)
    m_ColorSelected = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorSelectedBorder() As OLE_COLOR
    ColorSelectedBorder = m_ColorSelectedBorder
End Property
Public Property Let ColorSelectedBorder(ByVal NewColor As OLE_COLOR)
    m_ColorSelectedBorder = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorSelectedText() As OLE_COLOR
    ColorSelectedText = m_ColorSelectedText
End Property
Public Property Let ColorSelectedText(ByVal NewColor As OLE_COLOR)
    m_ColorSelectedText = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorSeparator() As OLE_COLOR
    ColorSeparator = m_ColorSeparator
End Property
Public Property Let ColorSeparator(ByVal NewColor As OLE_COLOR)
    m_ColorSeparator = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get ColorSeparatorBack() As OLE_COLOR
    ColorSeparatorBack = m_ColorSeparatorBack
End Property
Public Property Let ColorSeparatorBack(ByVal NewColor As OLE_COLOR)
    m_ColorSeparatorBack = NewColor
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get Font() As Font
    Set Font = m_Font
End Property
Public Property Set Font(ByRef NewValue As Font)
    Dim NewFont As New StdFont
    ' have to do it this way because otherwise we'd link with existing font object
    NewFont.Bold = NewValue.Bold
    NewFont.Charset = NewValue.Charset
    NewFont.Italic = NewValue.Italic
    NewFont.Name = NewValue.Name
    NewFont.Size = NewValue.Size
    NewFont.Strikethrough = NewValue.Strikethrough
    NewFont.Underline = NewValue.Underline
    NewFont.Weight = NewValue.Weight
    Set m_Font = NewFont
    If Not blnDesignTime Then Else PropertyChanged "Font"
    Refresh
End Property
Friend Function Friend_MenuItems() As Collection
    Dim MenuControl As Control, MenuItem As Menu
    Set Friend_MenuItems = New Collection
    If blnParentIsForm Then
        For Each MenuControl In Parent.Controls
            If TypeOf MenuControl Is Menu Then
                Set MenuItem = MenuControl
                Friend_MenuItems.Add MenuItem, CStr(ObjPtr(MenuItem))
            End If
        Next MenuControl
        Set MenuItem = Nothing
    End If
End Function
Public Property Get In3D() As Boolean
    In3D = Styles.In3D
End Property
Public Property Let In3D(ByVal NewValue As Boolean)
    Styles.In3D = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Property Get Panel() As Boolean
    Panel = Styles.Panel
End Property
Public Property Let Panel(ByVal NewValue As Boolean)
    Styles.Panel = NewValue
    Style = [Style Custom]
    Refresh
End Property
Public Sub Refresh()
    Dim lngA As Long, udtMenuItem As MENUITEMINFO
    Dim lngTopCount As Long
    If lngMenuhWnd Then
        ' get number of items
        lngTopCount = GetMenuItemCount(lngMenuhWnd) - 1
        ' set ending ID
        lngMenuTopEndID = MENU_TOP_ID + lngTopCount
        ' check if to skip the first item
        lngA = CLng(Abs(UserControl.Parent.WindowState = vbMaximized))
        ' loop through all top level menu items
        For lngA = lngA To lngTopCount
            With udtMenuItem
                ' nullify type to get string length
                .fMask = MIIM_STRING
                .dwTypeData = vbNullString
                .cch = 0
                .cbSize = Len(udtMenuItem)
                ' get string length info
                GetMenuItemInfo lngMenuhWnd, lngA, True, udtMenuItem
                ' restore ownerdraw?
                If Not (.fType And MF_OWNERDRAW) Then
                    ' initialize information to get the required data
                    .fMask = MIIM_TYPE ' MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or
                    .dwTypeData = String$(.cch, vbNullChar)
                    .cch = .cch + 1
                    .cbSize = Len(udtMenuItem)
                    ' get the data
                    GetMenuItemInfo lngMenuhWnd, lngA, True, udtMenuItem
                    ' save the menu string
                    Private_AddMenuCaption MENU_TOP_ID + lngA, .dwTypeData
                    ' set menu ownerdrawn
                    ModifyMenu lngMenuhWnd, lngA, .fType Or MF_OWNERDRAW Or MF_BYPOSITION, MENU_TOP_ID + lngA, ByVal 2&
                    .fMask = MIIM_ID
                    .wid = MENU_TOP_ID + lngA
                    SetMenuItemInfo lngMenuhWnd, lngA, True, udtMenuItem
                End If
            End With
        Next lngA
    End If
End Sub
Public Sub SetColorStyle(ByVal NewStyle As MENU_STYLE)
    Select Case NewStyle
        Case [Style Classic]
            m_ColorArrow = vbMenuText
            m_ColorArrowSelected = vbHighlightText
            m_ColorBar = vbButtonFace
            m_ColorBarBorder = m_ColorBar
            m_ColorBarHot = vbHighlight
            m_ColorBarHotBorder = m_ColorBarHot
            m_ColorBarHotText = m_ColorArrowSelected
            m_ColorBarSelected = m_ColorBarHot
            m_ColorBarSelectedBorder = m_ColorBarHot
            m_ColorBarSelectedText = m_ColorArrowSelected
            m_ColorBarText = m_ColorArrow
            m_ColorBorderBack = vbMenuBar
            m_ColorBorderInner = m_ColorBorderBack
            m_ColorBorderOuter = vb3DShadow
            m_ColorCheck = m_ColorArrow
            m_ColorCheckBack = m_ColorBorderBack
            m_ColorCheckBorder = m_ColorBorderBack
            m_ColorDisabled = m_ColorBorderBack
            m_ColorDisabledText = vbGrayText
            m_ColorImageBack = m_ColorBorderBack
            m_ColorImageBorder = m_ColorBorderBack
            m_ColorImageShadow = m_ColorDisabledText
            m_ColorItem = m_ColorBorderBack
            m_ColorItemBorder = m_ColorBorderBack
            m_ColorItemText = m_ColorArrow
            m_ColorSelected = m_ColorBarHot
            m_ColorSelectedBorder = m_ColorBarHot
            m_ColorSelectedText = m_ColorArrowSelected
            m_ColorSeparator = m_ColorDisabledText
            m_ColorSeparatorBack = m_ColorBorderBack
            m_SmoothArrow = 0
            m_SmoothArrowSelected = 0
            m_SmoothBar = 0
            m_SmoothBarBorder = 0
            m_SmoothBarHot = 0
            m_SmoothBarHotBorder = 0
            m_SmoothBarHotText = 0
            m_SmoothBarSelected = 0
            m_SmoothBarSelectedBorder = 0
            m_SmoothBarSelectedText = 0
            m_SmoothBarText = 0
            m_SmoothBorderBack = 0
            m_SmoothBorderInner = 0
            m_SmoothBorderOuter = 0
            m_SmoothCheck = 0
            m_SmoothCheckBack = 0
            m_SmoothCheckBorder = 0
            m_SmoothDisabled = 0
            m_SmoothDisabledText = 0
            m_SmoothImageBack = 0
            m_SmoothImageBorder = 0
            m_SmoothImageShadow = 0
            m_SmoothItem = 0
            m_SmoothItemBorder = 0
            m_SmoothItemText = 0
            m_SmoothSelected = 0
            m_SmoothSelectedBorder = 0
            m_SmoothSelectedText = 0
            m_SmoothSeparator = 0
            m_SmoothSeparatorBack = 0
        Case [Style Office97]
            m_ColorArrow = vbMenuText
            m_ColorArrowSelected = vbHighlightText
            m_ColorBar = vbButtonFace
            m_ColorBarBorder = m_ColorBar
            m_ColorBarHot = m_ColorBar
            m_ColorBarHotBorder = m_ColorBarHot
            m_ColorBarHotText = m_ColorArrow
            m_ColorBarSelected = m_ColorBar
            m_ColorBarSelectedBorder = m_ColorBar
            m_ColorBarSelectedText = m_ColorArrow
            m_ColorBarText = m_ColorArrow
            m_ColorBorderBack = vbMenuBar
            m_ColorBorderInner = m_ColorBorderBack
            m_ColorBorderOuter = vb3DShadow
            m_ColorCheck = m_ColorArrow
            m_ColorCheckBack = m_ColorBorderBack
            m_ColorCheckBorder = m_ColorBorderBack
            m_ColorDisabled = m_ColorBorderBack
            m_ColorDisabledText = vbGrayText
            m_ColorImageBack = m_ColorBorderBack
            m_ColorImageBorder = m_ColorBorderBack
            m_ColorImageShadow = m_ColorDisabledText
            m_ColorItem = m_ColorBorderBack
            m_ColorItemBorder = m_ColorBorderBack
            m_ColorItemText = m_ColorArrow
            m_ColorSelected = vbHighlight
            m_ColorSelectedBorder = m_ColorSelected
            m_ColorSelectedText = m_ColorArrowSelected
            m_ColorSeparator = m_ColorDisabledText
            m_ColorSeparatorBack = m_ColorBorderBack
            m_SmoothArrow = 0
            m_SmoothArrowSelected = 0
            m_SmoothBar = 0
            m_SmoothBarBorder = 0
            m_SmoothBarHot = 0
            m_SmoothBarHotBorder = 0
            m_SmoothBarHotText = 0
            m_SmoothBarSelected = 0
            m_SmoothBarSelectedBorder = 0
            m_SmoothBarSelectedText = 0
            m_SmoothBarText = 0
            m_SmoothBorderBack = 0
            m_SmoothBorderInner = 0
            m_SmoothBorderOuter = 0
            m_SmoothCheck = 0
            m_SmoothCheckBack = 2
            m_SmoothCheckBorder = 0
            m_SmoothDisabled = 0
            m_SmoothDisabledText = 0
            m_SmoothImageBack = 0
            m_SmoothImageBorder = 0
            m_SmoothImageShadow = 0
            m_SmoothItem = 0
            m_SmoothItemBorder = 0
            m_SmoothItemText = 0
            m_SmoothSelected = 0
            m_SmoothSelectedBorder = 0
            m_SmoothSelectedText = 0
            m_SmoothSeparator = 0
            m_SmoothSeparatorBack = 0
        Case [Style OfficeXP]
            m_ColorArrow = vbHighlight
            m_ColorArrowSelected = vbMenuText
            m_ColorBar = vbButtonFace
            m_ColorBarBorder = m_ColorBar
            m_ColorBarHot = m_ColorArrow
            m_ColorBarHotBorder = m_ColorArrow
            m_ColorBarHotText = vbMenuText
            m_ColorBarSelected = vbButtonFace
            m_ColorBarSelectedBorder = vb3DDKShadow
            m_ColorBarSelectedText = m_ColorBarHotText
            m_ColorBarText = m_ColorBarHotText
            m_ColorBorderBack = vbButtonFace
            m_ColorBorderInner = m_ColorBarSelected
            m_ColorBorderOuter = m_ColorBarSelectedBorder
            m_ColorCheck = m_ColorArrow
            m_ColorCheckBack = m_ColorBorderBack
            m_ColorCheckBorder = m_ColorArrow
            m_ColorDisabled = m_ColorBorderBack
            m_ColorDisabledText = vbButtonShadow
            m_ColorImageBack = m_ColorBarSelected
            m_ColorImageBorder = m_ColorBarSelected
            m_ColorImageShadow = m_ColorArrow
            m_ColorItem = m_ColorBorderBack
            m_ColorItemBorder = m_ColorBorderBack
            m_ColorItemText = m_ColorBarHotText
            m_ColorSelected = m_ColorBarHot
            m_ColorSelectedBorder = m_ColorArrow
            m_ColorSelectedText = m_ColorBarHotText
            m_ColorSeparator = vbButtonShadow
            m_ColorSeparatorBack = m_ColorBorderBack
            m_SmoothArrow = 1
            m_SmoothArrowSelected = 1
            m_SmoothBar = 0
            m_SmoothBarBorder = 1
            m_SmoothBarHot = 2
            m_SmoothBarHotBorder = m_SmoothArrow
            m_SmoothBarHotText = 0
            m_SmoothBarSelected = 1
            m_SmoothBarSelectedBorder = 1
            m_SmoothBarSelectedText = m_SmoothBarHotText
            m_SmoothBarText = m_SmoothBarHotText
            m_SmoothBorderBack = 3
            m_SmoothBorderInner = m_SmoothBarSelected
            m_SmoothBorderOuter = m_SmoothBarSelectedBorder
            m_SmoothCheck = 1
            m_SmoothCheckBack = m_SmoothBorderBack
            m_SmoothCheckBorder = 0
            m_SmoothDisabled = m_SmoothBorderBack
            m_SmoothDisabledText = 1
            m_SmoothImageBack = m_SmoothBarSelected
            m_SmoothImageBorder = m_SmoothBarSelected
            m_SmoothImageShadow = m_SmoothArrow
            m_SmoothItem = m_SmoothBorderBack
            m_SmoothItemBorder = m_SmoothBorderBack
            m_SmoothItemText = m_SmoothBarHotText
            m_SmoothSelected = m_SmoothBarHot
            m_SmoothSelectedBorder = m_SmoothArrow
            m_SmoothSelectedText = m_SmoothBarHotText
            m_SmoothSeparator = 1
            m_SmoothSeparatorBack = m_SmoothBorderBack
    End Select
    Private_SetColors
End Sub
Public Sub SetStyle(ByVal NewStyle As MENU_STYLE, Optional Private_SetColors As Boolean = False, Optional ByVal In3D As Boolean, Optional ByVal Panel As Boolean, Optional ByVal ImageSize As Byte = m_def_ImageSize)
    With Styles
        Select Case NewStyle
            Case [Style Classic]
                .BarBorder = False
                .BarHotBorder = False
                .BarSelectedBorder = False
                .CheckBorder = False
                .ImageBorder = False
                .In3D = True
                .ItemBorder = False
                .Panel = False
                .SelectedBorder = False
            Case [Style Office97]
                .BarBorder = False
                .BarHotBorder = True
                .BarSelectedBorder = True
                .CheckBorder = True
                .ImageBorder = True
                .In3D = True
                .ItemBorder = False
                .Panel = False
                .SelectedBorder = False
            Case [Style OfficeXP]
                .BarBorder = False
                .BarHotBorder = True
                .BarSelectedBorder = True
                .CheckBorder = True
                .ImageBorder = False
                .In3D = False
                .ItemBorder = False
                .Panel = True
                .SelectedBorder = True
            Case Else
                If Not IsMissing(In3D) Then .In3D = In3D
                If Not IsMissing(Panel) Then .Panel = Panel
        End Select
        If Not IsMissing(ImageSize) And ImageSize > 0 Then .ImageSize = ImageSize
        If Private_SetColors Then SetColorStyle NewStyle
    End With
    m_Style = NewStyle
    RaiseEvent StyleChange
End Sub
Public Property Get SmoothArrow() As Byte
    SmoothArrow = m_SmoothArrow
End Property
Public Property Let SmoothArrow(ByVal NewSmooth As Byte)
    m_SmoothArrow = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothArrowSelected() As Byte
    SmoothArrowSelected = m_SmoothArrowSelected
End Property
Public Property Let SmoothArrowSelected(ByVal NewSmooth As Byte)
    m_SmoothArrowSelected = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBar() As Byte
    SmoothBar = m_SmoothBar
End Property
Public Property Let SmoothBar(ByVal NewSmooth As Byte)
    m_SmoothBar = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarBorder() As Byte
    SmoothBarBorder = m_SmoothBarBorder
End Property
Public Property Let SmoothBarBorder(ByVal NewSmooth As Byte)
    m_SmoothBarBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarHot() As Byte
    SmoothBarHot = m_SmoothBarHot
End Property
Public Property Let SmoothBarHot(ByVal NewSmooth As Byte)
    m_SmoothBarHot = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarHotBorder() As Byte
    SmoothBarHotBorder = m_SmoothBarHotBorder
End Property
Public Property Let SmoothBarHotBorder(ByVal NewSmooth As Byte)
    m_SmoothBarHotBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarHotText() As Byte
    SmoothBarHotText = m_SmoothBarHotText
End Property
Public Property Let SmoothBarHotText(ByVal NewSmooth As Byte)
    m_SmoothBarHotText = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarSelected() As Byte
    SmoothBarSelected = m_SmoothBarSelected
End Property
Public Property Let SmoothBarSelected(ByVal NewSmooth As Byte)
    m_SmoothBarSelected = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarSelectedBorder() As Byte
    SmoothBarSelectedBorder = m_SmoothBarSelectedBorder
End Property
Public Property Let SmoothBarSelectedBorder(ByVal NewSmooth As Byte)
    m_SmoothBarSelectedBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarSelectedText() As Byte
    SmoothBarSelectedText = m_SmoothBarSelectedText
End Property
Public Property Let SmoothBarSelectedText(ByVal NewSmooth As Byte)
    m_SmoothBarSelectedText = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBarText() As Byte
    SmoothBarText = m_SmoothBarText
End Property
Public Property Let SmoothBarText(ByVal NewSmooth As Byte)
    m_SmoothBarText = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBorderBack() As Byte
    SmoothBorderBack = m_SmoothBorderBack
End Property
Public Property Let SmoothBorderBack(ByVal NewSmooth As Byte)
    m_SmoothBorderBack = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBorderinner() As Byte
    SmoothBorderinner = m_SmoothBorderInner
End Property
Public Property Let SmoothBorderinner(ByVal NewSmooth As Byte)
    m_SmoothBorderInner = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothBorderOuter() As Byte
    SmoothBorderOuter = m_SmoothBorderOuter
End Property
Public Property Let SmoothBorderOuter(ByVal NewSmooth As Byte)
    m_SmoothBorderOuter = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothCheck() As Byte
    SmoothCheck = m_SmoothCheck
End Property
Public Property Let SmoothCheck(ByVal NewSmooth As Byte)
    m_SmoothCheck = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothCheckBack() As Byte
    SmoothCheckBack = m_SmoothCheckBack
End Property
Public Property Let SmoothCheckBack(ByVal NewSmooth As Byte)
    m_SmoothCheckBack = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothCheckBorder() As Byte
    SmoothCheckBorder = m_SmoothCheckBorder
End Property
Public Property Let SmoothCheckBorder(ByVal NewSmooth As Byte)
    m_SmoothCheckBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothDisabled() As Byte
    SmoothDisabled = m_SmoothDisabled
End Property
Public Property Let SmoothDisabled(ByVal NewSmooth As Byte)
    m_SmoothDisabled = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothDisabledText() As Byte
    SmoothDisabledText = m_SmoothDisabledText
End Property
Public Property Let SmoothDisabledText(ByVal NewSmooth As Byte)
    m_SmoothDisabledText = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothImageBack() As Byte
    SmoothImageBack = m_SmoothImageBack
End Property
Public Property Let SmoothImageBack(ByVal NewSmooth As Byte)
    m_SmoothImageBack = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothImageBorder() As Byte
    SmoothImageBorder = m_SmoothImageBorder
End Property
Public Property Let SmoothImageBorder(ByVal NewSmooth As Byte)
    m_SmoothImageBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothImageShadow() As Byte
    SmoothImageShadow = m_SmoothImageShadow
End Property
Public Property Let SmoothImageShadow(ByVal NewSmooth As Byte)
    m_SmoothImageShadow = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothItem() As Byte
    SmoothItem = m_SmoothItem
End Property
Public Property Let SmoothItem(ByVal NewSmooth As Byte)
    m_SmoothItem = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothItemBorder() As Byte
    SmoothItemBorder = m_SmoothItemBorder
End Property
Public Property Let SmoothItemBorder(ByVal NewSmooth As Byte)
    m_SmoothItemBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothItemText() As Byte
    SmoothItemText = m_SmoothItemText
End Property
Public Property Let SmoothItemText(ByVal NewSmooth As Byte)
    m_SmoothItemText = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothSelected() As Byte
    SmoothSelected = m_SmoothSelected
End Property
Public Property Let SmoothSelected(ByVal NewSmooth As Byte)
    m_SmoothSelected = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothSelectedBorder() As Byte
    SmoothSelectedBorder = m_SmoothSelectedBorder
End Property
Public Property Let SmoothSelectedBorder(ByVal NewSmooth As Byte)
    m_SmoothSelectedBorder = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothSelectedText() As Byte
    SmoothSelectedText = m_SmoothSelectedText
End Property
Public Property Let SmoothSelectedText(ByVal NewSmooth As Byte)
    m_SmoothSelectedText = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothSeparator() As Byte
    SmoothSeparator = m_SmoothSeparator
End Property
Public Property Let SmoothSeparator(ByVal NewSmooth As Byte)
    m_SmoothSeparator = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get SmoothSeparatorBack() As Byte
    SmoothSeparatorBack = m_SmoothSeparatorBack
End Property
Public Property Let SmoothSeparatorBack(ByVal NewSmooth As Byte)
    m_SmoothSeparatorBack = NewSmooth
    Private_SetColors
    If blnDesignTime Then StartingStyle = [Style Custom] Else Style = [Style Custom]
    Refresh
End Property
Public Property Get StartingStyle() As MENU_STYLE
Attribute StartingStyle.VB_Description = "WARNING! Changing this value will reset all color and style settings!"
    StartingStyle = m_StartingStyle
End Property
Public Property Let StartingStyle(ByVal NewStyle As MENU_STYLE)
    Dim intQuery As Integer
    ' can't change in runtime
    If blnDesignTime Then
        ' no point changing to what have already...
        If m_StartingStyle = NewStyle Then Exit Property
        ' show warning message?
        If m_StartingStyle = [Style Custom] Then
            intQuery = MsgBox("Your current style is custom. If you change style, you lose all current color and style settings." & vbCrLf & vbCrLf & "Proceed?", vbQuestion Or vbYesNo, "StartingStyle")
            If intQuery = vbNo Then Exit Property
        End If
        ' set the style!
        m_StartingStyle = NewStyle
        SetStyle NewStyle, True
    End If
End Property
' this function can be called to convert a BSTR string to UTF-8 string
Public Function StrToUTF8(ByRef Text As String) As String
    If LenB(Text) Then StrToUTF8 = StrConv(Private_UTF16toUTF8(Text), vbUnicode)
    'Dim bytUTF8() As Byte, lngA As Long
    'If LenB(Text) Then
    '    bytUTF8 = Private_UTF16toUTF8(Text)
    '    ReDim Preserve bytUTF8(UBound(bytUTF8) * 2 + 1)
    '    Debug.Print UBound(bytUTF8)
    '    For lngA = (UBound(bytUTF8) - 1) To 0 Step -2
    '        bytUTF8(lngA) = bytUTF8(lngA \ 2)
    '        bytUTF8(lngA + 1) = 0
    '    Next lngA
    '    StrToUTF8 = bytUTF8
    'End If
End Function
Public Property Get Style() As MENU_STYLE
    Style = m_Style
End Property
Public Property Let Style(ByVal NewStyle As MENU_STYLE)
    SetStyle NewStyle
    'm_Style = NewStyle
    'RaiseEvent StyleChange
End Property
Public Property Get Unicode() As MENU_UNICODE
    Unicode = m_Unicode
End Property
Public Property Let Unicode(ByVal NewValue As MENU_UNICODE)
    m_Unicode = NewValue
    Private_DetectUnicode
    Refresh
End Property
' this function can be called to convert a UTF-8 string to BSTR string
Public Function UTF8toStr(ByRef Text As String) As String
    If LenB(Text) Then UTF8toStr = Private_UTF8toUTF16(StrConv(Text, vbFromUnicode))
    'Dim bytUTF8() As Byte, lngA As Long
    'If LenB(Text) Then
    '    bytUTF8 = Text
    '    For lngA = 2 To UBound(bytUTF8) - 1 Step 2
    '        bytUTF8(lngA \ 2) = bytUTF8(lngA)
    '    Next lngA
    '    ReDim Preserve bytUTF8((lngA \ 2) - 1)
    '    UTF8toStr = Private_UTF8toUTF16(CStr(bytUTF8))
    'End If
End Function
Private Sub chkChecked_Click()
    If Not m_MenuCurrent Is Nothing Then
        If m_MenuCurrent.Caption = "-" Then
            chkChecked.Value = vbUnchecked
        Else
            On Error Resume Next
            m_MenuCurrent.Checked = chkChecked.Value = vbChecked
            If Err.Number = 0 Then On Error GoTo 0 Else On Error GoTo 0: chkChecked.Value = Abs(m_MenuCurrent.Checked)
        End If
    End If
End Sub
Private Sub chkEnabled_Click()
    If Not m_MenuCurrent Is Nothing Then
        If m_MenuCurrent.Caption = "-" Then
            chkEnabled.Value = vbChecked
        Else
            m_MenuCurrent.Enabled = chkEnabled.Value = vbChecked
        End If
    End If
End Sub
Private Sub chkVisible_Click()
    If Not m_MenuCurrent Is Nothing Then m_MenuCurrent.Visible = chkVisible.Value = vbChecked
End Sub
Private Sub cmbMenu_Click()
    Dim lngPtr As Long, lngA As Long
    Set m_MenuCurrent = Nothing
    If cmbMenu.ListIndex >= -1 Then
        lngPtr = cmbMenu.ItemData(cmbMenu.ListIndex)
        For lngA = 0 To Parent.Controls.Count - 1
            If ObjPtr(Parent.Controls.Item(lngA)) = lngPtr Then Set m_MenuCurrent = Parent.Controls.Item(lngA): Exit For
        Next lngA
    End If
    If Not m_MenuCurrent Is Nothing Then
        txtCaption.Text = UTF8toStr(m_MenuCurrent.Caption)
        chkChecked.Value = Abs(m_MenuCurrent.Checked)
        chkEnabled.Value = Abs(m_MenuCurrent.Enabled)
        chkVisible.Value = Abs(m_MenuCurrent.Visible)
        chkWindowList.Value = Abs(m_MenuCurrent.WindowList)
    Else
        txtCaption.Text = vbNullString
        chkChecked.Value = vbUnchecked
        chkEnabled.Value = vbUnchecked
        chkVisible.Value = vbUnchecked
        chkWindowList.Value = vbUnchecked
    End If
End Sub
Private Sub cmdCaption_Click()
    txtCaption_Change
End Sub
Private Sub tmrDesignTime_Timer()
    ' disable the timer
    tmrDesignTime.Enabled = False
    ' and then enable for design time
    ' Private_Enable
End Sub
Private Sub txtCaption_Change()
    Dim strCaption As String
    If Not m_MenuCurrent Is Nothing Then
        strCaption = txtCaption.Text
        If strCaption = "-" Then
            If m_MenuCurrent.Checked Then chkChecked.Value = vbUnchecked
            If Not m_MenuCurrent.Enabled Then chkEnabled.Value = vbChecked
        End If
        On Error Resume Next
        m_MenuCurrent.Caption = StrToUTF8(strCaption)
        If Err.Number = 0 Then On Error GoTo 0 Else On Error GoTo 0: txtCaption.Text = UTF8toStr(m_MenuCurrent.Caption)
    End If
End Sub

Private Sub txtCaption_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdCaption_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Select Case PropertyName
        Case "DisplayName"
                lblTitle.Caption = Ambient.DisplayName
        Case "ScaleUnits"
            UserControl.ScaleMode = Private_GetScaleMode
    End Select
End Sub

Private Sub UserControl_EnterFocus()
    Dim MenuItem As Menu
    blnEditMode = blnDesignTime
    If blnEditMode Then
        For Each MenuItem In Friend_MenuItems
            cmbMenu.AddItem MenuItem.Name & IIf(MenuItem.Index > -1, " (" & MenuItem.Index & ")", vbNullString)
            cmbMenu.ItemData(cmbMenu.NewIndex) = ObjPtr(MenuItem)
        Next MenuItem
        If cmbMenu.ListCount Then cmbMenu.ListIndex = 0
        fraMenu.Visible = True
        lblTitle.Visible = False
    End If
End Sub

Private Sub UserControl_ExitFocus()
    If blnEditMode Then
        cmbMenu.Clear
        blnEditMode = False
        fraMenu.Visible = False
        lblTitle.Visible = True
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim lngA As Long, lngNew As Long
    blnInIDE = (App.LogMode = 0)
    ' detect Windows
    m_WINNT = (Environ$("OS") = "Windows_NT")
    ' here we initialize the smooth color values
    For lngA = 0 To 255
        lngNew = lngA + 76 - ((lngA + 32) / 76) * 19
        If lngNew < 0 Then lngNew = 0
        If lngNew > 255 Then lngNew = 255
        SmoothColors(lngA) = lngNew
    Next lngA
    ' set default image size
    Styles.ImageSize = m_def_ImageSize
End Sub
' just so you don't forget: this happens only when the control is created the first time
Private Sub UserControl_InitProperties()
    blnDesignTime = Not Ambient.UserMode
    blnParentIsForm = TypeOf Parent Is Form
    UserControl.ScaleMode = Private_GetScaleMode
    lblTitle.Caption = UserControl.Ambient.DisplayName
    ' set font
    Set m_Font = UserControl.Ambient.Font
    ' set Unicode mode
    m_Unicode = [Detect Windows]
    Private_DetectUnicode
    ' set initial settings and then enable it
    m_StartingStyle = [Style OfficeXP]
    SetStyle m_StartingStyle, True, , , m_def_ImageSize
    Private_Enable
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    blnDesignTime = Not Ambient.UserMode
    blnParentIsForm = TypeOf Parent Is Form
    UserControl.ScaleMode = Private_GetScaleMode
    lblTitle.Caption = UserControl.Ambient.DisplayName
    ' read properties...
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Unicode = PropBag.ReadProperty("Unicode", [Detect Windows])
    ' set Styles.InANSI by calling DetectUnicode
    Private_DetectUnicode
    ' set image size
    Styles.ImageSize = CByte(PropBag.ReadProperty("ImageSize", m_def_ImageSize))
    ' set the default color theme
    m_StartingStyle = PropBag.ReadProperty("StartingStyle", [Style OfficeXP])
    SetStyle m_StartingStyle
    ' set styles
    Styles.BarBorder = PropBag.ReadProperty("BorderBar", Styles.BarBorder)
    Styles.BarHotBorder = PropBag.ReadProperty("BorderBarHot", Styles.BarHotBorder)
    Styles.BarSelectedBorder = PropBag.ReadProperty("BorderBarSelected", Styles.BarSelectedBorder)
    Styles.CheckBorder = PropBag.ReadProperty("BorderCheck", Styles.CheckBorder)
    Styles.ImageBorder = PropBag.ReadProperty("BorderImage", Styles.ImageBorder)
    Styles.ItemBorder = PropBag.ReadProperty("BorderItem", Styles.ItemBorder)
    Styles.SelectedBorder = PropBag.ReadProperty("BorderSelected", Styles.SelectedBorder)
    Styles.In3D = PropBag.ReadProperty("In3D", Styles.In3D)
    Styles.Panel = PropBag.ReadProperty("Panel", Styles.Panel)
    Styles.ImageSize = PropBag.ReadProperty("ImageSize", m_def_ImageSize)
    ' get colors!
    m_ColorArrow = CLng(PropBag.ReadProperty("ColorArrow", m_ColorArrow))
    m_ColorArrowSelected = CLng(PropBag.ReadProperty("ColorArrowSelected", m_ColorArrowSelected))
    m_ColorBar = CLng(PropBag.ReadProperty("ColorBar", m_ColorBar))
    m_ColorBarBorder = CLng(PropBag.ReadProperty("ColorBarBorder", m_ColorBarBorder))
    m_ColorBarHot = CLng(PropBag.ReadProperty("ColorBarHot", m_ColorBarHot))
    m_ColorBarHotBorder = CLng(PropBag.ReadProperty("ColorBarHotBorder", m_ColorBarHotBorder))
    m_ColorBarHotText = CLng(PropBag.ReadProperty("ColorBarHotText", m_ColorBarHotText))
    m_ColorBarSelected = CLng(PropBag.ReadProperty("ColorBarSelected", m_ColorBarSelected))
    m_ColorBarSelectedBorder = CLng(PropBag.ReadProperty("ColorBarSelectedBorder", m_ColorBarSelectedBorder))
    m_ColorBarSelectedText = CLng(PropBag.ReadProperty("ColorBarSelectedText", m_ColorBarSelectedText))
    m_ColorBarText = CLng(PropBag.ReadProperty("ColorBarText", m_ColorBarText))
    m_ColorBorderBack = CLng(PropBag.ReadProperty("ColorBorderBack", m_ColorBorderBack))
    m_ColorBorderInner = CLng(PropBag.ReadProperty("ColorBorderInner", m_ColorBorderInner))
    m_ColorBorderOuter = CLng(PropBag.ReadProperty("ColorBorderOuter", m_ColorBorderOuter))
    m_ColorCheck = CLng(PropBag.ReadProperty("ColorCheck", m_ColorCheck))
    m_ColorCheckBack = CLng(PropBag.ReadProperty("ColorCheckBack", m_ColorCheckBack))
    m_ColorCheckBorder = CLng(PropBag.ReadProperty("ColorCheckBorder", m_ColorCheckBorder))
    m_ColorDisabled = CLng(PropBag.ReadProperty("ColorDisabled", m_ColorDisabled))
    m_ColorDisabledText = CLng(PropBag.ReadProperty("ColorDisabledText", m_ColorDisabledText))
    m_ColorImageBack = CLng(PropBag.ReadProperty("ColorImageBack", m_ColorImageBack))
    m_ColorImageBorder = CLng(PropBag.ReadProperty("ColorImageBorder", m_ColorImageBorder))
    m_ColorImageShadow = CLng(PropBag.ReadProperty("ColorImageShadow", m_ColorImageShadow))
    m_ColorItem = CLng(PropBag.ReadProperty("ColorItem", m_ColorItem))
    m_ColorItemBorder = CLng(PropBag.ReadProperty("ColorItemBorder", m_ColorItemBorder))
    m_ColorItemText = CLng(PropBag.ReadProperty("ColorItemText", m_ColorItemText))
    m_ColorSelected = CLng(PropBag.ReadProperty("ColorSelected", m_ColorSelected))
    m_ColorSelectedBorder = CLng(PropBag.ReadProperty("ColorSelectedBorder", m_ColorSelectedBorder))
    m_ColorSelectedText = CLng(PropBag.ReadProperty("ColorSelectedText", m_ColorSelectedText))
    m_ColorSeparator = CLng(PropBag.ReadProperty("ColorSeparator", m_ColorSeparator))
    m_ColorSeparatorBack = CLng(PropBag.ReadProperty("ColorSeparatorBack", m_ColorSeparatorBack))
    ' get smooth settings!
    m_SmoothArrow = CByte(PropBag.ReadProperty("SmoothArrow", m_SmoothArrow))
    m_SmoothArrowSelected = CByte(PropBag.ReadProperty("SmoothArrowSelected", m_SmoothArrowSelected))
    m_SmoothBar = CByte(PropBag.ReadProperty("SmoothBar", m_SmoothBar))
    m_SmoothBarBorder = CByte(PropBag.ReadProperty("SmoothBarBorder", m_SmoothBarBorder))
    m_SmoothBarHot = CByte(PropBag.ReadProperty("SmoothBarHot", m_SmoothBarHot))
    m_SmoothBarHotBorder = CByte(PropBag.ReadProperty("SmoothBarHotBorder", m_SmoothBarHotBorder))
    m_SmoothBarHotText = CByte(PropBag.ReadProperty("SmoothBarHotText", m_SmoothBarHotText))
    m_SmoothBarSelected = CByte(PropBag.ReadProperty("SmoothBarSelected", m_SmoothBarSelected))
    m_SmoothBarSelectedBorder = CByte(PropBag.ReadProperty("SmoothBarSelectedBorder", m_SmoothBarSelectedBorder))
    m_SmoothBarSelectedText = CByte(PropBag.ReadProperty("SmoothBarSelectedText", m_SmoothBarSelectedText))
    m_SmoothBarText = CByte(PropBag.ReadProperty("SmoothBarText", m_SmoothBarText))
    m_SmoothBorderBack = CByte(PropBag.ReadProperty("SmoothBorderBack", m_SmoothBorderBack))
    m_SmoothBorderInner = CByte(PropBag.ReadProperty("SmoothBorderInner", m_SmoothBorderInner))
    m_SmoothBorderOuter = CByte(PropBag.ReadProperty("SmoothBorderOuter", m_SmoothBorderOuter))
    m_SmoothCheck = CByte(PropBag.ReadProperty("SmoothCheck", m_SmoothCheck))
    m_SmoothCheckBack = CByte(PropBag.ReadProperty("SmoothCheckBack", m_SmoothCheckBack))
    m_SmoothCheckBorder = CByte(PropBag.ReadProperty("SmoothCheckBorder", m_SmoothCheckBorder))
    m_SmoothDisabled = CByte(PropBag.ReadProperty("SmoothDisabled", m_SmoothDisabled))
    m_SmoothDisabledText = CByte(PropBag.ReadProperty("SmoothDisabledText", m_SmoothDisabledText))
    m_SmoothImageBack = CByte(PropBag.ReadProperty("SmoothImageBack", m_SmoothImageBack))
    m_SmoothImageBorder = CByte(PropBag.ReadProperty("SmoothImageBorder", m_SmoothImageBorder))
    m_SmoothImageShadow = CByte(PropBag.ReadProperty("SmoothImageShadow", m_SmoothImageShadow))
    m_SmoothItem = CByte(PropBag.ReadProperty("SmoothItem", m_SmoothItem))
    m_SmoothItemBorder = CByte(PropBag.ReadProperty("SmoothItemBorder", m_SmoothItemBorder))
    m_SmoothItemText = CByte(PropBag.ReadProperty("SmoothItemText", m_SmoothItemText))
    m_SmoothSelected = CByte(PropBag.ReadProperty("SmoothSelected", m_SmoothSelected))
    m_SmoothSelectedBorder = CByte(PropBag.ReadProperty("SmoothSelectedBorder", m_SmoothSelectedBorder))
    m_SmoothSelectedText = CByte(PropBag.ReadProperty("SmoothSelectedText", m_SmoothSelectedText))
    m_SmoothSeparator = CByte(PropBag.ReadProperty("SmoothSeparator", m_SmoothSeparator))
    m_SmoothSeparatorBack = CByte(PropBag.ReadProperty("SmoothSeparatorBack", m_SmoothSeparatorBack))
    ' initialize colors!
    Private_SetColors
    ' we have no use of parent subclassing in design time...
    If Not blnDesignTime Then
        ' we subclass the parent window and wait for WM_SHOWWINDOW message...
        ' then we can really subclass a menu, because before WM_SHOWWINDOW we have no menu
        If ssc_Subclass(UserControl.Parent.hWnd, , , , , True) Then
            ssc_AddMsg UserControl.Parent.hWnd, MSG_AFTER, WM_SHOWWINDOW
            Debug.Print "UniMenu (WM_SHOWWINDOW): Started subclassing! " & Hex$(UserControl.Parent.hWnd)
        End If
    Else
        ' support for design time
        tmrDesignTime.Interval = 1
        tmrDesignTime.Enabled = True
    End If
End Sub
Private Sub UserControl_Resize()
    If Not blnDesignTime Then
    Else
        UserControl.Size ScaleX(fraMenu.Width + fraMenu.Left * 2, ScaleMode, vbTwips), ScaleY(fraMenu.Height + fraMenu.Top * 2, ScaleMode, vbTwips)
        shpTitle.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        lblTitle.Move (UserControl.ScaleWidth - lblTitle.Width) \ 2, (UserControl.ScaleHeight - lblTitle.Height) \ 2
    End If
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    ' make sure we end subclassing, hooking and so on
    Private_Disable
    ' clear temporary string array
    Erase Captions
    Set m_Font = Nothing
    ' make sure we unhook the menu when not using the class anymore
    If Not blnInIDE Then Private_ClearMenuHooks
    ' remove owner-drawing (in design time)
    If blnDesignTime Then Private_ClearOwnerDrawing
    ssc_Terminate
    shk_TerminateHooks
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim MenuItem As Object
    PropBag.WriteProperty "BorderBar", Styles.BarBorder
    PropBag.WriteProperty "BorderBarHot", Styles.BarHotBorder
    PropBag.WriteProperty "BorderBarSelected", Styles.BarSelectedBorder
    PropBag.WriteProperty "BorderCheck", Styles.CheckBorder
    PropBag.WriteProperty "BorderImage", Styles.ImageBorder
    PropBag.WriteProperty "BorderItem", Styles.ItemBorder
    PropBag.WriteProperty "BorderSelected", Styles.SelectedBorder
    PropBag.WriteProperty "In3D", Styles.In3D
    PropBag.WriteProperty "Panel", Styles.Panel
    PropBag.WriteProperty "ImageSize", Styles.ImageSize
    PropBag.WriteProperty "ColorArrow", m_ColorArrow
    PropBag.WriteProperty "ColorArrowSelected", m_ColorArrowSelected
    PropBag.WriteProperty "ColorBar", m_ColorBar
    PropBag.WriteProperty "ColorBarBorder", m_ColorBarBorder
    PropBag.WriteProperty "ColorBarHot", m_ColorBarHot
    PropBag.WriteProperty "ColorBarHotBorder", m_ColorBarHotBorder
    PropBag.WriteProperty "ColorBarHotText", m_ColorBarHotText
    PropBag.WriteProperty "ColorBarSelected", m_ColorBarSelected
    PropBag.WriteProperty "ColorBarSelectedBorder", m_ColorBarSelectedBorder
    PropBag.WriteProperty "ColorBarSelectedText", m_ColorBarSelectedText
    PropBag.WriteProperty "ColorBarText", m_ColorBarText
    PropBag.WriteProperty "ColorBorderBack", m_ColorBorderBack
    PropBag.WriteProperty "ColorBorderInner", m_ColorBorderInner
    PropBag.WriteProperty "ColorBorderOuter", m_ColorBorderOuter
    PropBag.WriteProperty "ColorCheck", m_ColorCheck
    PropBag.WriteProperty "ColorCheckBack", m_ColorCheckBack
    PropBag.WriteProperty "ColorCheckBorder", m_ColorCheckBorder
    PropBag.WriteProperty "ColorDisabled", m_ColorDisabled
    PropBag.WriteProperty "ColorDisabledText", m_ColorDisabledText
    PropBag.WriteProperty "ColorImageBack", m_ColorImageBack
    PropBag.WriteProperty "ColorImageBorder", m_ColorImageBorder
    PropBag.WriteProperty "ColorImageShadow", m_ColorImageShadow
    PropBag.WriteProperty "ColorItem", m_ColorItem
    PropBag.WriteProperty "ColorItemBorder", m_ColorItemBorder
    PropBag.WriteProperty "ColorItemText", m_ColorItemText
    PropBag.WriteProperty "ColorSelected", m_ColorSelected
    PropBag.WriteProperty "ColorSelectedBorder", m_ColorSelectedBorder
    PropBag.WriteProperty "ColorSelectedText", m_ColorSelectedText
    PropBag.WriteProperty "ColorSeparator", m_ColorSeparator
    PropBag.WriteProperty "ColorSeparatorBack", m_ColorSeparatorBack
    PropBag.WriteProperty "Font", m_Font, Ambient.Font
    PropBag.WriteProperty "ImageSize", Styles.ImageSize, m_def_ImageSize
    PropBag.WriteProperty "SmoothArrow", m_SmoothArrow
    PropBag.WriteProperty "SmoothArrowSelected", m_SmoothArrowSelected
    PropBag.WriteProperty "SmoothBar", m_SmoothBar
    PropBag.WriteProperty "SmoothBarBorder", m_SmoothBarBorder
    PropBag.WriteProperty "SmoothBarHot", m_SmoothBarHot
    PropBag.WriteProperty "SmoothBarHotBorder", m_SmoothBarHotBorder
    PropBag.WriteProperty "SmoothBarHotText", m_SmoothBarHotText
    PropBag.WriteProperty "SmoothBarSelected", m_SmoothBarSelected
    PropBag.WriteProperty "SmoothBarSelectedBorder", m_SmoothBarSelectedBorder
    PropBag.WriteProperty "SmoothBarSelectedText", m_SmoothBarSelectedText
    PropBag.WriteProperty "SmoothBarText", m_SmoothBarText
    PropBag.WriteProperty "SmoothBorderBack", m_SmoothBorderBack
    PropBag.WriteProperty "SmoothBorderInner", m_SmoothBorderInner
    PropBag.WriteProperty "SmoothBorderOuter", m_SmoothBorderOuter
    PropBag.WriteProperty "SmoothCheck", m_SmoothCheck
    PropBag.WriteProperty "SmoothCheckBack", m_SmoothCheckBack
    PropBag.WriteProperty "SmoothCheckBorder", m_SmoothCheckBorder
    PropBag.WriteProperty "SmoothDisabled", m_SmoothDisabled
    PropBag.WriteProperty "SmoothDisabledText", m_SmoothDisabledText
    PropBag.WriteProperty "SmoothImageBack", m_SmoothImageBack
    PropBag.WriteProperty "SmoothImageBorder", m_SmoothImageBorder
    PropBag.WriteProperty "SmoothImageShadow", m_SmoothImageShadow
    PropBag.WriteProperty "SmoothItem", m_SmoothItem
    PropBag.WriteProperty "SmoothItemBorder", m_SmoothItemBorder
    PropBag.WriteProperty "SmoothItemText", m_SmoothItemText
    PropBag.WriteProperty "SmoothSelected", m_SmoothSelected
    PropBag.WriteProperty "SmoothSelectedBorder", m_SmoothSelectedBorder
    PropBag.WriteProperty "SmoothSelectedText", m_SmoothSelectedText
    PropBag.WriteProperty "SmoothSeparator", m_SmoothSeparator
    PropBag.WriteProperty "SmoothSeparatorBack", m_SmoothSeparatorBack
    PropBag.WriteProperty "StartingStyle", m_StartingStyle, [Style OfficeXP]
    PropBag.WriteProperty "Unicode", m_Unicode, [Detect Windows]
End Sub

'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&
      ssc_Subclass = True                                                     'Indicate success
      
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        
    End If

 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'ssc_Subclass has failed after memory allocation, so release the memory
      
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks SubclassThunk
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    zUnThunk lng_hWnd, SubclassThunk
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)                           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                              ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                         'If the message is to be removed from the before original WndProc table...
              zDelMsg Messages(M), IDX_BTABLE, z_ScMem                          'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                                          'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(M), IDX_ATABLE, z_ScMem                          'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'-SelfHook code------------------------------------------------------------------------------------
'-The following routines are exclusively for the shk_SetHook routines----------------------------
Private Function shk_SetHook(ByVal HookType As eHookType, _
                    Optional ByVal bGlobal As Boolean, _
                    Optional ByVal When As eMsgWhen = MSG_BEFORE, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False) As Boolean 'Setting specified hook

    '*************************************************************************************************
    '* HookType - One of the eHookType enumerators
    '* bGlobal - If False, then hook applies to app's thread else it applies Globally (only supported by WH_KEYBOARD_LL & WH_MOUSE_LL)
    '* When - either MSG_AFTER, MSG_BEFORE or MSG_BEFORE_AFTER
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is no reason to set this to false
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '*************************************************************************************************
    ' Hook procedure must be declared identical to the one near the end of this class (Sample at Ordinal #2)

    ' \\LaVolpe - The ASM for this procedure rewritten to mirror Paul Caton's SelfSub ASM
    '       Therefore, it appears to be crash proof and allows a choice of whether you want
    '       hook messages before and/or after the VB gets the message
    
    Dim z_Sc(0 To 66) As Long                   'Thunk machine-code initialised here
    Const MEM_LEN      As Long = 4 * 67         'Thunk length in bytes (last # must be = UBound zSc() + 1)
    
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3             'Thunk data index of the EbMode function address
    Const IDX_CNH       As Long = 4             'Thunk data index of the CallNextHook function address
    Const IDX_UNW       As Long = 5             'Thunk data index of the UnhookWindowsEx function address
    Const IDX_OBJCHK    As Long = 6             'Thunk data index of the callback validation token
    Const IDX_BADPTR    As Long = 7             'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8             'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10            'Thunk data index of the callback method address
    Const IDX_BTABLE    As Long = 11            'Thunk data index of the Before flag
    Const IDX_ATABLE    As Long = 12            'Thunk data index of the After flag
    Const IDX_EBX       As Long = 16            'Thunk code patch index of the thunk data
    Const PROC_OFF      As Long = &H38          'Thunk offset to the HookProc execution address
    Const SUB_NAME      As String = "shk_SetHook" 'This routine's name
    
    Dim nAddr           As Long
    Dim nID             As Long
    Dim nMyID           As Long
    Dim bIDE            As Boolean
    Dim z_ScMem         As Long                 'Thunk base address

    If oCallback Is Nothing Then Set oCallback = Me 'If the user hasn't specified the callback owner
      
    nAddr = zAddressOf(oCallback, nOrdinal)         'Get the address of the specified ordinal method
    If nAddr = 0 Then                               'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
    End If
       
    If Not bGlobal Then nID = App.ThreadID                      ' thread ID to be used if not global hook
      
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)    'Allocate executable memory
    
    If z_ScMem <> 0 Then                                        'Ensure the allocation succeeded
      
        If z_hkFunk Is Nothing Then Set z_hkFunk = New Collection   'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                                'Catch double subclassing
          z_hkFunk.Add z_ScMem, "h" & HookType                      'Add the hook/thunk-address to the collection
        On Error GoTo 0
        
        ' create the thunk; zSc(16) filled in below along with z_Sc(2-13)
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H3FFF631: z_Sc(18) = &H75047339: z_Sc(19) = &H2873FF2F: z_Sc(20) = &H751C53FF: z_Sc(21) = &HC43391F: z_Sc(22) = &H7B8B2E74: z_Sc(23) = &H8478B20: z_Sc(24) = &H75184339: z_Sc(25) = &HC53FF0F: z_Sc(26) = &H13D&: z_Sc(27) = &H85197400: z_Sc(28) = &H900975C0: z_Sc(29) = &H443C7: z_Sc(30) = &H90000080: z_Sc(31) = &H47E8&: z_Sc(32) = &H2AE900: z_Sc(33) = &H90900000: z_Sc(34) = &H742C7339: z_Sc(35) = &H75894A0F: z_Sc(36) = &H46E81C: z_Sc(37) = &H75390000: z_Sc(38) = &H90157514: z_Sc(39) = &H27E8&: z_Sc(40) = &H30733900: z_Sc(41) = &HD2310A74: z_Sc(42) = &H2FE8&: z_Sc(43) = &H90909000: z_Sc(44) = &H33390BFF: z_Sc(45) = &H73390E75: z_Sc(46) = &HFF097404: z_Sc(47) = &H53FF2473: z_Sc(48) = &H90909014: z_Sc(49) = &HCC261: z_Sc(50) = &HFF2C75FF: z_Sc(51) = &H75FF2875: z_Sc(52) = &H2473FF24
        z_Sc(53) = &H891053FF: z_Sc(54) = &H90C31C45: z_Sc(55) = &H2873FF52: z_Sc(56) = &H5A1C53FF: z_Sc(57) = &H438D2275: z_Sc(58) = &H144D8D34: z_Sc(59) = &H1C458D50: z_Sc(60) = &HFF0873FF: z_Sc(61) = &H75FF2C75: z_Sc(62) = &H2475FF28: z_Sc(63) = &HFF525150: z_Sc(64) = &H53FF2073: z_Sc(65) = &H90909028: z_Sc(66) = &HC3&

        z_Sc(IDX_EBX) = z_ScMem                         'Patch the thunk data address
        z_Sc(IDX_INDEX) = HookType                      'Store the hook type in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)             'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                      'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                'Store the lParamUser callback parameter in the thunk data
        
        ' get a piece of the oCallback to use as a validation token
        RtlMoveMemory VarPtr(z_Sc(IDX_OBJCHK)), z_Sc(IDX_OWNER) + 8&, 4&

        ' validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(GetDesktopWindow) <> 0&)
        
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)    'Store the IsBadCodePtr function address in the thunk data
        z_Sc(IDX_CNH) = zFnAddr("user32", "CallNextHookEx", bUnicode)       'Store CallWindowProc function address in the thunk data
        z_Sc(IDX_UNW) = zFnAddr("user32", "UnhookWindowsHookEx", bUnicode)  'Store the SetWindowLong function address in the thunk data
        
        If bIdeSafety = True Then                                           'If the user wants IDE protection
            Debug.Assert zInIDE(bIDE)
            If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
        End If
        
        If (When And MSG_BEFORE) = MSG_BEFORE Then z_Sc(IDX_BTABLE) = 1     ' non-zero flag if Before messages desired
        If (When And MSG_AFTER) = MSG_AFTER Then z_Sc(IDX_ATABLE) = 1       ' non-zero flag if After messages desired
        
        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), MEM_LEN                     'Copy the thunk code/data to the allocated memory
        'Set the new HookProc, return the address of the original HookProc
        If bUnicode Then
            z_Sc(IDX_PREVPROC) = SetWindowsHookExW(HookType, z_ScMem + PROC_OFF, App.hInstance, nID)
        Else
            z_Sc(IDX_PREVPROC) = SetWindowsHookExA(HookType, z_ScMem + PROC_OFF, App.hInstance, nID)
        End If
        
        If z_Sc(IDX_PREVPROC) = 0 Then                                              'Ensure the new HookProc was set correctly
          zError SUB_NAME, "SetWindowsHookEx failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
        End If
        RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&    'Store the callback address
            
        shk_SetHook = True                                                          'Indicate success
      Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
      End If
      
      Exit Function
    
CatchDoubleSub:
      zError SUB_NAME, "Hook is already established"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'shk_SetHook has failed after memory allocation, so release the memory
    
End Function

'Call the next hook proc
Private Function shk_CallNextHook(ByVal HookType As eHookType, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your hook procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(HookType, HookThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        shk_CallNextHook = CallNextHookEx(zData(IDX_PREVPROC, z_ScMem), nCode, wParam, lParam) 'Call the next hook proc
    End If
End Function

Private Function shk_UnHook(ByVal HookType As eHookType) As Boolean
    ' can be made public, can be removed & zUnThunk can be called instead
    zUnThunk HookType, HookThunk
End Function

Private Sub shk_TerminateHooks()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks HookThunk
End Sub

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType) As Long
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zGet_lParamUser = zData(IDX_PARM_USER, z_ScMem)               'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType, ByVal NewValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, z_ScMem) = NewValue                      'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                                                'Get the current table entry count
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = 0 Then                                           'If the element is free...
            zData(i, nBase) = uMsg                                              'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(i, nBase) = uMsg Then                                    'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
    
        nCount = i                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
        
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
        colID = "h" & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then zError "zMap_VFunction", "Thunk type for " & vType & " does not exist"
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)

    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1            'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE, z_ScMem   'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE, z_ScMem   'Delete all after messages
        End If
        z_scFunk.Remove "h" & thunkID                   'Remove the specified thunk from the collection
        
    Case HookThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            ' if not unhooked, then unhook now
            If zData(IDX_SHUTDOWN, z_ScMem) = 0 Then UnhookWindowsHookEx zData(IDX_PREVPROC, z_ScMem)
            If zData(0, z_ScMem) = 0 Then               ' not recursing then
                VirtualFree z_ScMem, 0, MEM_RELEASE     'Release allocated memory
                z_hkFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
            Else
                zData(IDX_SHUTDOWN, z_ScMem) = 1        ' Set the shutdown indicator
                zData(IDX_ATABLE, z_ScMem) = 0          ' want no more After messages
                zData(IDX_BTABLE, z_ScMem) = 0          ' want no more Before messages
                ' when zTerminate is called this thunk's memory will be released
            End If
        Else
            z_hkFunk.Remove "h" & thunkID       'Remove the specified thunk from the collection
        End If
    Case CallbackThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            VirtualFree z_ScMem, 0, MEM_RELEASE 'Release allocated memory
        End If
        z_cbFunk.Remove "h" & ObjPtr(oCallback) & "." & thunkID           'Remove the specified thunk from the collection
    End Select

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)

    ' Terminates all thunks of a specific type
    ' Any subclassing, hooking, recurring callbacks should have already been canceled

    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case HookThunk
        Set thunkCol = z_hkFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case HookThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), HookThunk        'Unhook
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    RtlMoveMemory VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&
                    zUnThunk zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback ' release callback
                    ' remove the object pointer reference
                    RtlMoveMemory VarPtr(oCallback), VarPtr(INDX_OWNER), 4&
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If


End Sub

' this function is called from Enable and iSubClass_Before (WM_INITPOPUPMENU)
Private Function Private_AddMenuCaption(ByVal Index As Long, ByRef Caption As String) As Integer
    Dim barText() As Byte, strText() As String, intA As Integer, lngA As Long, blnInit As Boolean
    ' do not add empty strings
    If LenB(Caption) = 0 Then Private_AddMenuCaption = -1: Exit Function
    blnInit = Not ((Not Captions) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' check if any items exist
    If blnInit Then
        ' check if we know about this item already
        For intA = 0 To UBound(Captions)
            If Captions(intA).Index = Index Then Exit For
        Next intA
        ' check if adding a completely new item
        If intA > UBound(Captions) Then ReDim Preserve Captions(intA)
    Else
        ' add the first item
        ReDim Captions(0)
    End If
    ' set item data
    With Captions(intA)
        .Index = Index
        ' convert UTF-8 to Unicode
        .Caption = UTF8toStr(Caption)
    End With
    ' return array index
    Private_AddMenuCaption = intA
End Function
' this function is called from iHook_Before (WM_CREATE)
Private Function Private_AddMenuHook(ByVal hWnd As Long) As Integer
    Dim intA As Integer, blnInit As Boolean
    blnInit = Not ((Not HookSub) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' initialized?
    If blnInit Then
        ' seek for a free item
        For intA = 0 To UBound(HookSub)
            If HookSub(intA).hWnd = 0 Then Exit For
        Next intA
        ' if no free item was found, reserve memory for a new one
        If intA > UBound(HookSub) Then ReDim Preserve HookSub(intA)
    Else
        ' add the first item
        ReDim HookSub(0)
    End If
    'frmMain.Caption = Str(lngLastMenuParent <> 0) & " : " & Hex$(lngLastMenuParent) & " : " & Hex$(lngMenuhWnd)
    With HookSub(intA)
        ' initialize settings
        .hWnd = hWnd
        If ssc_Subclass(hWnd, , , , Not blnDesignTime, True) Then
            ssc_AddMsg hWnd, MSG_AFTER, WM_DESTROY, WM_NCPAINT
            Debug.Print "UniMenu (hook): Started hooking " & Hex$(hWnd)
        End If
    End With
    Private_AddMenuHook = intA
End Function
Private Sub Private_AddOwnerDraw(ByVal hWnd As Long, ByVal Index As Long)
    Dim intA As Integer, blnInit As Boolean
    blnInit = Not ((Not OwnerDraw) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' initialized?
    If blnInit Then
        ' seek for a free item
        For intA = 0 To UBound(OwnerDraw)
            If OwnerDraw(intA).hWnd = hWnd Then Exit For
        Next intA
        ' if no free item was found, reserve memory for a new one
        If intA > UBound(OwnerDraw) Then ReDim Preserve OwnerDraw(intA)
    Else
        ' add the first item
        ReDim OwnerDraw(0)
    End If
    With OwnerDraw(intA)
        .hWnd = hWnd
        blnInit = Not ((Not .Index) = -1&)
        If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
        If blnInit Then
            For intA = 0 To UBound(.Index)
                If .Index(intA) = Index Then Exit Sub
            Next intA
            If intA > UBound(.Index) Then ReDim Preserve .Index(intA)
        Else
            intA = 0
            ReDim .Index(0)
        End If
        .Index(intA) = Index
    End With
End Sub
' this sub is called from iSubClass_After (WM_DESTROY)
Private Sub Private_ClearMenuHook(ByVal hWnd As Long)
    Dim intA As Integer, blnInit As Boolean
    blnInit = Not ((Not HookSub) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' are there any?
    If Not blnInit Then Exit Sub
    ' loop through every item
    For intA = 0 To UBound(HookSub)
        With HookSub(intA)
            ' check if hWnd matches
            If .hWnd = hWnd Then
                ' remove subclassing
                .hWnd = 0
                ssc_DelMsg hWnd, MSG_AFTER, WM_DESTROY, WM_NCPAINT
                ssc_UnSubclass hWnd
                Debug.Print "UniMenu (hook): Removed " & Hex$(hWnd)
                ' done, exit
                Exit Sub
            End If
        End With
    Next intA
End Sub
' this sub is called from Class_Terminate if running compiled code
Private Sub Private_ClearMenuHooks()
    Dim intA As Integer, blnInit As Boolean
    blnInit = Not ((Not HookSub) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' are there any?
    If Not blnInit Then Exit Sub
    ' loop through every item
    For intA = 0 To UBound(HookSub)
        With HookSub(intA)
            ' check if the item isn't removed
            If .hWnd <> 0 Then
                ' unsubclass
                ssc_UnSubclass .hWnd
                Debug.Print "UniMenu (hook): Removed " & Hex$(.hWnd)
                .hWnd = 0
            End If
        End With
    Next intA
    ' clear it all
    Erase HookSub
End Sub
Private Sub Private_ClearOwnerDrawing()
    Dim intA As Integer, intB As Integer, blnInit As Boolean
    Dim udtMenuItem As MENUITEMINFO
    blnInit = Not ((Not OwnerDraw) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    If Not blnInit Then Exit Sub
    For intA = 0 To UBound(OwnerDraw)
        With OwnerDraw(intA)
            blnInit = Not ((Not .Index) = -1&)
            If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
            If blnInit Then
                For intB = 0 To UBound(.Index)
                    udtMenuItem.fMask = MIIM_TYPE
                    ' get the data
                    GetMenuItemInfo .hWnd, .Index(intB), True, udtMenuItem
                    ' remove owner-drawn flag
                    udtMenuItem.fType = udtMenuItem.fType And Not MF_OWNERDRAW
                    udtMenuItem.fMask = MIIM_TYPE
                    ' set the data
                    SetMenuItemInfo .hWnd, .Index(intB), True, udtMenuItem
                    'Debug.Print "Removed owner-draw: " & Hex$(.hWnd) & vbTab & .Index(intB)
                Next intB
            End If
        End With
    Next intA
    ' clear array
    Erase OwnerDraw
End Sub
Private Sub Private_DetectUnicode()
    Select Case m_Unicode
        Case [Detect Windows]
            Styles.InANSI = Not m_WINNT
        Case [Always Unicode]
            Styles.InANSI = False
        Case [Always ANSI]
            Styles.InANSI = True
    End Select
End Sub
Private Sub Private_Disable()
    If lngOwnerhWnd = 0 Then Exit Sub
    ' unsubclass
    ssc_UnSubclass lngOwnerhWnd
    Debug.Print "UniMenu: Ended subclassing! " & Hex$(lngOwnerhWnd)
    Private_ClearMenuHooks
    m_ShowWindow = False
    ' unhook
    If Not blnDesignTime Then
        shk_UnHook WH_CALLWNDPROC
        Debug.Print "UniMenu: Ended hooking!"
    End If
    ' clear some other stuff...
    lngOwnerhWnd = 0
    lngSyshWnd = 0
    lngMenuhWnd = 0
    shk_TerminateHooks
    ssc_Terminate
End Sub
Private Sub Private_DrawMenuBorder(ByVal hWnd As Long, ByVal Region As Long)
    Dim lngDC As Long, udtRECT As RECT, lngPen As Long, lngBorder(3) As Long
    ' get the hDC to our use
    lngDC = GetWindowDC(hWnd)
    'lngDC = GetDCEx(hWnd, Region, 3)
    GetClipBox lngDC, udtRECT
    ' set DC draw color
    lngPen = SelectObject(lngDC, CreatePen(0, 1, Colors.BorderOuter))
    ' draw outer border
    Rectangle lngDC, udtRECT.Left, udtRECT.Top, udtRECT.Right, udtRECT.Top + 1
    Rectangle lngDC, udtRECT.Left, udtRECT.Top + 1, udtRECT.Left + 1, udtRECT.Bottom
    Rectangle lngDC, udtRECT.Right - 1, udtRECT.Top + 1, udtRECT.Right, udtRECT.Bottom
    Rectangle lngDC, udtRECT.Left + 1, udtRECT.Bottom - 1, udtRECT.Right - 1, udtRECT.Bottom
    ' release DC pen
    DeleteObject SelectObject(lngDC, lngPen)
    ' shrink the border
    With udtRECT
        .Left = .Left + 1
        .Top = .Top + 1
        .Right = .Right - 1
        .Bottom = .Bottom - 1
    End With
    ' set DC draw color
    lngPen = SelectObject(lngDC, CreatePen(0, 1, Colors.BorderBack))
    ' draw "back color"
    Rectangle lngDC, udtRECT.Left, udtRECT.Top, udtRECT.Right, udtRECT.Top + 1
    Rectangle lngDC, udtRECT.Left, udtRECT.Top + 1, udtRECT.Left + 1, udtRECT.Bottom
    Rectangle lngDC, udtRECT.Right - 1, udtRECT.Top + 1, udtRECT.Right, udtRECT.Bottom
    Rectangle lngDC, udtRECT.Left + 1, udtRECT.Bottom - 1, udtRECT.Right - 1, udtRECT.Bottom
    ' release DC pen
    DeleteObject SelectObject(lngDC, lngPen)
    ' shrink the border
    With udtRECT
        .Left = .Left + 1
        .Top = .Top + 1
        .Right = .Right - 1
        .Bottom = .Bottom - 1
    End With
    ' set DC draw color
    lngPen = SelectObject(lngDC, CreatePen(0, 1, Colors.BorderInner))
    ' draw inner border
    Rectangle lngDC, udtRECT.Left, udtRECT.Top, udtRECT.Right, udtRECT.Top + 1
    Rectangle lngDC, udtRECT.Left, udtRECT.Top + 1, udtRECT.Left + 1, udtRECT.Bottom
    Rectangle lngDC, udtRECT.Right - 1, udtRECT.Top + 1, udtRECT.Right, udtRECT.Bottom
    Rectangle lngDC, udtRECT.Left + 1, udtRECT.Bottom - 1, udtRECT.Right - 1, udtRECT.Bottom
    ' release DC pen
    DeleteObject SelectObject(lngDC, lngPen)
    ' release DC
    ReleaseDC hWnd, lngDC
End Sub
Private Sub Private_DrawMenuItem(ByVal hWnd As Long, ByRef DrawItem As DRAWITEMSTRUCT, ByRef MenuItem As MENUITEMINFO)
    ' settings
    Dim IsBar As Boolean, IsBarBreak As Boolean, IsBreak As Boolean, IsRadio As Boolean, _
        IsRightJustify As Boolean, IsRightOrder As Boolean, IsSeparator As Boolean, IsSubMenu As Boolean
    Dim IsSelected As Boolean, IsGrayed As Boolean, IsDisabled As Boolean, IsChecked As Boolean, _
        IsFocus As Boolean, IsDefault As Boolean, IsHotTrack As Boolean
    ' for drawing
    Dim BackBitmap As Long, BackBuffer As RECT, BackDC As Long, BoxBuffer As RECT, OwnerBuffer As RECT
    ' general helper variables and for drawing
    Dim lngA As Long, lngBrush As Long, lngPen As Long, udtPoint As POINTAPI
    ' for text output
    Dim strText() As String, lngFont As Long, lngOFont As Long, udtCalcRECT As RECT
    ' for checkbox and image drawing
    Dim lngDC1 As Long, lngDC2 As Long, lngTempDC As Long, lngTempBitmap As Long
    Dim picBitmap1 As StdPicture, picBitmap2 As StdPicture, udtBitmap As BITMAP
    ' detect settings
    IsBar = (DrawItem.ItemData = 2)
    IsBarBreak = (MenuItem.fType And MF_MENUBARBREAK) <> 0
    IsBreak = (MenuItem.fType And MF_MENUBREAK) <> 0
    IsRadio = (MenuItem.fType And MF_RADIOCHECK) <> 0
    IsRightOrder = (MenuItem.fType And MF_RIGHTORDER) <> 0
    If Not IsBar Then
        ' menu bar can't be a separator
        IsSeparator = (MenuItem.fType And MF_SEPARATOR) <> 0
    Else
        ' as can't a regular menu item be right justified
        IsRightJustify = (MenuItem.fType And MF_RIGHTJUSTIFY) <> 0
    End If
    IsSubMenu = (MenuItem.hSubMenu <> 0)
    IsSelected = (DrawItem.ItemState And ODS_SELECTED) <> 0
    IsGrayed = (DrawItem.ItemState And ODS_GRAYED) <> 0
    IsDisabled = (DrawItem.ItemState And ODS_DISABLED) <> 0
    IsChecked = (DrawItem.ItemState And ODS_CHECKED) <> 0
    IsFocus = (DrawItem.ItemState And ODS_FOCUS) <> 0
    IsDefault = (DrawItem.ItemState And ODS_DEFAULT) <> 0
    IsHotTrack = (DrawItem.ItemState And ODS_HOTTRACK) <> 0
    ' create a backbuffer for the menuitem (prevents flickering)
    With BackBuffer
        ' determine size
        .Right = DrawItem.rcItem.Right - DrawItem.rcItem.Left
        .Bottom = DrawItem.rcItem.Bottom - DrawItem.rcItem.Top
        ' create DC
        BackDC = CreateCompatibleDC(DrawItem.hDC)
        ' create a bitmap we can place to the DC
        BackBitmap = CreateCompatibleBitmap(DrawItem.hDC, .Right, .Bottom)
        ' replace the automatically created 1 x 1 image with the new bitmap
        DeleteObject SelectObject(BackDC, BackBitmap)
    End With
    ' we leave menu bar item drawing to later as it is a shorter and less complex code
    If Not IsBar Then
        ' set color depending on state
        If Not IsDisabled Then
            If IsSelected Then
                lngBrush = CreateSolidBrush(Colors.Selected)
            Else
                lngBrush = CreateSolidBrush(Colors.Item)
            End If
        Else
            lngBrush = CreateSolidBrush(Colors.Disabled)
        End If
        ' set to correct position and paint
        SetBrushOrgEx BackDC, -DrawItem.rcItem.Left, -DrawItem.rcItem.Top, udtPoint
        lngA = SelectObject(BackDC, lngBrush)
        FillRect BackDC, BackBuffer, lngBrush
        SelectObject BackDC, lngA
        DeleteObject lngBrush
        ' check if we draw XP style panel for checkbox/image area
        If Styles.Panel Then
            With BoxBuffer
                .Left = 0
                .Top = 0
                .Right = Styles.ImageSize + 7
                .Bottom = BackBuffer.Bottom
            End With
            lngBrush = CreateSolidBrush(Colors.ImageBack)
            FillRect BackDC, BoxBuffer, lngBrush
            DeleteObject lngBrush
        End If
        ' draw a separator or a regular menu item?
        If IsSeparator Then
            If Styles.Panel Then
                With BoxBuffer
                    .Left = Styles.ImageSize + 7
                    .Top = BackBuffer.Top
                    .Bottom = BackBuffer.Bottom
                    .Right = BackBuffer.Right
                End With
                ' draw separator bg
                lngBrush = CreateSolidBrush(Colors.SeparatorBack)
                FillRect BackDC, BoxBuffer, lngBrush
                DeleteObject lngBrush
            Else
                With BoxBuffer
                    .Left = BackBuffer.Left
                    .Top = BackBuffer.Top
                    .Bottom = BackBuffer.Bottom
                    .Right = BackBuffer.Right
                End With
                ' draw separator bg
                lngBrush = CreateSolidBrush(Colors.SeparatorBack)
                FillRect BackDC, BoxBuffer, lngBrush
                DeleteObject lngBrush
            End If
            If Not Styles.In3D Then
                With BoxBuffer
                    .Left = .Left + 2
                    .Top = .Top + (.Bottom - .Top) \ 2
                    .Bottom = .Top + 1
                    .Right = .Right - 2
                End With
                ' draw simple line
                lngBrush = CreateSolidBrush(Colors.Separator)
                FillRect BackDC, BoxBuffer, lngBrush
                DeleteObject lngBrush
            Else
                With BoxBuffer
                    .Left = .Left + 2
                    .Top = .Top + (.Bottom - .Top) \ 2
                    .Bottom = .Top + 2
                    .Right = .Right - 2
                End With
                ' draw 3D line
                DrawEdge BackDC, BoxBuffer, BDR_SUNKENOUTER, BF_RECT
            End If
        Else
            ' draw item selected except if it is disabled
            If IsSelected And Not IsDisabled Then
                lngBrush = SelectObject(BackDC, CreateSolidBrush(Colors.Selected))
                ' draw border?
                If Styles.SelectedBorder Then
                    lngPen = SelectObject(BackDC, CreatePen(0, 1, Colors.SelectedBorder))
                Else
                    lngPen = SelectObject(BackDC, CreatePen(0, 1, Colors.Selected))
                End If
                Rectangle BackDC, BackBuffer.Left, BackBuffer.Top, BackBuffer.Right, BackBuffer.Bottom
                DeleteObject SelectObject(BackDC, lngPen)
                DeleteObject SelectObject(BackDC, lngBrush)
                DeleteObject lngBrush
            ElseIf Styles.ItemBorder Then
                lngBrush = SelectObject(BackDC, CreateSolidBrush(Colors.ItemBorder))
                With BoxBuffer
                    .Left = BackBuffer.Left
                    .Top = BackBuffer.Top
                    .Right = BackBuffer.Right
                    .Bottom = BackBuffer.Top + 1
                End With
                FillRect BackDC, BoxBuffer, lngBrush
                With BoxBuffer
                    .Left = BackBuffer.Left
                    .Top = BackBuffer.Top + 1
                    .Right = BackBuffer.Left + 1
                    .Bottom = BackBuffer.Bottom
                End With
                FillRect BackDC, BoxBuffer, lngBrush
                With BoxBuffer
                    .Left = BackBuffer.Right - 1
                    .Top = BackBuffer.Top + 1
                    .Right = BackBuffer.Right
                    .Bottom = BackBuffer.Bottom
                End With
                FillRect BackDC, BoxBuffer, lngBrush
                With BoxBuffer
                    .Left = BackBuffer.Left + 1
                    .Top = BackBuffer.Bottom - 1
                    .Right = BackBuffer.Right - 1
                    .Bottom = BackBuffer.Bottom
                End With
                FillRect BackDC, BoxBuffer, lngBrush
                DeleteObject lngBrush
            End If
            ' if we draw it checked
            If IsChecked Then
                With BoxBuffer
                    .Left = BackBuffer.Left + (Styles.ImageSize - 6) \ 2 - 1
                    .Top = BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - Styles.ImageSize + 4) \ 2 - 1
                    .Right = .Left + 15
                    .Bottom = .Top + 15
                End With
                If Styles.CheckBorder Then
                    SetTextColor BackDC, Colors.Check
                    SetBkColor BackDC, Colors.CheckBack
                    If Not Styles.In3D Then
                        lngBrush = CreateSolidBrush(Colors.CheckBorder)
                        FillRect BackDC, BoxBuffer, lngBrush
                        DeleteObject lngBrush
                    Else
                        DrawEdge BackDC, BoxBuffer, BDR_SUNKENOUTER, BF_RECT
                    End If
                ElseIf IsSelected Then
                    SetTextColor BackDC, Colors.Check
                    SetBkColor BackDC, Colors.Selected
                Else
                    SetTextColor BackDC, Colors.Check
                    SetBkColor BackDC, Colors.Item
                End If
                With BoxBuffer
                    .Left = .Left + 1
                    .Top = .Top + 1
                    .Right = .Right - 1
                    .Bottom = .Bottom - 1
                End With
                lngTempBitmap = LoadBitmap(0&, ByVal OBM_CHECK)
                lngTempDC = CreateCompatibleDC(BackDC)
                DeleteObject SelectObject(lngTempDC, lngTempBitmap)
                BitBlt BackDC, BoxBuffer.Left, BoxBuffer.Top, 13, 13, lngTempDC, 0, 0, vbSrcCopy
                DeleteDC lngTempDC
                DeleteObject lngTempBitmap
                'Set picBitmap1 = LoadResPicture("checked", vbResBitmap)
                'Set picBitmap2 = LoadResPicture("checked", vbResBitmap)
                'GetObject picBitmap1, Len(udtBitmap), udtBitmap
                'lngTempDC = CreateCompatibleDC(DrawItem.hDC)
                'lngTempBitmap = CreateCompatibleBitmap(DrawItem.hDC, udtBitmap.bmWidth, udtBitmap.bmHeight)
                'DeleteObject SelectObject(lngTempDC, lngTempBitmap)
                'With BoxBuffer
                '    .Left = 0
                '    .Top = 0
                '    .Right = udtBitmap.bmWidth
                '    .Bottom = udtBitmap.bmHeight
                'End With
                'If IsSelected Then
                '    lngBrush = CreateSolidBrush(Colors.Selected)
                'Else
                '    lngBrush = CreateSolidBrush(Colors.CheckBack)
                'End If
                'FillRect lngTempDC, BoxBuffer, lngBrush
                'DeleteObject lngBrush
                'lngDC1 = CreateCompatibleDC(DrawItem.hDC)
                'lngDC2 = CreateCompatibleDC(DrawItem.hDC)
                'If Not IsDisabled Then
                '    DeleteObject SelectObject(lngDC1, picBitmap1.Handle)
                '    DeleteObject SelectObject(lngDC2, picBitmap2.Handle)
                'End If
                'BitBlt lngDC1, 0, 0, udtBitmap.bmWidth, udtBitmap.bmHeight, lngTempDC, 0, 0, vbSrcPaint
                'DeleteDC lngTempDC
                'DeleteObject lngTempBitmap
            End If
            ' draw raised edge for image
            If Not picBitmap1 Is Nothing Then
                If picBitmap1.Handle <> 0 Then
                    If IsSelected And Not (IsChecked Or IsDisabled) Then
                        With BoxBuffer
                            .Top = BackBuffer.Top
                            .Left = BackBuffer.Left
                            .Right = .Left + Styles.ImageSize + 2
                            .Bottom = BackBuffer.Bottom
                        End With
                        DrawEdge BackDC, BoxBuffer, BDR_RAISEDINNER, BF_RECT
                    End If
                End If
            End If
            ' text?
            If LenB(Private_GetMenuCaption(MenuItem.wid)) > 0 Then
                ' separate texts by tab
                strText = Split(Private_GetMenuCaption(MenuItem.wid), vbTab)
            Else
                ' empty string
                ReDim strText(0)
            End If
            ' check length
            If LenB(strText(0)) > 0 Then
                ' set font
                lngFont = Private_GetFont(BackDC, IsDefault)
                lngOFont = SelectObject(BackDC, lngFont)
                ' set text color
                If Not IsDisabled Then
                    If IsSelected Then
                        SetTextColor BackDC, Colors.SelectedText
                        SetBkColor BackDC, Colors.Selected
                    Else
                        SetTextColor BackDC, Colors.ItemText
                        SetBkColor BackDC, Colors.Item
                    End If
                Else
                    SetTextColor BackDC, Colors.DisabledText
                    SetBkColor BackDC, Colors.Disabled
                End If
                ' set transparent background
                SetBkMode BackDC, 0&
                ' calculate the size of the text
                Private_DrawTextAuto BackDC, StrPtr(strText(0)), LenB(strText(0)), udtCalcRECT, DT_CALCRECT
                ' set text area
                With BoxBuffer
                    If Styles.Panel Then
                        .Left = Styles.ImageSize + 14
                    Else
                        .Left = Styles.ImageSize + 7
                    End If
                    .Top = BackBuffer.Top + ((BackBuffer.Bottom - BackBuffer.Top - udtCalcRECT.Bottom) \ 2)
                    .Right = BackBuffer.Right
                    .Bottom = BackBuffer.Bottom
                End With
                ' draw the text
                If Not (IsDisabled And Styles.In3D) Then
                    Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(0)), LenB(strText(0)), BoxBuffer.Left, BoxBuffer.Top, 0, 0, DST_PREFIXTEXT Or DSS_NORMAL
                    If UBound(strText) > 0 Then
                        If LenB(strText(1)) > 0 Then
                            Private_DrawTextAuto BackDC, StrPtr(strText(1)), LenB(strText(1)), udtCalcRECT, DT_CALCRECT
                            Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(1)), LenB(strText(1)), BoxBuffer.Right - 10 - udtCalcRECT.Right, BoxBuffer.Top, 0, 0, DST_PREFIXTEXT Or DSS_NORMAL
                        End If
                    End If
                Else
                    Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(0)), LenB(strText(0)), BoxBuffer.Left, BoxBuffer.Top, 0, 0, DST_PREFIXTEXT Or DSS_DISABLED
                    If UBound(strText) > 0 Then
                        If LenB(strText(1)) > 0 Then
                            Private_DrawTextAuto BackDC, StrPtr(strText(1)), LenB(strText(1)), udtCalcRECT, DT_CALCRECT
                            Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(1)), LenB(strText(1)), BoxBuffer.Right - 10 - udtCalcRECT.Right, BoxBuffer.Top, 0, 0, DST_PREFIXTEXT Or DSS_DISABLED
                        End If
                    End If
                End If
                ' draw arrow
                If IsSubMenu Then
                    ' determine colors
                    If IsSelected Then
                        SetTextColor BackDC, Colors.ArrowSelected
                        SetBkColor BackDC, Colors.Selected
                    Else
                        SetTextColor BackDC, Colors.Arrow
                        SetBkColor BackDC, Colors.Item
                    End If
                    ' load menu arrow bitmap, create DC and set bitmap to the DC
                    lngTempBitmap = LoadBitmap(0&, ByVal OBM_MNARROW)
                    lngTempDC = CreateCompatibleDC(BackDC)
                    ' and delete the 1 x 1 bitmap that is created automatically...
                    DeleteObject SelectObject(lngTempDC, lngTempBitmap)
                    ' draw the arrow from temporary bitmap buffer to the backbuffer
                    BitBlt BackDC, BackBuffer.Right - 15, (BackBuffer.Bottom - 13) \ 2, 13, 13, lngTempDC, 0, 0, vbSrcCopy
                    ' and clear the temporary bitmap from memory
                    DeleteDC lngTempDC
                    DeleteObject lngTempBitmap
                End If
                If Not (picBitmap1 Is Nothing) Then
                    If picBitmap1.Handle <> 0 Then
                        If Not IsDisabled Then
                            BitBlt BackDC, (7 + Styles.ImageSize - udtBitmap.bmWidth) \ 2, BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtBitmap.bmHeight) \ 2, udtBitmap.bmWidth, udtBitmap.bmHeight, lngDC1, 0, 0, vbMergePaint
                            BitBlt BackDC, (7 + Styles.ImageSize - udtBitmap.bmWidth) \ 2, BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtBitmap.bmHeight) \ 2, udtBitmap.bmWidth, udtBitmap.bmHeight, lngDC2, 0, 0, vbSrcAnd
                        Else
                            Private_DrawStateAuto BackDC, 0, 0, picBitmap1.Handle, 0, (7 + Styles.ImageSize - udtBitmap.bmWidth) \ 2, BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtBitmap.bmHeight) \ 2, udtBitmap.bmWidth, udtBitmap.bmHeight, DST_BITMAP Or DSS_DISABLED
                        End If
                        ' clean memory
                        DeleteDC lngDC1
                        DeleteObject picBitmap1.Handle
                        DeleteDC lngDC2
                        DeleteObject picBitmap2.Handle
                    End If
                End If
            End If
        End If
    Else
        '' see if we draw the rest of the menu bar
        'If IsBarBreak Or IsBreak Or DrawItem.itemID = lngMenuTopEndID Then
        '    ' set color
        '    lngBrush = CreateSolidBrush(Colors.Bar)
        '    ' get form rectangle
        '    GetWindowRect hWnd, OwnerBuffer
        '    ' set draw area
        '    With BoxBuffer
        '        .Left = DrawItem.rcItem.Right
        '        .Top = DrawItem.rcItem.Top
        '        .Right = OwnerBuffer.Right - OwnerBuffer.Left - (GetSystemMetrics(SM_CXBORDER) * 2) - 2
        '        .Bottom = DrawItem.rcItem.Bottom
        '    End With
        '    ' draw the remaining menu bar
        '    FillRect DrawItem.hDC, BoxBuffer, lngBrush
        '    ' delete brush
        '    DeleteObject lngBrush
        'End If
        ' set color depending on state
        If Not IsDisabled Then
            If IsHotTrack Then
                lngBrush = CreateSolidBrush(Colors.BarHot)
            ElseIf IsSelected Then
                lngBrush = CreateSolidBrush(Colors.BarSelected)
            Else
                lngBrush = CreateSolidBrush(Colors.Bar)
            End If
        Else
            lngBrush = CreateSolidBrush(Colors.Bar)
        End If
        ' set to correct position and paint
        SetBrushOrgEx BackDC, -DrawItem.rcItem.Left, -DrawItem.rcItem.Top, udtPoint
        lngA = SelectObject(BackDC, lngBrush)
        FillRect BackDC, BackBuffer, lngBrush
        SelectObject BackDC, lngA
        DeleteObject lngBrush
        ' text?
        If LenB(Private_GetMenuCaption(MenuItem.wid)) > 0 Then
            ' separate texts by tab
            strText = Split(Private_GetMenuCaption(MenuItem.wid), vbTab)
        Else
            ' empty string
            ReDim strText(0)
        End If
        ' set font
        lngFont = Private_GetFont(BackDC, IsDefault)
        lngOFont = SelectObject(BackDC, lngFont)
        If Not IsDisabled Then
            If IsHotTrack Then
                SetTextColor BackDC, Colors.BarHotText
                If Styles.BarHotBorder Then
                    If Not Styles.In3D Then
                        lngBrush = SelectObject(BackDC, CreateSolidBrush(Colors.BarHot))
                        lngPen = SelectObject(BackDC, CreatePen(0, 1, Colors.BarHotBorder))
                        Rectangle BackDC, BackBuffer.Left, BackBuffer.Top, BackBuffer.Right, BackBuffer.Bottom
                        DeleteObject SelectObject(BackDC, lngPen)
                        DeleteObject SelectObject(BackDC, lngBrush)
                        DeleteObject lngBrush
                    Else
                        DrawEdge BackDC, BackBuffer, BDR_RAISEDINNER, BF_RECT
                    End If
                End If
            ElseIf IsSelected Then
                SetTextColor BackDC, Colors.BarSelectedText
                If Styles.BarSelectedBorder Then
                    If Not Styles.In3D Then
                        lngBrush = SelectObject(BackDC, CreateSolidBrush(Colors.BarSelected))
                        lngPen = SelectObject(BackDC, CreatePen(0, 1, Colors.BarSelectedBorder))
                        If IsSubMenu Then
                            Rectangle BackDC, BackBuffer.Left, BackBuffer.Top, BackBuffer.Right, BackBuffer.Bottom + 1
                        Else
                            Rectangle BackDC, BackBuffer.Left, BackBuffer.Top, BackBuffer.Right, BackBuffer.Bottom
                        End If
                        DeleteObject SelectObject(BackDC, lngPen)
                        DeleteObject SelectObject(BackDC, lngBrush)
                        DeleteObject lngBrush
                    Else
                        DrawEdge BackDC, BackBuffer, BDR_SUNKENOUTER, BF_RECT
                    End If
                End If
            Else
                SetTextColor BackDC, Colors.barText
                If Styles.BarBorder Then
                    lngBrush = SelectObject(BackDC, CreateSolidBrush(Colors.Bar))
                    lngPen = SelectObject(BackDC, CreatePen(0, 1, Colors.BarBorder))
                    Rectangle BackDC, BackBuffer.Left, BackBuffer.Top, BackBuffer.Right, BackBuffer.Bottom
                    DeleteObject SelectObject(BackDC, lngPen)
                    DeleteObject SelectObject(BackDC, lngBrush)
                    DeleteObject lngBrush
                End If
            End If
        Else
            SetTextColor BackDC, Colors.DisabledText
        End If
        ' transparent text
        SetBkMode BackDC, 0&
        ' text size
        Private_DrawTextAuto BackDC, StrPtr(strText(0)), LenB(strText(0)), udtCalcRECT, DT_CALCRECT
        ' draw text
        If Not (Styles.In3D And IsSelected And Styles.BarSelectedBorder) Then
            If Not (IsDisabled And Styles.In3D) Then
                Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(0)), LenB(strText(0)), BackBuffer.Left + (BackBuffer.Right - BackBuffer.Left - udtCalcRECT.Right) \ 2, BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtCalcRECT.Bottom) \ 2, 0, 0, DST_PREFIXTEXT Or DSS_NORMAL
            Else
                Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(0)), LenB(strText(0)), BackBuffer.Left + (BackBuffer.Right - BackBuffer.Left - udtCalcRECT.Right) \ 2, BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtCalcRECT.Bottom) \ 2, 0, 0, DST_PREFIXTEXT Or DSS_DISABLED
            End If
        Else
            If Not (IsDisabled And Styles.In3D And Styles.BarSelectedBorder) Then
                Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(0)), LenB(strText(0)), BackBuffer.Left + (BackBuffer.Right - BackBuffer.Left - udtCalcRECT.Right) \ 2 + 1, 1 + BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtCalcRECT.Bottom) \ 2, 0, 0, DST_PREFIXTEXT Or DSS_NORMAL
            Else
                Private_DrawStateAuto BackDC, 0, 0, StrPtr(strText(0)), LenB(strText(0)), BackBuffer.Left + (BackBuffer.Right - BackBuffer.Left - udtCalcRECT.Right) \ 2 + 1, 1 + BackBuffer.Top + (BackBuffer.Bottom - BackBuffer.Top - udtCalcRECT.Bottom) \ 2, 0, 0, DST_PREFIXTEXT Or DSS_DISABLED
            End If
        End If
    End If
    If lngOFont <> 0 Then
        SelectObject BackDC, lngOFont
        DeleteObject lngFont
    End If
    ' draw from the backbuffer on the menu item
    BitBlt DrawItem.hDC, DrawItem.rcItem.Left, DrawItem.rcItem.Top, BackBuffer.Right, BackBuffer.Bottom, BackDC, 0, 0, vbSrcCopy
    ' prevent Windows from adding the arrow...
    If Not IsBar Then ExcludeClipRect DrawItem.hDC, DrawItem.rcItem.Left, DrawItem.rcItem.Top, BackBuffer.Right, BackBuffer.Bottom
    ' clear backbuffer
    DeleteDC BackDC
    DeleteObject BackBitmap
    ' delete string array
    Erase strText
End Sub
' wanted to keep code clear and short instead of concentrating to speed... the result: DrawStateAuto
Private Sub Private_DrawStateAuto(ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal Flags As Long)
    Dim barTemp() As Byte, lngA As Long
    ' draw in Unicode?
    If Not Styles.InANSI Then
        ' wParam is used for text length in characters... but it can be zero if used for bitmaps
        If wParam > 0 Then wParam = wParam \ 2
        ' are we under Windows NT/2k/XP? (native Unicode support)
        If m_WINNT Then
            DrawState hDC, hBrush, lpDrawStateProc, lParam, wParam, X, Y, cX, cY, Flags
        Else ' use unicows.dll
            DrawStateUnicows hDC, hBrush, lpDrawStateProc, lParam, wParam, X, Y, cX, cY, Flags
        End If
    Else ' draw in ANSI
        If wParam > 0 Then
            ReDim barTemp(wParam * 2 - 1)
            lngA = WideCharToMultiByte(ByVal GetACP, ByVal 0&, ByVal lParam, ByVal wParam, ByVal VarPtr(barTemp(0)), ByVal UBound(barTemp) + 1, ByVal 0&, ByVal 0&)
            If lngA > 0 Then
                ReDim Preserve barTemp(lngA - 1)
                DrawStateANSI hDC, hBrush, lpDrawStateProc, VarPtr(barTemp(0)), lngA \ 2, X, Y, cX, cY, Flags
            End If
            Erase barTemp
        Else
            DrawStateANSI hDC, hBrush, lpDrawStateProc, lParam, 0&, X, Y, cX, cY, Flags
        End If
    End If
End Sub
' wanted to keep code clear and short instead of concentrating to speed... the result: DrawTextAuto
Private Sub Private_DrawTextAuto(ByVal hDC As Long, ByVal lpStrPtr As Long, ByVal nCount As Long, ByRef lpRECT As RECT, ByVal wFormat As Long)
    Dim barTemp() As Byte, lngA As Long
    If nCount = 0 Then Exit Sub
    ' draw in Unicode?
    If Not Styles.InANSI Then
        ' correct text length (given in characters, comes in bytes...)
        nCount = nCount \ 2
        ' are we under Windows NT/2k/XP? (native Unicode support)
        If m_WINNT Then
            DrawText hDC, lpStrPtr, nCount, lpRECT, wFormat
        Else ' use unicows.dll
            DrawTextUnicows hDC, lpStrPtr, nCount, lpRECT, wFormat
        End If
    Else ' draw in ANSI
        ReDim barTemp(nCount * 2 - 1)
        lngA = WideCharToMultiByte(ByVal GetACP, ByVal 0&, ByVal lpStrPtr, ByVal nCount, ByVal VarPtr(barTemp(0)), ByVal nCount * 2, ByVal 0&, ByVal 0&)
        If lngA > 0 Then
            ReDim Preserve barTemp(lngA - 1)
            DrawTextANSI hDC, VarPtr(barTemp(0)), lngA \ 2, lpRECT, wFormat
        End If
        Erase barTemp
    End If
End Sub
Private Sub Private_Enable()
    Dim lngA As Long, udtMenuItem As MENUITEMINFO
    Dim lngTopCount As Long, blnInit As Boolean
    ' do not subclass twice
    If lngOwnerhWnd <> 0 Then Exit Sub
    RaiseEvent StyleChange
    ' check if is a form
    If TypeOf UserControl.Parent Is Form Then
        ' get hWnd and system menu hWnd
        lngOwnerhWnd = UserControl.Parent.hWnd
        lngSyshWnd = GetSystemMenu(lngOwnerhWnd, 0&)
        ' get top menu hWnd
        lngMenuhWnd = GetMenu(lngOwnerhWnd)
        ' get number of items
        lngTopCount = GetMenuItemCount(lngMenuhWnd) - 1
        ' set ending ID
        lngMenuTopEndID = MENU_TOP_ID + lngTopCount
        ' check if to skip the first item
        lngA = CLng(Abs(UserControl.Parent.WindowState = vbMaximized))
        ' loop through all top level menu items
        For lngA = lngA To lngTopCount
            With udtMenuItem
                ' nullify type to get string length
                .fMask = MIIM_STRING
                .dwTypeData = vbNullString
                .cch = 0
                .cbSize = Len(udtMenuItem)
                ' get string length info
                GetMenuItemInfo lngMenuhWnd, lngA, True, udtMenuItem
                ' initialize information to get the required data
                .fMask = MIIM_TYPE ' MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or
                .dwTypeData = String$(.cch, vbNullChar)
                .cch = .cch + 1
                .cbSize = Len(udtMenuItem)
                ' get the data
                GetMenuItemInfo lngMenuhWnd, lngA, True, udtMenuItem
                ' save the menu string
                Private_AddMenuCaption MENU_TOP_ID + lngA, .dwTypeData
                ' set menu ownerdrawn
                ModifyMenu lngMenuhWnd, lngA, .fType Or MF_OWNERDRAW Or MF_BYPOSITION, MENU_TOP_ID + lngA, ByVal 2&
                .fMask = MIIM_ID
                .wid = MENU_TOP_ID + lngA
                SetMenuItemInfo lngMenuhWnd, lngA, True, udtMenuItem
            End With
        Next lngA
        blnInit = Not ((Not Captions) = -1&)
        If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
        If Not blnInit Then
            intMenuTopEndID = -1
        Else
            ' the upper bound of top menu items
            intMenuTopEndID = UBound(Captions)
        End If
    Else
        ' no form, no top menu
        lngMenuhWnd = 0
        lngMenuTopEndID = 0
    End If
    If lngOwnerhWnd <> 0 Then
        If Not blnDesignTime Then shk_SetHook WH_CALLWNDPROC, , MSG_AFTER, , 2, , Ambient.UserMode
        If Not m_ShowWindow Then m_ShowWindow = ssc_Subclass(lngOwnerhWnd, , , , True)
        If m_ShowWindow Then
            ssc_AddMsg lngOwnerhWnd, MSG_BEFORE, WM_DRAWITEM, WM_ERASEBKGND, WM_INITMENUPOPUP, WM_MEASUREITEM, WM_MENUSELECT
            Debug.Print "UniMenu: Started subclassing! " & Hex$(lngOwnerhWnd)
        Else
            Debug.Print "UniMenu: Subclassing failed!"
        End If
        DrawMenuBar lngOwnerhWnd
    End If
End Sub
' this function is called from iSubClass_Before (WM_DRAWITEM and WM_MEASUREITEM)
Private Function Private_GetFont(ByVal hDC As Long, Optional ByVal ForceBold As Boolean) As Long
    Dim uLF As LOGFONT, lngLen As Long
    ' initialize font settings
    With m_Font
        ' determine length of font name
        If Len(.Name) >= LF_FACESIZE Then lngLen = LF_FACESIZE Else lngLen = Len(.Name)
        ' copy maximum allowed length
        CopyMemory uLF.lfFaceName(0), ByVal .Name, lngLen
        ' set other font settings
        uLF.lfHeight = -MulDiv(.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)
        uLF.lfItalic = .Italic
        If Not .Bold And Not ForceBold Then uLF.lfWeight = FW_NORMAL Else uLF.lfWeight = FW_BOLD
        uLF.lfUnderline = .Underline
        uLF.lfStrikeOut = .Strikethrough
        uLF.lfCharSet = .Charset
    End With
    Private_GetFont = CreateFontIndirect(uLF)
'    Dim udtMetrics As NONCLIENTMETRICS
'    Dim SystemMenuFont As String, SystemMenuSize As Long
'    Dim lngCaps As Long, sngPixelConv As Single, Divider As Single
'    lngCaps = GetDeviceCaps(hDC, LOGPIXELSY)
'    ' sometimes I can't figure out why IDE makes stupid error messages...
'    ' yes, this is all because IDE raises an error on sngPixelConv = CSng(lngCaps / 80) !!!
'    If Not blnInIDE Then
'        If m_Font Is Nothing Then Divider = 80 Else Divider = 69.5
'        sngPixelConv = CSng(lngCaps / Divider)
'    Else
'        On Error Resume Next
'TryAgain:
'        If m_Font Is Nothing Then Divider = 80 Else Divider = 69.5
'        sngPixelConv = CSng(lngCaps / Divider)
'        If Err.Number <> 0 Then Err.Clear: GoTo TryAgain
'        On Error GoTo 0
'    End If
'    If Not m_Font Is Nothing Then
'        Dim udtLogFont  As LOGFONT, strFont As String * 32
'        With udtLogFont
'            strFont = m_Font.Name
'            CopyMemory .lfFaceName(0), ByVal StrPtr(StrConv(strFont, vbFromUnicode)), Len(strFont)
'            .lfCharSet = m_Font.Charset
'            .lfHeight = m_Font.Size * -sngPixelConv
'            .lfItalic = m_Font.Italic
'            .lfStrikeOut = m_Font.Strikethrough
'            .lfUnderline = m_Font.Underline
'            If ForceBold Then
'                .lfWeight = m_Font.Weight * 2
'            Else
'                .lfWeight = m_Font.Weight
'            End If
'        End With
'        Private_GetFont = CreateFontIndirect(udtLogFont)
'        Exit Function
'    End If
'    udtMetrics.cbSize = Len(udtMetrics)
'    SystemParametersInfo SPI_GETNONCLIENTMETRICS, udtMetrics.cbSize, udtMetrics, 0
'    With udtMetrics.lfMenuFont
'        If m_Font Is Nothing Then
'            SystemMenuFont = StrConv(Left$(.lfFaceName, InStr(.lfFaceName, vbNullChar) - 1), vbUnicode)
'            Private_GetFont = CreateFont(-sngPixelConv * .lfHeight, .lfWidth, .lfEscapement, .lfOrientation, .lfWeight, .lfItalic, .lfUnderline, .lfStrikeOut, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, StrPtr(SystemMenuFont))
'        Else
'            Private_GetFont = CreateFont(-sngPixelConv * m_Font.Size, 0, .lfEscapement, .lfOrientation, m_Font.Weight, m_Font.Italic, m_Font.Underline, m_Font.Strikethrough, m_Font.Charset, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, StrPtr(m_Font.Name))
'        End If
'    End With
End Function
' this function is called from iSubClass_Before (WM_DRAWITEM and WM_MEASUREITEM)
Private Function Private_GetMenuCaption(ByVal Index As Long) As String
    Dim intA As Integer, blnInit As Boolean
    blnInit = Not ((Not Captions) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' if we have none, we can't return anything
    If Not blnInit Then Exit Function
    For intA = 0 To UBound(Captions)
        ' check if the index matches and return the string if so
        With Captions(intA)
            If .Index = Index Then Private_GetMenuCaption = .Caption: Exit Function
        End With
    Next intA
End Function
' this function is called from iSubClass_Before (WM_INITPOPUPMENU)
Private Function Private_GetMenuHookCount() As Integer
    Dim intA As Integer, intCount As Integer, blnInit As Boolean
    blnInit = Not ((Not HookSub) = -1&)
    If Not blnInIDE Then Else On Error Resume Next: Debug.Assert True Xor CLng(0#): On Error GoTo 0
    ' none, return zero
    If Not blnInit Then Exit Function
    ' loop through all items
    For intA = 0 To UBound(HookSub)
        ' if subclassed, then increase the count
        If HookSub(intA).hWnd <> 0 Then intCount = intCount + 1
    Next intA
    ' return the result
    Private_GetMenuHookCount = intCount
End Function
Private Function Private_GetScaleMode() As ScaleModeConstants
    Select Case Ambient.ScaleUnits
        Case "Twip"
            Private_GetScaleMode = vbTwips
        Case "Point"
            Private_GetScaleMode = vbPoints
        Case "Pixel"
            Private_GetScaleMode = vbPixels
        Case "Character"
            Private_GetScaleMode = vbCharacters
        Case "Inch"
            Private_GetScaleMode = vbInches
        Case "Millimeter"
            Private_GetScaleMode = vbMillimeters
        Case "Centimeter"
            Private_GetScaleMode = vbCentimeters
        Case "User"
            ' prevent user scalemode
            Parent.ScaleMode = vbTwips
            Private_GetScaleMode = vbTwips
    End Select
End Function
' this function is called from SetColorStyle
Private Function Private_GetSmoothColor(ByVal Color As Long, Optional ByVal Count As Byte = 0) As Long
    Dim lngCount As Long, Red As Long, Green As Long, Blue As Long
    ' check if a system color
    If Color < 0 Then Color = GetSysColor(Color And Not &H80000000)
    If Count = 0 Then
        ' just return the color as-is
        Private_GetSmoothColor = Color
    Else
        ' get RGB
        Red = Color And &HFF&
        Green = (Color And &HFF00&) \ &H100&
        Blue = (Color And &HFF0000) \ &H10000
        ' do smooth Count times
        Do Until lngCount = Count
            ' return smooth color value from smooth colors array
            Red = SmoothColors(Red)
            Green = SmoothColors(Green)
            Blue = SmoothColors(Blue)
            ' increase counter
            lngCount = lngCount + 1
        Loop
        ' and now return the result
        Private_GetSmoothColor = (Blue * &H10000) Or (Green * &H100&) Or Red
    End If
End Function
' this function is called by iHook_Before
Private Function Private_IsMenuClass(ByVal hWnd As Long) As Boolean
    Dim strClassName As String, lngClassLen As Long
    ' init string buffer
    strClassName = String$(128, vbNullChar)
    ' get class name
    lngClassLen = GetClassName(hWnd, strClassName, 128)
    ' the length doesn't match, return False
    If lngClassLen <> 6 Then Exit Function
    ' compare the class name to menu class name (always #32768)
    Private_IsMenuClass = (Left$(strClassName, lngClassLen) = "#32768")
End Function
Private Sub Private_SetColors()
    With Colors
        .Arrow = Private_GetSmoothColor(m_ColorArrow, m_SmoothArrow)
        .ArrowSelected = Private_GetSmoothColor(m_ColorArrowSelected, m_SmoothArrowSelected)
        .Bar = Private_GetSmoothColor(m_ColorBar, m_SmoothBar)
        .BarBorder = Private_GetSmoothColor(m_ColorBarBorder, m_SmoothBarBorder)
        .BarHot = Private_GetSmoothColor(m_ColorBarHot, m_SmoothBarHot)
        .BarHotBorder = Private_GetSmoothColor(m_ColorBarHotBorder, m_SmoothBarHotBorder)
        .BarHotText = Private_GetSmoothColor(m_ColorBarHotText, m_SmoothBarHotText)
        .BarSelected = Private_GetSmoothColor(m_ColorBarSelected, m_SmoothBarSelected)
        .BarSelectedBorder = Private_GetSmoothColor(m_ColorBarSelectedBorder, m_SmoothBarSelectedBorder)
        .BarSelectedText = Private_GetSmoothColor(m_ColorBarSelectedText, m_SmoothBarSelectedText)
        .barText = Private_GetSmoothColor(m_ColorBarText, m_SmoothBarText)
        .BorderBack = Private_GetSmoothColor(m_ColorBorderBack, m_SmoothBorderBack)
        .BorderInner = Private_GetSmoothColor(m_ColorBorderInner, m_SmoothBorderInner)
        .BorderOuter = Private_GetSmoothColor(m_ColorBorderOuter, m_SmoothBorderOuter)
        .Check = Private_GetSmoothColor(m_ColorCheck, m_SmoothCheck)
        .CheckBack = Private_GetSmoothColor(m_ColorCheckBack, m_SmoothCheckBack)
        .CheckBorder = Private_GetSmoothColor(m_ColorCheckBorder, m_SmoothCheckBorder)
        .Disabled = Private_GetSmoothColor(m_ColorDisabled, m_SmoothDisabled)
        .DisabledText = Private_GetSmoothColor(m_ColorDisabledText, m_SmoothDisabledText)
        .ImageBack = Private_GetSmoothColor(m_ColorImageBack, m_SmoothImageBack)
        .ImageBorder = Private_GetSmoothColor(m_ColorImageBorder, m_SmoothImageBorder)
        .ImageShadow = Private_GetSmoothColor(m_ColorImageShadow, m_SmoothImageShadow)
        .Item = Private_GetSmoothColor(m_ColorItem, m_SmoothItem)
        .ItemBorder = Private_GetSmoothColor(m_ColorItemBorder, m_SmoothItemBorder)
        .ItemText = Private_GetSmoothColor(m_ColorItemText, m_SmoothItemText)
        .Selected = Private_GetSmoothColor(m_ColorSelected, m_SmoothSelected)
        .SelectedBorder = Private_GetSmoothColor(m_ColorSelectedBorder, m_SmoothSelectedBorder)
        .SelectedText = Private_GetSmoothColor(m_ColorSelectedText, m_SmoothSelectedText)
        .Separator = Private_GetSmoothColor(m_ColorSeparator, m_SmoothSeparator)
        .SeparatorBack = Private_GetSmoothColor(m_ColorSeparatorBack, m_SmoothSeparatorBack)
    End With
End Sub
Private Property Get Private_UTF16toUTF8(ByRef Text As String, Optional lFlags As Long) As String
    Static tmpArr() As Byte
    Dim tmpLen As Long, textLen As Long
    If LenB(Text) <> 0 Then
        textLen = Len(Text)
        tmpLen = LenB(Text) * 2 + 1
        ReDim Preserve tmpArr(tmpLen - 1)
        tmpLen = WideCharToMultiByte(65001, lFlags, ByVal StrPtr(Text), textLen, ByVal VarPtr(tmpArr(0)), tmpLen, ByVal 0&, ByVal 0&)
        If tmpLen > 0 Then
            If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen - 1)
            Private_UTF16toUTF8 = CStr(tmpArr)
        End If
    End If
End Property
Private Property Get Private_UTF8toUTF16(ByRef Text As String, Optional lFlags As Long) As String
    Static tmpArr() As Byte
    Dim tmpLen As Long, textLen As Long
    If LenB(Text) <> 0 Then
        textLen = LenB(Text)
        tmpLen = textLen * 2
        ReDim Preserve tmpArr(tmpLen + 1)
        tmpLen = MultiByteToWideChar(65001, lFlags, ByVal StrPtr(Text), textLen, ByVal VarPtr(tmpArr(0)), tmpLen) * 2
        If tmpLen > 0 Then
            If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen - 1)
            Private_UTF8toUTF16 = CStr(tmpArr)
        End If
    End If
End Property
Private Sub Private_WndHookProcKeyboard(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long): End Sub
Private Sub Private_WndHookProcMenu(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long)
    Static m_MenuHook As Long
    Dim CWP As CWPSTRUCT
    ' check the code
    If nCode <> HC_ACTION Then
    Else
        ' get the structure
        CopyMemory CWP, ByVal lParam, Len(CWP)
        ' check the hook message
        Select Case CWP.Message
            ' creating something
            Case WM_CREATE
                ' we don't want to subclass system menu
                If CWP.hWnd <> lngSyshWnd Then
                    ' check if it is a menu class
                    bHandled = Private_IsMenuClass(CWP.hWnd)
                    ' initialize subclassing for the hook if it is
                    If bHandled Then
                        Private_AddMenuHook CWP.hWnd
                        If m_MenuHook = 0 Then
                            m_MenuHook = CWP.hWnd
                            shk_SetHook WH_KEYBOARD, , MSG_BEFORE, , 3
                        End If
                    End If
                End If
            Case WM_DESTROY
                If m_MenuHook = 0 Then Else If m_MenuHook = CWP.hWnd Then m_MenuHook = 0: shk_UnHook WH_KEYBOARD
        End Select
    End If
End Sub
Private Sub Private_WndProcMenu(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
    Dim lngTemp As Long
    Dim udtDrawItem As DRAWITEMSTRUCT, udtMeasure As MEASUREITEMSTRUCT, udtMenuItem As MENUITEMINFO, udtCalcRECT As RECT
    Dim lngFont As Long, lngOFont As Long, strText() As String
    Dim lngA As Long, lngDC As Long, lngHeight As Long, lngWidth As Long
    If uMsg = WM_NCPAINT Then Debug.Print "WM_NCPAINT: " & Hex$(lng_hWnd), IIf(bBefore, "BEFORE", "AFTER")
    If bBefore Then
        Select Case uMsg
            Case WM_DRAWITEM
                CopyMemory udtDrawItem, ByVal lParam, Len(udtDrawItem)
                If udtDrawItem.CtlType = ODT_MENU Then
                    ' draw the menu item
                    With udtMenuItem
                        .fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
                        .cbSize = Len(udtMenuItem)
                        ' get the data
                        GetMenuItemInfo udtDrawItem.hwndItem, udtDrawItem.itemID, False, udtMenuItem
                    End With
                    Private_DrawMenuItem lng_hWnd, udtDrawItem, udtMenuItem
                    ' don't process more
                    lReturn = 0
                    bHandled = True
                End If
            Case WM_ERASEBKGND
                '
            Case WM_INITMENUPOPUP
                If wParam <> lngSyshWnd Then
                    For lngA = 0 To GetMenuItemCount(wParam) - 1
                        With udtMenuItem
                            .fMask = MIIM_STRING
                            .dwTypeData = vbNullString
                            .cch = 0
                            .cbSize = Len(udtMenuItem)
                            GetMenuItemInfo wParam, lngA, True, udtMenuItem
                            ' initialize information to get data
                            .fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
                            .cbSize = Len(udtMenuItem)
                            .dwTypeData = String$(.cch, vbNullChar)
                            .cch = .cch + 1
                            ' get the data
                            GetMenuItemInfo wParam, lngA, True, udtMenuItem
                            ' store caption
                            Private_AddMenuCaption .wid, .dwTypeData
                            ' set to owner-drawn
                            .fType = .fType Or MF_OWNERDRAW
                            .fMask = MIIM_ID Or MIIM_TYPE
                            ' set the data
                            SetMenuItemInfo wParam, lngA, True, udtMenuItem
                            ' record owner-draw so we can remove it... (in design time)
                            If blnDesignTime Then Private_AddOwnerDraw wParam, lngA
                        End With
                    Next lngA
                    ' don't process more
                    lReturn = 0
                    bHandled = True
                End If
            Case WM_MEASUREITEM
                CopyMemory udtMeasure, ByVal lParam, Len(udtMeasure)
                If udtMeasure.CtlType = ODT_MENU Then
                    With udtMenuItem
                        ' initialize information to get data
                        .fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
                        .cbSize = Len(udtMenuItem)
                        ' get the data
                        GetMenuItemInfo lngMenuhWnd, udtMeasure.itemID, False, udtMenuItem
                        If Not (.fType And MF_SEPARATOR) = MF_SEPARATOR Then
                            ' check if the string is of any size
                            If LenB(Private_GetMenuCaption(udtMeasure.itemID)) > 0 Then
                                ' separate texts by tab
                                strText = Split(Private_GetMenuCaption(udtMeasure.itemID), vbTab)
                                ' get DC
                                lngDC = GetDC(lng_hWnd)
                                lngHeight = udtMeasure.itemHeight
                                lngFont = Private_GetFont(lngDC)
                                lngOFont = SelectObject(lngDC, lngFont)
                                ' get string lengths
                                For lngA = 0 To UBound(strText)
                                    Private_DrawTextAuto lngDC, StrPtr(strText(lngA)), LenB(strText(lngA)), udtCalcRECT, DT_CALCRECT
                                    lngWidth = lngWidth + udtCalcRECT.Right
                                    If lngHeight < udtCalcRECT.Bottom Then lngHeight = udtCalcRECT.Bottom
                                Next lngA
                                If lngOFont <> 0 Then
                                    SelectObject lngDC, lngOFont
                                    DeleteObject lngFont
                                End If
                                ' release DC
                                ReleaseDC lng_hWnd, lngDC
                                ' set width
                                If .dwItemData <> 2 Then
                                    lngWidth = lngWidth + (UBound(strText) + 1) * 7 + 30 - 4
                                    ' set height
                                    udtMeasure.itemHeight = lngHeight + 6
                                Else
                                    lngWidth = lngWidth ' + (UBound(strText) + 1) * 4
                                    ' set height
                                    udtMeasure.itemHeight = lngHeight
                                End If
                                If udtMeasure.itemWidth < lngWidth Then udtMeasure.itemWidth = lngWidth
                                ' send structure back
                                CopyMemory ByVal lParam, udtMeasure, Len(udtMeasure)
                                ' don't process more
                                lReturn = 0
                                bHandled = True
                            End If
                        Else
                            udtMeasure.itemWidth = Styles.ImageSize
                            udtMeasure.itemHeight = 3
                            CopyMemory ByVal lParam, udtMeasure, Len(udtMeasure)
                            ' don't process more
                            lReturn = 0
                            bHandled = True
                        End If
                    End With
                End If
            Case WM_MENUSELECT
                'With udtMenuItem
                '    ' initialize information to get data
                '    .fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
                '    .cbSize = Len(udtMenuItem)
                '    ' get the data
                '    GetMenuItemInfo lngMenuhWnd, wParam And &HFFFF&, False, udtMenuItem
                '    If .wid = 0 Then
                '        If (wParam \ &H10000 And &HFFFF&) = &HFFFF& Then
                '            lngLastMenuParent = 0
                '        Else
                '            lngLastMenuParent = lParam
                '        End If
                '    End If
                '    'Debug.Print .dwItemData & " (" & CStr(.wid) & " | " & CStr(wParam And &HFFFF&) & ")", Hex$(wParam \ &H10000 And &HFFFF&), Hex$(lParam)
                'End With
        End Select
    ' AFTER
    Else
        ' this handles the subclassing of the border around the menuitems
        ' with the exception of WM_SHOWWINDOW, which is used to see if we can start owner-draw
        Select Case uMsg
            Case WM_SHOWWINDOW
                ' unsubclass because we don't need to watch for WM_SHOWWINDOW anymore
                ssc_DelMsg lng_hWnd, MSG_AFTER, WM_SHOWWINDOW
                'ssc_UnSubclass lng_hWnd
                'ssc_Terminate
                Debug.Print "UniMenu (WM_SHOWWINDOW): Ended subclassing! " & Hex$(lng_hWnd)
                ' this will begin the owner-drawing
                m_ShowWindow = True
                Private_Enable
            Case WM_DESTROY
                ' unsubclass when the menu is destroyed
                Private_ClearMenuHook lng_hWnd
            Case WM_NCPAINT
                ' this draws all the menus created by the program... because all those are hooked!
                ' this includes system/control/window menu (each word is the same thing) which can't be owner-drawn (afaik)
                If Not Styles.In3D Then Private_DrawMenuBorder lng_hWnd, wParam
        End Select
    End If
End Sub
