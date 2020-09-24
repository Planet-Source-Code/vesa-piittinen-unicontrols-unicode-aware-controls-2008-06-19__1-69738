Attribute VB_Name = "Module1"
Option Explicit

'=====[CONSTANTS]==============================================================================

' Class styles
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000

' Window styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  ' WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)

' ExWindowStyles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_CLIENTEDGE = &H200&

' Color constants
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20

' Window messages
Public Enum WND_MSG
    WM_ACTIVATE = &H6
    WM_ACTIVATEAPP = &H1C
    WM_APPCOMMAND = &H319
    WM_ASKCBFORMATNAME = &H30C
    WM_CANCELJOURNAL = &H4B
    WM_CANCELMODE = &H1F
    WM_CAPTURECHANGED = &H1F
    WM_CAPTURECHANGED_R = &H215
    WM_CHANGECBCHAIN = &H30D
    WM_CHAR = &H102
    WM_CHARTOITEM = &H2F
    WM_CHILDACTIVATE = &H22
    WM_CHOOSEFONT_GETLOGFONT = &H401
    WM_CHOOSEFONT_SETFLAGS = (&H400 + 102)
    WM_CHOOSEFONT_SETLOGFONT = (&H400 + 101)
    WM_CLEAR = &H303
    WM_CLOSE = &H10
    WM_COMMAND = &H111
    WM_COMPACTING = &H41
    WM_COMPAREITEM = &H39
    WM_CONTEXTMENU = &H7B
    WM_CONVERTREQUESTEX = &H108
    WM_COPY = &H301
    WM_COPYDATA = &H4A
    WM_CREATE = &H1
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138
    WM_CUT = &H300
    WM_DDE_ACK = (&H3E0 + 4)
    WM_DDE_ADVISE = (&H3E0 + 2)
    WM_DDE_DATA = (&H3E0 + 5)
    WM_DDE_EXECUTE = (&H3E0 + 8)
    WM_DDE_FIRST = &H3E0
    WM_DDE_INITIATE = &H3E0
    WM_DDE_LAST = (&H3E0 + 8)
    WM_DDE_POKE = (&H3E0 + 7)
    WM_DDE_REQUEST = (&H3E0 + 6)
    WM_DDE_TERMINATE = (&H3E0 + 1)
    WM_DDE_UNADVISE = (&H3E0 + 3)
    WM_DEADCHAR = &H103
    WM_DELETEITEM = &H2D
    WM_DESTROY = &H2
    WM_DESTROYCLIPBOARD = &H307
    WM_DEVICECHANGE = &H219
    WM_DEVMODECHANGE = &H1B
    WM_DRAWCLIPBOARD = &H308
    WM_DRAWITEM = &H2B
    WM_DROPFILES = &H233
    WM_ENABLE = &HA
    WM_ENDSESSION = &H16
    WM_ENTERIDLE = &H121
    WM_ENTERSIZEMOVE = &H231
    WM_ENTERMENULOOP = &H211
    WM_ERASEBKGND = &H14
    WM_EXITMENULOOP = &H212
    WM_EXITSIZEMOVE = &H232
    WM_FONTCHANGE = &H1D
    WM_GETDLGCODE = &H87
    WM_GETFONT = &H31
    WM_GETHOTKEY = &H33
    WM_GETMINMAXINFO = &H24
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_HELP = &H53
    WM_HOTKEY = &H312
    WM_HSCROLL = &H114
    WM_HSCROLLCLIPBOARD = &H30E
    WM_ICONERASEBKGND = &H27
    WM_IME_CHAR = &H286
    WM_IME_COMPOSITION = &H10F
    WM_IME_COMPOSITIONFULL = &H284
    WM_IME_CONTROL = &H283
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_KEYDOWN = &H290
    WM_IME_KEYLAST = &H10F
    WM_IME_KEYUP = &H291
    WM_IME_NOTIFY = &H282
    WM_IME_SELECT = &H285
    WM_IME_SETCONTEXT = &H281
    WM_IME_STARTCOMPOSITION = &H10D
    WM_INITDIALOG = &H110
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_INPUTLANGCHANGEREQUEST = &H50
    WM_INPUTLANGCHANGE = &H51
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_KILLFOCUS = &H8
    WM_LBUTTONDBLCLK = &H203
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_MBUTTONDBLCLK = &H209
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MDIACTIVATE = &H222
    WM_MDICASCADE = &H227
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIGETACTIVE = &H229
    WM_MDIICONARRANGE = &H228
    WM_MDIMAXIMIZE = &H225
    WM_MDINEXT = &H224
    WM_MDIREFRESHMENU = &H234
    WM_MDIRESTORE = &H223
    WM_MDISETMENU = &H230
    WM_MDITILE = &H226
    WM_MEASUREITEM = &H2C
    WM_MENUCHAR = &H120
    WM_MENUSELECT = &H11F
    WM_MENURBUTTONUP = &H122
    WM_MENUDRAG = &H123
    WM_MENUGETOBJECT = &H124
    WM_MENUCOMMAND = &H126
    WM_MOUSEACTIVATE = &H21
    WM_MOUSEHOVER = &H2A1
    WM_MOUSELEAVE = &H2A3
    WM_MOUSEMOVE = &H200
    WM_MOUSEWHEEL = &H20A
    WM_MOVE = &H3
    WM_MOVING = &H216
    WM_NCACTIVATE = &H86
    WM_NCCALCSIZE = &H83
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCHITTEST = &H84
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCMBUTTONDBLCLK = &HA9
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMOUSEMOVE = &HA0
    WM_NCPAINT = &H85
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NEXTDLGCTL = &H28
    WM_NEXTMENU = &H213
    WM_NOTIFY = &H210
    WM_NULL = &H0
    WM_PAINT = &HF
    WM_PAINTCLIPBOARD = &H309
    WM_PAINTICON = &H26
    WM_PALETTECHANGED = &H311
    WM_PALETTEISCHANGING = &H310
    WM_PASTE = &H302
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    WM_POWER = &H48
    WM_POWERBROADCAST = &H218
    WM_PRINT = &H317
    WM_PRINTCLIENT = &H318
    WM_PSD_ENVSTAMPRECT = (&H400 + 5)
    WM_PSD_FULLPAGERECT = (&H400 + 1)
    WM_PSD_GREEKTEXTRECT = (&H400 + 4)
    WM_PSD_MARGINRECT = (&H400 + 3)
    WM_PSD_MINMARGINRECT = (&H400 + 2)
    WM_PSD_PAGESETUPDLG = (&H400)
    WM_PSD_YAFULLPAGERECT = (&H400 + 6)
    WM_QUERYDRAGICON = &H37
    WM_QUERYENDSESSION = &H11
    WM_QUERYNEWPALETTE = &H30F
    WM_QUERYOPEN = &H13
    WM_QUEUESYNC = &H23
    WM_QUIT = &H12
    WM_RBUTTONDBLCLK = &H206
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RENDERALLFORMATS = &H306
    WM_RENDERFORMAT = &H305
    WM_SETCURSOR = &H20
    WM_SETFOCUS = &H7
    WM_SETFONT = &H30
    WM_SETHOTKEY = &H32
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_SETTINGCHANGE = &H1A
    WM_SHOWWINDOW = &H18
    WM_SIZE = &H5
    WM_SIZING = &H214
    WM_SIZECLIPBOARD = &H30B
    WM_SPOOLERSTATUS = &H2A
    WM_SYSCHAR = &H106
    WM_SYSCOLORCHANGE = &H15
    WM_SYSCOMMAND = &H112
    WM_SYSDEADCHAR = &H107
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_TIMECHANGE = &H1E
    WM_TIMER = &H113
    WM_UNDO = &H304
    WM_USER = &H400
    WM_VKEYTOITEM = &H2E
    WM_VSCROLL = &H115
    WM_VSCROLLCLIPBOARD = &H30A
    WM_WINDOWPOSCHANGED = &H47
    WM_WINDOWPOSCHANGING = &H46
    WM_WININICHANGE = &H1A
End Enum

' ShowWindow commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

' Standard ID's of cursors
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&
Public Const GWL_WNDPROC = -4

Public Const IDC_MAIN_MDI = 1001
Public Const ID_MDI_FIRSTCHILD = 1004

Public Const CW_USEDEFAULT = &H80000000
Public Const MDIS_ALLCHILDSTYLES = &H1&

Public Const MF_BYCOMMAND = &H0&

Public Const SC_MAXIMIZE As Long = &HF030&
Public Const SC_MOVE As Long = &HF010&
Public Const SC_SIZE As Long = &HF000&

Public Const GWL_STYLE As Long = (-16)

Public Const CF_UNICODETEXT = 13
Public Const GMEM_DDESHARE = &H2000&
Public Const GMEM_MOVEABLE = &H2&

Public Const MK_CONTROL = &H8&
Public Const MK_LBUTTON = &H1&
Public Const MK_MBUTTON = &H10&
Public Const MK_RBUTTON = &H2&

' border style
Public Const BF_LEFT = &H1&
Public Const BF_TOP = &H2&
Public Const BF_RIGHT = &H4&
Public Const BF_BOTTOM = &H8&
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' scroll info
Public Const SIF_RANGE = &H1&
Public Const SIF_PAGE = &H2&
Public Const SIF_POS = &H4&
Public Const SIF_DISABLENOSCROLL = &H8&
Public Const SIF_TRACKPOS = &H10&
Public Const SIF_ALL = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS

Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5

' for SetIcon: http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGACOLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000&

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

Public Const IMAGE_ICON = 1

Public Const SM_CXICON = 11
Public Const SM_CYICON = 12

Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50

Public Const WM_SETICON = &H80

Public Const GW_OWNER = 4

' WM_SIZE (for WindowState)
Public Enum WM_SIZE_wParam
    SIZE_RESTORED = 0&
    SIZE_MINIMIZED = 1&
    SIZE_MAXIMIZED = 2&
    SIZE_MAXSHOW = 3&
    SIZE_MAXHIDE = 4&
End Enum


'=====[USER DEFINED TYPES]=====================================================================

' Standalone User Defined Types
Public Type CLIENTCREATESTRUCT
    hWindowMenu As Long
    idFirstChild As Long
End Type

Public Type MDICREATESTRUCT
    szClass As String
    szTitle As String
    hOwner As Long
    X As Long
    Y As Long
    cX As Long
    cY As Long
    Style As Long
    lParam As Long
End Type

Public Type NMHDR
    hWndFrom As Long
    IDFrom As Long
    Code As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Public Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

' User Defined Types that require the above UDTs
Public Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type


'=====[API DECLARATIONS]=======================================================================

' API declarations
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemoryFromRect Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As RECT, ByVal Length As Long)
Public Declare Sub CopyMemoryToRect Lib "kernel32" Alias "RtlMoveMemory" (Destination As RECT, ByVal Source As Long, ByVal Length As Long)
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal Bool As Boolean) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function TranslateMDISysAccel Lib "user32" (ByVal hWndClient As Long, lpMsg As Msg) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

' API declarations required by xProcAPI and xProcReplace procedures
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLen As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Function AddOf(ByVal AddressOfProcedure As Long) As Long
    AddOf = AddressOfProcedure
End Function
'Public Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Function CreateWindowExW(ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
    xProcAPI AddressOf Module1.CreateWindowExW, "CreateWindowExW", "user32.dll"
    CreateWindowExW = CreateWindowExW(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam)
End Function
'Public Declare Function DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hWnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Function DefFrameProc(ByVal hWnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    xProcAPI AddressOf Module1.DefFrameProc, "DefFrameProcA", "user32.dll"
    DefFrameProc = DefFrameProc(hWnd, hWndMDIClient, wMsg, wParam, lParam)
End Function
'Public Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Function DefMDIChildProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    xProcAPI AddressOf Module1.DefMDIChildProc, "DefMDIChildProcA", "user32.dll"
    DefMDIChildProc = DefMDIChildProc(hWnd, wMsg, wParam, lParam)
End Function
'Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Function DefWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    xProcAPI AddressOf Module1.DefWindowProc, "DefWindowProcW", "user32.dll"
    DefWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End Function
'Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Function FreeLibrary(ByVal hLibModule As Long) As Long
    xProcAPI AddressOf Module1.FreeLibrary, "FreeLibrary", "kernel32.dll"
    FreeLibrary = FreeLibrary(hLibModule)
End Function
' for SetIcon
'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Function GetSystemMetrics(ByVal nIndex As Long) As Long
    xProcAPI AddressOf Module1.GetSystemMetrics, "GetSystemMetrics", "user32.dll"
    GetSystemMetrics = GetSystemMetrics(nIndex)
End Function
'Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Function GetTickCount() As Long
    xProcAPI AddressOf Module1.GetTickCount, "GetTickCount", "kernel32.dll"
    GetTickCount = GetTickCount
End Function
' for SetIcon
'Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Function GetWindow(ByVal hWnd As Long, ByVal wCmd As Long) As Long
    xProcAPI AddressOf Module1.GetWindow, "GetWindow", "user32.dll"
    GetWindow = GetWindow(hWnd, wCmd)
End Function
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Function GetWindowText(ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
    xProcAPI AddressOf Module1.GetWindowText, "GetWindowTextW", "user32.dll"
    GetWindowText = GetWindowText(hWnd, lpString, cch)
End Function
'Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Public Function GetWindowTextLength(ByVal hWnd As Long) As Long
    xProcAPI AddressOf Module1.GetWindowTextLength, "GetWindowTextLengthW", "user32.dll"
    GetWindowTextLength = GetWindowTextLength(hWnd)
End Function
'Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Function GlobalAlloc(ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    xProcAPI AddressOf Module1.GlobalAlloc, "GlobalAlloc", "kernel32.dll"
    GlobalAlloc = GlobalAlloc(wFlags, dwBytes)
End Function
'Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Function GlobalFree(ByVal hMem As Long) As Long
    xProcAPI AddressOf Module1.GlobalFree, "GlobalFree", "kernel32.dll"
    GlobalFree = GlobalFree(hMem)
End Function
'Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Function GlobalLock(ByVal hMem As Long) As Long
    xProcAPI AddressOf Module1.GlobalLock, "GlobalLock", "kernel32.dll"
    GlobalLock = GlobalLock(hMem)
End Function
'Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Function GlobalUnlock(ByVal hMem As Long) As Long
    xProcAPI AddressOf Module1.GlobalUnlock, "GlobalUnlock", "kernel32.dll"
    GlobalUnlock = GlobalUnlock(hMem)
End Function
' for SetIcon
'Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Public Function LoadImageAsString(ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
    xProcAPI AddressOf Module1.LoadImageAsString, "LoadImageW", "user32.dll"
    LoadImageAsString = LoadImageAsString(hInst, lpsz, uType, cxDesired, cyDesired, fuLoad)
End Function
'Public Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long
Public Function LoadLibraryW(ByVal lpLibFileName As String) As Long
    xProcAPI AddressOf Module1.LoadLibraryW, "LoadLibraryW", "kernel32.dll"
    LoadLibraryW lpLibFileName
End Function
'Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Function MultiByteToWideChar(ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    xProcAPI AddressOf Module1.MultiByteToWideChar, "MultiByteToWideChar", "kernel32.dll"
    MultiByteToWideChar = MultiByteToWideChar(CodePage, dwFlags, lpMultiByteStr, cchMultiByte, lpWideCharStr, cchWideChar)
End Function
Public Function SendLong(ByVal hWnd As Long, wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    xProcAPI AddressOf Module1.SendLong, "SendMessageW", "user32.dll"
    SendLong = SendLong(hWnd, wMsg, wParam, lParam)
End Function
Public Function SendUnicode(ByVal hWnd As Long, wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    xProcAPI AddressOf Module1.SendUnicode, "SendMessageW", "user32.dll"
    SendUnicode = SendUnicode(hWnd, wMsg, wParam, lParam)
End Function
' Customized from: http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp
Public Sub SetIcon(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = False)
    Dim lhWndTop As Long, lhWnd As Long
    Dim cX As Long, cY As Long
    Dim hIconLarge As Long, hIconSmall As Long
    
    If bSetAsAppIcon Then
        lhWnd = hWnd
        lhWndTop = lhWnd
        Do While lhWnd <> 0
            lhWnd = GetWindow(lhWnd, GW_OWNER)
            If lhWnd <> 0 Then lhWndTop = lhWnd
        Loop
    End If
    
    cX = GetSystemMetrics(SM_CXICON)
    cY = GetSystemMetrics(SM_CYICON)
    
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cX, cY, LR_SHARED)
    
    If bSetAsAppIcon Then SendMessage lhWndTop, WM_SETICON, ICON_BIG, ByVal hIconLarge
    SendMessage hWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge
    
    cX = GetSystemMetrics(SM_CXSMICON)
    cY = GetSystemMetrics(SM_CYSMICON)
    
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cX, cY, LR_SHARED)
    
    If bSetAsAppIcon Then SendMessage lhWndTop, WM_SETICON, ICON_SMALL, ByVal hIconSmall
    SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall
End Sub
'Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextW" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Function SetWindowText(ByVal hWnd As Long, ByVal lpString As String) As Long
    xProcAPI AddressOf Module1.SetWindowText, "SetWindowTextW", "user32.dll"
    SetWindowText = SetWindowText(hWnd, lpString)
End Function
'Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Function URLDownloadToFile(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    xProcAPI AddressOf Module1.URLDownloadToFile, "URLDownloadToFileW", "urlmon.dll"
    URLDownloadToFile = URLDownloadToFile(pCaller, szURL, szFileName, dwReserved, lpfnCB)
End Function
'Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
Public Function WideCharToMultiByte(ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    xProcAPI AddressOf Module1.WideCharToMultiByte, "WideCharToMultiByte", "kernel32.dll"
    WideCharToMultiByte = WideCharToMultiByte(CodePage, dwFlags, lpWideCharStr, cchWideChar, lpMultiByteStr, cchMultiByte, lpDefaultChar, lpUsedDefaultChar)
End Function
' preserve Enum cases
Private Sub xEnumCase()
    Dim WM_ACTIVATE As Long
    Dim WM_ACTIVATEAPP As Long
    Dim WM_APPCOMMAND As Long
    Dim WM_ASKCBFORMATNAME As Long
    Dim WM_CANCELJOURNAL As Long
    Dim WM_CANCELMODE As Long
    Dim WM_CAPTURECHANGED As Long
    Dim WM_CAPTURECHANGED_R As Long
    Dim WM_CHANGECBCHAIN As Long
    Dim WM_CHAR As Long
    Dim WM_CHARTOITEM As Long
    Dim WM_CHILDACTIVATE As Long
    Dim WM_CHOOSEFONT_GETLOGFONT As Long
    Dim WM_CHOOSEFONT_SETFLAGS As Long
    Dim WM_CHOOSEFONT_SETLOGFONT As Long
    Dim WM_CLEAR As Long
    Dim WM_CLOSE As Long
    Dim WM_COMMAND As Long
    Dim WM_COMPACTING As Long
    Dim WM_COMPAREITEM As Long
    Dim WM_CONTEXTMENU As Long
    Dim WM_CONVERTREQUESTEX As Long
    Dim WM_COPY As Long
    Dim WM_COPYDATA As Long
    Dim WM_CREATE As Long
    Dim WM_CTLCOLORBTN As Long
    Dim WM_CTLCOLORDLG As Long
    Dim WM_CTLCOLOREDIT As Long
    Dim WM_CTLCOLORLISTBOX As Long
    Dim WM_CTLCOLORMSGBOX As Long
    Dim WM_CTLCOLORSCROLLBAR As Long
    Dim WM_CTLCOLORSTATIC As Long
    Dim WM_CUT As Long
    Dim WM_DDE_ACK As Long
    Dim WM_DDE_ADVISE As Long
    Dim WM_DDE_DATA As Long
    Dim WM_DDE_EXECUTE As Long
    Dim WM_DDE_FIRST As Long
    Dim WM_DDE_INITIATE As Long
    Dim WM_DDE_LAST As Long
    Dim WM_DDE_POKE As Long
    Dim WM_DDE_REQUEST As Long
    Dim WM_DDE_TERMINATE As Long
    Dim WM_DDE_UNADVISE As Long
    Dim WM_DEADCHAR As Long
    Dim WM_DELETEITEM As Long
    Dim WM_DESTROY As Long
    Dim WM_DESTROYCLIPBOARD As Long
    Dim WM_DEVICECHANGE As Long
    Dim WM_DEVMODECHANGE As Long
    Dim WM_DRAWCLIPBOARD As Long
    Dim WM_DRAWITEM As Long
    Dim WM_DROPFILES As Long
    Dim WM_ENABLE As Long
    Dim WM_ENDSESSION As Long
    Dim WM_ENTERIDLE As Long
    Dim WM_ENTERSIZEMOVE As Long
    Dim WM_ENTERMENULOOP As Long
    Dim WM_ERASEBKGND As Long
    Dim WM_EXITMENULOOP As Long
    Dim WM_EXITSIZEMOVE As Long
    Dim WM_FONTCHANGE As Long
    Dim WM_GETDLGCODE As Long
    Dim WM_GETFONT As Long
    Dim WM_GETHOTKEY As Long
    Dim WM_GETMINMAXINFO As Long
    Dim WM_GETTEXT As Long
    Dim WM_GETTEXTLENGTH As Long
    Dim WM_HELP As Long
    Dim WM_HOTKEY As Long
    Dim WM_HSCROLL As Long
    Dim WM_HSCROLLCLIPBOARD As Long
    Dim WM_ICONERASEBKGND As Long
    Dim WM_IME_CHAR As Long
    Dim WM_IME_COMPOSITION As Long
    Dim WM_IME_COMPOSITIONFULL As Long
    Dim WM_IME_CONTROL As Long
    Dim WM_IME_ENDCOMPOSITION As Long
    Dim WM_IME_KEYDOWN As Long
    Dim WM_IME_KEYLAST As Long
    Dim WM_IME_KEYUP As Long
    Dim WM_IME_NOTIFY As Long
    Dim WM_IME_SELECT As Long
    Dim WM_IME_SETCONTEXT As Long
    Dim WM_IME_STARTCOMPOSITION As Long
    Dim WM_INITDIALOG As Long
    Dim WM_INITMENU As Long
    Dim WM_INITMENUPOPUP As Long
    Dim WM_INPUTLANGCHANGEREQUEST As Long
    Dim WM_INPUTLANGCHANGE As Long
    Dim WM_KEYDOWN As Long
    Dim WM_KEYUP As Long
    Dim WM_KILLFOCUS As Long
    Dim WM_LBUTTONDBLCLK As Long
    Dim WM_LBUTTONDOWN As Long
    Dim WM_LBUTTONUP As Long
    Dim WM_MBUTTONDBLCLK As Long
    Dim WM_MBUTTONDOWN As Long
    Dim WM_MBUTTONUP As Long
    Dim WM_MDIACTIVATE As Long
    Dim WM_MDICASCADE As Long
    Dim WM_MDICREATE As Long
    Dim WM_MDIDESTROY As Long
    Dim WM_MDIGETACTIVE As Long
    Dim WM_MDIICONARRANGE As Long
    Dim WM_MDIMAXIMIZE As Long
    Dim WM_MDINEXT As Long
    Dim WM_MDIREFRESHMENU As Long
    Dim WM_MDIRESTORE As Long
    Dim WM_MDISETMENU As Long
    Dim WM_MDITILE As Long
    Dim WM_MEASUREITEM As Long
    Dim WM_MENUCHAR As Long
    Dim WM_MENUSELECT As Long
    Dim WM_MENURBUTTONUP As Long
    Dim WM_MENUDRAG As Long
    Dim WM_MENUGETOBJECT As Long
    Dim WM_MENUCOMMAND As Long
    Dim WM_MOUSEACTIVATE As Long
    Dim WM_MOUSEHOVER As Long
    Dim WM_MOUSELEAVE As Long
    Dim WM_MOUSEMOVE As Long
    Dim WM_MOUSEWHEEL As Long
    Dim WM_MOVE As Long
    Dim WM_MOVING As Long
    Dim WM_NCACTIVATE As Long
    Dim WM_NCCALCSIZE As Long
    Dim WM_NCCREATE As Long
    Dim WM_NCDESTROY As Long
    Dim WM_NCHITTEST As Long
    Dim WM_NCLBUTTONDBLCLK As Long
    Dim WM_NCLBUTTONDOWN As Long
    Dim WM_NCLBUTTONUP As Long
    Dim WM_NCMBUTTONDBLCLK As Long
    Dim WM_NCMBUTTONDOWN As Long
    Dim WM_NCMBUTTONUP As Long
    Dim WM_NCMOUSEMOVE As Long
    Dim WM_NCPAINT As Long
    Dim WM_NCRBUTTONDBLCLK As Long
    Dim WM_NCRBUTTONDOWN As Long
    Dim WM_NCRBUTTONUP As Long
    Dim WM_NEXTDLGCTL As Long
    Dim WM_NEXTMENU As Long
    Dim WM_NOTIFY As Long
    Dim WM_NULL As Long
    Dim WM_PAINT As Long
    Dim WM_PAINTCLIPBOARD As Long
    Dim WM_PAINTICON As Long
    Dim WM_PALETTECHANGED As Long
    Dim WM_PALETTEISCHANGING As Long
    Dim WM_PASTE As Long
    Dim WM_PENWINFIRST As Long
    Dim WM_PENWINLAST As Long
    Dim WM_POWER As Long
    Dim WM_POWERBROADCAST As Long
    Dim WM_PRINT As Long
    Dim WM_PRINTCLIENT As Long
    Dim WM_PSD_ENVSTAMPRECT As Long
    Dim WM_PSD_FULLPAGERECT As Long
    Dim WM_PSD_GREEKTEXTRECT As Long
    Dim WM_PSD_MARGINRECT As Long
    Dim WM_PSD_MINMARGINRECT As Long
    Dim WM_PSD_PAGESETUPDLG As Long
    Dim WM_PSD_YAFULLPAGERECT As Long
    Dim WM_QUERYDRAGICON As Long
    Dim WM_QUERYENDSESSION As Long
    Dim WM_QUERYNEWPALETTE As Long
    Dim WM_QUERYOPEN As Long
    Dim WM_QUEUESYNC As Long
    Dim WM_QUIT As Long
    Dim WM_RBUTTONDBLCLK As Long
    Dim WM_RBUTTONDOWN As Long
    Dim WM_RBUTTONUP As Long
    Dim WM_RENDERALLFORMATS As Long
    Dim WM_RENDERFORMAT As Long
    Dim WM_SETCURSOR As Long
    Dim WM_SETFOCUS As Long
    Dim WM_SETFONT As Long
    Dim WM_SETHOTKEY As Long
    Dim WM_SETREDRAW As Long
    Dim WM_SETTEXT As Long
    Dim WM_SETTINGCHANGE As Long
    Dim WM_SHOWWINDOW As Long
    Dim WM_SIZE As Long
    Dim WM_SIZING As Long
    Dim WM_SIZECLIPBOARD As Long
    Dim WM_SPOOLERSTATUS As Long
    Dim WM_SYSCHAR As Long
    Dim WM_SYSCOLORCHANGE As Long
    Dim WM_SYSCOMMAND As Long
    Dim WM_SYSDEADCHAR As Long
    Dim WM_SYSKEYDOWN As Long
    Dim WM_SYSKEYUP As Long
    Dim WM_TIMECHANGE As Long
    Dim WM_TIMER As Long
    Dim WM_UNDO As Long
    Dim WM_USER As Long
    Dim WM_VKEYTOITEM As Long
    Dim WM_VSCROLL As Long
    Dim WM_VSCROLLCLIPBOARD As Long
    Dim WM_WINDOWPOSCHANGED As Long
    Dim WM_WINDOWPOSCHANGING As Long
    Dim WM_WININICHANGE As Long
    
    Dim SIZE_RESTORED As Long
    Dim SIZE_MINIMIZED As Long
    Dim SIZE_MAXIMIZED As Long
    Dim SIZE_MAXSHOW As Long
    Dim SIZE_MAXHIDE As Long
End Sub
' replace procedure with an API call
Private Sub xProcAPI(ByVal AddressOfDest As Long, ByRef API As String, ByRef Module As String)
    Dim lngModuleHandle As Long, AddressOfSrc As Long, lngProcessHandle As Long, lngBytesWritten As Long
    Dim lngJMPASM(1) As Long
    ' get handle for module
    lngModuleHandle = GetModuleHandle(Module)
    If lngModuleHandle = 0 Then lngModuleHandle = LoadLibrary(Module)
    ' if failed, we can't do anything
    If lngModuleHandle = 0 Then Exit Sub
    ' get address of function
    AddressOfSrc = GetProcAddress(lngModuleHandle, API)
    ' if failed, we can't do anything
    If AddressOfSrc = 0 Then Exit Sub
    ' get a handle for current process
    lngProcessHandle = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)
    ' if failed, we can't do anything
    If lngProcessHandle = 0 Then Exit Sub
    ' check if we are in the IDE
    If App.LogMode = 0 Then
        ' get the real location of the procedure
        CopyMemory AddressOfDest, ByVal AddressOfDest + &H16&, 4&
    End If
    ' set ASM JMP
    lngJMPASM(0) = &HE9000000
    ' set JMP parameter (how many bytes to jump)
    lngJMPASM(1) = AddressOfSrc - AddressOfDest - 5&
    ' replace original procedure with the JMP
    WriteProcessMemory lngProcessHandle, ByVal AddressOfDest, ByVal VarPtr(lngJMPASM(0)) + 3&, 5&, lngBytesWritten
    ' close handle for current process
    CloseHandle lngProcessHandle
End Sub
' replace procedure with another procedure or with custom code
Private Sub xProcReplace(ByVal AddressOfDest As Long, ByVal AddressOfSrc As Long)
    Dim lngProcessHandle As Long, lngBytesWritten As Long
    Dim lngJMPASM(1) As Long
    ' get a handle for current process
    lngProcessHandle = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)
    ' if failed, we can't do anything
    If lngProcessHandle = 0 Then Exit Sub
    ' check if we are in the IDE
    If App.LogMode = 0 Then
        ' get the real locations of the procedures
        CopyMemory AddressOfDest, ByVal AddressOfDest + &H16&, 4&
        CopyMemory AddressOfSrc, ByVal AddressOfSrc + &H16&, 4&
    End If
    ' set ASM JMP
    lngJMPASM(0) = &HE9000000
    ' set JMP parameter (how many bytes to jump)
    lngJMPASM(1) = AddressOfSrc - AddressOfDest - 5&
    ' replace original procedure with the JMP
    WriteProcessMemory lngProcessHandle, ByVal AddressOfDest, ByVal VarPtr(lngJMPASM(0)) + 3&, 5&, lngBytesWritten
    ' close handle for current process
    CloseHandle lngProcessHandle
End Sub
Public Function WndProc(ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static blnPreventDefault As Boolean
    Select Case Message
        Case WM_SIZE
            Debug.Print "WM_SIZE"
        Case WM_SIZING
            Debug.Print "WM_SIZING"
        Case WM_ENTERSIZEMOVE
            Debug.Print "WM_ENTERSIZEMOVE"
        Case WM_EXITSIZEMOVE
            Debug.Print "WM_EXITSIZEMOVE"
        Case WM_DESTROY
            SetWindowLong Form1.hWnd, GWL_WNDPROC, Form1.WndProc
            Form1.Show
            Unload Form1
            Debug.Print "WM_DESTROY"
    End Select
    ' use the default window procedure
    'WndProc = WndProc2(hWnd, Message, wParam, lParam)
    WndProc = DefWindowProc(hWnd, Message, wParam, lParam)
End Function
Public Function WndProc2(ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    xProcReplace AddressOf Module1.WndProc2, Form1.WndProc
    WndProc2 = WndProc2(hWnd, Message, wParam, lParam)
End Function
