VERSION 5.00
Begin VB.UserControl UniForm 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UniForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000                  ' WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_CHILDWINDOW = (WS_CHILD)

Private Const IDC_ARROW = 32512&
Private Const IDC_IBEAM = 32513&
Private Const IDC_WAIT = 32514&
Private Const IDC_CROSS = 32515&
Private Const IDC_UPARROW = 32516&
Private Const IDC_SIZE = 32640&
Private Const IDC_ICON = 32641&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZENESW = 32643&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEALL = 32646&
Private Const IDC_NO = 32648&
Private Const IDC_APPSTARTING = 32650&

Private Const GWL_EXSTYLE = -20
Private Const GWL_STYLE = -16
Private Const GWL_WNDPROC = -4

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassW" (Class As WNDCLASS) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
End Type

Private m_ClientRect As RECT
Private m_RegisterClass As Boolean
Private m_hWnd As Long
Private m_Parent As Form
Private m_Wnd As WNDCLASS
Private m_WndClass As String

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim lngX As Long, lngY As Long, lngWidth As Long, lngHeight As Long
    If Ambient.UserMode Then
        m_WndClass = UserControl.Name & "_" & Right$("0000000" & Hex$(ObjPtr(Me)), 8) & Hex$(Timer)
        With m_Wnd
            .hInstance = App.hInstance
            .Style = 0
            .hbrBackground = vbWindowBackground And &H7FFFFFFF
            .hCursor = LoadCursor(0, IDC_ARROW)
            .lpszClassName = StrPtr(m_WndClass)
            .lpfnwndproc = AddOf(AddressOf Module1.WndProc)
        End With
        m_RegisterClass = RegisterClass(m_Wnd) <> 0&
        If m_RegisterClass Then
            lngX = Parent.Left \ Screen.TwipsPerPixelX
            lngY = Parent.Top \ Screen.TwipsPerPixelY
            lngWidth = Parent.Width \ Screen.TwipsPerPixelX
            lngHeight = Parent.Height \ Screen.TwipsPerPixelY
            If lngWidth < 1 Then lngWidth = 1
            If lngHeight < 1 Then lngHeight = 1
            m_hWnd = CreateWindowExW(0&, m_Wnd.lpszClassName, 0&, WS_OVERLAPPEDWINDOW, lngX, lngY, lngWidth, lngHeight, 0&, 0&, App.hInstance, ByVal 0&)
            If m_hWnd Then
                Set m_Parent = Parent
                'SetWindowTextW m_hWnd, StrPtr("ÅÄÖ" & ChrW$(&H3041) & ChrW$(&H3043))
                SendMessageW m_hWnd, WM_SETTEXT, 0&, ByVal StrPtr(ChrW$(&H3041) & ChrW$(&H3043) & "ÅÄÖ")
                SetWindowLongA Parent.hWnd, GWL_STYLE, GetWindowLongA(Parent.hWnd, GWL_STYLE) And Not WS_OVERLAPPEDWINDOW Or WS_CHILD
                SetParent Parent.hWnd, m_hWnd
                GetClientRect m_hWnd, m_ClientRect
                MoveWindow Parent.hWnd, 0, 0, m_ClientRect.Right - m_ClientRect.Left, m_ClientRect.Bottom - m_ClientRect.Top, -1&
                ShowWindow m_hWnd, SW_SHOWNORMAL
            End If
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    If Not m_Parent Is Nothing Then Unload m_Parent
    If m_hWnd Then DestroyWindow m_hWnd
    If m_RegisterClass Then UnregisterClass m_Wnd.lpszClassName, App.hInstance
End Sub
