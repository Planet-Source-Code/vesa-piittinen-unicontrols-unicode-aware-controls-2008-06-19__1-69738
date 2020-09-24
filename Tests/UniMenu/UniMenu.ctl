VERSION 5.00
Begin VB.UserControl UniMenu 
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UniMenu.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "UniMenu.ctx":0282
   Windowless      =   -1  'True
End
Attribute VB_Name = "UniMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum UniMenuParent
    [Unsupported Parent]
    [Form Parent]
    [MDI Parent]
End Enum

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

'Private Declare Function AppendMenuW Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal Bool As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function InsertMenuItemW Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal Bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal Bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long

Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Private m_RightToLeft As Boolean

Private WithEvents m_Menu As UniMenuItem
Attribute m_Menu.VB_VarHelpID = -1
Private m_MenuWnd As Long

Private m_Parent As UniMenuParent
Private m_ParentWnd As Long
Private m_ParentControl As Object

' for fixing XP Theme problem with a certain version of comctl32.dll
Private m_FreeShell32 As Boolean
Private m_Shell32 As Long

Dim blnDesignTime As Boolean                ' True if in IDE design time

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
    
    Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowUnicode Lib "USER32.DLL" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    '-SelfHook specific declarations----------------------------------------------------------------------------
    Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "USER32.DLL" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
    
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

Public Function Add(Optional ByRef ChildOf As UniMenuItem, Optional ByRef Key As String, Optional ByRef Caption As String = "-", Optional ByVal Checked As Boolean, Optional ByVal Enabled As Boolean, Optional ByVal RadioGroup As Byte, Optional HelpContextID As Boolean, Optional ByVal WindowList As Boolean, Optional ByVal Visible As Boolean = True) As UniMenuItem
    Set Add = m_Menu.MenuItemAdd(Key, Caption, Checked, Enabled, HelpContextID, RadioGroup, WindowList, Visible)
End Function
Public Function AddByKey(Optional ByRef ParentKey As String, Optional ByRef Key As String, Optional ByRef Caption As String = "-", Optional ByVal Checked As Boolean, Optional ByVal Enabled As Boolean, Optional ByVal RadioGroup As Byte, Optional HelpContextID As Boolean, Optional ByVal WindowList As Boolean, Optional ByVal Visible As Boolean = True) As UniMenuItem
    Set AddByKey = m_Menu.MenuItemAdd(Key, Caption, Checked, Enabled, HelpContextID, RadioGroup, WindowList, Visible)
End Function
Public Function Item(ByVal Index As Long) As UniMenuItem
    Set Item = m_Menu.MenuItem(Index)
End Function
Public Function ItemByKey(ByRef Key As String) As UniMenuItem
    Set ItemByKey = m_Menu.MenuItemByKey(Key)
End Function
Private Sub Private_GetParentForm()
    Set m_ParentControl = UserControl.Parent
    If m_ParentControl Is Nothing Then
        m_Parent = [Unsupported Parent]
        m_ParentWnd = 0
    ElseIf TypeOf m_ParentControl Is Form Then
        m_Parent = [Form Parent]
        m_ParentWnd = m_ParentControl.hWnd
    ElseIf TypeOf m_ParentControl Is MDIForm Then
        m_Parent = [MDI Parent]
        m_ParentWnd = m_ParentControl.hWnd
    Else
        Set m_ParentControl = Nothing
        m_Parent = [Unsupported Parent]
        m_ParentWnd = 0
    End If
End Sub
Private Sub Private_Init()
    If m_ParentWnd Then
        If m_MenuWnd Then
            DestroyMenu m_MenuWnd
            m_MenuWnd = CreateMenu
            SetMenu m_ParentWnd, m_MenuWnd
        Else
            m_MenuWnd = GetMenu(m_ParentWnd)
            If m_MenuWnd = 0 Then
                m_MenuWnd = CreateMenu
                SetMenu m_ParentWnd, m_MenuWnd
            End If
        End If
        m_Menu.hWnd = m_MenuWnd
        
    End If
End Sub
Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property
Public Property Let RightToLeft(ByVal NewValue As Boolean)
    m_RightToLeft = NewValue
    m_Menu.RightToLeft = NewValue
End Property
Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    '
End Sub
Private Sub UserControl_Initialize()
    Set m_Menu = New UniMenuItem
End Sub
Private Sub UserControl_InitProperties()
    Private_GetParentForm
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Private_GetParentForm
End Sub
Private Sub UserControl_Resize()
    If UserControl.ScaleWidth <> 32 Then UserControl.Width = UserControl.ScaleX(32, vbPixels, vbTwips)
    If UserControl.ScaleHeight <> 32 Then UserControl.Height = UserControl.ScaleY(32, vbPixels, vbTwips)
End Sub
Private Sub UserControl_Terminate()
    Set m_Menu = Nothing
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
End Sub

'-SelfSub code------------------------------------------------------------------------------------
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
        If thunkID <> 0 Then z_scFunk.Remove "h" & thunkID                   'Remove the specified thunk from the collection
        
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
