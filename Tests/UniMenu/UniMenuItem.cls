VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event MouseEnter()
Public Event MouseLeave()

Private Const MF_APPEND = &H100&
Private Const MF_BITMAP = &H4&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_CALLBACKS = &H8000000
Private Const MF_CHANGE = &H80&
Private Const MF_CHECKED = &H8&
Private Const MF_CONV = &H40000000
Private Const MF_DELETE = &H200&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_END = &H80
Private Const MF_ERRORS = &H10000000
Private Const MF_GRAYED = &H1&
Private Const MF_HELP = &H4000&
Private Const MF_HILITE = &H80&
Private Const MF_HSZ_INFO = &H1000000
Private Const MF_INSERT = &H0&
Private Const MF_LINKS = &H20000000
Private Const MF_MASK = &HFF000000
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_MOUSESELECT = &H8000&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_POPUP = &H10&
Private Const MF_POSTMSGS = &H4000000
Private Const MF_REMOVE = &H1000&
Private Const MF_SENDMSGS = &H2000000
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_SYSMENU = &H2000&
Private Const MF_UNCHECKED = &H0&
Private Const MF_UNHILITE = &H0&
Private Const MF_USECHECKBITMAPS = &H200&
Private Const MFCOMMENT = 15

Private Const MFS_GRAYED = &H3&
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE

Private Const MFT_STRING = MF_STRING
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200&
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_RIGHTORDER = &H2000&

Private Const MIIM_STATE As Long = &H1
Private Const MIIM_ID As Long = &H2
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_CHECKMARKS As Long = &H8
Private Const MIIM_TYPE As Long = &H10
Private Const MIIM_DATA As Long = &H20
Private Const MIIM_STRING As Long = &H40
Private Const MIIM_BITMAP As Long = &H80
Private Const MIIM_FTYPE As Long = &H100

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

Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal Bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long

Private m_Caption As String
Private m_Checked As Boolean
Private m_Enabled As Boolean
Private m_HelpContextID As Long
Private m_Key As String
Private m_RadioGroup As Byte
Private m_MouseOver As Boolean
Private m_WindowList As Boolean
Private m_Visible As Boolean

Private m_hWnd As Long
Private m_ID As Long
Private m_MenuItems As Collection
Private m_MenuParent As UniMenuItem
Private m_MII As MENUITEMINFO
Private m_RightToLeft As Boolean

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByRef NewValue As String)
    If m_hWnd = 0 Then
        m_Caption = NewValue
        If m_Caption = "-" Then
            m_MII.fType = m_MII.fType Or MFT_SEPARATOR And Not MFT_BITMAP And Not MFT_STRING
            m_MII.dwTypeData = 0
            m_MII.cch = 0
        Else
            m_MII.fType = m_MII.fType Or MFT_STRING And Not MFT_SEPARATOR
            m_MII.dwTypeData = StrPtr(m_Caption)
            m_MII.cch = Len(m_Caption)
        End If
        m_MII.fMask = m_MII.fMask Or MIIM_FTYPE Or MIIM_STRING
        Me.Refresh
    End If
End Property
Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property
Public Property Let Checked(ByVal NewValue As Boolean)
    If m_hWnd = 0 Then
        m_Checked = NewValue
        If m_Checked Then
            m_MII.fState = m_MII.fState Or MFS_CHECKED And Not MFS_UNCHECKED
        Else
            m_MII.fState = m_MII.fState Or MFS_UNCHECKED And Not MFS_CHECKED
        End If
        m_MII.fMask = m_MII.fMask Or MIIM_STATE
        Me.Refresh
    End If
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
    If m_hWnd = 0 Then
        m_Enabled = NewValue
        If m_Enabled Then
            m_MII.fState = m_MII.fState And Not MFS_DISABLED
        Else
            m_MII.fState = m_MII.fState Or MFS_DISABLED
        End If
        m_MII.fMask = m_MII.fMask Or MIIM_STATE
        Me.Refresh
    End If
End Property
Public Property Get HelpContextID() As Long
    HelpContextID = m_HelpContextID
End Property
Public Property Let HelpContextID(ByVal NewValue As Long)
    m_HelpContextID = NewValue
End Property
Friend Property Get hWnd() As Long
    hWnd = m_hWnd
    If m_hWnd = 0 Then Set m_MenuItems = Nothing
End Property
Friend Property Let hWnd(ByVal NewValue As Long)
    If m_MenuParent Is Nothing Then
        m_hWnd = NewValue
        If m_hWnd Then
            Set m_MenuItems = New Collection
            m_Key = vbNullChar
        End If
    End If
End Property
Friend Property Get ID() As Long
    ID = m_ID
End Property
Friend Property Let ID(ByVal NewValue As Long)
    m_ID = NewValue
End Property
Friend Sub Init(ByRef Key As String, ByRef Caption As String, ByVal Checked As Boolean, ByVal Enabled As Boolean, ByVal HelpContextID As Long, ByVal RadioGroup As Byte, ByVal RightToLeft As Boolean, ByVal WindowList As Boolean)
    If m_hWnd = 0 Then
        m_Key = Key
        m_Caption = Caption
        m_Checked = Checked
        m_Enabled = Enabled
        m_HelpContextID = HelpContextID
        m_RadioGroup = RadioGroup
        m_RightToLeft = RightToLeft
        m_WindowList = WindowList
    End If
End Sub
Public Property Get Key() As String
    If m_hWnd = 0 Then Key = m_Key
End Property
Public Property Let Key(ByRef NewValue As String)
    If m_hWnd = 0 And Not m_MenuParent Is Nothing Then
        If m_MenuParent.KeyReplace(m_Key, NewValue) Then
            m_Key = NewValue
        End If
    End If
End Property
Friend Function KeyExists(ByRef Key As String) As Boolean
    If m_hWnd Then
        ' simple InStr to see whether key has been stored
        KeyExists = InStr(m_Key, vbNullChar & Key & vbNullChar) > 0
    End If
End Function
Friend Function KeyReplace(ByRef OldKey As String, ByRef NewKey As String) As Boolean
    If m_hWnd Then
        ' make sure old key exists
        If InStr(m_Key, vbNullChar & OldKey, vbNullChar) > 0 Then
            ' make sure new key does not exist
            If InStr(m_Key, vbNullChar & NewKey, vbNullChar) < 1 Then
                ' replace the key
                m_Key = Replace(m_Key, vbNullChar & OldKey & vbNullChar, vbNullChar & NewKey & vbNullChar)
                KeyReplace = True
            End If
        End If
    End If
End Function
Friend Function MenuItem(ByVal Index As Long) As UniMenuItem
    If m_hWnd Then
        If Index >= 1 And Index <= m_MenuItems.Count Then
            Set MenuItem = m_MenuItems(Index)
        End If
    End If
End Function
Friend Function MenuItemAdd(ByRef Key As String, ByRef Caption As String, ByVal Checked As Boolean, ByVal Enabled As Boolean, ByVal HelpContextID As Long, ByVal RadioGroup As Byte, ByVal WindowList As Boolean, ByVal Visible As Boolean) As UniMenuItem
    Dim MenuItem As UniMenuItem, KeyIndex As Long
    If m_hWnd Then
        If LenB(Key) = 0 Or InStr(Key, vbNullChar) > 0 Then
            Do While InStr(m_Key, vbNullChar & "Item" & CStr(KeyIndex) & vbNullChar) > 0
                KeyIndex = KeyIndex + 1
            Loop
            Key = "Item" & CStr(KeyIndex)
        End If
        Set MenuItem = New UniMenuItem
        m_MenuItems.Add MenuItem, Key
        MenuItem.Init Key, Caption, Checked, Enabled, HelpContextID, RadioGroup, m_RightToLeft, WindowList
        Set MenuItemAdd = MenuItem
        Set MenuItem = Nothing
    End If
End Function
Friend Function MenuItemByKey(ByRef Key As String) As UniMenuItem
    If m_hWnd Then
        If InStr(m_Key, vbNullChar & Key & vbNullChar) > 0 Then
            Set MenuItemByKey = m_MenuItems(Key)
        End If
    End If
End Function
Friend Sub MenuParent(ByRef MenuItem As UniMenuItem)
    If Not MenuItem Is Nothing Then
        If MenuItem.hWnd Then Set m_MenuParent = MenuItem
    Else
        Set m_MenuParent = Nothing
    End If
End Sub
Public Property Get RadioGroup() As Byte
    RadioGroup = m_RadioGroup
End Property
Public Property Let RadioGroup(ByVal NewValue As Byte)
    If m_hWnd = 0 Then
        m_RadioGroup = NewValue
        If m_RadioGroup > 0 Then
            m_MII.fType = m_MII.fType Or MFT_RADIOCHECK
        Else
            m_MII.fType = m_MII.fType And Not MFT_RADIOCHECK
        End If
        m_MII.fMask = m_MII.fMask Or MIIM_STATE
        Me.Refresh
    End If
End Property
Friend Sub Refresh()
    ' check if can update
    If m_hWnd = 0 And Not m_MenuParent Is Nothing And m_MII.fMask <> 0 Then
        ' update
        SetMenuItemInfoW m_MenuParent.hWnd, m_ID, 0&, m_MII
        ' prevent further updates unless something has been changed
        m_MII.fMask = 0
    End If
End Sub
Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property
Friend Property Let RightToLeft(ByVal NewValue As Boolean)
    m_RightToLeft = NewValue
    If NewValue Then
        m_MII.fType = m_MII.fType Or MFT_RIGHTORDER
    Else
        m_MII.fType = m_MII.fType And Not MFT_RIGHTORDER
    End If
End Property
Friend Sub SetRightToLeft(ByVal NewValue As Boolean)
    Dim CurrentItem As UniMenuItem
    If m_hWnd Then
        For Each CurrentItem In m_MenuItems
            CurrentItem.RightToLeft = NewValue
        Next CurrentItem
    End If
End Sub
Friend Sub SetWindowList(ByRef MenuItem As UniMenuItem)
    Dim CurrentItem As UniMenuItem
    If m_hWnd Then
        For Each CurrentItem In m_MenuItems
            If MenuItem <> CurrentItem Then CurrentItem.WindowList = False
        Next CurrentItem
    End If
End Sub
Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property
Public Property Let Visible(ByVal NewValue As Boolean)
    m_Visible = NewValue
End Property
Public Property Get WindowList() As Boolean
    WindowList = m_WindowList
End Property
Public Property Let WindowList(ByVal NewValue As Boolean)
    If Not m_MenuParent Is Nothing Then
        If NewValue Then
            m_MenuParent.SetWindowList Me
            m_WindowList = True
            ' SET WINDOWLIST STYLE
        ElseIf m_WindowList Then
            ' REMOVE WINDOWLIST STYLE
            m_WindowList = False
        End If
    End If
End Property
Private Sub Class_Initialize()
    m_MII.cbSize = LenB(m_MII)
End Sub
Private Sub Class_Terminate()
    If m_hWnd Then
        Do While m_MenuItems.Count
            m_MenuItems.Remove 1
        Loop
    Else
        Set m_MenuParent = Nothing
    End If
End Sub
