VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3075
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Pleh"
      Begin VB.Menu mnuTest4 
         Caption         =   "Pleh4"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTest2 
         Caption         =   "Pleh2"
      End
      Begin VB.Menu mnuTest3 
         Caption         =   "Pleh3"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oIconUp As IPictureDisp, oIconDown As IPictureDisp
 
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

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

Private Declare Function GetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long

Private Sub Command1_Click()
    Dim Control As Control
    For Each Control In Controls
        If TypeOf Control Is Menu Then
            On Error Resume Next
            Control.Visible = Not Control.Visible
            On Error GoTo 0
        End If
    Next Control
End Sub

Private Sub Form_Load()
    Dim Control As Control
    For Each Control In Controls
        If TypeOf Control Is Menu Then
            Debug.Print Control.Name
        End If
    Next Control
    Debug.Print FormMenuSetIcon(Me, 0&, 0&, Me.Picture, Me.Picture)
    Debug.Print FormMenuSetIcon(Me, 0&, 1&, Me.Picture, Me.Picture)
    
End Sub
Function FormMenuSetIcon(frmMenu As Form, lSubItem As Long, lMenuItem As Long, oPicChecked As IPictureDisp, oPicUnchecked As IPictureDisp) As Boolean
    Const MF_BYCOMMAND = &H0&
    Dim MII As MENUITEMINFO, strBuffer As String
 
    Dim lhWndMenu As Long, lhWndSubMenu As Long, lhWndMenuItem As Long
 
    'Get the menu handle
    lhWndMenu = GetMenu(frmMenu.hWnd)
    'Get the handle of the submenu
    lhWndSubMenu = GetSubMenu(lhWndMenu, lSubItem)
    'Get the handle of the menu item
    lhWndMenuItem = GetMenuItemID(lhWndSubMenu, lMenuItem)
   
    If lhWndMenuItem Then
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_STRING
        If GetMenuItemInfoW(lhWndMenu, lhWndMenuItem, 0&, MII) Then
            strBuffer = Space$(MII.cch)
            MII.dwTypeData = StrPtr(strBuffer)
            MII.cch = MII.cch + 1
            If GetMenuItemInfoW(lhWndMenu, lhWndMenuItem, 0&, MII) Then
                Debug.Print """" & strBuffer & """"
                MII.cch = 3
                strBuffer = ChrW$(&H3041) & ChrW$(&H3043)
                MII.dwTypeData = StrPtr(strBuffer)
                If SetMenuItemInfoW(lhWndMenu, lhWndMenuItem, 0&, MII) Then
                
                Else
                    Debug.Print "SetMenuItemInfoW failed"
                End If
            Else
                Debug.Print "GetMenuItemInfoW second phase failed"
            End If
        Else
            Debug.Print "GetMenuItemInfoW failed"
        End If
        'The form has a sub menu, add the picture/icon
        FormMenuSetIcon = CBool(SetMenuItemBitmaps(lhWndMenu, lhWndMenuItem, MF_BYCOMMAND, oPicChecked.Handle, oPicUnchecked.Handle))
        If FormMenuSetIcon Then
            'Added successfully, repaint the menu
            Call DrawMenuBar(lhWndMenu)
        End If
    End If
End Function
