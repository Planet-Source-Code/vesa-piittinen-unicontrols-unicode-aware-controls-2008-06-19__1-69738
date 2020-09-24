VERSION 5.00
Begin VB.UserControl SelfControl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "SelfControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' SelfIPAO requirements

' structure that is passed to our custom IOleInPlaceActivate procedures
Private Type IOleHook
    VTablePtr As Long
    IPAO As IOleInPlaceActiveObject
    Ctrl As SelfControl                  ' <- CHANGE "SelfControl" TO THE CONTROL NAME
    Pointer As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' structure for convenience (no need to Dim so much)
Private Type IOle
    Hook As IOleHook                        ' Hook (passed in procedures)
    IID As GUID                             ' GUID
    OLE As IOleObject                       ' reference to user control
    VTable(9) As Long                       ' VTable replacement to custom OLE procedures
    InPlaceSite As IOleInPlaceSite          ' user control's InPlaceSite
    InPlaceFrame As IOleInPlaceFrame        ' user control's InPlaceFrame
    InPlaceUIWindow As IOleInPlaceUIWindow  ' user control's InPlaceUIWindow
    Pos As RECT                             ' Position information
    Clip As RECT                            ' Clipping information
    FrameInfo As OLEINPLACEFRAMEINFO        ' Frame information
End Type

Private Type IKey
    CaptureEnter As Boolean
    CaptureEsc As Boolean
    CaptureNavigation As Boolean
    CaptureTab As Boolean
    Locked As Boolean
End Type

Private m_Key As IKey
Private m_OLE As IOle
Private m_SelfCallbackCode() As Long        ' base machine code without procedure specific data

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsEqualGUID Lib "ole32" (IID1 As GUID, IID2 As GUID) As Long


' SelfCallback requirements

Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_RWX As Long = &H40&

Private Const IDX_CALLBACKORDINAL As Long = 36
' memory bytes required for the callback thunk
Private Const MEM_LEN As Long = IDX_CALLBACKORDINAL * 4 + 4
' thunk data index of the Owner object's vTable address
Private Const INDX_OWNER As Long = 0
' thunk data index of the callback procedure address
Private Const INDX_CALLBACK As Long = 1
' thunk data index of the EbMode function address
Private Const INDX_EBMODE As Long = 2
' thunk data index of the IsBadCodePtr function address
Private Const INDX_BADPTR As Long = 3
' thunk data index of the KillTimer function address
Private Const INDX_KT As Long = 4
' thunk code patch index of the thunk data
Private Const INDX_EBX As Long = 6
' thunk code patch index of the number of parameters expected in callback
Private Const INDX_PARAMS As Long = 18
' thunk code patch index of the bytes to be released after callback
Private Const INDX_PARAMLEN As Long = 24
' thunk offset to the callback execution address
Private Const PROC_OFF As Long = &H14

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Function ScAddressOfOrdinal(ByVal Ordinal As Long) As Long
    Dim bytValue As Byte, bytSignature As Byte, lngA As Long, lngAddress As Long
    ' get address of this class module instance
    RtlMoveMemory VarPtr(lngAddress), ObjPtr(Me), 4&
    ' probe for UserControl method
    If ScProbe(lngAddress + &H7A4&, lngA, bytSignature) Then
        ' scan up to 256 vTable entries
        For lngA = lngA + 4 To lngA + 1024 Step 4
            ' get address in vTable
            RtlMoveMemory VarPtr(lngAddress), lngA, 4&
            ' invalid code address?
            If IsBadCodePtr(lngAddress) Then
                ' return this vTable address
                RtlMoveMemory VarPtr(ScAddressOfOrdinal), lngA - (Ordinal * 4&), 4&
                Exit For
            End If
            ' get byte pointed to vTable address
            RtlMoveMemory VarPtr(bytValue), lngAddress, 1&
            ' if does not match the expected value...
            If bytValue <> bytSignature Then
                ' return this vTable address
                RtlMoveMemory VarPtr(ScAddressOfOrdinal), lngA - (Ordinal * 4&), 4&
                Exit For
            End If
        Next lngA
    End If
End Function
Private Property Get ScData(ByVal Index As Long, ByVal ThunkPtr As Long) As Long
    RtlMoveMemory VarPtr(ScData), ThunkPtr + (Index * 4&), 4&
End Property
Private Property Let ScData(ByVal Index As Long, ByVal ThunkPtr As Long, ByVal NewValue As Long)
    RtlMoveMemory ThunkPtr + (Index * 4&), VarPtr(NewValue), 4&
End Property
Private Function ScProbe(ByVal Address As Long, ByRef Method As Long, ByRef Signature As Byte) As Boolean
    Dim bytValue As Byte, lngVTableEntry As Long
    ' probe eight entries
    For Address = Address To Address + 32 Step 4
        ' get vTable entry
        RtlMoveMemory VarPtr(lngVTableEntry), Address, 4&
        ' if not an implemented interface
        If lngVTableEntry Then
            ' get the value pointed at by the vTable entry
            RtlMoveMemory VarPtr(bytValue), lngVTableEntry, 1&
            ' if native or P-code signature...
            If (bytValue = &H33) Or (bytValue = &HE9) Then
                ' return this information
                Method = Address
                Signature = bytValue
                ' success, exit loop
                ScProbe = True
                Exit For
            End If
        End If
    Next Address
End Function
Private Function ScProcedureAddress(ByVal DynamicLinkLibrary As String, ByVal Procedure As String, ByVal Unicode As Boolean) As Long
    ' get the procedure address
    If Unicode Then
        ScProcedureAddress = GetProcAddress(GetModuleHandleW(StrPtr(DynamicLinkLibrary)), Procedure)
    Else
        ScProcedureAddress = GetProcAddress(GetModuleHandleA(DynamicLinkLibrary), Procedure)
    End If
    ' in IDE, verify we got it
    Debug.Assert ScProcedureAddress
End Function
Private Sub SiFocus()
    With m_OLE
        Set .InPlaceSite = m_OLE.OLE.GetClientSite
        If Not .InPlaceSite Is Nothing Then
            .InPlaceSite.GetWindowContext .InPlaceFrame, .InPlaceUIWindow, VarPtr(.Pos), VarPtr(.Clip), VarPtr(.FrameInfo)
            If .InPlaceFrame Is Nothing Then
                .InPlaceFrame.SetActiveObject .Hook.Pointer, vbNullString
                ' clear reference
                Set .InPlaceFrame = Nothing
            End If
            If Not .InPlaceUIWindow Is Nothing Then
                .InPlaceUIWindow.SetActiveObject .Hook.Pointer, vbNullString
                ' clear reference
                Set .InPlaceUIWindow = Nothing
            Else
                m_OLE.OLE.DoVerb OLEIVERB_UIACTIVATE, 0&, .InPlaceSite, 0&, UserControl.hWnd, VarPtr(.Pos)
            End If
            ' clear reference
            Set .InPlaceSite = Nothing
        End If
    End With
End Sub
Private Sub SiInit()
    Dim objIPAO As IOleInPlaceActiveObject
    ' initialize machine code for SelfCallback
    ReDim m_SelfCallbackCode(0 To IDX_CALLBACKORDINAL) As Long
    ' create the base machine code array
    m_SelfCallbackCode(5) = &HBB60E089
    m_SelfCallbackCode(7) = &H73FFC589
    m_SelfCallbackCode(8) = &HC53FF04
    m_SelfCallbackCode(9) = &H59E80A74
    m_SelfCallbackCode(10) = &HE9000000
    m_SelfCallbackCode(11) = &H30&
    m_SelfCallbackCode(12) = &H87B81
    m_SelfCallbackCode(13) = &H75000000
    m_SelfCallbackCode(14) = &H9090902B
    m_SelfCallbackCode(15) = &H42DE889
    m_SelfCallbackCode(16) = &H50000000
    m_SelfCallbackCode(17) = &HB9909090
    m_SelfCallbackCode(19) = &H90900AE3
    m_SelfCallbackCode(20) = &H8D74FF
    m_SelfCallbackCode(21) = &H9090FAE2
    m_SelfCallbackCode(22) = &H53FF33FF
    m_SelfCallbackCode(23) = &H90909004
    m_SelfCallbackCode(24) = &H2BADC261
    m_SelfCallbackCode(25) = &H3D0853FF
    m_SelfCallbackCode(26) = &H1&
    m_SelfCallbackCode(27) = &H23DCE74
    m_SelfCallbackCode(28) = &H74000000
    m_SelfCallbackCode(29) = &HAE807
    m_SelfCallbackCode(30) = &H90900000
    m_SelfCallbackCode(31) = &H4589C031
    m_SelfCallbackCode(32) = &H90DDEBFC
    m_SelfCallbackCode(33) = &HFF0C75FF
    m_SelfCallbackCode(34) = &H53FF0475
    m_SelfCallbackCode(35) = &HC310&
    m_SelfCallbackCode(INDX_BADPTR) = ScProcedureAddress("kernel32", "IsBadCodePtr", False)
    m_SelfCallbackCode(INDX_OWNER) = ObjPtr(Me)
    ' IDE safety
    If App.LogMode = 0 Then
        ' store the EbMode function address in the thunk data
        m_SelfCallbackCode(INDX_EBMODE) = ScProcedureAddress("vba6", "EbMode", False)
    End If
    ' set OLE object reference
    Set m_OLE.OLE = Me
    ' create callback procedures for VTable
    With m_OLE
        .VTable(0) = SiVTableInit(10, 3)    ' QueryInterface
        .VTable(1) = SiVTableInit(9, 1)     ' AddRef
        .VTable(2) = SiVTableInit(8, 1)     ' Release
        .VTable(3) = SiVTableInit(7, 2)     ' GetWindow
        .VTable(4) = SiVTableInit(6, 2)     ' ContextSensitiveHelp
        .VTable(5) = SiVTableInit(5, 2)     ' TranslateAccelerator
        .VTable(6) = SiVTableInit(4, 2)     ' OnFrameWindowActivate
        .VTable(7) = SiVTableInit(3, 2)     ' OnDocWindowActivate
        .VTable(8) = SiVTableInit(2, 4)     ' ResizeBorder
        .VTable(9) = SiVTableInit(1, 2)     ' EnableModeless
    End With
    ' init GUID
    With m_OLE.IID
        .Data1 = &H117&
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    ' set InPlaceActiveObject reference
    Set objIPAO = Me
    ' now create the interface...
    With m_OLE.Hook
        .VTablePtr = VarPtr(m_OLE.VTable(0))
        CopyMemory .IPAO, objIPAO, 4&
        CopyMemory .Ctrl, Me, 4&
        .Pointer = VarPtr(m_OLE.Hook)
    End With
    Debug.Print "Init"
End Sub
Private Sub SiTerminate()
    If Not m_OLE.OLE Is Nothing Then
        ' remove the interface
        With m_OLE.Hook
            CopyMemory .IPAO, 0&, 4&
            CopyMemory .Ctrl, 0&, 4&
        End With
        ' free VTable entries
        VirtualFree m_OLE.VTable(0), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(1), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(2), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(3), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(4), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(5), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(6), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(7), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(8), 0&, MEM_RELEASE
        VirtualFree m_OLE.VTable(9), 0&, MEM_RELEASE
        ' remove reference
        Set m_OLE.OLE = Nothing
        Debug.Print "Terminate"
    End If
End Sub
Private Function SiVTableInit(Optional ByVal Ordinal As Long = 1, Optional ByVal ParamCount = 1) As Long
    Dim lngCallback As Long, lngCallbackCode() As Long, lngScMem As Long
    ' get address of the procedure
    lngCallback = ScAddressOfOrdinal(Ordinal)
    ' verify we got it
    If lngCallback Then
        ' allocate executable memory
        lngScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
        ' verify we got it
        If lngScMem Then
            ' allocate and create the machine code array
            lngCallbackCode = m_SelfCallbackCode
            ' set callback address
            lngCallbackCode(INDX_CALLBACK) = lngCallback
            ' remember the ordinal used
            lngCallbackCode(IDX_CALLBACKORDINAL) = Ordinal
            ' parameter count
            lngCallbackCode(INDX_PARAMS) = ParamCount
            RtlMoveMemory VarPtr(lngCallbackCode(INDX_PARAMLEN)) + 2&, VarPtr(ParamCount * 4&), 2&
            ' special for timer callback:
            lngCallbackCode(INDX_KT) = ScProcedureAddress("user32", "KillTimer", False)
            ' set the data address relative to virtual memory pointer
            lngCallbackCode(INDX_EBX) = lngScMem
            ' copy thunk code to executable memory
            RtlMoveMemory lngScMem, VarPtr(lngCallbackCode(INDX_OWNER)), MEM_LEN
            ' return the procedure address
            SiVTableInit = lngScMem + PROC_OFF
            Debug.Print SiVTableInit
        End If
    End If
End Function

Private Sub Text1_GotFocus()
    SiFocus
End Sub

Private Sub UserControl_GotFocus()
    UserControl.BackColor = vbHighlight
    SiFocus
End Sub
Private Sub UserControl_Initialize()
    m_Key.CaptureEnter = True
    m_Key.CaptureEsc = True
    m_Key.CaptureNavigation = True
    m_Key.CaptureTab = True
    m_Key.Locked = False
End Sub
Private Sub UserControl_InitProperties()
    If Ambient.UserMode Then SiInit
End Sub
Private Sub UserControl_LostFocus()
    UserControl.BackColor = vbWindowBackground
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then SiInit
End Sub
Private Sub UserControl_Terminate()
    SiTerminate
End Sub

' ordinal #10
Private Function SiQueryInterface(This As IOleHook, rIID As GUID, pvObj As Long) As Long
    ' Install the interface if required
    If IsEqualGUID(rIID, m_OLE.IID) Then
        ' Install alternative IOleInPlaceActiveObject interface implemented here
        pvObj = This.Pointer
        SiAddRef This
        SiQueryInterface = 0
    Else
        ' Use the default support for the interface:
        SiQueryInterface = This.IPAO.QueryInterface(ByVal VarPtr(rIID), pvObj)
    End If
End Function
' ordinal #9
Private Function SiAddRef(This As IOleHook) As Long
    SiAddRef = This.IPAO.AddRef
End Function
' ordinal #8
Private Function SiRelease(This As IOleHook) As Long
    SiRelease = This.IPAO.Release
End Function
' ordinal #7
Private Function SiGetWindow(This As IOleHook, phwnd As Long) As Long
    SiGetWindow = This.IPAO.GetWindow(phwnd)
End Function
' ordinal #6
Private Function SiContextSensitiveHelp(This As IOleHook, ByVal fEnterMode As Long) As Long
    SiContextSensitiveHelp = This.IPAO.ContextSensitiveHelp(fEnterMode)
End Function
' ordinal #5
Private Function SiTranslateAccelerator(This As IOleHook, lpMsg As MSG) As Long
    Dim objControlSite As IOleControlSite
    Debug.Print "Translate"
    Select Case lpMsg.message
        Case WM_KEYDOWN, WM_KEYUP
            Select Case lpMsg.wParam
                Case vbKeyTab
                    If m_Key.CaptureTab Then
                        Debug.Print "Tab"
                        SiTranslateAccelerator = 1&
                    End If
                Case vbKeyReturn
                    If m_Key.CaptureEnter Then
                        
                        SiTranslateAccelerator = 1&
                    End If
                Case vbKeyEscape
                    If m_Key.CaptureEsc Then
                        
                        SiTranslateAccelerator = 1&
                    End If
                Case vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyUp, _
                    vbKeyEnd, vbKeyHome, vbKeyPageDown, vbKeyPageUp
                    If m_Key.CaptureNavigation Then
                        
                        SiTranslateAccelerator = 1&
                    End If
                Case Else
                    If m_Key.Locked Then SiTranslateAccelerator = 1&
            End Select
        Case WM_CHAR
            If m_Key.Locked Then SiTranslateAccelerator = 1&
    End Select
    ' pass to standard user control method?
    If SiTranslateAccelerator = 0& Then SiTranslateAccelerator = This.IPAO.TranslateAccelerator(ByVal VarPtr(lpMsg))
End Function
' ordinal #4
Private Function SiOnFrameWindowActivate(This As IOleHook, ByVal fActivate As Long) As Long
    SiOnFrameWindowActivate = This.IPAO.OnFrameWindowActivate(fActivate)
End Function
' ordinal #3
Private Function SiOnDocWindowActivate(This As IOleHook, ByVal fActivate As Long) As Long
    SiOnDocWindowActivate = This.IPAO.OnDocWindowActivate(fActivate)
End Function
' ordinal #2
Private Function SiResizeBorder(This As IOleHook, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    SiResizeBorder = This.IPAO.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function
' ordinal #1
Private Function SiEnableModeless(This As IOleHook, ByVal fEnable As Long) As Long
    SiEnableModeless = This.IPAO.EnableModeless(fEnable)
End Function
