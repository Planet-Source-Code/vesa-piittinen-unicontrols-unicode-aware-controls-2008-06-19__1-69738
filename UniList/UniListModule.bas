Attribute VB_Name = "UniListModule"
'*************************************************************************************************
'* UniListModule.bas - IOLEInPlaceActiveObject handler for UniList
'* ---------------------------------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* REQUIREMENTS
'* ------------
'* Note: TLBs are compiled to your program so you don't need to distribute the files
'* - OleGuids3.tlb      = Ole Guid and interface definitions 3.0
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.0.0 (2008-06-11)
'* - Customized and simplified version for UniList control
'*
'* mIOleInPlaceActivate.bas (1999-01-09)
'* - Author: Mike Gainer, Matt Curland and Bill Storage
'* - WWW: http://vbaccelerator.com
'*
'*************************************************************************************************
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type UniList_IPAOHook
    lpVTable As Long
    IPAOReal As IOleInPlaceActiveObject
    Ctl As UniList
    ThisPointer As Long
End Type

Private Const S_FALSE As Long = 1
Private Const S_OK As Long = 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As GUID, iid2 As GUID) As Long

Private IID_IOleInPlaceActiveObject As GUID
Private m_IPAOVTable(9) As Long

Private Function AddOf(ByVal AddressOfProcedure As Long) As Long
    AddOf = AddressOfProcedure
End Function

Private Function IPAO_AddRef(This As UniList_IPAOHook) As Long
    IPAO_AddRef = This.IPAOReal.AddRef
End Function

Private Function IPAO_ContextSensitiveHelp(This As UniList_IPAOHook, ByVal fEnterMode As Long) As Long
    IPAO_ContextSensitiveHelp = This.IPAOReal.ContextSensitiveHelp(fEnterMode)
End Function

Private Function IPAO_EnableModeless(This As UniList_IPAOHook, ByVal fEnable As Long) As Long
    IPAO_EnableModeless = This.IPAOReal.EnableModeless(fEnable)
End Function

Private Function IPAO_GetWindow(This As UniList_IPAOHook, phwnd As Long) As Long
    IPAO_GetWindow = This.IPAOReal.GetWindow(phwnd)
End Function

Private Function IPAO_OnDocWindowActivate(This As UniList_IPAOHook, ByVal fActivate As Long) As Long
    IPAO_OnDocWindowActivate = This.IPAOReal.OnDocWindowActivate(fActivate)
End Function

Private Function IPAO_OnFrameWindowActivate(This As UniList_IPAOHook, ByVal fActivate As Long) As Long
    IPAO_OnFrameWindowActivate = This.IPAOReal.OnFrameWindowActivate(fActivate)
End Function

Private Function IPAO_QueryInterface(This As UniList_IPAOHook, riid As GUID, pvObj As Long) As Long
    If IsEqualGUID(riid, IID_IOleInPlaceActiveObject) Then
        pvObj = This.ThisPointer
        IPAO_AddRef This
        IPAO_QueryInterface = 0
    Else
        IPAO_QueryInterface = This.IPAOReal.QueryInterface(ByVal VarPtr(riid), pvObj)
    End If
End Function

Private Function IPAO_Release(This As UniList_IPAOHook) As Long
    IPAO_Release = This.IPAOReal.Release
End Function

Private Function IPAO_ResizeBorder(This As UniList_IPAOHook, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    IPAO_ResizeBorder = This.IPAOReal.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function

Private Function IPAO_TranslateAccelerator(This As UniList_IPAOHook, lpMsg As MSG) As Long
    Dim CtlText As UniList
    If TypeOf This.Ctl Is UniList Then
        Set CtlText = This.Ctl
        If CtlText.TranslateAccel(lpMsg) Then IPAO_TranslateAccelerator = S_OK: Exit Function
    End If
    IPAO_TranslateAccelerator = This.IPAOReal.TranslateAccelerator(ByVal VarPtr(lpMsg))
End Function

Public Sub UniList_Init(UniList_IPAOHook As UniList_IPAOHook, Ctl As UniList)
    Dim IPAO As IOleInPlaceActiveObject
    If m_IPAOVTable(0) = 0 Then
        m_IPAOVTable(0) = AddOf(AddressOf IPAO_QueryInterface)
        m_IPAOVTable(1) = AddOf(AddressOf IPAO_AddRef)
        m_IPAOVTable(2) = AddOf(AddressOf IPAO_Release)
        m_IPAOVTable(3) = AddOf(AddressOf IPAO_GetWindow)
        m_IPAOVTable(4) = AddOf(AddressOf IPAO_ContextSensitiveHelp)
        m_IPAOVTable(5) = AddOf(AddressOf IPAO_TranslateAccelerator)
        m_IPAOVTable(6) = AddOf(AddressOf IPAO_OnFrameWindowActivate)
        m_IPAOVTable(7) = AddOf(AddressOf IPAO_OnDocWindowActivate)
        m_IPAOVTable(8) = AddOf(AddressOf IPAO_ResizeBorder)
        m_IPAOVTable(9) = AddOf(AddressOf IPAO_EnableModeless)
        With IID_IOleInPlaceActiveObject
           .Data1 = &H117&
           .Data4(0) = &HC0
           .Data4(7) = &H46
        End With
    End If
    With UniList_IPAOHook
        Set IPAO = Ctl
        CopyMemory .IPAOReal, IPAO, 4
        CopyMemory .Ctl, Ctl, 4
        .lpVTable = VarPtr(m_IPAOVTable(0))
        .ThisPointer = VarPtr(UniList_IPAOHook)
    End With
End Sub
Public Sub UniList_Terminate(UniList_IPAOHook As UniList_IPAOHook)
    With UniList_IPAOHook
        CopyMemory .IPAOReal, 0&, 4
        CopyMemory .Ctl, 0&, 4
    End With
End Sub
