VERSION 5.00
Begin VB.Form UniCaption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UniCaption; hit run to see"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UniCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniCaption - Unicode Form Caption
'* ---------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*************************************************************************************************
Option Explicit

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long

Private Const GWL_WNDPROC = -4

Private m_Caption As String

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
Private Sub Form_Load()
    ' some hiragana (you need Japanese fonts installed to see them)
    CaptionW = ChrW$(&H3042) & ChrW$(&H3044) & ChrW$(&H3046) & ChrW$(&H3048) & ChrW$(&H304A) & " ovat japanilaisia hiragana-merkkejä."
End Sub
