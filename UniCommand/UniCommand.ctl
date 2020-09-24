VERSION 5.00
Begin VB.UserControl UniCommand 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ToolboxBitmap   =   "UniCommand.ctx":0000
End
Attribute VB_Name = "UniCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniCommand 1.4.1 - Unicode command button user control
'* ------------------------------------------------------
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
'* No special requirements.
'*
'* HOW TO ADD TO YOUR PROGRAM
'* --------------------------
'* 1) Copy UniCommand.ctl and UniCommand.ctx to your project folder.
'* 2) In your project, add UniCommand.ctl.
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.4.1 (2008-06-19)
'* - Font improvements
'*
'* Version 1.4 (2007-12-08)
'* - some minor fixes
'*
'* Version 1.2 (2005-10-14)
'* - fixed ScaleMode problems with MDIForms
'* - fixes to AutoSize behavior
'* - added the forgotten MouseIcon and MousePointer
'*
'* Version 1.1 (2005-07-05)
'* - added access keys and keyboard behaviour (Enter, Esc)
'* - added default border indicator (set color by FillColor)
'*
'* Version 1.0 (2004-09-23)
'* - initial release
'*************************************************************************************************
Option Explicit

'constants for API calls
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_NOCLIP As Long = &H100
Private Const DT_RIGHT As Long = &H2
Private Const DT_WORDBREAK As Long = &H10

'constants for borderstyle
Private Const BDR_NONE As Long = &H0
Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_SUNKENINNER As Long = &H8
Private Const EDGE_RAISED As Long = &H5
Private Const EDGE_ETCHED As Long = &H6
Private Const EDGE_BUMP As Long = &H9
Private Const EDGE_SUNKEN As Long = &HA

Private Const BF_RECT = &HF

'custom types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'API declarations
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRECT As RECT) As Long
Private Declare Function DrawTextANSI Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, lpRECT As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextUnicode Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpArrPtr As Long, ByVal nCount As Long, lpRECT As RECT, ByVal wFormat As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'public enums
Public Enum uniBorderStyle
    ucbNone = BDR_NONE
    ucbRaised = EDGE_RAISED
    ucbSunken = EDGE_SUNKEN
    ucbRaisedOuter = BDR_RAISEDOUTER
    ucbRaisedInner = BDR_RAISEDINNER
    ucbSunkenOuter = BDR_SUNKENOUTER
    ucbSunkenInner = BDR_SUNKENINNER
    ucbBump = EDGE_BUMP
    ucbEtched = EDGE_ETCHED
End Enum

Public Enum uniClickDepth
    ucb0x0
    ucb0x1
    ucb1x0
    ucb1x1
    ucb1x2
    ucb2x1
    ucb2x2
End Enum

'events
Public Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event Change()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'default public properties
Private Const m_def_Alignment As Byte = vbCenter
Private Const m_def_AutoRedraw As Boolean = False
Private Const m_def_AutoSize As Boolean = False
Private Const m_def_BorderDown As Byte = ucbSunken
Private Const m_def_BorderUp As Byte = ucbRaised
Private Const m_def_ClickDepth As Byte = ucb1x1
Private Const m_def_ShowFocus As Boolean = False
Private Const m_def_WordWrap As Boolean = False

'public properties
Private m_Alignment As AlignmentConstants
Private m_AutoRedraw As Boolean
Private m_AutoSize As Boolean
Private m_BackColor As Long
Private m_BorderDown As uniBorderStyle
Private m_BorderUp As uniBorderStyle
Private m_CaptionB() As Byte
Private m_CaptionLen As Long
Private m_ClickDepth As uniClickDepth
Private m_FillColor As Long
Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Private m_ForeColor As Long
Private m_ShowFocus As Boolean
Private m_WordWrap As Boolean

'helper variables
Private m_BorderStyle As uniBorderStyle
Private m_Caption As String 'only used under Windows 95/98/98SE/ME
Private m_ClickX As Byte
Private m_ClickY As Byte
Private m_DEFAULTRECT As RECT
Private m_DTMODE As Long
Private m_ExtenderScale As ScaleModeConstants
Private m_FOCUSRECT As RECT
Private m_FULLRECT As RECT
Private m_HasFocus As Boolean
Private m_RECT As RECT
Private m_WINNT As Boolean
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property
Public Property Let Alignment(ByVal NewAlignment As AlignmentConstants)
    m_Alignment = NewAlignment
    'repaint
    UpdateRect
End Property
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = m_AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal NewMode As Boolean)
    Dim OldMode As Boolean, EmptyImage As IPictureDisp
    OldMode = m_AutoRedraw
    m_AutoRedraw = NewMode
    UserControl.AutoRedraw = NewMode
    'when autoredraw mode changes, old content is set as picture: clear with an empty image
    If NewMode = False And OldMode = True Then UserControl.Picture = EmptyImage
    UpdateRect
End Property
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property
Public Property Let AutoSize(ByVal NewMode As Boolean)
    m_AutoSize = NewMode
    'repaint
    UpdateRect
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    m_BackColor = NewColor
    'change backcolor
    UserControl.BackColor = m_BackColor
    'repaint
    UserControl_Paint
End Property
Public Property Get BorderDown() As uniBorderStyle
    BorderDown = m_BorderDown
End Property
Public Property Let BorderDown(ByVal NewMode As uniBorderStyle)
    If NewMode = m_BorderDown Then Exit Property
    'change border
    m_BorderDown = NewMode
    'repaint
    UserControl_Paint
End Property
Public Property Get BorderUp() As uniBorderStyle
    BorderUp = m_BorderUp
End Property
Public Property Let BorderUp(ByVal NewMode As uniBorderStyle)
    If NewMode = m_BorderUp Then Exit Property
    'change border
    m_BorderUp = NewMode
    'repaint
    UserControl_Paint
End Property
Public Property Get Caption() As String
    Caption = m_CaptionB
End Property
Public Property Let Caption(ByVal NewCaption As String)
    'check for null string
    If LenB(NewCaption) > 0 Then
        'non-null string
        m_CaptionB = NewCaption
        m_CaptionLen = (UBound(m_CaptionB) + 1) \ 2
    Else
        'null string
        Erase m_CaptionB
        m_CaptionLen = 0
    End If
    SetAccessKeys
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
End Property
'only called under Windows 95/98/98SE/ME
Private Sub CaptionChange()
    'check if the array is empty or not
    If (Not m_CaptionB) <> True Then
        m_Caption = m_CaptionB
    Else
        m_Caption = vbNullString
    End If
End Sub
Public Property Get CaptionLen() As Long
    'return the length of the text
    CaptionLen = m_CaptionLen
End Property
Public Property Let CaptionLen(ByVal NewLength As Long)
    Dim NewUBound As Long
    'check for invalid length
    If NewLength < 0 Then NewLength = 0
    'if somebody is really doing something this silly...
    If NewLength > &H3FFFFFFF Then NewLength = &H3FFFFFFF
    'and of course a check for this, no need to update!
    If m_CaptionLen = NewLength Then Exit Property
    'change it
    m_CaptionLen = NewLength
    If NewLength = 0 Then
        'null string
        Erase m_CaptionB
        'clear control
        UserControl.BackColor = m_BackColor
    Else
        'new array size
        NewUBound = NewLength + NewLength - 1
        'change byte array / string size
        ReDim Preserve m_CaptionB(NewUBound)
    End If
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
End Property
Public Property Get CaptionAscB(ByVal Index As Long) As Byte
    'no out of bounds checking...
    CaptionAscB = m_CaptionB(Index + Index)
End Property
Public Property Let CaptionAscB(ByVal Index As Long, ByVal NewCode As Byte)
    'we have no out of bounds checking here...
    m_CaptionB(Index + Index) = NewCode
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
End Property
Public Property Get CaptionAscW(ByVal Index As Long) As Integer
    Dim CurIndex As Long
    'check we are not out of bounds
    If Index < 0 Or Index > m_CaptionLen - 1 Then Exit Property
    'very minor speed optimization
    CurIndex = Index + Index
    'is the highest bit active?
    If (m_CaptionB(CurIndex + 1) And &H80) = 0 Then
        'not active
        'convert two bytes into an integer
        CaptionAscW = m_CaptionB(CurIndex) Or (CInt(m_CaptionB(CurIndex + 1)) * &H100)
    Else
        'active
        'convert two bytes into an integer and mark highest bit active
        CaptionAscW = m_CaptionB(CurIndex) Or (CInt(m_CaptionB(CurIndex + 1) And &H7F) * &H100) Or &H8000
    End If
End Property
Public Property Let CaptionAscW(ByVal Index As Long, ByVal NewCode As Integer)
    Dim Byte1 As Byte, Byte2 As Byte, CurIndex As Long
    'check we are not out of bounds
    If Index < 0 Or Index > m_CaptionLen - 1 Then Exit Property
    'rip lower byte
    Byte1 = CByte(NewCode And &HFF)
    'rip higher byte: check if the highest bit is active
    If NewCode < 0 Then
        'highest bit active
        Byte2 = ((NewCode And &H7F00) \ &H100) Or &H80
    Else
        'highest bit not active
        Byte2 = (NewCode And &H7F00) \ &H100
    End If
    'very minor speed optimization
    CurIndex = Index + Index
    'update data in array
    m_CaptionB(CurIndex) = Byte1
    m_CaptionB(CurIndex + 1) = Byte2
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
End Property
Public Property Get ClickDepth() As uniClickDepth
    ClickDepth = m_ClickDepth
End Property
Public Property Let ClickDepth(ByVal NewMode As uniClickDepth)
    If NewMode = m_ClickDepth Then Exit Property
    'change click depth
    m_ClickDepth = NewMode
    'change X and Y according to new mode
    Select Case NewMode
        Case ucb0x1
            m_ClickX = 0
            m_ClickY = 1
        Case ucb1x0
            m_ClickX = 1
            m_ClickY = 0
        Case ucb1x1
            m_ClickX = 1
            m_ClickY = 1
        Case ucb1x2
            m_ClickX = 1
            m_ClickY = 2
        Case ucb2x1
            m_ClickX = 2
            m_ClickY = 1
        Case ucb2x2
            m_ClickX = 2
            m_ClickY = 2
        Case Else
            m_ClickX = 0
            m_ClickY = 0
    End Select
    'repaint
    UserControl_Paint
End Property
Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColor
End Property
Public Property Let FillColor(ByVal NewColor As OLE_COLOR)
    m_FillColor = NewColor
    UserControl.FillColor = m_FillColor
    'repaint
    UserControl_Paint
End Property
Public Property Get Font() As Font
    Set Font = m_Font
End Property
Public Property Set Font(ByVal NewValue As Font)
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
    Set UserControl.Font = NewFont
    m_Font_FontChanged vbNullString
    'repaint
    UpdateRect
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = m_Font.Bold
End Property
Public Property Let FontBold(ByVal NewValue As Boolean)
    m_Font.Bold = NewValue
    If Ambient.UserMode Then Else PropertyChanged "Font"
    UpdateRect
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = m_Font.Italic
End Property
Public Property Let FontItalic(ByVal NewValue As Boolean)
    m_Font.Italic = NewValue
    If Ambient.UserMode Then Else PropertyChanged "Font"
    UpdateRect
End Property
Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = m_Font.Name
End Property
Public Property Let FontName(ByRef NewValue As String)
    m_Font.Name = NewValue
    If Ambient.UserMode Then Else PropertyChanged "Font"
    UpdateRect
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = m_Font.Size
End Property
Public Property Let FontSize(ByVal NewValue As Single)
    m_Font.Size = NewValue
    If Ambient.UserMode Then Else PropertyChanged "Font"
    UpdateRect
End Property
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = m_Font.Strikethrough
End Property
Public Property Let FontStrikethru(ByVal NewValue As Boolean)
    m_Font.Strikethrough = NewValue
    If Ambient.UserMode Then Else PropertyChanged "Font"
    UpdateRect
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = m_Font.Underline
End Property
Public Property Let FontUnderline(ByVal NewValue As Boolean)
    m_Font.Underline = NewValue
    If Ambient.UserMode Then Else PropertyChanged "Font"
    UpdateRect
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    m_ForeColor = NewColor
    UserControl.ForeColor = m_ForeColor
    'repaint
    UserControl_Paint
End Property
Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByRef NewIcon As IPictureDisp)
    Set UserControl.MouseIcon = NewIcon
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal NewPointer As MousePointerConstants)
    UserControl.MousePointer = NewPointer
End Property
Private Sub SetAccessKeys()
    Dim A As Long, TempStr As String * 1
    AccessKeys = vbNullString
    If (Not m_CaptionB) = True Then Exit Sub
    'accesskey
    For A = 0 To UBound(m_CaptionB) - 3 Step 2
        If m_CaptionB(A) = 38 And m_CaptionB(A + 1) = 0 Then
            TempStr = ChrW$(m_CaptionB(A + 2) Or (m_CaptionB(A + 3) * &H100))
            If AscW(TempStr) <> 38 Then
                AccessKeys = TempStr
                Exit For
            End If
        End If
    Next A
End Sub
Public Function GetCaptionB() As Byte()
    GetCaptionB = m_CaptionB
End Function
Public Sub SetCaptionB(ByRef NewCaption() As Byte)
    'check if the array is empty
    If (Not NewCaption) <> True Then
        'array with data
        m_CaptionB = NewCaption
        m_CaptionLen = (UBound(m_CaptionB) + 1) \ 2
    Else
        'empty array
        Erase m_CaptionB
        m_CaptionLen = 0
    End If
    SetAccessKeys
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
End Sub
Public Property Get ShowFocus() As Boolean
    ShowFocus = m_ShowFocus
End Property
Public Property Let ShowFocus(ByVal NewFocus As Boolean)
    If NewFocus = m_ShowFocus Then Exit Property
    'change focus status
    m_ShowFocus = NewFocus
    'repaint if control currently has focus
    If m_HasFocus Then UserControl_Paint
End Property
Private Sub UpdateRect()
    Dim TempRect As RECT
    Static HereAlready As Boolean
    'check if this sub is running already
    If HereAlready Then Exit Sub
    'mark we are running this sub
    HereAlready = True
    'alignment for painting
    Select Case Alignment
        Case vbLeftJustify
            'paint left justified
            m_DTMODE = DT_LEFT
        Case vbCenter
            'paint centered
            m_DTMODE = DT_CENTER
        Case vbRightJustify
            'paint right justified
            m_DTMODE = DT_RIGHT
    End Select
    'autosize mode?
    If Not m_AutoSize Then
        'no autosize, use control width and height as the painting area
        With m_RECT
            .Top = 3
            .Left = 3
            .Bottom = UserControl.ScaleHeight - 3
            .Right = UserControl.ScaleWidth - 3
        End With
        'set wordwrapping settings
        If m_WordWrap Then
            'paint wordwrapped
            m_DTMODE = m_DTMODE Or DT_WORDBREAK
            TempRect.Right = UserControl.ScaleWidth
        End If
        If (Not m_CaptionB) <> True Then
            'get text width and height
            If m_WINNT Then
                DrawTextUnicode UserControl.hDC, VarPtr(m_CaptionB(0)), m_CaptionLen, TempRect, DT_CALCRECT Or m_DTMODE
            Else
                DrawTextANSI UserControl.hDC, VarPtr(m_CaptionB(0)), m_CaptionLen, TempRect, DT_CALCRECT Or m_DTMODE
            End If
            'center text vertically
            With m_RECT
                .Top = Int((m_FULLRECT.Bottom - TempRect.Bottom) / 2)
                .Bottom = .Bottom - .Top + 3
            End With
        End If
    Else
        With m_RECT
            'reset all of these
            .Top = 3
            .Left = 3
            .Bottom = 3
            .Right = UserControl.ScaleWidth - 3
            If m_WINNT Then 'UNICODE
                'set wordwrapping settings
                If m_WordWrap Then
                    'paint wordwrapped
                    m_DTMODE = m_DTMODE Or DT_WORDBREAK
                    'get paint area height
                    DrawTextUnicode UserControl.hDC, VarPtr(m_CaptionB(0)), m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control height
                    UserControl.Height = UserControl.ScaleY(.Bottom + 6, vbPixels, m_ExtenderScale)
                    'set paint area width (correct height is returned in drawtext
                    .Right = UserControl.ScaleWidth - 3
                Else
                    'no wordwrapping
                    m_DTMODE = m_DTMODE Or DT_NOCLIP
                    'get paint area width and height
                    DrawTextUnicode UserControl.hDC, VarPtr(m_CaptionB(0)), m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control width to the same as painting area width
                    UserControl.Width = UserControl.ScaleX(.Right + 6, vbPixels, m_ExtenderScale)
                    'set control height to the same as painting area height
                    UserControl.Height = UserControl.ScaleY(.Bottom + 6, vbPixels, m_ExtenderScale)
                    With m_RECT
                        .Left = 5
                        .Top = 4
                    End With
                End If
            Else 'ANSI
                'set wordwrapping settings
                If m_WordWrap Then
                    'paint wordwrapped
                    m_DTMODE = m_DTMODE Or DT_WORDBREAK
                    'get paint area height
                    DrawTextANSI UserControl.hDC, m_Caption, m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control height
                    UserControl.Height = UserControl.ScaleY(.Bottom + 6, vbPixels, m_ExtenderScale)
                    'set paint area width (correct height is returned in drawtext
                    .Right = UserControl.ScaleWidth - 3
                Else
                    'no wordwrapping
                    m_DTMODE = m_DTMODE Or DT_NOCLIP
                    'get paint area width and height
                    DrawTextANSI UserControl.hDC, m_CaptionB, m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control width to the same as painting area width
                    UserControl.Width = UserControl.ScaleX(.Right + 6, vbPixels, m_ExtenderScale)
                    'set control height to the same as painting area height
                    UserControl.Height = UserControl.ScaleY(.Bottom + 6, vbPixels, m_ExtenderScale)
                    With m_RECT
                        .Left = 5
                        .Top = 4
                    End With
                End If
            End If
        End With
    End If
    'mark we are done
    HereAlready = False
    'repaint
    UserControl_Paint
End Sub
Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property
Public Property Let WordWrap(ByVal NewMode As Boolean)
    m_WordWrap = NewMode
    UpdateRect
End Property
Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    UpdateRect
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub
Private Sub UserControl_AmbientChanged(PropertyName As String)
    Select Case PropertyName
        Case "DisplayAsDefault"
            UserControl_Paint
    End Select
End Sub
Private Sub UserControl_DblClick()
    'mouse button pressed down
    m_BorderStyle = m_BorderDown
    UserControl_Paint
    RaiseEvent DblClick
End Sub
Private Sub UserControl_ExitFocus()
    If m_BorderStyle = m_BorderDown Then
        m_BorderStyle = m_BorderUp
        UserControl_Paint
    End If
End Sub
Private Sub UserControl_GotFocus()
    If m_ShowFocus Then
        m_HasFocus = True
        UserControl_Paint
    End If
End Sub
Private Sub UserControl_Initialize()
    m_WINNT = (Environ$("OS") = "Windows_NT")
    On Error Resume Next
    m_ExtenderScale = UserControl.Extender.Container.ScaleMode
    If Err.Number Then
        Err.Clear
        m_ExtenderScale = UserControl.Parent.ScaleMode
        If Err.Number Then
            Err.Clear
            m_ExtenderScale = vbTwips
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub UserControl_InitProperties()
    'get default settings
    m_Alignment = m_def_Alignment
    m_AutoRedraw = m_def_AutoRedraw
    m_AutoSize = m_def_AutoSize
    m_BackColor = UserControl.Extender.Container.BackColor
    m_BorderDown = m_def_BorderDown
    m_BorderUp = m_def_BorderUp
    m_CaptionB = UserControl.Ambient.DisplayName
    m_CaptionLen = Len(UserControl.Name)
    m_ClickDepth = m_def_ClickDepth
    m_FillColor = UserControl.Extender.Container.FillColor
    Set m_Font = UserControl.Ambient.Font
    Set UserControl.Font = m_Font
    m_ForeColor = UserControl.Extender.Container.ForeColor
    m_ShowFocus = m_def_ShowFocus
    m_WordWrap = m_def_WordWrap
    
    m_BorderStyle = m_BorderUp
    With m_RECT
        .Top = 0
        .Left = 0
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    'set full button area size
    With m_FULLRECT
        .Left = 0
        .Top = 0
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    'change click depth X and Y according to the mode
    Select Case m_def_ClickDepth
        Case ucb0x1
            m_ClickX = 0
            m_ClickY = 1
        Case ucb1x0
            m_ClickX = 1
            m_ClickY = 0
        Case ucb1x1
            m_ClickX = 1
            m_ClickY = 1
        Case ucb1x2
            m_ClickX = 1
            m_ClickY = 2
        Case ucb2x1
            m_ClickX = 2
            m_ClickY = 1
        Case ucb2x2
            m_ClickX = 2
            m_ClickY = 2
        Case Else
            m_ClickX = 0
            m_ClickY = 0
    End Select
    
    SetAccessKeys
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        'mouse button pressed down
        m_BorderStyle = m_BorderDown
        UserControl_Paint
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        'mouse button released
        m_BorderStyle = m_BorderUp
        UserControl_Paint
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_LostFocus()
    If m_ShowFocus Then m_HasFocus = False
    UserControl_Paint
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        'mouse button pressed down
        m_BorderStyle = m_BorderDown
        UserControl_Paint
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TestCondition As Boolean
    If Button = vbLeftButton Then
        'check if we are within the area
        TestCondition = (X < 0 Or Y < 0 Or X >= UserControl.ScaleWidth Or Y >= UserControl.ScaleHeight)
        If TestCondition And (m_BorderStyle = m_BorderDown) Then
            'not in the area and button is down
            'make button go up
            m_BorderStyle = m_BorderUp
            UserControl_Paint
        ElseIf (Not TestCondition) And (m_BorderStyle = m_BorderUp) Then
            'in the area and button is up
            'make button go down
            m_BorderStyle = m_BorderDown
            UserControl_Paint
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        'mouse button released
        m_BorderStyle = m_BorderUp
        UserControl_Paint
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If (X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight) And (Button = vbLeftButton) Then RaiseEvent Click
End Sub
Private Sub UserControl_Paint()
    Static HereAlready As Boolean
    'if we wanted to prevent blinking we'd set AutoRedraw = True
    'and do nothing here and do painting in UpdateRect
    
    If HereAlready Then Exit Sub
    'check we are not running this sub already to prevent running this forever
    HereAlready = True
    'clear before redraw
    UserControl.Cls
    'check if not a null string and also that the colors differ
    If m_CaptionLen > 0 And m_BackColor <> m_ForeColor Then
        'button pushed?
        If m_BorderStyle = m_BorderDown Then
            With m_RECT
                .Bottom = .Bottom + m_ClickY
                .Left = .Left + m_ClickX
                .Right = .Right + m_ClickX
                .Top = .Top + m_ClickY
            End With
        End If
        'check OS
        If m_WINNT Then
            'Windows NT/2000/XP
            DrawTextUnicode UserControl.hDC, VarPtr(m_CaptionB(0)), m_CaptionLen, m_RECT, m_DTMODE
        Else
            'Windows 95/98/98SE/ME (no Unicode support)
            DrawTextANSI UserControl.hDC, m_Caption, m_CaptionLen, m_RECT, m_DTMODE
        End If
        'button pushed?
        If m_BorderStyle = m_BorderDown Then
            With m_RECT
                .Bottom = .Bottom - m_ClickY
                .Left = .Left - m_ClickX
                .Right = .Right - m_ClickX
                .Top = .Top - m_ClickY
            End With
        End If
    End If
    'draw button edge
    If m_BorderStyle <> BDR_NONE Then
        If Ambient.DisplayAsDefault Then
            DrawEdge UserControl.hDC, m_DEFAULTRECT, m_BorderStyle, BF_RECT
            UserControl.ForeColor = m_FillColor
            Rectangle UserControl.hDC, 0, 0, 1, m_FULLRECT.Bottom
            Rectangle UserControl.hDC, 1, 0, m_FULLRECT.Right, 1
            Rectangle UserControl.hDC, m_DEFAULTRECT.Right, 1, m_FULLRECT.Right, m_FULLRECT.Bottom
            Rectangle UserControl.hDC, 1, m_DEFAULTRECT.Bottom, m_DEFAULTRECT.Right, m_FULLRECT.Bottom
            UserControl.ForeColor = m_ForeColor
        Else
            DrawEdge UserControl.hDC, m_FULLRECT, m_BorderStyle, BF_RECT
        End If
    End If
    If m_ShowFocus And m_HasFocus Then DrawFocusRect UserControl.hDC, m_FOCUSRECT
    'make change visible (this is the main reason we use HereAlready)
    If m_AutoRedraw Then UserControl.Refresh
    'mark we are done
    HereAlready = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'get all saved properties
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_BackColor = PropBag.ReadProperty("BackColor", UserControl.Extender.Container.BackColor)
    m_BorderDown = PropBag.ReadProperty("BorderDown", m_def_BorderDown)
    m_BorderUp = PropBag.ReadProperty("BorderUp", m_def_BorderUp)
    m_CaptionB = PropBag.ReadProperty("CaptionB", UserControl.Ambient.DisplayName)
    m_CaptionLen = PropBag.ReadProperty("CaptionLen", (UBound(m_CaptionB) + 1) \ 2)
    m_ClickDepth = PropBag.ReadProperty("ClickDepth", m_def_ClickDepth)
    m_FillColor = PropBag.ReadProperty("FillColor", vbWindowText)
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    Set m_Font = UserControl.Font
    m_ForeColor = PropBag.ReadProperty("ForeColor", UserControl.Ambient.ForeColor)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    m_ShowFocus = PropBag.ReadProperty("ShowFocus", m_def_ShowFocus)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
    m_BorderStyle = m_BorderUp
    'use the settings
    With UserControl
        .AutoRedraw = m_AutoRedraw
        .BackColor = m_BackColor
        .ForeColor = m_ForeColor
    End With
    'change click depth X and Y according to the mode
    Select Case m_ClickDepth
        Case ucb0x1
            m_ClickX = 0
            m_ClickY = 1
        Case ucb1x0
            m_ClickX = 1
            m_ClickY = 0
        Case ucb1x1
            m_ClickX = 1
            m_ClickY = 1
        Case ucb1x2
            m_ClickX = 1
            m_ClickY = 2
        Case ucb2x1
            m_ClickX = 2
            m_ClickY = 1
        Case ucb2x2
            m_ClickX = 2
            m_ClickY = 2
        Case Else
            m_ClickX = 0
            m_ClickY = 0
    End Select
    'initial draw
    UpdateRect
    SetAccessKeys
End Sub
Private Sub UserControl_Resize()
    'set full button area size
    With m_FULLRECT
        .Left = 0
        .Top = 0
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    With m_DEFAULTRECT
        .Left = 1
        .Top = 1
        .Bottom = UserControl.ScaleHeight - 1
        .Right = UserControl.ScaleWidth - 1
    End With
    With m_FOCUSRECT
        .Left = 3
        .Top = 3
        .Bottom = m_FULLRECT.Bottom - 3
        .Right = m_FULLRECT.Right - 3
    End With
    'refresh
    UpdateRect
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save all properties
    PropBag.WriteProperty "Alignment", m_Alignment, m_def_Alignment
    PropBag.WriteProperty "AutoRedraw", m_AutoRedraw, m_def_AutoRedraw
    PropBag.WriteProperty "AutoSize", m_AutoSize, m_def_AutoSize
    PropBag.WriteProperty "BackColor", m_BackColor, UserControl.Extender.Container.BackColor
    PropBag.WriteProperty "BorderDown", m_BorderDown, m_def_BorderDown
    PropBag.WriteProperty "BorderUp", m_BorderUp, m_def_BorderUp
    PropBag.WriteProperty "CaptionB", m_CaptionB, UserControl.Name
    PropBag.WriteProperty "CaptionLen", m_CaptionLen, Len(UserControl.Name)
    PropBag.WriteProperty "ClickDepth", m_ClickDepth, m_def_ClickDepth
    PropBag.WriteProperty "FillColor", m_FillColor, vbWindowText
    PropBag.WriteProperty "Font", m_Font, UserControl.Ambient.Font
    PropBag.WriteProperty "ForeColor", m_ForeColor, UserControl.Ambient.ForeColor
    PropBag.WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
    PropBag.WriteProperty "ShowFocus", m_ShowFocus, m_def_ShowFocus
    PropBag.WriteProperty "WordWrap", m_WordWrap, m_def_WordWrap
End Sub
