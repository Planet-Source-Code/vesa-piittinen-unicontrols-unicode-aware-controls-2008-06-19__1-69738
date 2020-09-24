VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UniLabel testing"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   StartUpPosition =   1  'CenterOwner
   Begin UniLabelTest.UniLabel lblEvents 
      Height          =   345
      Left            =   7320
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":000C
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":005E
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   0   'False
      UseMnemonic     =   0   'False
      WordWrap        =   -1  'True
   End
   Begin UniLabelTest.UniLabel lblFeatures 
      Height          =   345
      Left            =   4680
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":007A
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":00D4
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   0   'False
      UseMnemonic     =   0   'False
      WordWrap        =   -1  'True
   End
   Begin UniLabelTest.UniLabel UniLabel1 
      Height          =   345
      Index           =   4
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   609
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":00F0
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":016C
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   0   'False
      WordWrap        =   0   'False
   End
   Begin UniLabelTest.UniLabel UniLabel1 
      Height          =   435
      Index           =   3
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   767
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":0188
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":01CE
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   0   'False
      WordWrap        =   0   'False
   End
   Begin UniLabelTest.UniLabel UniLabel1 
      Height          =   345
      Index           =   2
      Left            =   4680
      TabIndex        =   1
      Top             =   480
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   609
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":01EA
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":0260
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   0   'False
      WordWrap        =   0   'False
   End
   Begin UniLabelTest.UniLabel UniLabel1 
      Height          =   345
      Index           =   1
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   609
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":027C
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":02EA
      MousePointer    =   0
      PaddingBottom   =   5
      PaddingLeft     =   5
      PaddingRight    =   5
      PaddingTop      =   5
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   0   'False
      WordWrap        =   0   'False
   End
   Begin UniLabelTest.UniLabel UniLabel1 
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   529
      Alignment       =   0
      AutoSize        =   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      Caption         =   "frmTest.frx":0306
      DesignTimeSafe  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MouseIcon       =   "frmTest.frx":033A
      MousePointer    =   99
      PaddingBottom   =   2
      PaddingLeft     =   20
      PaddingRight    =   2
      PaddingTop      =   2
      RightToLeft     =   0   'False
      UseEvents       =   -1  'True
      UseMnemonic     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' UnicodeStrings.res and sample code below by Zhu Jin Yong
    Dim lngA As Long
    Dim strText As String
    For lngA = 102 To 116
        strText = strText & LoadResString(lngA) & "|" & LoadResString(lngA + 20) & IIf(lngA < 116, vbNewLine, vbNullString)
    Next lngA
    ' change the text
    UniLabel1(0).Caption = strText
    ' resize form to label height
    Me.Height = Me.Height - ScaleY(Me.ScaleHeight, Me.ScaleMode, vbTwips) + ScaleY(UniLabel1(0).Top * 2 + UniLabel1(0).Height, Me.ScaleMode, vbTwips)
    
    ' the below tries to apply Japanese on Windows 9X/ME
    UniLabel1(3).Caption = ChrW$(&H3042) & ChrW$(&H3044) & ChrW$(&H3046) & ChrW$(&H3048) & ChrW$(&H304A)
    ' set to Japanese character set
    UniLabel1(3).Font.Charset = 128
    ' try using MS Gothic
    On Error Resume Next: UniLabel1(3).Font.Name = "MS Gothic": On Error GoTo 0
    ' For Windows 9X/ME testers:
    ' The only difference to normal is that you can see Japanese
    ' even if your Windows locale setting is not Japanese.
    ' This, of course, applies to all languages. You need the font installed for this to work.
    ' And the end result is not real Unicode, you are dependant of the character set.
    
    lblFeatures.Caption = "Properties: " & _
        vbNewLine & "- Alignment" & _
        vbNewLine & "- AutoSize" & _
        vbNewLine & "- BackColor" & _
        vbNewLine & "- BackStyle (not working)" & _
        vbNewLine & "- Caption" & _
        vbNewLine & "- Font" & _
        vbNewLine & "- ForeColor" & _
        vbNewLine & "- MouseIcon" & _
        vbNewLine & "- MousePointer" & _
        vbNewLine & "- PaddingBottom" & _
        vbNewLine & "- PaddingLeft" & _
        vbNewLine & "- PaddingRight" & _
        vbNewLine & "- PaddingTop" & _
        vbNewLine & "- RightToLeft" & _
        vbNewLine & "- UseEvents" & _
        vbNewLine & "- UseMnemonic" & _
        vbNewLine & "- WordWrap"
        
    lblEvents.Caption = "Events: " & _
        vbNewLine & "- Change" & _
        vbNewLine & "- Click" & _
        vbNewLine & "- DblClick" & _
        vbNewLine & "- MouseDown" & _
        vbNewLine & "- MouseEnter" & _
        vbNewLine & "- MouseLeave" & _
        vbNewLine & "- MouseMove" & _
        vbNewLine & "- MouseUp"
End Sub

Private Sub UniLabel1_Click(Index As Integer, Button As UniLabelMouseButtonConstants)
    Me.Caption = "Clicked!"
    Select Case Button
        Case vbLeftButton
            Debug.Print "Left click!"
        Case vbRightButton
            Debug.Print "Right click!"
        Case vbMiddleButton
            Debug.Print "Middle click!"
    End Select
End Sub
Private Sub UniLabel1_DblClick(Index As Integer, Button As UniLabelMouseButtonConstants)
    Me.Caption = "Double clicked!"
End Sub
Private Sub UniLabel1_MouseDown(Index As Integer, Button As UniLabelMouseButtonConstants, Shift As UniLabelShiftConstants, X As Single, Y As Single)
    UniLabel1(Index).BackColor = vbHighlight
    UniLabel1(Index).ForeColor = vbHighlightText
End Sub
Private Sub UniLabel1_MouseEnter(Index As Integer)
    With UniLabel1(Index)
        .BackColor = vbWhite
        .ForeColor = vbRed
    End With
End Sub
Private Sub UniLabel1_MouseLeave(Index As Integer)
    With UniLabel1(Index)
        .BackColor = Me.BackColor
        .ForeColor = Me.ForeColor
    End With
End Sub
Private Sub UniLabel1_MouseMove(Index As Integer, Button As UniLabelMouseButtonConstants, Shift As UniLabelShiftConstants, X As Single, Y As Single)
    Me.Caption = X & " x " & Y
End Sub
Private Sub UniLabel1_MouseUp(Index As Integer, Button As UniLabelMouseButtonConstants, Shift As UniLabelShiftConstants, X As Single, Y As Single)
    UniLabel1(Index).BackColor = vbHighlightText
    UniLabel1(Index).ForeColor = vbHighlight
End Sub
