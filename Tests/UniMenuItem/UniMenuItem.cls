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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private m_Caption As String
Private m_ChildCount As Long
Private m_Childs() As UniMenuItem
Private m_Key As String
Private m_Keys As String
Private m_Level As Long
Private m_Parent As UniMenuItem
Private m_Root As UniMenuItem

Public Function Add(Optional ByRef Caption As String = "-", Optional ByVal Index As Long = -1) As UniMenuItem
    ' add new item to the end
    ReDim Preserve m_Childs(m_ChildCount)
    ' where we insert item?
    If Index < 0 Or Index >= m_ChildCount Then
        ' add to the end
        Index = m_ChildCount
    Else
        ' move the existing items
        CopyMemory ByVal VarPtr(m_Childs(Index + 1)), ByVal VarPtr(m_Childs(Index)), (m_ChildCount - Index) * 4& + 4&
        ' clean up the duplicate
        CopyMemory ByVal VarPtr(m_Childs(Index)), 0&, 4&
    End If
    ' increase number of children
    m_ChildCount = m_ChildCount + 1
    ' create the new item
    Set m_Childs(Index) = New UniMenuItem
    With m_Childs(Index)
        ' set tree deepness
        .Level = m_Level + 1
        ' reference the root
        Set .Root = m_Root
        ' reference the parent
        Set .Parent = Me
        ' set caption
        .Caption = Caption
    End With
    ' return the new item
    Set Add = m_Childs(Index)
End Function
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByRef NewValue As String)
    If Not Parent Is Nothing Then
        m_Caption = NewValue
        If Not Parent Is Nothing Then
            Debug.Print "I am " & NewValue & " and my parent is " & Parent.Caption & "."
        Else
            Debug.Print "I am " & NewValue & " and I have no parent."
        End If
    Else
        Debug.Print "I am the root element."
    End If
End Property
Public Property Get Child(ID) As UniMenuItem
    If VarType(ID) = vbString Then
        
    ElseIf IsNumeric(ID) Then
        If ID >= 0 And ID < m_ChildCount Then
            Set Child = m_Childs(ID)
        End If
    End If
End Property
Public Property Get Level() As Long
    Level = m_Level
End Property
Friend Property Let Level(ByVal NewValue As Long)
    m_Level = NewValue
End Property
Public Property Get Parent() As UniMenuItem
    ' here is one of the interesting features: if we have no parent, we must be a root element
    If Not m_Parent Is Nothing Then
        ' return parent
        Set Parent = m_Parent
    Else
        ' return ourselves, we are the parent!
        Set Parent = Me
    End If
End Property
Friend Property Set Parent(ByRef NewValue As UniMenuItem)
    Set m_Parent = NewValue
End Property
Public Function Remove(ByVal Index As Long) As Boolean
    ' validate range
    If Index >= 0 And Index < m_ChildCount Then
        ' decrease children count
        m_ChildCount = m_ChildCount - 1
        ' remove the item
        Set m_Childs(Index) = Nothing
        If m_ChildCount > 0 Then
            ' was it in the end of the array?
            If Index < m_ChildCount Then
                ' nope, we need to move stuff
                CopyMemory ByVal VarPtr(m_Childs(Index)), ByVal VarPtr(m_Childs(Index + 1)), (m_ChildCount - Index) * 4& + 4&
                ' erase the last item that will be removed
                CopyMemory ByVal VarPtr(m_Childs(m_ChildCount)), 0&, 4&
            End If
            ' free the item
            ReDim Preserve m_Childs(m_ChildCount - 1)
        Else
            ' it was the last item
            Erase m_Childs
        End If
    End If
End Function
Public Property Get Root() As UniMenuItem
    Set Root = m_Root
End Property
Friend Property Set Root(ByRef NewValue As UniMenuItem)
    Set m_Root = NewValue
End Property
Private Sub Class_Initialize()
    m_Keys = vbNullChar
    ' always start up as a root
    Set m_Root = Me
End Sub
Private Sub Class_Terminate()
    Dim lngA As Long
    ' clean up all references
    Set m_Root = Nothing
    Set m_Parent = Nothing
    For lngA = 0 To m_ChildCount - 1
        Set m_Childs(lngA) = Nothing
    Next lngA
    Erase m_Childs
End Sub
