VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSizeGrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CSizeGrip.cls
'
' Implements themed (uxtheme.dll) size grip.
'   - Uses legacy code when theming not enabled.
'   - Supports Bottom Left Orientation
'

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Const TS_MIN As Long = 0
Private Const TS_TRUE As Long = 1
Private Const TS_DRAW As Long = 2
Private Const DFC_SCROLL As Long = 3
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT As Long = &H10
Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3
Private Const SPB_CLASS As String = "Scrollbar"
Private Const SPB_SIZEBOX As Long = 10
Private Const SZB_RIGHTALIGN As Long = 1
Private Const SZB_LEFTALIGN As Long = 2
Private Const S_OK As Long = 0
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTBOTTOMRIGHT As Long = 17
Private Const HTBOTTOMLEFT As Long = 16

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, ByRef lprc As RECT, ByVal uType As Long, ByVal uState As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function GetThemePartSize Lib "uxtheme" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal prc As Long, ByVal eSize As Long, psz As SIZE) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long

Public Enum EGripOrientation
    gripBottomRight
    gripBottomLeft
End Enum

Private WithEvents m_Parent As Form
Attribute m_Parent.VB_VarHelpID = -1
Private WithEvents m_Grip As Label
Attribute m_Grip.VB_VarHelpID = -1
Private m_Orientation As EGripOrientation
'

Public Property Let Orientation(Value As EGripOrientation)
    m_Orientation = Value
    If m_Grip Is Nothing Then Exit Property
    m_Grip.MousePointer = vbSizeNWSE
    If m_Orientation = gripBottomLeft Then m_Grip.MousePointer = vbSizeNESW
    m_Parent_Resize
End Property

Public Property Get Orientation() As EGripOrientation
    Orientation = m_Orientation
End Property

Public Sub Attach(Parent As Form, Optional GripOrientation As EGripOrientation = gripBottomRight)
    Dim hTheme As Long
    Dim ControlSize As SIZE
    
    m_Orientation = GripOrientation
    Set m_Parent = Parent
    If m_Grip Is Nothing Then Set m_Grip = m_Parent.Controls.Add("VB.Label", "lblGrip" & CStr(Orientation) & m_Parent.Name)
    With m_Grip
        .Caption = vbNullString
        .BackStyle = vbTransparent
        .Visible = True
        .MousePointer = vbSizeNWSE
        If m_Orientation = gripBottomLeft Then .MousePointer = vbSizeNESW
    End With
    
    hTheme = OpenThemeData(m_Parent.hWnd, StrPtr(SPB_CLASS))
    If hTheme Then
        GetThemePartSize hTheme, m_Parent.hDC, SPB_SIZEBOX, SZB_RIGHTALIGN + m_Orientation, 0, TS_TRUE, ControlSize
        m_Grip.Width = ControlSize.cx
        m_Grip.Height = ControlSize.cy
        CloseThemeData hTheme
    Else
        m_Grip.Width = GetSystemMetrics(SM_CXVSCROLL)
        m_Grip.Height = GetSystemMetrics(SM_CYHSCROLL)
    End If
End Sub

Private Sub Class_Terminate()
    Set m_Parent = Nothing
    Set m_Grip = Nothing
End Sub

Private Sub m_Grip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage m_Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT - m_Orientation, 0
End Sub

Private Sub m_Parent_Paint()
    Dim hTheme As Long
    Dim ControlRect As RECT
    
    If m_Orientation = gripBottomRight Then ControlRect.Left = m_Parent.ScaleWidth - m_Grip.Width
    ControlRect.Top = m_Grip.Top
    ControlRect.Right = ControlRect.Left + m_Grip.Width
    ControlRect.Bottom = m_Grip.Top + m_Grip.Height
    
    hTheme = OpenThemeData(m_Parent.hWnd, StrPtr(SPB_CLASS))
    If hTheme Then
        DrawThemeBackground hTheme, m_Parent.hDC, SPB_SIZEBOX, SZB_RIGHTALIGN + m_Orientation, ControlRect, ControlRect
        CloseThemeData hTheme
    Else
        Select Case m_Orientation
        Case gripBottomRight:   DrawFrameControl m_Parent.hDC, ControlRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP
        Case gripBottomLeft:    DrawFrameControl m_Parent.hDC, ControlRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIPRIGHT
        End Select
    End If
End Sub

Private Sub m_Parent_Resize()
    Dim Left As Single
    If m_Orientation = gripBottomRight Then Left = m_Parent.ScaleWidth - m_Grip.Width
    m_Grip.Move Left, m_Parent.ScaleHeight - m_Grip.Height
End Sub
