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
' Implements a themed (uxtheme.dll) size grip.
'   - Uses DrawFrameControl when theming not enabled.
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

Private Enum THEMESIZE
    TS_MIN
    TS_TRUE
    TS_DRAW
End Enum

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const S_OK As Long = 0
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTBOTTOMRIGHT As Long = 17
Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3
Private Const DFC_SCROLL As Long = 3
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8
Private Const SPB_CLASS As String = "Scrollbar"
Private Const SPB_SIZEBOX As Long = 10
Private Const SZB_RIGHTALIGN As Long = 1
Private Const SZB_HALFBOTTOMRIGHTALIGN As Long = 5
 
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetVersionExW Lib "kernel32" (lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, ByRef lprc As RECT, ByVal uType As Long, ByVal uState As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function GetThemePartSize Lib "uxtheme" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal prc As Long, ByVal eSize As THEMESIZE, psz As SIZE) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long

Private WithEvents m_Parent As Form
Attribute m_Parent.VB_VarHelpID = -1
Private WithEvents m_Grip As Image
Attribute m_Grip.VB_VarHelpID = -1
'

Public Sub Attach(Parent As Form)
    Set m_Parent = Parent
    m_Parent.ScaleMode = vbPixels

    Set m_Grip = m_Parent.Controls.Add("VB.Image", "imgSizeGrip" & m_Parent.Controls.Count)
    With m_Grip
        .Width = GetSystemMetrics(SM_CXVSCROLL)
        .Height = GetSystemMetrics(SM_CYHSCROLL)
        .MousePointer = vbSizeNWSE
        .Visible = True
    End With
    
    Dim ThemeHandle As Long
    ThemeHandle = OpenThemeData(m_Parent.hWnd, StrPtr(SPB_CLASS))
    If ThemeHandle Then
        Dim GripSize As SIZE
        If GetThemePartSize(ThemeHandle, m_Parent.hdc, SPB_SIZEBOX, SZB_RIGHTALIGN, 0&, TS_TRUE, GripSize) = S_OK Then
            m_Grip.Width = GripSize.cx
            m_Grip.Height = GripSize.cy
        End If
        CloseThemeData ThemeHandle
        
        ' BUG:  Themed size is off 1 pixel in XP.
        ' FIX:  if OSVersion is XP then add 1 pixel to width
        Dim osvi As OSVERSIONINFO
        osvi.dwOSVersionInfoSize = LenB(osvi)
        If GetVersionExW(osvi) Then If osvi.dwMajorVersion = 5 Then m_Grip.Width = m_Grip.Width + 1
    End If
End Sub

Private Sub Class_Terminate()
    Set m_Parent = Nothing
    Set m_Grip = Nothing
End Sub

Private Sub m_Grip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessageA m_Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    
    ' BUG: MousePointer changes back to normal, even though it's still over the Grip
    ' FIX: Force MousePointer back to vbSizeNWSE when sizing complete
    DoEvents
    m_Grip.MousePointer = vbSizeNWSE
End Sub

Private Sub m_Parent_Paint()
    If Not m_Grip.Visible Then Exit Sub
    
    Dim ClipRect As RECT
    ClipRect.Left = m_Grip.Left
    ClipRect.Top = m_Grip.Top
    ClipRect.Right = m_Parent.ScaleWidth
    ClipRect.Bottom = m_Parent.ScaleHeight
    
    Dim ThemeHandle As Long
    ThemeHandle = OpenThemeData(m_Parent.hWnd, StrPtr(SPB_CLASS))
    If ThemeHandle Then
        DrawThemeBackground ThemeHandle, m_Parent.hdc, SPB_SIZEBOX, SZB_RIGHTALIGN, ClipRect, ClipRect
        CloseThemeData ThemeHandle
    Else
        DrawFrameControl m_Parent.hdc, ClipRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP
    End If
End Sub

Private Sub m_Parent_Resize()
    m_Grip.Visible = (m_Parent.WindowState = vbNormal)
    m_Grip.Move m_Parent.ScaleWidth - m_Grip.Width, m_Parent.ScaleHeight - m_Grip.Height
End Sub
