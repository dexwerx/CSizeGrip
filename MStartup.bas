Attribute VB_Name = "MStartup"
Option Explicit

Private Declare Function InitShell Lib "shell32" Alias "IsUserAnAdmin" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub Main()
    Dim f As Form

    InitShell
    InitCommonControls
    
    Set f = New FTestForm
    f.Show
End Sub
