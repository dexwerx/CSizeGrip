VERSION 5.00
Begin VB.Form FTestForm 
   Caption         =   "Themed Size Grip"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnFlip 
      Caption         =   "Flip Sides"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private SizeGrip As New CSizeGrip
Private Sub Form_Load()
    SizeGrip.Attach Me
End Sub
Private Sub btnFlip_Click()
    If SizeGrip.Orientation = gripBottomLeft Then
        SizeGrip.Orientation = gripBottomRight
    Else
        SizeGrip.Orientation = gripBottomLeft
    End If
End Sub


