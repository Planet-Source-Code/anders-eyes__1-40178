VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Eyes..."
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2235
   LinkTopic       =   "Form1"
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   149
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pMouse As POINTAPI
Dim pNewMouse As POINTAPI
Dim pCenter As POINTAPI
Dim pCenter2 As POINTAPI
Dim intRadius As Integer


Sub Run()
Dim picCenter As POINTAPI

picCenter.X = (Picture1.ScaleWidth / 2)
picCenter.Y = (Picture1.ScaleHeight / 2)

Do
Call Sleep(1)
pMouse = pNewMouse
Call GetCursorPos(pNewMouse)
If pNewMouse.X <> pMouse.X Or pNewMouse.Y <> pMouse.Y Then
Me.Cls
pNewMouse.X = pNewMouse.X - (Me.Left / Screen.TwipsPerPixelX) - 4
pNewMouse.Y = pNewMouse.Y - (Me.Top / Screen.TwipsPerPixelY) - 4
Call PaintAEye(pCenter, pNewMouse, picCenter)
Call PaintAEye(pCenter2, pNewMouse, picCenter)

End If
DoEvents
Loop
End Sub

Private Sub PaintAEye(pEyeToPaint As POINTAPI, pMousePos As POINTAPI, pCenterOfPic As POINTAPI)
Dim dblAngle As Double
Dim X As Long
Dim Y As Long
Dim i As Integer
dblAngle = GetAngleOfMouse(pEyeToPaint, pMousePos)
X = ((Cos(dblAngle) * intRadius) + pEyeToPaint.X) - pCenterOfPic.X
Y = ((Sin(dblAngle) * intRadius) + pEyeToPaint.Y) - pCenterOfPic.Y


Me.DrawWidth = 5
For i = 0 To intRadius + 10 Step Me.DrawWidth
Me.Circle (pEyeToPaint.X, pEyeToPaint.Y), i, RGB(255, 255, 255)
Next i
Me.DrawWidth = 8
Me.Circle (pEyeToPaint.X, pEyeToPaint.Y), intRadius + 10

Call PaintTransPictureAt(Picture1, Picture2, X, Y, Me.hdc)

End Sub



Private Sub Form_DblClick()
MsgBox "Mail: " & vbNewLine & "egon@olsen.vg", vbInformation, "Eyes..."
End Sub

Private Sub Form_Load()
Me.ScaleMode = vbPixels
intRadius = 25
Form1.Show
Call Run
End Sub

Private Sub Form_Resize()
On Error Resume Next
pCenter.X = Me.ScaleWidth / 2
pCenter.Y = Me.ScaleHeight / 2
pCenter2.X = pCenter.X + 35
pCenter2.Y = pCenter.Y
pCenter.X = pCenter.X - 35
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

