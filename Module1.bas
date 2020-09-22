Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Global Const SRCCOPY = &HCC0020
Global Const SRCAND = &H8800C6
Global Const SRCPAINT = &HEE0086

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Global Const pi = 3.14159265358979 ' (pi = 4 * Atn(1))

Public Function GetAngleOfMouse(ByRef pCenterPos As POINTAPI, ByRef pMousePos As POINTAPI) As Double
Dim lhKat As Long, laKat As Long
Dim v As Double

If (pMousePos.Y < pCenterPos.Y) Then
 If (pMousePos.X > pCenterPos.X) Then
   lhKat = (pMousePos.X - pCenterPos.X)
   laKat = (pCenterPos.Y - pMousePos.Y)
   v = Atn(lhKat / laKat) + (3 * pi / 2)
 ElseIf (pMousePos.X < pCenterPos.X) Then
  lhKat = (pCenterPos.X - pMousePos.X)
  laKat = (pCenterPos.Y - pMousePos.Y)
   v = Atn(laKat / lhKat) + pi
 ElseIf (pMousePos.X = pCenterPos.X) Then
  If pMousePos.Y < pCenterPos.Y Then v = ((3 * pi) / 2)
 End If
ElseIf (pMousePos.Y > pCenterPos.Y) Then
 If pMousePos.X > pCenterPos.X Then
  lhKat = (pMousePos.X - pCenterPos.X)
  laKat = (pCenterPos.Y - pMousePos.Y)
  v = Atn(lhKat / laKat) + (pi / 2)
 ElseIf pMousePos.X < pCenterPos.X Then v = (pi)
  lhKat = (pMousePos.X - pCenterPos.X)
  laKat = (pCenterPos.Y - pMousePos.Y)
  v = Atn(lhKat / laKat) + (pi / 2)
 ElseIf (pMousePos.X = pCenterPos.X) Then
  If pMousePos.Y > pCenterPos.Y Then v = (pi / 2)
 End If
ElseIf (pMousePos.Y = pCenterPos.Y) Then
  If pMousePos.X > pCenterPos.X Then v = 2 * pi
  If pMousePos.X < pCenterPos.X Then v = (pi)
End If
 GetAngleOfMouse = v '((v / pi) * 180)
End Function

Public Sub PaintTransPictureAt(picboxB, picboxW As Object, xval, yval, ByRef lnghDCTarget As Long)
Call BitBlt(lnghDCTarget, xval, yval, picboxB.ScaleWidth, picboxB.ScaleHeight, picboxB.hdc, 0, 0, SRCAND)
Call BitBlt(lnghDCTarget, xval, yval, picboxW.ScaleWidth, picboxW.ScaleHeight, picboxW.hdc, 0, 0, SRCPAINT)
End Sub
