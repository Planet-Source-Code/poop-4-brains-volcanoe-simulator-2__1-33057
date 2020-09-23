Attribute VB_Name = "Module1"
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest

Type Pixel
XS As Long
YS As Long
x As Long
y As Long
Count As Long
Act As Long
Color As Long
Hit As Long
End Type

Public P(200) As Pixel

Function MoveDrops(picture As PictureBox)
Dim i As Long, yspeed
For i = 0 To 200
If P(i).Act = True Then
If P(i).Hit < 15 Then
P(i).x = P(i).x + P(i).XS
P(i).y = P(i).y + P(i).YS

If P(i).x < 0 Then
P(i).Act = False
End If

If P(i).x > Form1.board.ScaleWidth Then
P(i).Act = False
End If

P(i).YS = P(i).YS + 1: If P(i).YS > 20 Then P(i).YS = 20

If P(i).y < 0 Then
P(i).y = 1
P(i).Hit = 0
yspeed = (P(i).y)
If yspeed < 0 Then yspeed = 1
P(i).YS = yspeed
End If

If P(i).y > picture.ScaleHeight Then
P(i).x = P(i).x - (P(i).XS + (P(i).XS * 0.99))
P(i).y = P(i).y - (P(i).YS + (P(i).YS * 0.99))
P(i).YS = -(P(i).YS - (P(i).YS * 0.99))
End If

If Form1.buffer.Point(P(i).x, P(i).y) = vbBlack Then
P(i).Hit = P(i).Hit + 1
Form1.fountainm.DrawWidth = 10
Form1.fountains.DrawWidth = 10
Form1.fountainm.PSet (P(i).x, (P(i).y - P(i).YS)), vbWhite
Form1.fountains.PSet (P(i).x, (P(i).y - P(i).YS) - 200), vbBlack
P(i).Hit = 15

P(i).y = P(i).y - (P(i).YS - (P(i).YS * 0.5))
P(i).x = P(i).x - (P(i).XS - (P(i).XS * 0.5))

P(i).YS = -(P(i).YS - (P(i).YS * 0.99))
End If

P(i).Count = P(i).Count - 1: If P(i).Count < 0 Then P(i).Act = False

Else
Dim n As Long
n = Int(Rnd * 255) + 1
If n < 100 Then n = 100
If n > 220 Then n = 220
Form1.fountains.DrawWidth = 10
Form1.fountains.PSet (P(i).x, (P(i).y - P(i).YS) - 200), RGB(n, n, n)
Form1.fountainm.DrawWidth = 10
Form1.fountainm.PSet (P(i).x, (P(i).y - P(i).YS)), vbBlack
P(i).Act = False
P(i).Hit = 0
End If
End If
DoEvents
Next i
End Function

Function MakeDroplet(picture As PictureBox)
Dim i As Long
For i = 0 To 200
If P(i).Act = False Then
P(i).YS = -(Int(Rnd * Int(Rnd * 5)) + Form1.sldPressure.Value)
P(i).XS = IIf(Int(Rnd * 2) = 0, (Int(Rnd * (Form1.sldWidth.Value)) + 1), -(Int(Rnd * (Form1.sldWidth.Value)) + 1))
P(i).Act = True
P(i).x = picture.ScaleWidth \ 2
P(i).y = picture.ScaleHeight - 90
P(i).Count = -((P(i).YS - 5) * (Int(Rnd * 10) + 1))

Select Case Int(Rnd * 3)
Case 0
P(i).Color = &HFF&
Case 1
P(i).Color = &H80FF&
Case 2
P(i).Color = &HC0&
End Select
Exit For
End If
Next i
End Function
