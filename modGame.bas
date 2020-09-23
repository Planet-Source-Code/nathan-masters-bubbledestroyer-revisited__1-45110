Attribute VB_Name = "modGame"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Public Const Pi = 3.14159



Public Function CircCol(x1 As Single, y1 As Single, Rad1 As Single, x2 As Single, y2 As Single, Rad2 As Single) As Boolean
Dim Distance As Single
Distance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
If Distance - Rad1 - Rad2 <= 0 Then CircCol = True
End Function


Public Function Num(Minimum As Single, Maximum As Single) As Single
Dim Range As Single, Temp As Single

Randomize
Temp = 0 + Int((Maximum - Minimum + 1) * Rnd())
Num = Minimum + Temp
End Function

Public Function GetDist(intX1 As Single, intY1 As Single, intX2 As Single, intY2 As Single) As Single
GetDist = Sqr((intX1 - intX2) * (intX1 - intX2) + (intY1 - intY2) * (intY1 - intY2))
End Function

Public Function GetAngle(intX1 As Single, intY1 As Single, intX2 As Single, intY2 As Single) As Single

Dim XComp As Single
Dim YComp As Single

XComp = (intX2 - intX1)
YComp = (intY1 - intY2)

If YComp > 0 Then GetAngle = Atn(XComp / YComp)
If YComp < 0 Then GetAngle = Atn(XComp / YComp) + Pi

End Function

Public Sub AddVectors(Mag1 As Single, Dir1 As Single, Mag2 As Single, Dir2 As Single, Optional ByRef MagResult As Single, Optional ByRef DirResult As Single)

Dim XComp As Single
Dim YComp As Single

XComp = Mag1 * Sin(Dir1) + Mag2 * Sin(Dir2)
YComp = Mag1 * Cos(Dir1) + Mag2 * Cos(Dir2)

MagResult = Sqr(XComp * XComp + YComp * YComp)

If Sgn(YComp) > 0 Then DirResult = Atn(XComp / YComp)
If Sgn(YComp) < 0 Then DirResult = Atn(XComp / YComp) + Pi

End Sub

