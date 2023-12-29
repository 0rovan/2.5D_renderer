Attribute VB_Name = "math"
Option Explicit
Private Const piOver180 = 0.017453292519943

Public Function rad(degree As Variant) As Double
    rad = degree * piOver180
End Function

Public Function deg(radians As Double) As Currency
    deg = radians / piOver180
End Function

Public Function normalizeDeg(degree As Currency) As Currency
    Dim dbl As Single
    dbl = degree - Int(degree)
    degree = Int(degree) Mod 360
    normalizeDeg = IIf(degree < 0, degree + 360, degree) + dbl
End Function

Public Function aproxDistance(dx As Long, dy As Long) As Long
    dx = Abs(dx)
    dy = Abs(dy)
    aproxDistance = dx + dy - (IIf(dx < dy, dx, dy) * 0.5)
End Function

Public Function tripletOrientation(v1 As clsVertex, v2 As clsVertex, v3 As clsVertex) As Byte
    '0 - coLinnear
    '1 - clockWise
    '2 - counterClockWise
    Select Case ((v2.y - v1.y) * (v3.x - v2.x)) - ((v2.x - v1.x) * (v3.y - v2.y))
        Case Is > 0
            tripletOrientation = 1
        Case Is < 0
            tripletOrientation = 2
    End Select
End Function


Public Function lineIntersect(l1 As clsLine, l2 As clsLine) As Boolean
    lineIntersect = (Not tripletOrientation(l1.a, l1.b, l2.a) = tripletOrientation(l1.a, l1.b, l2.b)) And (Not tripletOrientation(l2.a, l2.b, l1.a) = tripletOrientation(l2.a, l2.b, l1.b))
End Function
