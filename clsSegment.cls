VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objLine As clsLine
Private decAngle As Currency
Private objLinedef As clsLinedef
Private boolFlipped As Boolean
'_____FROM PLAYERS VIEW
Private decAngle2V1 As Currency
Private decAngle2V2 As Currency
Private lngDistance As Long 'shotest distance from player to infinate segment
Private decInitAngle As Integer 'angle to line.a before clipping

Public Sub init(line As clsLine, bamAngle As Integer, linedef As clsLinedef, orientation As Integer)
    Set objLine = line
    Set objLinedef = linedef
    boolFlipped = orientation
    decAngle = 0.0054931640625 * bamAngle
    If decAngle < 0 Then decAngle = decAngle + 360
End Sub
    
Public Property Get line() As clsLine
     Set line = objLine
End Property

Public Property Get linedef() As clsLinedef
     Set linedef = objLinedef
End Property

Public Property Get isFlipped() As Boolean
    isFlipped = boolFlipped
End Property

Public Property Get angle() As Currency
     angle = decAngle
End Property

Public Property Get angleToV1() As Currency
     angleToV1 = decAngle2V1
End Property

Public Property Get angleToV2() As Currency
     angleToV2 = decAngle2V2
End Property

Public Property Let angleToV1(val As Currency)
     decAngle2V1 = val
End Property

Public Property Let angleToV2(val As Currency)
     decAngle2V2 = val
End Property

Public Property Get distance() As Long
     distance = lngDistance
End Property

Public Property Let distance(val As Long)
     lngDistance = val
End Property

Public Property Let initAngle(val As Currency)
     decInitAngle = val
End Property

Public Function scaleAtAngle(degrees As Currency, playerAngle As Integer) As Single
     Dim dist1 As Integer
     Dim dist2 As Long
     
     'screen dist = 400 @ 90FOV
     'screenDist=screenWidth/tan(rad(HFOV))
     dist1 = 400 * Cos(rad(decAngle - degrees - playerAngle - 90))
     dist2 = lngDistance * Cos(rad(degrees))
     If dist2 = 0 Then dist2 = 1
     scaleAtAngle = Abs(dist1 / dist2) * 1
     'If scaleAtAngle < 0 Then scaleAtAngle = 0
End Function

Public Property Get backSector() As clsSector
    If boolFlipped Then
        Set backSector = objLinedef.front.sector
    Else
        Set backSector = objLinedef.back.sector
    End If
End Property

Public Property Get frontSector() As clsSector
    If boolFlipped Then
        Set frontSector = objLinedef.back.sector
    Else
        Set frontSector = objLinedef.front.sector
    End If
End Property
