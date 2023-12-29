VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private valX As Long
Private valY As Long

Public Function init(x As Long, y As Long) As clsVertex
    valX = x
    valY = y
    Set init = Me
End Function

Public Property Get x() As Long
     x = valX
End Property

Public Property Get y() As Long
     y = valY
End Property

Public Property Let x(val As Long)
    valX = val
End Property

Public Property Let y(val As Long)
    valY = val
End Property

Public Function angleTo(vertex As clsVertex) As Currency
    Dim dx As Long
    Dim dy As Long
    dx = vertex.x - valX
    dy = vertex.y - valY
    If Not (dx = 0 And dy = 0) Then angleTo = deg(WorksheetFunction.Atan2(dx, dy))
End Function

Public Function distanceTo(vertex As clsVertex) As Integer
    Dim dx As Long
    Dim dy As Long
    dx = Abs(vertex.x - valX)
    dy = Abs(vertex.y - valY)
    distanceTo = Sqr(dx * dx + dy * dy)
End Function
