VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngTop As Long
Private lngBottom As Long
Private lngLeft As Long
Private lngRight As Long

Public Sub init(top As Long, bottom As Long, left As Long, right As Long)
    lngTop = top
    lngBottom = bottom
    lngLeft = left
    lngRight = right
End Sub

'____________________________________________________4 get side props
Public Property Get left() As Long
     left = lngLeft
End Property

Public Property Get right() As Long
     right = lngRight
End Property

Public Property Get top() As Long
     top = lngTop
End Property

Public Property Get bottom() As Long
     bottom = lngBottom
End Property

'____________________________________________________4 get box line functions
Public Function leftLine() As clsLine
     Dim a As New clsVertex
     Dim b As New clsVertex
     Set leftLine = New clsLine
     leftLine.init a.init(lngLeft, lngTop), b.init(lngLeft, lngBottom)
End Function
Public Function rightLine() As clsLine
     Dim a As New clsVertex
     Dim b As New clsVertex
     Set rightLine = New clsLine
     rightLine.init a.init(lngRight, lngBottom), b.init(lngRight, lngTop)
End Function
Public Function topLine() As clsLine
     Dim a As New clsVertex
     Dim b As New clsVertex
     Set topLine = New clsLine
     topLine.init a.init(lngRight, lngTop), b.init(lngLeft, lngTop)
End Function
Public Function bottomLine() As clsLine
     Dim a As New clsVertex
     Dim b As New clsVertex
     Set bottomLine = New clsLine
     bottomLine.init a.init(lngLeft, lngBottom), b.init(lngRight, lngBottom)
End Function

'____________________________________________________2 get DX/DY props
Public Property Get dx() As Long
     dx = lngRight - lngLeft
End Property

Public Property Get dy() As Long
     dy = lngBottom - lngTop
End Property

'____________________________________________________4 LET side props
Public Property Let left(val As Long)
    lngLeft = val
End Property

Public Property Let right(val As Long)
    lngRight = val
End Property

Public Property Let top(val As Long)
    lngTop = val
End Property

Public Property Let bottom(val As Long)
    lngBottom = val
End Property


