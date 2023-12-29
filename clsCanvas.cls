VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BORDER = 10
Private intWidth As Integer
Private intHeight As Integer
Private intTop As Integer
Private intLeft As Integer
Private dblScale As Double
Private offsetX As Long
Private offsetY As Long


Public Sub init(left As Long, top As Long, width As Long, height As Long)
    intWidth = width
    intHeight = height
    intTop = top
    intLeft = left
End Sub

Public Property Get width() As Single
     width = intWidth
End Property

Public Property Get height() As Single
     height = intHeight
End Property

Public Sub mkBox(box As clsBox, Optional color As Long)
    Dim a As New clsVertex
    Dim b As New clsVertex
    Dim c As New clsVertex
    Dim d As New clsVertex
    Set a = reMap(a.init(box.left, box.top))
    Set b = reMap(b.init(box.right, box.top))
    Set c = reMap(c.init(box.right, box.bottom))
    Set d = reMap(d.init(box.left, box.bottom))
    With ThisWorkbook.Worksheets(1).Shapes
        .AddConnector(msoConnectorStraight, a.x, a.y, b.x, b.y).line.ForeColor.RGB = color
        .AddConnector(msoConnectorStraight, b.x, b.y, c.x, c.y).line.ForeColor.RGB = color
        .AddConnector(msoConnectorStraight, c.x, c.y, d.x, d.y).line.ForeColor.RGB = color
        .AddConnector(msoConnectorStraight, d.x, d.y, a.x, a.y).line.ForeColor.RGB = color
    End With
End Sub

Public Sub mkLine(line As clsLine, Optional color As Long)
    Dim a As clsVertex
    Dim b As clsVertex
    Set a = reMap(line.a)
    Set b = reMap(line.b)
    ThisWorkbook.Worksheets(1).Shapes.AddConnector(msoConnectorStraight, a.x, a.y, b.x, b.y).line.ForeColor.RGB = color
    'ThisWorkbook.Worksheets(1).Shapes.AddConnector msoConnectorStraight, a.x, a.y, b.x, b.y
End Sub

Public Sub mkDot(vertex As clsVertex, Optional color As Long, Optional width As Byte = 32)
    Dim v As clsVertex
    width = width * dblScale
    Set v = reMap(vertex)
    ThisWorkbook.Worksheets(1).Shapes.AddShape(msoShapeRectangle, v.x - width * 0.5, v.y - width * 0.5, width, width).line.ForeColor.RGB = color
    'ThisWorkbook.Worksheets(1).Shapes.AddShape(msoShapeRectangle, v.x - 1, v.y - 1, 2, 2).line.ForeColor.RGB = color
End Sub

Public Sub setScale(bbox As clsBox)
    Dim s1 As Double
    Dim s2 As Double
    s1 = (intWidth - 2 * BORDER) / bbox.dx
    s2 = (intHeight - 2 * BORDER) / bbox.dy
    dblScale = IIf(s1 < s2, s1, s2)
    offsetX = -bbox.left
    offsetY = -bbox.top
End Sub

Private Function reMap(v As clsVertex) As clsVertex
    Dim vertex As New clsVertex
    vertex.init intLeft + (v.x + offsetX) * dblScale + BORDER, intHeight + intTop - (v.y + offsetY) * dblScale - BORDER
    Set reMap = vertex
End Function

Public Sub clear()
    Dim s As Shape
    With ThisWorkbook.Worksheets(1)
        For Each s In ThisWorkbook.Worksheets(1).Shapes
            s.Delete
        Next s
        ''Shapes.AddConnector(msoConnectorStraight, intLeft, intHeight + intTop, intWidth + intLeft, intHeight + intTop).line.ForeColor.RGB = vbBlue
        'Shapes.AddConnector(msoConnectorStraight, intLeft + intWidth, intTop, intWidth + intLeft, intHeight + intTop).line.ForeColor.RGB = vbBlue
    End With
End Sub

