VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vertPosition As clsVertex
Private objOldPosition As New clsVertex
Private intAngle As Integer
Private valFlag As Byte
Private canvas As clsCanvas
Private screen As Boolean
Private intHeight As Integer


Public Sub init(thing As clsThing, canvos As clsCanvas)
    Set vertPosition = thing.position
    Set canvas = canvos
    intAngle = thing.angle
    intHeight = 15
End Sub

Public Property Get position() As clsVertex
     Set position = vertPosition
End Property

Public Property Set position(obj As clsVertex)
     Set vertPosition = obj
End Property

Public Property Get angle() As Integer
     angle = intAngle
End Property

Public Property Get height() As Integer
     height = intHeight
End Property

Public Property Let height(val As Integer)
     intHeight = val
End Property

Public Property Let angle(val As Integer)
     intAngle = val
End Property

Public Sub draw()
    Dim v1 As New clsVertex
    Dim v2 As New clsVertex
    Dim l1 As New clsLine
    Dim l2 As New clsLine
    
    canvas.mkDot vertPosition, , PLAYER_WIDTH
    '_________________________________FOW lines
    v1.init vertPosition.x + 500 * Cos(rad(intAngle + HFOV)), vertPosition.y + 500 * Sin(rad(intAngle + HFOV))
    v2.init vertPosition.x + 500 * Cos(rad(intAngle - HFOV)), vertPosition.y + 500 * Sin(rad(intAngle - HFOV))
    canvas.mkLine l1.init(vertPosition, v1), vbMagenta
    canvas.mkLine l2.init(vertPosition, v2), vbMagenta
End Sub

Public Sub move(key, steps)
    objOldPosition.init vertPosition.x, vertPosition.y
    Select Case key
        Case vbKeyUp, vbKeyW
            vertPosition.y = vertPosition.y + steps * Sin(rad(intAngle))
            vertPosition.x = vertPosition.x + steps * Cos(rad(intAngle))
        Case vbKeyDown, vbKeyS
            vertPosition.y = vertPosition.y - steps * Sin(rad(intAngle))
            vertPosition.x = vertPosition.x - steps * Cos(rad(intAngle))
        Case vbKeyA:
            vertPosition.y = vertPosition.y + steps * Sin(rad(intAngle + 90))
            vertPosition.x = vertPosition.x + steps * Cos(rad(intAngle + 90))
        Case vbKeyD:
            vertPosition.y = vertPosition.y - steps * Sin(rad(intAngle + 90))
            vertPosition.x = vertPosition.x - steps * Cos(rad(intAngle + 90))
    End Select
End Sub

Public Sub moveBack()
    vertPosition.init objOldPosition.x, objOldPosition.y
    Beep
End Sub

Public Sub turn(angle As Integer)
    intAngle = intAngle + angle
End Sub

Public Function seeBox(box As clsBox) As Boolean
    'If node's bounding box is not seen from players FOV, no need to render ANYTHING inside
    Dim lines(1) As clsLine
    Dim i As Byte
    Dim d1 As Integer    '          1 | 2 | 3
    Dim d2 As Integer    '          --+---+---
    Dim span As Integer  '          4 | 5 | 6
    Dim span1 As Integer '          --+---+---
    Dim span2 As Integer '          7 | 8 | 9
    '_________________________________________get lines you could see from player's position
    If vertPosition.y > box.top Then
        If vertPosition.x < box.left Then
            Set lines(0) = box.leftLine     '1
            Set lines(1) = box.topLine
        ElseIf vertPosition.x > box.right Then
            Set lines(0) = box.rightLine    '3
            Set lines(1) = box.topLine
        Else
            Set lines(0) = box.topLine      '2
        End If
    ElseIf vertPosition.y < box.bottom Then
        If vertPosition.x < box.left Then
            Set lines(0) = box.leftLine     '7
            Set lines(1) = box.bottomLine
        ElseIf vertPosition.x > box.right Then
            Set lines(0) = box.rightLine    '9
            Set lines(1) = box.bottomLine
        Else
            Set lines(0) = box.bottomLine   '8
        End If
    Else
        If vertPosition.x < box.left Then
            Set lines(0) = box.leftLine     '4
        ElseIf vertPosition.x > box.right Then
            Set lines(0) = box.rightLine    '6
        Else
            seeBox = True                   '5
            Exit Function
        End If
    End If
    '______________________________________is player turned towards ANY of collected lines
    For i = 0 To 1
        If Not lines(i) Is Nothing Then
            d1 = vertPosition.angleTo(lines(i).a)
            d2 = vertPosition.angleTo(lines(i).b)
            span = normalizeDeg(d1 - d2)
            d1 = d1 - intAngle
            span1 = normalizeDeg(d1 + HFOV)
            If span1 > FOV Then
                If span1 < span + FOV Then
                    seeBox = True
                    Exit Function
                End If
            Else
                seeBox = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function seeSegment(segment As clsSegment) As Boolean
    'check if and how much of segment should be rendered;
    Dim d1 As Currency
    Dim d2 As Currency
    Dim d0 As Currency
    Dim span As Currency
    Dim span1 As Currency
    Dim span2 As Currency
    '#Dim dist As Long 'while screen buffer not available; to reduce number of rendered sectors
    d1 = vertPosition.angleTo(segment.line.a)
    d2 = vertPosition.angleTo(segment.line.b)
    d0 = d1
    span = normalizeDeg(d1 - d2)
    If span > 180 Then Exit Function
    d1 = d1 - intAngle
    d2 = d2 - intAngle
    '______________________________________________________________can see D1?
    span1 = normalizeDeg(d1 + HFOV)
    If span1 > FOV Then
        If span1 >= span + FOV Then Exit Function
        d1 = HFOV
        '#dist = 1
    End If
    '_______________________________________________________________can see D2?
    span2 = normalizeDeg(HFOV - d2)
    If span2 > FOV Then
        If span2 >= span + FOV Then Exit Function
        d2 = -HFOV
        '#dist = 1
    End If
    'until we implement screen buffer, do not drow segments too far from player
    '#If dist = 0 Then dist = aproxDistance(vertPosition.x - segment.line.a.x, vertPosition.y - segment.line.a.y)
    '#span2 = aproxDistance(vertPosition.x - segment.line.b.x, vertPosition.y - segment.line.b.y)
    '#If dist < span2 Then dist = span2
    '#If dist > 4000 Then Exit Function
    '________________________________________classify segment from players view
    segment.angleToV1 = d1
    segment.angleToV2 = d2
    segment.distance = Cos(rad(segment.angle + 90 - d0)) * vertPosition.distanceTo(segment.line.a)
    segment.initAngle = d0
    seeSegment = True
End Function

