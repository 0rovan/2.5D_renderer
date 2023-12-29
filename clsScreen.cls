VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intWidth As Integer
Private intHeight As Integer
Private intTop As Integer
Private intLeft As Integer
Private intDist As Integer
Private objPlayer As clsPlayer
Private buffer() As bufColumn

Private Type bufColumn
    wall As Boolean
    top As Integer
    bottom As Integer
End Type

Private Type wallSpace
    x1 As Long
    x2 As Long
End Type

Private Type portalColumn
    ceilTop As Long
    ceilBot As Long
    floorTop As Long
    floorBot As Long
End Type

Public Sub init(left As Long, top As Long, width As Long, height As Long)
    intWidth = width
    intHeight = height
    intTop = top
    intLeft = left
    intDist = intWidth * 0.5 / Tan(rad(HFOV))
    'objBuffer.init intWidth
End Sub

Public Property Get width() As Single
     width = intWidth
End Property

Public Property Get height() As Single
     height = intHeight
End Property

Public Property Set player(obj As clsPlayer)
     Set objPlayer = obj
End Property

Public Sub mkLine(line As clsLine, Optional color As Long, Optional width As Byte = 2)
    Dim a As clsVertex
    Dim b As clsVertex
    Set a = line.a
    Set b = line.b
    With ThisWorkbook.Worksheets(1).Shapes.AddConnector(msoConnectorStraight, a.x + intLeft, a.y + intTop, b.x + intLeft, b.y + intTop).line
        .ForeColor.RGB = color
        .Weight = width
    End With
    'ThisWorkbook.Worksheets(1).Shapes.AddConnector msoConnectorStraight, a.x + intLeft, a.y + intTop, b.x + intLeft, b.y + intTop
End Sub

Public Function reMapX(degree As Currency) As Integer
    If degree > 0 Then
        reMapX = intWidth * 0.5 - intDist * Tan(rad(degree))
    Else
        reMapX = intWidth * 0.5 - Tan(rad(degree)) * intDist
    End If
End Function


Public Sub mkSolidWall(segment As clsSegment, Optional fakeWall = False)

    Dim line As New clsLine
    Dim v1 As New clsVertex
    Dim v2 As New clsVertex
    Dim y1c As Long
    Dim y2f As Long
    Dim y2c As Long
    Dim y1f As Long
    Dim clippedWall As wallSpace
    Dim wall As wallSpace
    Dim v1Scale As Single
    Dim v2Scale As Single
    Dim i As Integer
    Dim j As Integer
    Dim horizontalSpace() As wallSpace
    Dim color As Long
    'Dim noFloor As Boolean
    'Dim noCeiling As Boolean
    Dim lineWidth As Byte
    
    Dim floorLine() As clsLine
    Dim ceilingLine() As clsLine
    Dim step As Single
    'Dim vert As clsVertex
    

    wall.x1 = reMapX(segment.angleToV1)
    wall.x2 = reMapX(segment.angleToV2)
    'no need to draw invisibli thin wall; check again aftre clipping
    If wall.x1 = wall.x2 Then Exit Sub
    
    horizontalSpace = horizontalClip(wall)
    If horizontalSpace(0).x1 = horizontalSpace(0).x2 Then Exit Sub
    'if segment was clipped, find new angle needed for scaling factor
    For i = 0 To UBound(horizontalSpace)
        clippedWall = horizontalSpace(i)
        'scaling for left angle
        If clippedWall.x1 = wall.x1 Then
            v1Scale = segment.scaleAtAngle(segment.angleToV1, objPlayer.angle)
        Else
            v1Scale = segment.scaleAtAngle(deg(Atn((intWidth * 0.5 - clippedWall.x1) / intDist)), objPlayer.angle)
        End If
        'scaling for right angle
        If clippedWall.x2 = wall.x2 Then
            v2Scale = segment.scaleAtAngle(segment.angleToV2, objPlayer.angle)
        Else
            v2Scale = segment.scaleAtAngle(deg(Atn((intWidth * 0.5 - clippedWall.x2) / intDist)), objPlayer.angle)
        End If
        Select Case v1Scale + v2Scale
            Case Is > 3: lineWidth = 3
            Case Is > 0.8: lineWidth = 2
            Case Else: lineWidth = 1
        End Select
        'linedef won't back backsector
        y1c = intHeight * 0.5 - (segment.linedef.front.sector.ceiling - objPlayer.height) * v1Scale
        y2c = intHeight * 0.5 - (segment.linedef.front.sector.ceiling - objPlayer.height) * v2Scale
        y1f = intHeight * 0.5 - (segment.linedef.front.sector.floor - objPlayer.height) * v1Scale
        y2f = intHeight * 0.5 - (segment.linedef.front.sector.floor - objPlayer.height) * v2Scale
        'If y1c < -250 Then y1c = -250 ' cheesing - angles will not
        'If y2c < -250 Then y2c = -250 ' be represented trully
        
        If y1c <= buffer(clippedWall.x1).bottom Or y2c <= buffer(clippedWall.x1).bottom Then
            If y1f >= buffer(clippedWall.x1).top Or y2f >= buffer(clippedWall.x1).top Then
                '___________________________________FLOOR LINE
                If y1f <= intHeight Or y2f <= intHeight Then
                    color = IIf(fakeWall, &H66BBFF, vbRed)
                    step = (y1f - y2f) / (clippedWall.x2 - clippedWall.x1)
                    floorLine = verticalClip(v1.init(clippedWall.x1, y1f), clippedWall.x2, step)
                    For j = 0 To UBound(floorLine)
                        If Not floorLine(j).a.x = floorLine(j).b.x Then
                            mkLine floorLine(j), color, lineWidth
                        End If
                    Next j
                End If
                '____________________________________CEIL LINE
                If y1c >= 0 Or y2c >= 0 Then
                    color = IIf(fakeWall, vbCyan, vbBlue)
                    step = (y1c - y2c) / (clippedWall.x2 - clippedWall.x1)
                    ceilingLine = verticalClip(v1.init(clippedWall.x1, y1c), clippedWall.x2, step)
                    For j = 0 To UBound(ceilingLine)
                        If Not ceilingLine(j).a.x = ceilingLine(j).b.x Then
                            mkLine ceilingLine(j), color, lineWidth
                        End If
                    Next j
                End If
                '___________________________________SIDE LINES
                If y1f > buffer(clippedWall.x1).bottom Then y1f = buffer(clippedWall.x1).bottom
                If y1c < buffer(clippedWall.x1).top Then y1c = buffer(clippedWall.x1).top
                If y1f > y1c Then
                    mkLine line.init(v1.init(clippedWall.x1, y1c), v2.init(clippedWall.x1, y1f)), , lineWidth
                End If
                If y2f > buffer(clippedWall.x2).bottom Then y2f = buffer(clippedWall.x2).bottom
                If y2c < buffer(clippedWall.x2).top Then y2c = buffer(clippedWall.x2).top
                If y2f > y2c Then
                    mkLine line.init(v1.init(clippedWall.x2, y2c), v2.init(clippedWall.x2, y2f)), , lineWidth
                End If
                End If
        End If
    Next i
End Sub

Public Sub mkPortal(segment As clsSegment)

    Dim line As New clsLine
    Dim v1 As New clsVertex
    Dim v2 As New clsVertex
    Dim v1Scale As Single
    Dim v2Scale As Single
    Dim wall As wallSpace
    Dim horizontalSpace() As wallSpace
    Dim clippedWall As wallSpace
    Dim i As Byte
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim yf As Long
    Dim yc As Long
    Dim stalactite As Boolean
    Dim stalagmite As Boolean
    Dim floorDrop As Boolean
    Dim ceilDrop As Boolean
    Dim y1 As portalColumn
    Dim y2 As portalColumn
    Dim step1 As Single
    Dim step2 As Single
    Dim step As Single
    Dim floorLine() As clsLine
    Dim ceilingLine() As clsLine
    Dim lineWidth As Byte
    
    'Dim noRoof As Boolean
    'Dim noCeil As Boolean
    'Dim noDesk As Boolean
    'Dim noFloor As Boolean
    
    'treat portal as wall for simple checks
    wall.x1 = reMapX(segment.angleToV1)
    wall.x2 = reMapX(segment.angleToV2)
    'no need to draw invisibli thin wall; check again aftre clipping
    If wall.x1 = wall.x2 Then Exit Sub
    horizontalSpace = horizontalClip(wall, False)
    If horizontalSpace(0).x1 = horizontalSpace(0).x2 Then Exit Sub
    floorDrop = segment.frontSector.floor > segment.backSector.floor
    ceilDrop = segment.frontSector.ceiling > segment.backSector.ceiling
    For i = 0 To UBound(horizontalSpace)
        clippedWall = horizontalSpace(i)
        '__________________________________________________________________scaling factors
        If clippedWall.x1 = wall.x1 Then
            v1Scale = segment.scaleAtAngle(segment.angleToV1, objPlayer.angle)
        Else
            v1Scale = segment.scaleAtAngle(deg(Atn((intWidth * 0.5 - clippedWall.x1) / intDist)), objPlayer.angle)
        End If
        If clippedWall.x2 = wall.x2 Then
            v2Scale = segment.scaleAtAngle(segment.angleToV2, objPlayer.angle)
        Else
            v2Scale = segment.scaleAtAngle(deg(Atn((intWidth * 0.5 - clippedWall.x2) / intDist)), objPlayer.angle)
        End If
        '___________________________________________________________________assign heights
        y1.ceilBot = intHeight * 0.5 - (IIf(ceilDrop, segment.backSector.ceiling, segment.frontSector.ceiling) - objPlayer.height) * v1Scale
        y2.ceilBot = intHeight * 0.5 - (IIf(ceilDrop, segment.backSector.ceiling, segment.frontSector.ceiling) - objPlayer.height) * v2Scale
        y1.floorTop = intHeight * 0.5 - (IIf(floorDrop, segment.frontSector.floor, segment.backSector.floor) - objPlayer.height) * v1Scale
        y2.floorTop = intHeight * 0.5 - (IIf(floorDrop, segment.frontSector.floor, segment.backSector.floor) - objPlayer.height) * v2Scale
        Select Case v1Scale + v2Scale
            Case Is > 3: lineWidth = 3
            Case Is > 0.8: lineWidth = 2
            Case Else: lineWidth = 1
        End Select
        
        stalagmite = False
        stalactite = False
        
        If Not segment.frontSector.floor = segment.backSector.floor Then
            If y1.floorTop <= intHeight Or y2.floorTop <= intHeight Then stalagmite = True
        End If
        If Not segment.frontSector.ceiling = segment.backSector.ceiling Then
            If y1.ceilBot >= 0 Or y2.ceilBot >= 0 Then stalactite = True
        End If
        '===============================================================================================================
        '==================================================================================================== STALAGMITE
        If stalagmite Then
            y1.floorBot = intHeight * 0.5 - (IIf(floorDrop, segment.backSector.floor, segment.frontSector.floor) - objPlayer.height) * v1Scale
            y2.floorBot = intHeight * 0.5 - (IIf(floorDrop, segment.backSector.floor, segment.frontSector.floor) - objPlayer.height) * v2Scale
            
            '____________________________________________________________________________TOP LINE
            If y1.floorTop >= 0 Or y2.floorTop >= 0 Then
                step1 = (y1.floorTop - y2.floorTop) / (clippedWall.x2 - clippedWall.x1)
                ceilingLine = verticalClip(v1.init(clippedWall.x1, y1.floorTop), clippedWall.x2, step1)
                For j = 0 To UBound(ceilingLine)
                    If Not ceilingLine(j).a.x = ceilingLine(j).b.x Then
                        mkLine ceilingLine(j), &HFF8888, lineWidth
                    End If
                Next j
            End If
            If Not floorDrop Then
                '______________________________________________________________________BOTTOM LINE
                If y1.floorBot <= intHeight Or y2.floorBot <= intHeight Then
                    step = (y1.floorBot - y2.floorBot) / (clippedWall.x2 - clippedWall.x1)
                    floorLine = verticalClip(v1.init(clippedWall.x1, y1.floorBot), clippedWall.x2, step)
                    For j = 0 To UBound(floorLine)
                        If Not floorLine(j).a.x = floorLine(j).b.x Then
                            mkLine floorLine(j), &H8888FF, lineWidth
                        End If
                    Next j
                End If
                '________________________________________________________________________LEFT LINE
                j = clippedWall.x1 + IIf(clippedWall.x1 = 0, 0, 1)
                yf = IIf(y1.floorBot < buffer(j).bottom, y1.floorBot, buffer(j).bottom)
                yc = IIf(y1.floorTop > buffer(j).top, y1.floorTop, buffer(j).top)
                If yf > yc Then
                    mkLine line.init(v1.init(clippedWall.x1, yc), v2.init(clippedWall.x1, yf)), &HAAAAAA, lineWidth
                    If clippedWall.x1 = clippedWall.x2 Then Debug.Print "hoho"
                End If
                '_______________________________________________________________________RIGHT LINE
                j = clippedWall.x2 - IIf(clippedWall.x2 = intWidth, 0, 1)
                yf = IIf(y2.floorBot < buffer(j).bottom, y2.floorBot, buffer(j).bottom)
                yc = IIf(y2.floorTop > buffer(j).top, y2.floorTop, buffer(j).top)
                If yf > yc Then
                    mkLine line.init(v1.init(clippedWall.x2, yc), v2.init(clippedWall.x2, yf)), &HAAAAAA, lineWidth
                    If clippedWall.x1 = clippedWall.x2 Then Debug.Print "yoyo"
                End If
            End If
            '___________________________________________________________________________buffer fill
            If y1.floorTop < 0 And y2.floorTop < 0 Then
                For j = clippedWall.x1 To clippedWall.x2
                    buffer(j).wall = True
                Next j
            Else
                k = clippedWall.x1
                If y1.floorTop < 0 Then k = k + y1.floorTop / step1
                m = clippedWall.x2
                If y2.floorTop < 0 Then m = m + y2.floorTop / step1
                For j = k To m
                    yf = y1.floorTop - step1 * (j - clippedWall.x1)
                    If yf < buffer(j).bottom Then
                        If yf > buffer(j).top Then
                            buffer(j).bottom = yf
                        Else
                            buffer(j).wall = True
                        End If
                    End If
                Next j
                For j = clippedWall.x1 To k - 1
                    buffer(j).wall = True
                Next j
                For j = m + 1 To clippedWall.x2
                    buffer(j).wall = True
                Next j
            End If
        End If
        '===============================================================================================================
        '==================================================================================================== STALACTITE
        If stalactite Then
            y1.ceilTop = intHeight * 0.5 - (IIf(ceilDrop, segment.frontSector.ceiling, segment.backSector.ceiling) - objPlayer.height) * v1Scale
            y2.ceilTop = intHeight * 0.5 - (IIf(ceilDrop, segment.frontSector.ceiling, segment.backSector.ceiling) - objPlayer.height) * v2Scale
            
            '_________________________________________________________________BOTTOM LINE
            If y1.ceilBot <= intHeight Or y2.ceilBot <= intHeight Then
                step2 = (y1.ceilBot - y2.ceilBot) / (clippedWall.x2 - clippedWall.x1)
                floorLine = verticalClip(v1.init(clippedWall.x1, y1.ceilBot), clippedWall.x2, step2)
                For j = 0 To UBound(floorLine)
                    If Not floorLine(j).a.x = floorLine(j).b.x Then
                        mkLine floorLine(j), &H8888FF, lineWidth
                    End If
                Next j
            End If
            If ceilDrop Then
                '________________________________________________________________TOP LINE
                If y1.ceilTop >= 0 Or y2.ceilTop >= 0 Then
                    step = (y1.ceilTop - y2.ceilTop) / (clippedWall.x2 - clippedWall.x1)
                    ceilingLine = verticalClip(v1.init(clippedWall.x1, y1.ceilTop), clippedWall.x2, step)
                    For j = 0 To UBound(ceilingLine)
                        If Not ceilingLine(j).a.x = ceilingLine(j).b.x Then
                            mkLine ceilingLine(j), &HFF8888, lineWidth
                        End If
                    Next j
                End If
                '________________________________________________________________LEFT LINE
                j = clippedWall.x1 + IIf(clippedWall.x1 = 0, 0, 1)
                yc = IIf(y1.ceilTop < buffer(j).top, buffer(j).top, y1.ceilTop)
                yf = IIf(y1.ceilBot > buffer(j).bottom, buffer(j).bottom, y1.ceilBot)
                If yf > yc Then
                    mkLine line.init(v1.init(clippedWall.x1, yc), v2.init(clippedWall.x1, yf)), &HAAAAAA, lineWidth
                End If
                '_______________________________________________________________RIGHT LINE
                j = clippedWall.x2 - IIf(clippedWall.x2 = intWidth, 0, 1)
                yc = IIf(y2.ceilTop < buffer(j).top, buffer(j).top, y2.ceilTop)
                yf = IIf(y2.ceilBot > buffer(j).bottom, buffer(j).bottom, y2.ceilBot)
                If yf > yc Then
                    mkLine line.init(v1.init(clippedWall.x2, yc), v2.init(clippedWall.x2, yf)), &HAAAAAA, lineWidth
                End If
            End If
            '___________________________________________________________________buffer fill
            If y1.ceilBot >= intHeight And y2.ceilBot >= intHeight Then
                For j = clippedWall.x1 To clippedWall.x2
                    buffer(j).wall = True
                Next j
            Else
                k = clippedWall.x1
                If y1.ceilBot > intHeight Then
                    k = k + (intHeight - y2.ceilBot) / step2
                End If
                m = clippedWall.x2
                If y2.ceilBot > intHeight Then
                    m = m - (intHeight - y2.ceilBot) / step2
                End If
                For j = k To m
                    yc = y1.ceilBot - step2 * (j - clippedWall.x1)
                    If yc > buffer(j).top Then
                        If buffer(j).bottom > yc Then
                            buffer(j).top = yc
                        Else
                            buffer(j).wall = True
                        End If
                    End If
                Next j
                For j = clippedWall.x1 To k - 1
                    buffer(j).wall = True
                Next j
                For j = m + 1 To clippedWall.x2
                    buffer(j).wall = True
                Next j
            End If
        End If
    Next i
End Sub

Private Function horizontalClip(wall As wallSpace, Optional addToBuffer = True) As wallSpace()
    Dim i As Byte
    Dim w As wallSpace
    
    'even if no space if found, returned array has at least 1 element at index 0
    ReDim out(0) As wallSpace
    Do While Not wall.x1 = wall.x2
        w = firstHorizontalClip(wall, addToBuffer)
        If Not w.x1 = w.x2 Then
            
            If i > 0 Then ReDim Preserve out(i)
            out(i).x1 = w.x1
            out(i).x2 = w.x2
            wall.x1 = w.x2
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    horizontalClip = out
End Function

Private Function firstHorizontalClip(wall As wallSpace, Optional addToBuffer = True) As wallSpace
    Dim i As Integer
    
    For i = wall.x1 To wall.x2
        If Not buffer(i).wall Then Exit For
    Next i
    If i = intWidth + 1 Then Exit Function
    
    If wall.x1 < i Then i = i - 1
    firstHorizontalClip.x1 = i
    For i = i + 1 To wall.x2
        If buffer(i).wall Then
            i = i + 1
            Exit For
        End If
    Next i
    If i = wall.x2 Then i = i + 1
    firstHorizontalClip.x2 = i - 1
    If addToBuffer Then
        For i = firstHorizontalClip.x1 To firstHorizontalClip.x2
            buffer(i).wall = True
        Next i
    End If
End Function

Private Function verticalClip(ByVal start As clsVertex, ByVal x2 As Long, step As Single) As clsLine()
    Dim i As Integer
    Dim y2 As Integer
    Dim line As clsLine
    Dim vert As clsVertex
    Dim notInView As Boolean
    
    ReDim out(0) As clsLine
    
    'do not check part off the screen
    y2 = start.y - step * (x2 - start.x)
    If start.y < 0 Then
        start.x = start.x + start.y / step
        start.y = 0
    ElseIf start.y > intHeight Then
        start.x = start.x + (start.y - intHeight) / step
        start.y = intHeight
    End If
    If y2 < 0 Then
        x2 = x2 + y2 / step
    ElseIf y2 > intHeight Then
        x2 = x2 + (y2 - intHeight) / step
    End If
    
    If start.x >= x2 Then notInView = True
    If start.x > intWidth Or x2 < 0 Then notInView = True
    If notInView Then
        Set out(0) = New clsLine
        out(0).init start, start
        verticalClip = out
        Exit Function
    End If

    Do While Not start.x = x2
        Set line = firstVerticalClip(start, x2, step)
        If Not line.a Is line.b Then
            If i > 0 Then ReDim Preserve out(i)
            Set out(i) = line
            Set start = line.b
            i = i + 1
        Else
            If i = 0 Then Set out(0) = line
            Exit Do
        End If
    Loop
    verticalClip = out
End Function

Private Function firstVerticalClip(start As clsVertex, x2 As Long, step As Single) As clsLine
    Dim i As Long
    Dim y As Long
    Dim v1 As New clsVertex
    Dim line As New clsLine
    Dim v2 As New clsVertex
    
    For i = start.x To x2
        y = start.y + step * (start.x - i)
        If y < buffer(i).bottom Then
            If y > buffer(i).top Then
                Exit For
            End If
        End If
    Next i
    If i = x2 + 1 Then
        line.init start, start
        Set firstVerticalClip = line
        Exit Function
    End If
    v1.init i, y
    For i = i To x2  ' i + 1??
        y = v1.y + step * (v1.x - i)
        If y > buffer(i).bottom Or y < buffer(i).top Then
            i = i + 1
            Exit For
        End If
    Next i
    v2.init i - 1, y
    line.init v1, v2
    Set firstVerticalClip = line
End Function
Public Sub clearBuffer()
    Dim i As Integer
    ReDim buffer(intWidth)
    For i = 0 To intWidth
        buffer(i).bottom = intHeight
    Next i
End Sub

Public Function isFull() As Boolean
    Dim i As Integer
    For i = 0 To intWidth - 1
        If Not buffer(i).wall Then Exit Function
    Next i
    isFull = True
End Function

Public Sub dumpBuffer()
    Dim i As Long
    Dim line As New clsLine
    Dim v1 As New clsVertex
    Dim v2 As New clsVertex
    
    For i = 0 To intWidth
        If buffer(i).wall Then
            mkLine line.init(v1.init(i, -1000), v2.init(i, CLng(intWidth)))
        Else
            mkLine line.init(v1.init(i, CLng(buffer(i).top)), v2.init(i, CLng(buffer(i).bottom)))
        End If
    Next i
End Sub
