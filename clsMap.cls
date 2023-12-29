VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vertexes() As New clsVertex
Private things() As New clsThing
Private linedefs() As New clsLinedef
Private nodes() As New clsNode
Private lumps() As New clsLump
Private subSectors() As New clsSubSector
Private sidedefs() As New clsSidedef
Private segments() As New clsSegment
Private sectors() As New clsSector
Private blocks() As New clsBlock
Private bbox As clsBox ' reset when load
Private intLinedefCount As Integer
Private intNodeCount As Integer
Private objblockMapOrigin As New clsVertex
Private intBlockMapRows As Integer
Private intBlockMapColumns As Integer

Public Property Get linedef(idx As Integer) As clsLinedef
     Set linedef = linedefs(idx)
End Property

Public Property Get thing(idx As Integer) As clsThing
     Set thing = things(idx)
End Property

Public Property Get node(idx As Integer) As clsNode
     Set node = nodes(idx)
End Property

Public Property Get subSector(idx As Integer) As clsSubSector
     Set subSector = subSectors(idx)
End Property

Public Property Get segment(idx As Integer) As clsSegment
     Set segment = segments(idx)
End Property

Public Property Get sector(idx As Integer) As clsSector
     Set sector = sectors(idx)
End Property

Public Property Get block(idx As Integer) As clsBlock
     Set block = blocks(idx)
End Property

Public Property Get linedefCount() As Integer
     linedefCount = intLinedefCount
End Property

Public Property Get nodeCount() As Integer
     nodeCount = intNodeCount
End Property

Public Property Get blockMapOrigin() As clsVertex
     Set blockMapOrigin = objblockMapOrigin
End Property

Public Property Get boundryBox() As clsBox
     Set boundryBox = bbox
End Property

Public Sub load(wadPath As String, mapName As String)
    Dim i As Long
    Dim c As Long
    Dim o As Long
    Dim p As Long
    Dim line As clsLine
    Dim vertex As clsVertex
    Dim box1 As clsBox
    Dim box2 As clsBox
    Dim sidedef1 As clsSidedef
    Dim sidedef2 As clsSidedef
    Dim j As Integer
    Dim k As Integer
    
    Open wadPath For Binary Access Read As #1
    '_________________________________________________________________LUMPS
    c = readInt32(4)
    ReDim lumps(c - 1)
    o = readInt32(8)
    For i = 0 To c - 1
        lumps(i).init readInt32(o + 16 * i), readInt32(o + 16 * i + 4), readStr(o + 16 * i + 8, 8)
    Next i
    p = lumpIdx(mapName)
    '_____________________________________________________________VERTEXTES
    o = lumps(p + 4).filepos
    c = lumps(p + 4).size * 0.25
    ReDim vertexes(c - 1)
    For i = 0 To c - 1
        vertexes(i).init readInt16(o + 4 * i), readInt16(o + 4 * i + 2)
    Next i
    '____________________________________BBOX
    Set bbox = New clsBox
    bbox.left = bbox.right = vertexes(0).x
    bbox.top = bbox.bottom = vertexes(0).y
    For i = 1 To c - 1
        If vertexes(i).x < bbox.left Then bbox.left = vertexes(i).x
        If vertexes(i).x > bbox.right Then bbox.right = vertexes(i).x
        If vertexes(i).y < bbox.top Then bbox.top = vertexes(i).y
        If vertexes(i).y > bbox.bottom Then bbox.bottom = vertexes(i).y
    Next i
    '______________________________________________________________SECTORS
    o = lumps(p + 8).filepos
    c = lumps(p + 8).size / 26
    ReDim sectors(c - 1)
    For i = 0 To c - 1
        sectors(i).init readInt16(o + 26 * i), readInt16(o + 26 * i + 2)
    Next i
    '______________________________________________________________SIDEDEFS
    o = lumps(p + 3).filepos
    c = lumps(p + 3).size / 30
    ReDim sidedefs(c - 1)
    For i = 0 To c - 1
        sidedefs(i).init sectors(readInt16(o + 30 * i + 28))
    Next i
    '______________________________________________________________LINEDEFS
    o = lumps(p + 2).filepos
    c = lumps(p + 2).size / 14
    ReDim linedefs(c - 1)
    For i = 0 To c - 1
        Set line = New clsLine
        j = readInt16(o + 14 * i + 10)
        If j < 0 Then
            Set sidedef1 = Nothing
        Else
            Set sidedef1 = sidedefs(j)
        End If
        j = readInt16(o + 14 * i + 12)
        If j < 0 Then
            Set sidedef2 = Nothing
        Else
            Set sidedef2 = sidedefs(j)
        End If
        linedefs(i).init line.init(vertexes(readInt16(o + 14 * i)), vertexes(readInt16(o + 14 * i + 2))), readInt16(o + 14 * i + 6), readInt16(o + 14 * i + 8), sidedef1, sidedef2
    Next i
    
    intLinedefCount = i
    '________________________________________________________________THINGS
    o = lumps(p + 1).filepos
    c = lumps(p + 1).size / 10
    ReDim things(c - 1)
    For i = 0 To c - 1
        Set vertex = New clsVertex
        vertex.init readInt16(o + 10 * i), readInt16(o + 10 * i + 2)
        things(i).init vertex, readInt16(o + 10 * i + 4)
    Next i
    '_________________________________________________________________NODES
    o = lumps(p + 7).filepos
    c = lumps(p + 7).size / 28
    ReDim nodes(c - 1)
    For i = 0 To c - 1
        Set vertex = New clsVertex
        Set box1 = New clsBox
        Set box2 = New clsBox
        vertex.init readInt16(o + 28 * i), readInt16(o + 28 * i + 2)
        box1.init readInt16(o + 28 * i + 8), readInt16(o + 28 * i + 10), readInt16(o + 28 * i + 12), readInt16(o + 28 * i + 14)
        box2.init readInt16(o + 28 * i + 16), readInt16(o + 28 * i + 18), readInt16(o + 28 * i + 20), readInt16(o + 28 * i + 22)
        nodes(i).init vertex, readInt16(o + 28 * i + 4), readInt16(o + 28 * i + 6), box1, box2, readInt16(o + 28 * i + 24), readInt16(o + 28 * i + 26)
    Next i
    intNodeCount = i
    '______________________________________________________________SEGMENTS
    o = lumps(p + 5).filepos
    c = lumps(p + 5).size / 12
    ReDim segments(c - 1)
    For i = 0 To c - 1
        Set line = New clsLine
        line.init vertexes(readInt16(o + 12 * i)), vertexes(readInt16(o + 12 * i + 2))
        segments(i).init line, readInt16(o + 12 * i + 4), linedef(readInt16(o + 12 * i + 6)), readInt16(o + 12 * i + 8)
    Next i
    '____________________________________________________________SUBSECTORS
    o = lumps(p + 6).filepos
    c = lumps(p + 6).size * 0.25
    ReDim subSectors(c - 1)
    For i = 0 To c - 1
        subSectors(i).init readInt16(o + 4 * i), readInt16(o + 4 * i + 2)
    Next i
    '______________________________________________________________BLOCKMAP
    o = lumps(p + 10).filepos
    'c = lumps(p + 10).size * 0.25
    objblockMapOrigin.init readInt16(o), readInt16(o + 2)
    intBlockMapColumns = readInt16(o + 4)
    intBlockMapRows = readInt16(o + 6)
    c = intBlockMapRows * intBlockMapColumns
    ReDim blocks(c - 1)
    For i = 0 To c - 1
        c = readInt16(o + 8 + i * 2) * 2
        'k = readInt16(o + c) ' fisr one has always value 0;
        k = 0 ' see above
        j = 2
        Do While Not k = -1
            k = readInt16(o + c + j)
            If Not k = -1 Then
                blocks(i).addLinedef linedefs(k)
            End If
            j = j + 2
        Loop
        
    Next i
    '__________________________________________________________________DONE
    Close #1
End Sub

Function readStr(offset As Long, length As Byte) As String
    Dim str As String
    str = String(length, " ")
    Seek #1, offset + 1
    Get #1, , str
    readStr = Replace(str, Chr(0), "")
End Function

Function readInt16(offset As Long) As Integer
    Dim h As Integer
    Seek #1, offset + 1
    Get #1, , h
    readInt16 = h
End Function

Function readInt32(offset As Long) As Long
    Dim i As Long
    Seek #1, offset + 1
    Get #1, , i
    readInt32 = i
End Function

Private Function lumpIdx(name As String, Optional offset As Long) As Integer
    For lumpIdx = offset To UBound(lumps)
        If lumps(lumpIdx).name = name Then Exit For
    Next lumpIdx
End Function

Public Function findBlock(ByVal position As clsVertex) As Integer
    Dim x As Integer
    Dim y As Integer
    x = Int((position.x - objblockMapOrigin.x) / 128)
    y = Int((position.y - objblockMapOrigin.y) / 128)
    findBlock = intBlockMapColumns * y + x
End Function

Public Function checkPosition(ByVal position As clsVertex, height As Integer) As Boolean
    Dim i As Byte
    Dim j As Integer
    Dim k As Byte
    Dim block As clsBlock
    Dim blockIds(3) As Integer
    Dim hWidth As Byte
    Dim v1 As New clsVertex
    Dim v2 As New clsVertex
    Dim v3 As New clsVertex
    Dim v4 As New clsVertex
    Dim crossLines(1) As New clsLine
    Dim ceil As Integer
    Dim floor As Integer
    
    '______________________________________________________________________________________________________________off the grid completelly
    If position.x - objblockMapOrigin.x < 0 Or position.y - objblockMapOrigin.y < 0 Then
        checkPosition = False
        Exit Function
    End If
    If position.x > objblockMapOrigin.x + 128 * intBlockMapColumns Or position.y > objblockMapOrigin.y + 128 * intBlockMapRows Then
        checkPosition = False
        Exit Function
    End If
    hWidth = 0.5 * PLAYER_WIDTH
    '_____________________________________________________________________________________________________________GET IDs OF BLOCKS TO CHECK
    For i = 0 To 3
        blockIds(i) = -1
    Next i
    v1.init position.x - hWidth, position.y - hWidth
    blockIds(0) = findBlock(v1)                                                                                         'left bottom
    If (v1.x - objblockMapOrigin.x) Mod 128 >= 128 - PLAYER_WIDTH Then blockIds(1) = blockIds(0) + 1                    'right bottom
    If (v1.y - objblockMapOrigin.y) Mod 128 >= 128 - PLAYER_WIDTH Then blockIds(2) = blockIds(0) + intBlockMapColumns   'right top
    If Not blockIds(1) = -1 And Not blockIds(2) = -1 Then blockIds(3) = blockIds(2) + 1                                 'left top
    '___________________________________________________________________________________________________________PLAYER'S CROSS SECTION LINES
    v2.init position.x + hWidth, position.y + hWidth
    crossLines(0).init v1, v2
    v3.init position.x - hWidth, position.y + hWidth
    v4.init position.x + hWidth, position.y - hWidth
    crossLines(1).init v3, v4
                        
    '____________________________________________________________________________________________________________LOOP ALL LINES IN ALL BLOCK
    For i = 0 To 3
        If Not blockIds(i) = -1 Then
            Set block = blocks(blockIds(i))
            If block.lineCount > 0 Then
                For j = 0 To block.lineCount - 1
                    For k = 0 To 1
                        If lineIntersect(block.linedef(j).line, crossLines(k)) Then
                            '____________________________________________________________________STANDING IN SOLID WALL
                            If block.linedef(j).solidWall Then
                                checkPosition = False
                                Exit Function
                            End If
                            '_______________________________________________________________________STEP TO TALL TO WALK
                            floor = IIf(block.linedef(j).front.sector.floor > block.linedef(j).back.sector.floor, block.linedef(j).front.sector.floor, block.linedef(j).back.sector.floor)
                            If height - PLAYER_HEIGHT + 40 < floor Then
                                checkPosition = False
                                Exit Function
                            End If
                            '__________________________________________________PLAYER WONT FIT BETWEEN FLOOR AND CEILING
                            ceil = IIf(block.linedef(j).front.sector.ceiling < block.linedef(j).back.sector.ceiling, block.linedef(j).front.sector.ceiling, block.linedef(j).back.sector.ceiling)
                            '__________________UNTIL WE CAN@T OPEN DOORS LETS WALK TROUGH THOSE
                            If Not floor = ceil Then
                                If floor + PLAYER_HEIGHT > ceil Then
                                    checkPosition = False
                                    Exit Function
                                End If
                            End If 'door walking
                        End If
                    Next k
                Next j
            End If
        End If
    Next i
    checkPosition = True
End Function

Private Function isLineInArr(line As clsLine, arr() As clsLine)
    isLineInArr = False
End Function
