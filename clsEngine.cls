VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private canvas As New clsCanvas
Private screen As New clsScreen
Private player As New clsPlayer
Private map As New clsMap
Private boolPlayerFound As Boolean
Private wadPath As String
Dim cfg As config

Public Sub init()
    Dim cell As Range
    Dim confNames As Variant
    
    wadPath = Environ("temp") & "\levels.wad"
    If Dir(wadPath) = "" Then extractMaps
    
    canvas.init 920, 50, 500, 400
    screen.init 55, 22, 800, 600
    'screen.clear
    map.load wadPath, "E1M1"
    canvas.setScale map.boundryBox
    player.init map.thing(0), canvas
    Set screen.player = player
    loadConfig
    engine.draw
End Sub

Public Sub draw()
    Dim i As Integer
    Dim j As Integer
    Dim clock As Single
    Dim line As clsLine
    Dim v1 As clsVertex
    Dim v2 As clsVertex
    clock = Timer()
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(1).DrawingObjects.Delete
    If cfg.linedefs Then
        For i = 0 To map.linedefCount - 1
            canvas.mkLine map.linedef(i).line, &HAAAAAA
        Next i
    End If
    If cfg.blockMap Then
        Set v1 = New clsVertex
        Set v2 = New clsVertex
        Set line = New clsLine
        For i = 0 To 23
            v1.init map.blockMapOrigin.x, map.blockMapOrigin.y + i * 128
            v2.init map.blockMapOrigin.x + 36 * 128, map.blockMapOrigin.y + i * 128
            line.init v1, v2
            canvas.mkLine line
        Next i
        For i = 0 To 36
            v1.init map.blockMapOrigin.x + i * 128, map.blockMapOrigin.y
            v2.init map.blockMapOrigin.x + i * 128, map.blockMapOrigin.y + 23 * 128
            line.init v1, v2
            canvas.mkLine line
        Next i
    End If
    If cfg.player Then player.draw
    boolPlayerFound = False
    screen.clearBuffer
    bsp map.nodeCount - 1
    Application.ScreenUpdating = True
    If (Timer() - clock) = 0 Then
        ThisWorkbook.Worksheets(1).Cells(2, 20).Value = "FPS"
        Exit Sub
    End If
    ThisWorkbook.Worksheets(1).Cells(2, 20).Value = CStr(Round(1 / (Timer() - clock), 1)) & " FPS"
End Sub

Public Sub keyInput(key As Byte)
    Select Case key
        Case vbKeyUp, vbKeyW, vbKeyDown, vbKeyS, vbKeyA, vbKeyD
            player.move key, 30
            If Not map.checkPosition(player.position, player.height) Then player.moveBack
        Case vbKeyLeft
            player.turn 5
        Case vbKeyRight
            player.turn -5
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            map.load wadPath, "E1M" & CStr(key)
            Set player.position = map.thing(0).position
            player.angle = map.thing(0).angle
    End Select
    draw
End Sub

Private Sub bsp(nodeID As Integer)
    Dim node As clsNode
    Dim subSector As clsSubSector
    Dim segment As clsSegment
    Dim i As Integer
    
    If screen.isFull Then Exit Sub
    If nodeID < 0 Then
        Set subSector = map.subSector(-(&H8000 - nodeID))
        For i = 0 To subSector.segCount - 1
            Set segment = map.segment(i + subSector.firstSeg)
            If Not boolPlayerFound Then
                player.height = segment.frontSector.floor + PLAYER_HEIGHT
                boolPlayerFound = True
            End If
            If player.seeSegment(segment) Then
                If cfg.segments Then
                    canvas.mkLine segment.line, IIf(segment.linedef.solidWall, vbRed, vbBlue)
                End If
                If segment.linedef.solidWall Then
                    If cfg.solidWalls Then screen.mkSolidWall segment
                ElseIf segment.backSector.ceiling = segment.backSector.floor Then
                    If cfg.solidWalls Then screen.mkSolidWall segment, True
                Else
                     If cfg.portals Then screen.mkPortal segment
                End If
            End If
        Next i
        Exit Sub
    End If
    'this is just another plain node
    Set node = map.node(nodeID)
    If node.isRight(player.position) Then
        bsp node.right
        If player.seeBox(node.lBbox) Then
            bsp node.left
        End If
    Else
        bsp node.left
        If player.seeBox(node.rBbox) Then
            bsp node.right
        End If
    End If
End Sub

Private Sub extractMaps()
    Dim ShellApp As Object
    Dim fso As Object
    Dim tmp As String
    
    tmp = Environ("temp") & "\"
    MkDir tmp & "wad"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.copyfile ThisWorkbook.FullName, tmp & "wad\wad.zip"
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace((tmp) & "wad").CopyHere ShellApp.Namespace(tmp & "wad\wad.zip").items
    Set ShellApp = Nothing
    FileCopy tmp & "wad\xl\theme\theme2.xml", wadPath
    fso.deletefolder tmp & "wad"
    Set fso = Nothing
End Sub

Public Sub loadConfig()
    With ThisWorkbook.Worksheets(1).Range("AA32:AA37")
        cfg.player = .Cells(1, 1)
        cfg.linedefs = .Cells(2, 1)
        cfg.segments = .Cells(3, 1)
        cfg.blockMap = .Cells(4, 1)
        cfg.solidWalls = .Cells(5, 1)
        cfg.portals = .Cells(6, 1)
    End With
End Sub

