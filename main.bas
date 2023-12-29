Attribute VB_Name = "main"
Option Explicit
Public Const FOV = 90
Public Const HFOV = FOV * 0.5
Public Const PLAYER_HEIGHT = 56
Public Const PLAYER_WIDTH = 31
Public engine As New clsEngine

Public Type config
    player As Boolean
    linedefs As Boolean
    segments As Boolean
    solidWalls As Boolean
    portals As Boolean
    blockMap As Boolean
End Type

Private Sub key_up()
    engine.keyInput vbKeyUp
End Sub
Private Sub key_down()
    engine.keyInput vbKeyDown
End Sub
Private Sub key_left()
    engine.keyInput vbKeyLeft
End Sub
Private Sub key_right()
    engine.keyInput vbKeyRight
End Sub
Private Sub key_w()
    engine.keyInput vbKeyW
End Sub
Private Sub key_s()
    engine.keyInput vbKeyS
End Sub
Private Sub key_a()
    engine.keyInput vbKeyA
End Sub
Private Sub key_d()
    engine.keyInput vbKeyD
End Sub
Private Sub key_1()
    engine.keyInput 1
End Sub
Private Sub key_2()
    engine.keyInput 2
End Sub
Private Sub key_3()
    engine.keyInput 3
End Sub
Private Sub key_4()
    engine.keyInput 4
End Sub
Private Sub key_5()
    engine.keyInput 5
End Sub
Private Sub key_6()
    engine.keyInput 6
End Sub
Private Sub key_7()
    engine.keyInput 7
End Sub
Private Sub key_8()
    engine.keyInput 8
End Sub
Private Sub key_9()
    engine.keyInput 9
End Sub
Sub StartUp()
    Dim i As Integer
    Dim s As String * 1
    
    Application.OnKey "{UP}", "key_up"
    Application.OnKey "{DOWN}", "key_down"
    Application.OnKey "{LEFT}", "key_left"
    Application.OnKey "{RIGHT}", "key_right"
    Application.OnKey "w", "key_w"
    Application.OnKey "s", "key_s"
    Application.OnKey "a", "key_a"
    Application.OnKey "d", "key_d"
    For i = 1 To 9
        s = CStr(i)
        Application.OnKey s, "key_" & s
    Next i
    Application.OnKey "{ESC}", "halt"
    With ThisWorkbook.Worksheets(1)
        .Unprotect
        With .Range("T34:W42")
            .Font.color = 0
        End With
        .Protect , False
    End With
    
    engine.init
End Sub

Sub halt()
    Dim i As Integer
    Set engine = Nothing
    
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "w"
    Application.OnKey "s"
    Application.OnKey "a"
    Application.OnKey "d"
    For i = 1 To 9
        Application.OnKey CStr(i)
    Next i
    Application.OnKey "{ESC}"
    With ThisWorkbook.Worksheets(1)
        With ThisWorkbook.Worksheets(1).Buttons.add(952, 550, 114, 41)
            .OnAction = "StartUp"
            With .Characters
                .Text = "START"
                With .Font
                    .name = "Calibri"
                    .size = 12
                End With
            End With
        End With
        .Unprotect
        With .Range("T34:W42")
            .Font.color = .Interior.color
        End With
        .Protect
    End With
End Sub


