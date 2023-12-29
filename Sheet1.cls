VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Row < 31 Then Exit Sub
    If Not WorksheetFunction.IsLogical(Target.Value) And Not Target.Value = "" Then
        Target.Value = ""
    Else
        engine.loadConfig
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
