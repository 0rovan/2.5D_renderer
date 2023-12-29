VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intLineCount As Integer
Private linedefs() As clsLinedef

'Private objOrigin As New clsVertex
'Private intColumnCount As Integer
'Private intRowCount As Integer


Public Sub addLinedef(linedef As clsLinedef)
    ReDim Preserve linedefs(lineCount)
    Set linedefs(lineCount) = linedef
    intLineCount = intLineCount + 1
End Sub



'Public Property Let right(val As Long)
'    lngRight = val
'End Property''

Public Property Get lines() As clsLinedef()
    lines = linedefs
End Property

Public Property Get lineCount() As Integer
    lineCount = intLineCount
End Property

Public Property Get linedef(idx As Integer) As clsLinedef
    Set linedef = linedefs(idx)
End Property

'Public Property Get rowCount() As Integer
'    rowCount = intRowCount
'End Property


'Public Property Get columnCount() As Integer
'    columnCount = intColumnCount
'End Property
