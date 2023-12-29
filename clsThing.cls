VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsThing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vertPosition As clsVertex
Private intAngle As Integer
Private valFlag As Byte

Public Sub init(vertex As clsVertex, angle As Integer) ', typ As Byte, flags As Byte)
    Set vertPosition = vertex
    intAngle = angle
End Sub

Public Property Get position() As clsVertex
     Set position = vertPosition
End Property

Public Property Get angle() As Integer
     angle = intAngle
End Property

