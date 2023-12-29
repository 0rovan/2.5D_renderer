VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vertPosition As clsVertex
Private lngDx As Long
Private lngDy As Long
Private r_BBox As clsBox
Private l_BBox As clsBox
Private lelfChild As Integer
Private rightChild As Integer

Public Sub init(vertex As clsVertex, ch_x As Long, ch_y As Long, rBbox As clsBox, lBbox As clsBox, rightID As Integer, leftID As Integer)
    Set vertPosition = vertex
    lngDx = ch_x
    lngDy = ch_y
    Set r_BBox = rBbox
    Set l_BBox = lBbox
    lelfChild = leftID
    rightChild = rightID
End Sub

Public Property Get left() As Integer
     left = lelfChild
End Property

Public Property Get right() As Integer
     right = rightChild
End Property

Public Property Get lBbox() As clsBox
     Set lBbox = l_BBox
End Property

Public Property Get rBbox() As clsBox
     Set rBbox = r_BBox
End Property

Public Function isRight(vertex As clsVertex) As Boolean
    Dim dx As Long
    Dim dy As Long
    
    'as simple as cross product property?
    dx = vertex.x - vertPosition.x
    dy = vertex.y - vertPosition.y
    isRight = dx * lngDy - dy * lngDx > 0
End Function
