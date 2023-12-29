VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private valA As clsVertex
Private valB As clsVertex

Public Function init(a As clsVertex, b As clsVertex) As clsLine
    Set valA = a
    Set valB = b
    Set init = Me
End Function

Public Property Get a() As clsVertex
     Set a = valA
End Property

Public Property Get b() As clsVertex
     Set b = valB
End Property


'Public Property Set a(val As clsVertex)
'    Set valA = val
'End Property

'Public Property Set b(val As clsVertex)
'    Set valB = val
'End Property
