VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private valPos As Long
Private valSize As Long
Private valName As String

Public Function init(pos As Long, size As Long, name As String)
    valPos = pos
    valSize = size
    valName = name
End Function

Public Property Get filepos() As Long
     filepos = valPos
End Property

Public Property Get size() As Long
     size = valSize
End Property

Public Property Get name() As String
     name = valName
End Property
