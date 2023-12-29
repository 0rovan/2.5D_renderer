VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intFloor As Integer
Private intCeiling As Integer
'Private intlight As Integer

Public Function init(foorHeight As Integer, ceilingHeight As Integer) As clsSector
    intFloor = foorHeight
    intCeiling = ceilingHeight
    Set init = Me
End Function

Public Property Get floor() As Integer
     floor = intFloor
End Property

Public Property Get ceiling() As Integer
     ceiling = intCeiling
End Property

