VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLinedef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objLine As clsLine
Private objFront As clsSidedef
Private objBack As clsSidedef
Private boolSolidWall As Boolean
Private intSpecial As Byte
Private intSector As Integer

Public Sub init(line As clsLine, special, sector, Optional frontSidedef As clsSidedef, Optional backSidedef As clsSidedef)
    Set objLine = line
    intSpecial = special
    intSector = sector
    Set objFront = frontSidedef
    Set objBack = backSidedef
    boolSolidWall = frontSidedef Is Nothing Or backSidedef Is Nothing
End Sub

Public Property Get line() As clsLine
     Set line = objLine
End Property


Public Property Get front() As clsSidedef
     Set front = objFront
End Property

Public Property Get back() As clsSidedef
     Set back = objBack
End Property

Public Property Get solidWall() As Boolean
     solidWall = boolSolidWall
End Property
