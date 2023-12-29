VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSidedef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objSector As clsSector

Public Function init(sector As clsSector) As clsSidedef
    Set objSector = sector
    Set init = Me
End Function

Public Property Get sector() As clsSector
     Set sector = objSector
End Property



