VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSubSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intSegCount As Integer
Private intFirstSeg As Integer


Public Sub init(SegmentCount As Integer, FirstSegID As Integer)
    intSegCount = SegmentCount
    intFirstSeg = FirstSegID
End Sub
Public Property Get segCount() As Integer
     segCount = intSegCount
End Property

Public Property Get firstSeg() As Integer
     firstSeg = intFirstSeg
End Property
