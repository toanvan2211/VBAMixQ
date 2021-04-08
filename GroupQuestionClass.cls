VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupQuestionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHeader As Range
Private mCollQ As Collection

Public Property Get Header() As Range
    Set Header = mHeader
End Property

Public Property Let Header(value As Range)
    Set mHeader = value
End Property

Public Property Get CollQ() As Collection
    Set CollQ = mCollQ
End Property

Public Property Let CollQ(value As Collection)
    Set mCollQ = value
End Property

Private Sub Class_Initialize()
    
End Sub

