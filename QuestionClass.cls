VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuestionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mParaIndex As Integer
Private mCorrectAns As Integer
Private mCollRAns As Collection
'Doan bat dau
Public Property Get ParaIndex() As Integer
    ParaIndex = mParaIndex
End Property

Public Property Let ParaIndex(value As Integer)
    mParaIndex = value
End Property
' Dap an dung
Public Property Get CorrectAns() As Integer
    CorrectAns = mCorrectAns
End Property

Public Property Let CorrectAns(value As Integer)
    mCorrectAns = value
End Property
' Range Answer
Public Property Get CollRAns() As Collection
    Set CollRAns = mCollRAns
End Property

Public Property Let CollRAns(value As Collection)
    mCollRAns = value
End Property

Private Sub Class_Initialize()
    mParaIndex = -1
    mCorrectAns = -1
    Set mCollRAns = New Collection
End Sub
