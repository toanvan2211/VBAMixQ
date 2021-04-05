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
Private mAnsPerRow As Integer
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

' So cau tra loi tren 1 hang
Public Property Get AnsPerRow() As Integer
    AnsPerRow = mAnsPerRow
End Property

Public Property Let AnsPerRow(value As Integer)
    mAnsPerRow = value
End Property
Private Sub Class_Initialize()
    mParaIndex = -1
    mCorrectAns = -1
    Set mCollRAns = New Collection
    mAnsPerRow = 0
End Sub

Public Sub MixAns()
    
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    Dim rndN, i As Integer
    
    Dim posA As Integer
    posA = 0
    posA = Int((CollRAns.Count - 2 + 1) * Rnd + 2) ' new position

    i = 0
    Dim r As Range
    Set r = wrdDoc.Range( _
        Start:=wrdDoc.Paragraphs(ParaIndex).Range.Characters(1).Start, _
        End:=wrdDoc.Paragraphs(ParaIndex).Range.Characters(1).End)
        
    r.Move Unit:=wdCharacter, Count:=-1
    r.Text = "_temp"

    For Each Item In CollRAns

        Item.Font.Underline = False
        i = i + 1

        If i + 64 = CorrectAns Then
            CorrectAns = posA + 64
        ElseIf posA + 64 = CorrectAns Then
            CorrectAns = i + 64
        End If

        Item.Copy

        r.Paste

        CollRAns(posA).Copy
        Item.Paste

        r.Copy
        CollRAns(posA).Paste

        r.Delete

        If posA > CollRAns.Count Then
            posA = 1
        End If
    Next
        i = 64
    For Each Item In CollRAns
        i = i + 1
        Item.Words(1) = Chr(i)
        If i = CorrectAns Then
            Item.Words(1).Font.Underline = True
        End If
    Next

End Sub

Public Sub MarkRedCA()
    CollRAns(CorrectAns - 64).Words(1).Font.Color = vbRed
End Sub
Public Sub UnMarkRedCA()
    CollRAns(CorrectAns - 64).Words(1).Font.Color = vbBlack
End Sub
Public Sub MarkUnderlineCA()
    CollRAns(CorrectAns - 64).Words(1).Font.Underline = True
End Sub
Public Sub UnMarkUnderlineCA()
    CollRAns(CorrectAns - 64).Words(1).Font.Underline = False
End Sub
