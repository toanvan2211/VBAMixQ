VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuestionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mRangeQ As Range
Private mParaStartIndex As Integer
Private mCorrectAns As Integer
Private mCollRAns As Collection
Private mAnsPerRow As Integer
Private mTypeMarkCA As Integer '1 gach chan, 2 la to do, 3 la ca hai
'Range Cau hoi
Public Property Get RangeQ() As Range
    Set RangeQ = mRangeQ
End Property

Public Property Let RangeQ(value As Range)
    mRangeQ = value
End Property

' Doan bat dau

Public Property Get ParaStartIndex() As Integer
    ParaStartIndex = mParaStartIndex
End Property

Public Property Let ParaStartIndex(value As Integer)
    mParaStartIndex = value
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
Public Property Get TypeMarkCA() As Integer
    TypeMarkCA = mTypeMarkCA
End Property
Public Property Let TypeMarkCA(value As Integer)
    mTypeMarkCA = value
End Property
Private Sub Class_Initialize()
    mParaIndex = -1
    mCorrectAns = -1
    Set mCollRAns = New Collection
    mAnsPerRow = 0
    mTypeMarkCA = -1
End Sub
Public Function SetRangeQ(startQ As Integer, endQ As Integer, r As Range)
    Set mRangeQ = ActiveDocument.Range( _
        Start:=r.Paragraphs(startQ).Range.Start, _
        End:=r.Paragraphs(endQ).Range.Characters(r.Paragraphs(endQ).Range.Characters.Count - 1).End)
End Function
Public Function MixAns(rangeMix As Range)
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    Dim rndN, i As Integer
    
    Dim posA As Integer
    posA = 0
    posA = Int((CollRAns.Count - 2 + 1) * Rnd + 2) ' new position

    i = 0
    
    Dim r As Range
    Set r = wrdDoc.Range( _
        Start:=rangeMix.Paragraphs(ParaStartIndex).Range.Characters(1).Start, _
        End:=rangeMix.Paragraphs(ParaStartIndex).Range.Characters(1).End)
        
    r.Move Unit:=wdCharacter, Count:=-1
    
    
    For Each Item In CollRAns

        Item.Font.Underline = False
        Item.Font.Color = vbBlack
        i = i + 1

        If i + 64 = CorrectAns Then
            CorrectAns = posA + 64
        ElseIf posA + 64 = CorrectAns Then
            CorrectAns = i + 64
        End If
        
        If posA <> i Then
            
            Item.Copy
            r.Paste
    
            CollRAns(posA).Copy
            Item.Paste
            
            r.Cut
            CollRAns(posA).Paste
        End If
        
        posA = posA + 1
        If posA > CollRAns.Count Then
            posA = 1
        End If
    Next
    
    
    i = 64
    For Each Item In CollRAns
        i = i + 1
        Item.Words(1) = Chr(i)
        If i = CorrectAns Then
            If TypeMarkCA = 1 Then
                Item.Words(1).Font.Underline = True
            ElseIf TypeMarkCA = 2 Then
                Item.Words(1).Font.Color = vbRed
            ElseIf TypeMarkCA = 3 Then
                Item.Words(1).Font.Underline = True
                Item.Words(1).Font.Color = vbRed
            End If
        End If
    Next
    
End Function

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
