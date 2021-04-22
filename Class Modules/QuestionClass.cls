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
    Set mRangeQ = value
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
    Set mCollRAns = value
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
    Set r = wrdDoc.Range
        
    r.Move Unit:=wdCharacter, Count:=1
    
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
    
    Set mRangeQ = ActiveDocument.Range( _
        Start:=RangeQ.Characters(1).Start, _
        End:=CollRAns(CollRAns.Count).End)
    RangeQ.Select
    
    
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
'Them dau cham o cuoi cau tra loi
Public Sub DotMarkEndAns()
    For Each Item In CollRAns
    
        If Item.Characters(Item.Characters.Count) <> "." Then
            Item.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.Text = "."
        End If
    Next
End Sub

Public Sub ScanQuestion()
    Dim f, chrC, ansR, i, charcode As Integer
    f = 1
    ansR = 0
    charcode = 65
    For i = 1 To RangeQ.Paragraphs.Count
        chrC = 0
        If RangeQ.Paragraphs(i).Range.Words(1) = Chr(charcode) Then
            'Tim range cua cac cau tra loi va gan vao collection collRAns
            For Each ch In RangeQ.Paragraphs(i).Range.Characters
                chrC = chrC + 1
                If chrC = RangeQ.Paragraphs(i).Range.Characters.Count - 1 Then
                    'Debug.Print "Cuoi Hang o para " & i
                    Set r = ActiveDocument.Range( _
                        Start:=RangeQ.Paragraphs(i).Range.Characters(f).Start, _
                        End:=RangeQ.Paragraphs(i).Range.Characters(chrC).End)
                    
                    If r.Words(1) = Chr(charcode) Then
                        ansR = ansR + 1
                        CollRAns.Add r  'Gan tap cac range cau tra loi cho collection collRQ
                        charcode = charcode + 1
                    End If
                    Exit For
                ElseIf ch = Chr(9) Then
                    'Debug.Print "Tab o para " & i
                    If chrC - 1 = 0 Then
                        f = chrC + 1
                    ElseIf chrC - 1 > 0 Then
                        Set r = ActiveDocument.Range( _
                            Start:=RangeQ.Paragraphs(i).Range.Characters(f).Start, _
                            End:=RangeQ.Paragraphs(i).Range.Characters(chrC - 1).End)
                        
                        If r.Words(1) = Chr(charcode) Then
                            ansR = ansR + 1
                            CollRAns.Add r
                            charcode = charcode + 1
                        End If
                        f = chrC + 1
                    End If
                    
                    If f > RangeQ.Paragraphs(i).Range.Characters.Count Then
                        Exit For
                    End If
                End If
                
            Next
            If ansR > AnsPerRow Then
                AnsPerRow = ansR
            End If
        End If
        
        Dim oFound As Range
        Set oFound = RangeQ.Paragraphs(i).Range
        
        Dim oFound1 As Range
        Set oFound1 = RangeQ.Paragraphs(i).Range
        
        With oFound.Find
            .Font.Underline = True
            .Wrap = wdFindStop
        End With
        
        With oFound1.Find
            .Font.Color = vbRed
            .Wrap = wdFindStop
        End With
        
        oFound.Find.Execute
        oFound1.Find.Execute
        If oFound1.Find.Found = True Or oFound.Find.Found = True Then
            If oFound1.Find.Found = True And oFound.Find.Found = True Then
                TypeMarkCA = 3
                CorrectAns = Asc(Mid(oFound.Text, 1, 1))
            ElseIf oFound1.Find.Found = True Then
                TypeMarkCA = 2
                CorrectAns = Asc(Mid(oFound1.Text, 1, 1))
            ElseIf oFound.Find.Found = True Then
                TypeMarkCA = 1
                CorrectAns = Asc(Mid(oFound.Text, 1, 1))
            End If
        End If
        
    Next i
End Sub
