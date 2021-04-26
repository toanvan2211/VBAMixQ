Attribute VB_Name = "ModuleScanDocument"
'Ham nay chi tra ve collectionQ chua gia tri RangeQ va lastIndex, uu diem la duyet nhanh
Function FindQuestion_v2(ByRef lastIndexQ As Integer, rangeFind As Range) As Collection
    Dim i, cPQ, charcode As Integer
    ' Collection cau hoi
    Dim collQ As Collection
    Set collQ = New Collection
    
    Dim question As QuestionClass
    
    Dim r As Range
    i = 0
    cPQ = -1
    charcode = 65
    For Each Item In rangeFind.Paragraphs
        i = i + 1
        If Item.Range.Words(1) = "Câu " Then
            If cPQ <> -1 Then
                lastIndexQ = i - 1
                Set question = New QuestionClass
                Set r = ActiveDocument.Range( _
                    Start:=rangeFind.Paragraphs(cPQ).Range.Start, _
                    End:=rangeFind.Paragraphs(lastIndexQ).Range.Characters(rangeFind.Paragraphs(lastIndexQ).Range.Characters.Count - 1).End)
                question.RangeQ = r
                question.ParaStartIndex = cPQ
                collQ.Add question
                charcode = 65
            End If
            cPQ = i
        Else
            Dim c As Integer
            For c = 0 To 4
                If Item.Range.Words(1) = Chr(c + charcode) Then
                    charcode = charcode + c
                    lastIndexQ = i
                End If
            Next c
        End If
    Next
    
    Set question = New QuestionClass
    Set r = ActiveDocument.Range( _
        Start:=rangeFind.Paragraphs(cPQ).Range.Start, _
        End:=rangeFind.Paragraphs(lastIndexQ).Range.Characters(rangeFind.Paragraphs(lastIndexQ).Range.Characters.Count - 1).End)
    question.RangeQ = r
    question.ParaStartIndex = cPQ
    collQ.Add question
    
    Set FindQuestion_v2 = collQ
    
End Function
'Print to document the value of start paragraphs index in special format, last of string of value is index of paragraph belong to last question
Sub SaveValueForIndex()
    ActiveDocument.Range.Select
    Selection.Move Unit:=wdCharacter, Count:=-1
    Selection.TypeText ActiveDocument.Range.Paragraphs.Count + 3
    Selection.TypeParagraph
    ActiveDocument.Range.Select
    Selection.Move Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Selection.TypeParagraph
    Dim coll As Collection
    Dim lIndex As Integer
    Set coll = FindQuestion_v2(lIndex, ActiveDocument.Range)
    For Each Item In coll
        Selection.TypeText Item.ParaStartIndex & ";"
    Next
    Selection.TypeText lIndex + 1
End Sub
'Read value from the document and convert it into array contain start paragraph index, lastindex is end of last question
Function GetArrayIndexParaQuestion() As Variant
    Dim tempArr() As String
    Dim startValueIndex As Integer
    Dim stringValueIndex As String
    
    startValueIndex = ActiveDocument.Paragraphs(1).Range.Text
    stringValueIndex = ActiveDocument.Paragraphs(startValueIndex).Range.Text
    tempArr = Split(stringValueIndex, ";")
    GetArrayIndexParaQuestion = tempArr
End Function
'Convert array to collection of questionClass
Function ConvertArrayToCollQ() As Collection
    Dim arr() As String
    Dim i As Integer
    
    arr = GetArrayIndexParaQuestion
    Dim coll As Collection
    Set coll = New Collection
    Dim question As QuestionClass
    For i = 0 To UBound(arr) - 1
        Set question = New QuestionClass
        
        question.ParaStartIndex = CInt(arr(i))
        question.SetRangeQ CInt(arr(i)), CInt(arr(i + 1) - 1), ActiveDocument.Range
        coll.Add question
    Next i
    Set ConvertArrayToCollQ = coll
    
End Function
Sub T2()
    Dim coll As Collection
    Set coll = ConvertArrayToCollQ
    Debug.Print coll.Count
End Sub

Sub T()
    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion_v2(lIndex, rangeFind)
    Debug.Print coll.Count
End Sub
'Just mix the question index of whole document
Sub MixQuestionFromDataIndex()
    Dim rndN, i As Integer
    Dim tempR As Range
    Set tempR = ActiveDocument.Range
    tempR.Move Unit:=wdCharacter, Count:=1
    
    Dim coll As Collection
    Set coll = ConvertArrayToCollQ()
    For Each Item In coll
            i = i + 1
            rndN = Int(coll.Count * Rnd) + 1
            If rndN <> i Then
                Item.RangeQ.Copy
                tempR.Paste
    
                coll(rndN).RangeQ.Copy
                Item.RangeQ.Paste
                
                tempR.Cut
                coll(rndN).RangeQ.Paste
            End If
        Next
    i = 1
    'Sua lai stt cau hoi
    For Each Item In coll
        If Item.RangeQ.Words(1) = "Câu " Then
            Item.RangeQ.Words(2).Text = i
            i = i + 1
        End If
    Next
End Sub
'Because after mix the question, index paragraphs of question will be change, so need this function to update the string value at end of docum
Function UpdateStringValueIndex(coll)
    
End Function
