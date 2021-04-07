Attribute VB_Name = "ModuleFunction"
Function Mix(rangeMix As Range, coll As Collection, lIndex As Integer, typeMix As Integer)

    Dim rndN, i As Integer
    Dim mRange As Range
    'Tao mang chua gia tri vi tri moi
    Dim collNewIndex As Collection
    Set collNewIndex = New Collection
    'Duyet mang & gan gia tri moi
    For i = 1 To coll.Count
        rndN = Int(1 + Rnd * (4 - 1 + 1))
        collNewIndex.Add rndN
    Next i

    If typeMix = 1 Then 'Tron cau hoi
        Dim paraC1, paraC2, paraC3, paraC4 As Integer
        paraC1 = 0
        paraC2 = 0
        paraC3 = 0
        paraC4 = 0

        For i = 1 To coll.Count
            If i = coll.Count Then
                Set mRange = ActiveDocument.Range( _
                    Start:=rangeMix.Paragraphs(coll(i).ParaIndex).Range.Start, _
                    End:=rangeMix.Paragraphs(lIndex).Range.End)
            ElseIf i <= coll.Count - 1 Then
                Set mRange = ActiveDocument.Range( _
                    Start:=rangeMix.Paragraphs(coll(i).ParaIndex).Range.Start, _
                    End:=rangeMix.Paragraphs(coll(i + 1).ParaIndex - 1).Range.End)
            End If

            Select Case collNewIndex(i)
                Case 1:
                    rangeMix.Paragraphs(lIndex + paraC1).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC1 = paraC1 + mRange.Paragraphs.Count
                    paraC2 = paraC2 + mRange.Paragraphs.Count
                    paraC3 = paraC3 + mRange.Paragraphs.Count
                    paraC4 = paraC4 + mRange.Paragraphs.Count
                Case 2:
                    rangeMix.Paragraphs(lIndex + paraC2).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC2 = paraC2 + mRange.Paragraphs.Count
                    paraC3 = paraC3 + mRange.Paragraphs.Count
                    paraC4 = paraC4 + mRange.Paragraphs.Count
                Case 3:
                    rangeMix.Paragraphs(lIndex + paraC3).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC3 = paraC3 + mRange.Paragraphs.Count
                    paraC4 = paraC4 + mRange.Paragraphs.Count
                Case 4:
                    rangeMix.Paragraphs(lIndex + paraC4).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC4 = paraC4 + mRange.Paragraphs.Count
            End Select
        Next i

        'Xoa phan cau hoi cu
        Set mRange = ActiveDocument.Range( _
            Start:=rangeMix.Paragraphs(coll(1).ParaIndex).Range.Start, _
            End:=rangeMix.Paragraphs(lIndex).Range.End)
        mRange.Select
        Selection.Delete

        i = 1
        Set mRange = ActiveDocument.Range( _
            Start:=rangeMix.Paragraphs(coll(1).ParaIndex).Range.Start, _
            End:=rangeMix.Paragraphs(lIndex).Range.End)
        
        'Sua lai stt cau hoi
        For Each Paragraph In mRange.Paragraphs
            If Paragraph.Range.Words(1) = "Câu " Then
                Paragraph.Range.Words(2).Text = i
                i = i + 1
            End If
        Next

    ElseIf typeMix = 2 Then 'Tron cau tat ca tra loi
        For Each Item In coll
            Item.MixAns
        Next
    End If
End Function

'Function duyet document tim cau hoi
Function FindQuestion(ByRef lastIndexQ As Integer, rangeFind As Range) As Collection
    
    Dim i As Integer
    ' Collection cau hoi
    Dim collQ As Collection
    Set collQ = New Collection
    
    Dim question As QuestionClass
    
    ' Duyet cac doan trong document
    For Each Paragraph In rangeFind.Paragraphs
        i = i + 1
        'Kiem tra xem doan van co phai cau hoi hay khong
        If Paragraph.Range.Words(1) = "Câu " Then
            Set question = New QuestionClass
            Dim r As Range
            Dim t, c As Integer
            t = 1
            c = 0 'So cau tra loi tim thay
            lastIndexQ = i - 1
            'Duyet cac doan cho den cau hoi tiep theo
            Do While i + t <= rangeFind.Paragraphs.Count 'Duyet den cau hoi tiep theo
                If rangeFind.Paragraphs(i + t).Range.Words(1) = "Câu " Then
                    Exit Do
                End If
                
                Dim f, chrC, ansR As Integer
                f = 1
                chrC = 0
                ansR = 0
                
                'Tim range cua cac cau tra loi va gan vao collection collRAns
                For Each ch In rangeFind.Paragraphs(i + t).Range.Characters
                    chrC = chrC + 1
                    If chrC = rangeFind.Paragraphs(i + t).Range.Characters.Count - 1 Then
                        'Debug.Print "Cuoi Hang o para " & i + t
                        Set r = ActiveDocument.Range( _
                            Start:=rangeFind.Paragraphs(i + t).Range.Characters(f).Start, _
                            End:=rangeFind.Paragraphs(i + t).Range.Characters(chrC).End)
                        
                        
                        If r.Words(1) = Chr(65 + c) Then
                            ansR = ansR + 1
                            question.CollRAns.Add r  'Gan tap cac range cau tra loi cho collection collRQ
                            c = c + 1
                            lastIndexQ = i + t
                        End If
                        Exit For
                    ElseIf ch = Chr(9) Then
                        'Debug.Print "Tab o para " & i + t
                        If chrC - 1 = 0 Then
                            f = chrC + 1
                        ElseIf chrC - 1 > 0 Then
                            Set r = ActiveDocument.Range( _
                                Start:=rangeFind.Paragraphs(i + t).Range.Characters(f).Start, _
                                End:=rangeFind.Paragraphs(i + t).Range.Characters(chrC - 1).End)
                                           
                            If r.Words(1) = Chr(65 + c) Then
                                ansR = ansR + 1
                                question.CollRAns.Add r
                                c = c + 1
                                lastIndexQ = i + t
                            End If
                            f = chrC + 1
                        End If
                        
                        If f > rangeFind.Paragraphs(i + t).Range.Characters.Count Then
                            Exit For
                        End If
                    End If
                Next
                If ansR > question.AnsPerRow Then
                    question.AnsPerRow = ansR
                End If
                t = t + 1
            Loop
            
            If c >= 2 Then
                question.ParaIndex = i
                
                Dim lo As Integer
                For lo = i + 1 To i + t - 1
                    Dim oFound As Range
                    Set oFound = rangeFind.Paragraphs(lo).Range
                    
                    Dim oFound1 As Range
                    Set oFound1 = rangeFind.Paragraphs(lo).Range
                    
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
                            question.TypeMarkCA = 3
                            question.CorrectAns = Asc(Mid(oFound.Text, 1, 1))
                        ElseIf oFound1.Find.Found = True Then
                            question.TypeMarkCA = 2
                            question.CorrectAns = Asc(Mid(oFound1.Text, 1, 1))
                        ElseIf oFound.Find.Found = True Then
                            question.TypeMarkCA = 1
                            question.CorrectAns = Asc(Mid(oFound.Text, 1, 1))
                        End If
                        Exit For
                    End If
                Next
                collQ.Add question
            End If
        End If
    Next
    
    Set FindQuestion = collQ
End Function
