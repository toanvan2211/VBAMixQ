Attribute VB_Name = "ModuleFunction"
Function Mix(rangeMix As Range, coll As Collection, lIndex As Integer, typeMix As Integer)
    Dim rndN, i As Integer
    Dim mRange As Range
    
    If typeMix = 1 Then 'Tron cau hoi
        Dim doc As Document
        Set doc = New Document
        doc.ActiveWindow.Visible = False
        Dim tempR As Range
        Set tempR = doc.Range
        tempR.Move Unit:=wdCharacter, Count:=1
        i = 0
        For Each Item In coll
            i = i + 1
            rndN = Int(coll.Count * Rnd) + 1
            If rndN <> i Then
                Dim tempAns As Integer
                
            
                Item.RangeQ.Copy
                tempAns = Item.CorrectAns
                tempR.Paste
    
                coll(rndN).RangeQ.Copy
                Item.CorrectAns = coll(rndN).CorrectAns
                Item.RangeQ.Paste
                
                tempR.Cut
                coll(rndN).RangeQ.Paste
                coll(rndN).CorrectAns = tempAns
            End If
        Next
        doc.Close SaveChanges:=wdDoNotSaveChanges
        i = 1
        'Sua lai stt cau hoi
        For Each Item In coll
            If Item.RangeQ.Words(1) = "C�u " Then
                Item.RangeQ.Words(2).Text = i
                i = i + 1
            End If
        Next
    ElseIf typeMix = 2 Then 'Tron cau tat ca tra loi
        For Each Item In coll
            Call Item.MixAns(rangeMix)
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
    Dim cPQ As Integer
    
    ' Duyet cac doan trong document
    For cPQ = 1 To rangeFind.Paragraphs.Count
        i = cPQ
        'Kiem tra xem doan van co phai cau hoi hay khong
        If rangeFind.Paragraphs(cPQ).Range.Words(1) = "C�u " Then
            Set question = New QuestionClass
            Dim r As Range
            Dim T, c, endQ As Integer
            T = 1
            c = 0 'So cau tra loi tim thay
            lastIndexQ = i - 1
            'Duyet cac doan cho den cau hoi tiep theo
            Do While i + T <= rangeFind.Paragraphs.Count 'Duyet den cau hoi tiep theo
                cPQ = cPQ + 1
                If rangeFind.Paragraphs(i + T).Range.Words(1) = "C�u " Then
                    cPQ = i + T - 1
                    Exit Do
                End If
                
                Dim f, chrC, ansR As Integer
                f = 1
                chrC = 0
                ansR = 0
                
                'Tim range cua cac cau tra loi va gan vao collection collRAns
                For Each ch In rangeFind.Paragraphs(i + T).Range.Characters
                    chrC = chrC + 1
                    If chrC = rangeFind.Paragraphs(i + T).Range.Characters.Count - 1 Then
                        'Debug.Print "Cuoi Hang o para " & i + t
                        Set r = ActiveDocument.Range( _
                            Start:=rangeFind.Paragraphs(i + T).Range.Characters(f).Start, _
                            End:=rangeFind.Paragraphs(i + T).Range.Characters(chrC).End)
                        
                        If r.Words(1) = Chr(65 + c) Then
                            endQ = i + T
                            ansR = ansR + 1
                            question.CollRAns.Add r  'Gan tap cac range cau tra loi cho collection collRQ
                            c = c + 1
                            lastIndexQ = i + T
                        End If
                        Exit For
                    ElseIf ch = Chr(9) Then
                        'Debug.Print "Tab o para " & i + t
                        If chrC - 1 = 0 Then
                            f = chrC + 1
                        ElseIf chrC - 1 > 0 Then
                            Set r = ActiveDocument.Range( _
                                Start:=rangeFind.Paragraphs(i + T).Range.Characters(f).Start, _
                                End:=rangeFind.Paragraphs(i + T).Range.Characters(chrC - 1).End)
                            
                            If r.Words(1) = Chr(65 + c) Then
                                endQ = i + T
                                ansR = ansR + 1
                                question.CollRAns.Add r
                                c = c + 1
                                lastIndexQ = i + T
                            End If
                            f = chrC + 1
                        End If
                        
                        If f > rangeFind.Paragraphs(i + T).Range.Characters.Count Then
                            Exit For
                        End If
                    End If
                    
                Next
                If ansR > question.AnsPerRow Then
                    question.AnsPerRow = ansR
                End If
                T = T + 1
                If cPQ < rangeFind.Paragraphs.Count Then
                    cPQ = cPQ - 1
                End If
            Loop
            
            If c >= 2 Then
                question.ParaStartIndex = i
                Call question.SetRangeQ(i, endQ, rangeFind)
                
                Dim lo As Integer
                For lo = i + 1 To i + T - 1
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
        
    Next cPQ
    
    Set FindQuestion = collQ
End Function
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
        If Item.Range.Words(1) = "C�u " Then
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

'Tao bang dap an
Function CreateTableOfAns(coll As Collection, wrdDoc As Document) As Object
    
    Dim objTable
    
    
    intNoOfRows = Fix(coll.Count / 4)
    If coll.Count Mod 4 <> 0 Then
        intNoOfRows = intNoOfRows + 1
    End If
    
    
    intNoOfColumns = 4
    

    Set objRange = wrdDoc.Range( _
        Start:=wrdDoc.Paragraphs(wrdDoc.Paragraphs.Count).Range.Characters(wrdDoc.Paragraphs(wrdDoc.Paragraphs.Count).Range.Characters.Count).Start, _
        End:=wrdDoc.Paragraphs(wrdDoc.Paragraphs.Count).Range.Characters(wrdDoc.Paragraphs(wrdDoc.Paragraphs.Count).Range.Characters.Count).End)
        

    wrdDoc.Tables.Add objRange, intNoOfRows, intNoOfColumns

    Set objTable = wrdDoc.Tables(wrdDoc.Tables.Count)

    objTable.Borders.Enable = True
    Dim ansCount As Integer
    ansCount = 0
    For i = 1 To intNoOfRows
        For j = 1 To intNoOfColumns
            ansCount = ansCount + 1
            If ansCount > coll.Count Then
                Exit For
            End If
            objTable.Cell(i, j).Range.Text = ansCount & ". " & Chr(coll(ansCount).CorrectAns)
        Next
    Next
    
    Set CreateTableOfAns = objTable
    
End Function
