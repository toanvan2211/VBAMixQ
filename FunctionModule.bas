Attribute VB_Name = "ModuleFunction"
Function Mix(coll As Collection, typeMix As Integer)
    Dim rndN, i As Integer
    Dim mRange As Range
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    'Tao mang chua gia tri vi tri moi
    Dim collNewIndex As Collection
    Set collNewIndex = New Collection
    'Duyet mang & gan gia tri moi
    For i = 1 To coll.Count
        rndN = Int(1 + rnd * (4 - 1 + 1))
        collNewIndex.Add rndN
    Next i

    If typeMix = 1 Then 'Tron cau hoi
        Dim paraC1, paraC2, paraC3, paraC4 As Integer
        paraC1 = 0
        paraC2 = 0
        paraC3 = 0
        paraC4 = 0

        For i = 1 To coll.Count - 1
            If i = coll.Count - 1 Then
                Set mRange = wrdDoc.Range( _
                    Start:=wrdDoc.Paragraphs(coll(i)).Range.Start, _
                    End:=wrdDoc.Paragraphs(coll(i + 1)).Range.End)
            ElseIf i < coll.Count - 1 Then
                Set mRange = wrdDoc.Range( _
                    Start:=wrdDoc.Paragraphs(coll(i)).Range.Start, _
                    End:=wrdDoc.Paragraphs(coll(i + 1) - 1).Range.End)
            End If

            Select Case collNewIndex(i)
                Case 1:
                    wrdDoc.Paragraphs(coll(coll.Count) + paraC1).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC1 = paraC1 + mRange.Paragraphs.Count
                    paraC2 = paraC2 + mRange.Paragraphs.Count
                    paraC3 = paraC3 + mRange.Paragraphs.Count
                    paraC4 = paraC4 + mRange.Paragraphs.Count
                Case 2:
                    wrdDoc.Paragraphs(coll(coll.Count) + paraC2).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC2 = paraC2 + mRange.Paragraphs.Count
                    paraC3 = paraC3 + mRange.Paragraphs.Count
                    paraC4 = paraC4 + mRange.Paragraphs.Count
                Case 3:
                    wrdDoc.Paragraphs(coll(coll.Count) + paraC3).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC3 = paraC3 + mRange.Paragraphs.Count
                    paraC4 = paraC4 + mRange.Paragraphs.Count
                Case 4:
                    wrdDoc.Paragraphs(coll(coll.Count) + paraC4).Range.Select
                    Selection.MoveRight
                    mRange.Copy
                    Selection.Paste
                    paraC4 = paraC4 + mRange.Paragraphs.Count
            End Select
        Next i

        'Xoa phan cau hoi cu
        Set mRange = wrdDoc.Range( _
            Start:=wrdDoc.Paragraphs(coll(1)).Range.Start, _
            End:=wrdDoc.Paragraphs(coll(coll.Count)).Range.End)
        mRange.Select
        Selection.Delete

        i = 1
        Set mRange = wrdDoc.Range( _
            Start:=wrdDoc.Paragraphs(coll(1)).Range.Start, _
            End:=wrdDoc.Paragraphs(coll(coll.Count)).Range.End)
        
        'Sua lai stt cau hoi
        For Each Paragraph In mRange.Paragraphs
            Debug.Print Paragraph.Range.Words(1)
            If Paragraph.Range.Words(1) = "Câu " Then
                Paragraph.Range.Words(2).Text = i
                i = i + 1
            End If
        Next

    ElseIf typeMix = 2 Then 'Tron cau tra loi

    End If
End Function
Function FindAnswer(fIndex As Integer, lIndex As Integer) As Collection
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    'Tao bien chua collection range cua tung cau tra loi
    Dim collA As Collection
    Set collA = New Collection
    
    Dim i, aC As Integer
    aC = 0
    For i = fIndex To lIndex
        Dim tabA As Variant
        'Duyet cac cau tra loi cung hang
        tabA = Split(wrdDoc.Paragraphs(i).Range, vbTab)
        For Each a In tabA
            If Left(a, 1) = Chr(65 + c) Then
                c = c + 1
                lastIndexQ = i + t
            End If
        Next
    Next i
    
End Function
'Function duyet document tim cau hoi
Function FindQuestion(ByRef collCorrectAns As Collection, ByRef collRAnsQ As Collection) As Collection
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    Dim i As Integer
    ' Collection cau hoi
    Dim collQ As Collection
    Set collQ = New Collection
    ' Collection cau tra loi dung
    Dim collC As Collection
    Set collC = New Collection
    'Collection range cau hoi
    Dim collRQ As Collection
    Set collRQ = New Collection
    
    Dim lastIndexQ As Integer
    ' Duyet cac doan trong document
    For Each Paragraph In wrdDoc.Paragraphs
        i = i + 1
        'Kiem tra xem doan van co phai cau hoi hay khong
        If Paragraph.Range.Words(1) = "Câu " Then
            Dim collRAns As Collection
            Set collRAns = New Collection
            Dim r As Range
            Dim t, c As Integer
            t = 1
            c = 0 'So cau tra loi tim thay
            lastIndexQ = i - 1
            'Duyet cac doan cho den cau hoi tiep theo
            Do While i + t <= wrdDoc.Paragraphs.Count 'Duyet den cau hoi tiep theo
                If wrdDoc.Paragraphs(i + t).Range.Words(1) = "Câu " Then
                    Exit Do
                End If
                
                
                Dim f, chrC As Integer
                f = 1
                chrC = 0
                
                'Tim range cua cac cau tra loi va gan vao collection collRAns
                For Each ch In wrdDoc.Paragraphs(i + t).Range.Characters
                    chrC = chrC + 1
                    If chrC = wrdDoc.Paragraphs(i + t).Range.Characters.Count Then
                        'Debug.Print "Cuoi Hang o para " & i + t
                        Set r = wrdDoc.Range( _
                            Start:=wrdDoc.Paragraphs(i + t).Range.Characters(f).Start, _
                            End:=wrdDoc.Paragraphs(i + t).Range.Characters(chrC).End)
                        
                        collRAns.Add r
                        If collRAns(collRAns.Count).Words(1) = Chr(65 + c) Then
                            c = c + 1
                            lastIndexQ = i + t
                        Else
                            collRAns.Remove (collRAns.Count)
                        End If
                        Exit For
                    ElseIf ch = Chr(9) Then
                        'Debug.Print "Tab o para " & i + t
                        If chrC - 1 = 0 Then
                            f = chrC + 1
                        ElseIf chrC - 1 > 0 Then
                            Set r = wrdDoc.Range( _
                                Start:=wrdDoc.Paragraphs(i + t).Range.Characters(f).Start, _
                                End:=wrdDoc.Paragraphs(i + t).Range.Characters(chrC - 1).End)
                                                            
                            collRAns.Add r
                            If collRAns(collRAns.Count).Words(1) = Chr(65 + c) Then
                                c = c + 1
                                lastIndexQ = i + t
                            Else
                                collRAns.Remove (collRAns.Count)
                            End If
                            f = chrC + 1
                        End If
                        
                        If f > wrdDoc.Paragraphs(i + t).Range.Characters.Count Then
                            Exit For
                        End If
                    End If
                Next
                t = t + 1
            Loop
            
            If c >= 2 Then
                collQ.Add i
                collRQ.Add collRAns 'Gan tap cac range cau tra loi cho collection collRQ
                Dim lo As Integer
                For lo = i + 1 To i + t - 1
                    Dim oFound As Range
                    Set oFound = wrdDoc.Paragraphs(lo).Range
                    With oFound.Find
                        .Font.Underline = True
                        .Wrap = wdFindStop
                    End With
                    oFound.Find.Execute
                    If oFound.Find.Found = True Then
                        collC.Add CStr(Asc(Mid(oFound.Text, 1, 1)))
                        Exit For
                    End If
                Next
                
            End If
        End If
    Next
    Set collCorrectAns = collC
    Set collRAnsQ = collRQ
    collQ.Add lastIndexQ
    Set FindQuestion = collQ
End Function
Sub Test()

    Dim collQ As Collection
    Dim collCA As Collection
    Dim collRAnsQ As Collection
    

    Set collQ = FindQuestion(collCA, collRAnsQ)
    'Call Mix(collQ, 1)
    
    Dim i As Integer
    For i = 1 To collQ.Count - 1
        Debug.Print "Question Para Index: " & collQ(i) & ", Answer Correct: " & Chr(collCA(i))
    Next

End Sub

Sub FindUnderlineInDoc()
    
    Dim oFound As Range
    Set oFound = ActiveDocument.Content
    Dim value As String
    value = "nothing"
    With oFound.Find
        .Font.Underline = True
        .Wrap = wdFindStop
        .Execute
    End With
    Debug.Print oFound.Text & " - " & CStr(Asc(Mid(oFound.Text, 1, 1)))
End Sub

Sub Format()
    ActiveDocument.Paragraphs.TabStops.ClearAll
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(0.5)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(4.77)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(9.07)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(13.36)
End Sub
