Attribute VB_Name = "NewMacros"
Sub MixAnswer()
' Tron cau tra loi
Dim arr1(4)
arr1(0) = ActiveDocument.Paragraphs(1).Range.Words(1)
arr1(1) = ActiveDocument.Paragraphs(2).Range.Words(1)
arr1(2) = ActiveDocument.Paragraphs(3).Range.Words(1)
arr1(3) = ActiveDocument.Paragraphs(4).Range.Words(1)

Dim oRang As Range
Set oRang = ActiveDocument.Paragraphs(5).Range
Dim iAnswer As Integer, iPosition As Integer
iAnswer = Int((3 - 0 + 1) * rnd + 0)
iPosition = Int((3 - 0 + 1) * rnd + 0)

If iAnswer = iPosition Then
 iPosition = iPosition + 1
 If iPosition > 3 Then
    iPosition = 0
    End If
End If
'oRang(0).Text = arr1(iAnswer)
'oRang(0).Bold = True
Dim i As Integer
'Selection.TypeText "iAnswer: " & iAnswer & " - iPosition: " & iPosition & "." & vbCrLf
Dim aPick As Integer
For i = 1 To 4 Step 1
    Dim arr(4) As Integer
    arr(0) = -1
    arr(iPosition) = iAnswer
    iPosition = iPosition + 1
    iAnswer = iAnswer + 1
    If iPosition > 3 Then
    iPosition = 0
    End If
    If iAnswer > 3 Then
    iAnswer = 0
    End If
    If arr(0) <> -1 Then
        aPick = arr(0)
        Exit For
    End If
Next i
For i = 1 To 4 Step 1
    Selection.TypeText Chr(64 + i) & ". "
    Selection.TypeText arr1(aPick) & vbCrLf
    aPick = aPick + 1
    If aPick > 3 Then
    aPick = 0
    End If
Next i
End Sub
Function TronCauTraLoi(arrA(), a) As Variant
    'Tao bien de random
    Dim rndN As Integer
    Dim i, o As Integer
    
    Dim collRandomSortAnswer1 As Collection
    Set collRandomSortAnswer1 = New Collection
    Dim collRandomSortAnswer2 As Collection
    Set collRandomSortAnswer2 = New Collection
    Dim collRandomSortAnswer3 As Collection
    Set collRandomSortAnswer3 = New Collection
    Dim collRandomSortAnswer4 As Collection
    Set collRandomSortAnswer4 = New Collection
    
    For i = 0 To a - 1 Step 1
       rndNumber = Int(1 + rnd * (4 - 1 + 1))
       'Luu vao collection vi tri cua cac cau tra loi
       If rndNumber = 1 Then
            collRandomSortAnswer1.Add i
        ElseIf rndNumber = 2 Then
            collRandomSortAnswer2.Add i
        ElseIf rndNumber = 3 Then
            collRandomSortAnswer3.Add i
        Else
            collRandomSortAnswer4.Add i
        End If
    Next i
    o = 0
    'Tao array temp
    Dim arrTemp()
    ReDim arrTemp(UBound(arrA, 1))
    'in ra cau hoi voi thu tu da tron
    For i = 1 To collRandomSortAnswer1.Count Step 1
        arrTemp(o) = arrA(collRandomSortAnswer1(i))
        o = o + 1
    Next i
    For i = 1 To collRandomSortAnswer2.Count Step 1
        arrTemp(o) = arrA(collRandomSortAnswer2(i))
        o = o + 1
    Next i
    For i = 1 To collRandomSortAnswer3.Count Step 1
        arrTemp(o) = arrA(collRandomSortAnswer3(i))
        o = o + 1
    Next i
    For i = 1 To collRandomSortAnswer4.Count Step 1
        arrTemp(o) = arrA(collRandomSortAnswer4(i))
        o = o + 1
    Next i
    TronCauTraLoi = arrTemp
End Function


Sub PrintTheRandomParagraphs()
Dim rndPara As Integer, paraCnt As Integer
paraCnt = ActiveDocument.Paragraphs.Count

Dim myRange As Range
Set myRange = ActiveDocument.Content
myRange.Font.Bold = False
myRange.Font.Color = wdColorBlack


rndPara = Int((paraCnt * rnd) + 1)
Dim paraRange As Range
Set paraRange = ActiveDocument.Paragraphs(rndPara).Range
paraRange.Copy
paraRange.Font.Bold = True
paraRange.Font.Color = wdColorRed
Selection.TypeText vbCrLf
Selection.Paste

End Sub

Sub FindAndMarkTheKeyWord()
Dim keyWord As String
Dim myRange As Word.Range
keyWord = "a"
Set myRange = ActiveDocument.Content
With myRange.Find
    .Text = keyWord
    .Forward = True
    .Execute
    If .Found = True Then .Parent.Bold = True
    .Parent.Font.Color = wdColorRed
End With
Selection.Find.Execute
End Sub
Sub TronCauHoi_Temp()
Dim arr(3, 3) As Integer
' Index of Answer
Dim arrIndexAns As Collection
Set arrIndexAns = New Collection

Dim rnd As Integer, i As Integer, o As Integer
o = 0
i = 0
Dim wrdDoc As Document
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Set wrdDoc = Application.ActiveDocument
For Each Paragraph In ActiveDocument.Paragraphs
    i = i + 1
    If Paragraph.Range.Words(1) = "C�u " Then
        arrIndexAns.Add i
        o = o + 1
        Debug.Print arrIndexAns(o)
    End If
Next

' Define the range collection for pharse of ans
Dim rangeColl As Collection
Set rangeColl = New Collection

' A simple range
Dim r As Range
For i = 1 To o - 1 Step 1
    Set r = wrdDoc.Range( _
     Start:=wrdDoc.Paragraphs(arrIndexAns(i)).Range.Start, _
     End:=wrdDoc.Paragraphs(arrIndexAns(i + 1) - 1).Range.End)
     'Add range to coll
    rangeColl.Add r
Next i
   ' Print the range coll
   Selection.EndKey Unit:=wdStory
    rangeColl(3).Copy
    Selection.Paste
End Sub
Sub TronCauHoi()
    ' Khai bao bien co ban
    Dim mRange As Range
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    wrdDoc.Range.ListFormat.ConvertNumbersToText
    Dim i As Integer, rndNumber As Integer
    i = 0
    'Khai bao collection cau hoi
    Dim collQuestion As Collection 'Dung de luu vi tri doan bat dau cua cau hoi
    Set collQuestion = New Collection
    ' Duyet cac doan trong document
    For Each Paragraph In wrdDoc.Paragraphs
        i = i + 1
        'Kiem tra xem doan van co phai cau hoi hay khong
        If Paragraph.Range.Words(1) = "C�u " Then
            collQuestion.Add i
        End If
    Next
    'Khai bao bien chua so luong cau hoi
    Dim qCount As Integer
    qCount = collQuestion.Count
    'Khai bao collection chua range cho cau hoi
    Dim collRangeQ As Collection
    Set collRangeQ = New Collection
    'Set range cho collection Range vua tao
    ' Gan so thu tu ngau nhien cho cac cau hoi
    Dim collRandomSortQuestion1 As Collection
    Set collRandomSortQuestion1 = New Collection
    Dim collRandomSortQuestion2 As Collection
    Set collRandomSortQuestion2 = New Collection
    Dim collRandomSortQuestion3 As Collection
    Set collRandomSortQuestion3 = New Collection
    Dim collRandomSortQuestion4 As Collection
    Set collRandomSortQuestion4 = New Collection
    
    For i = 1 To qCount Step 1
       rndNumber = Int(1 + rnd * (4 - 1 + 1))
       'Debug.Print ("Random: " & rndNumber)
       'Debug.Print ("Start i: " & collQuestion(i))
       'Debug.Print "End i: " & collQuestion(i + 1) - 1
       'Debug.Print collQuestion(i)
       If i < qCount Then
            Set mRange = wrdDoc.Range( _
            Start:=wrdDoc.Paragraphs(collQuestion(i)).Range.Start, _
            End:=wrdDoc.Paragraphs(collQuestion(i + 1) - 1).Range.End)
        Else
            Set mRange = wrdDoc.Range( _
            Start:=wrdDoc.Paragraphs(collQuestion(i)).Range.Start, _
            End:=wrdDoc.Paragraphs(wrdDoc.Paragraphs.Count).Range.End) 'Sua sau
        End If
        ' add range vao collection range
        collRangeQ.Add mRange
       If rndNumber = 1 Then
            collRandomSortQuestion1.Add i
        ElseIf rndNumber = 2 Then
            collRandomSortQuestion2.Add i
        ElseIf rndNumber = 3 Then
            collRandomSortQuestion3.Add i
        Else
            collRandomSortQuestion4.Add i
        End If
    Next i
    ' In ra cac cau hoi theo thu tu moi
    Dim idQ As Integer ' Dung de danh stt moi cho cau hoi
    idQ = 1
    Selection.EndKey Unit:=wdStory
    For i = 1 To collRandomSortQuestion1.Count Step 1
        collRangeQ(collRandomSortQuestion1(i)).Words(2).Text = idQ
        collRangeQ(collRandomSortQuestion1(i)).Copy
        Selection.Paste
        idQ = idQ + 1
    Next i
    For i = 1 To collRandomSortQuestion2.Count Step 1
        collRangeQ(collRandomSortQuestion2(i)).Words(2).Text = idQ
        collRangeQ(collRandomSortQuestion2(i)).Copy
        Selection.Paste
        idQ = idQ + 1
    Next i
    For i = 1 To collRandomSortQuestion3.Count Step 1
        collRangeQ(collRandomSortQuestion3(i)).Words(2).Text = idQ
        collRangeQ(collRandomSortQuestion3(i)).Copy
        Selection.Paste
        idQ = idQ + 1
    Next i
    For i = 1 To collRandomSortQuestion4.Count Step 1
        collRangeQ(collRandomSortQuestion4(i)).Words(2).Text = idQ
        collRangeQ(collRandomSortQuestion4(i)).Copy
        Selection.Paste
        idQ = idQ + 1
    Next i
End Sub

Sub DemSoCauHoi(idParaStart As Integer, idParaEnd As Integer)
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    Dim i, a, correctA, indexAS As Integer
    'So charcode bat dau tu A
    Dim cNum, pCount As Integer
    cNum = 65
    pCount = idParaEnd - idParaStart
    Debug.Print pCount
    'Tao mang chua cac cau tra loi
    Dim arrA()
    ReDim arrA(pCount - 1)
    a = 0 'So cau hoi da tim thay
    'Duyet cau hoi & gan cac cau hoi vao bien & tim cau hoi dung
    For i = 0 To pCount Step 1
        Dim s1, s2 As String
        s1 = wrdDoc.Paragraphs(idParaStart + i).Range.Words(1) & "."
        s2 = Chr(cNum) & "."
        Debug.Print "i= " & i
        Debug.Print s1
        Debug.Print s2
        If StrComp(s1, s2) = 0 Then
            If s2 = "A." Then
                indexAS = idParaStart + i
            End If
            arrA(a) = wrdDoc.Paragraphs(idParaStart + i).Range
            'Danh dau cau tra loi dung
            If wrdDoc.Paragraphs(idParaStart + i).Range.Words(1).Font.Underline = wdUnderlineSingle Or _
                wrdDoc.Paragraphs(idParaStart + i).Range.Words(1).Font.Color = wdColorRed Then
                'Xac dinh vi tri cau tra loi dung
                correctA = a
            End If
            a = a + 1
            cNum = cNum + 1
        End If
    Next i
    ' Goi function tron cau tra loi
    arrA = TronCauTraLoi(arrA(), a)
    ' Cap nhat lai cau tra loi dung sau khi tron
    For i = 0 To a - 1 Step 1
        If Left(arrA(i), 1) = Chr(65 + correctA) Then
            correctA = i
            Exit For
        End If
    Next i
    Dim o As Integer
    o = 0
    'Sua lai cau tra loi theo thu tu moi
    For i = indexAS To indexAS + a - 1 Step 1
        wrdDoc.Paragraphs(i).Range.Text = Chr(65 + o) & Right(arrA(o), Len(arrA(o)) - 1)
        wrdDoc.Paragraphs(i).Format.Style = wdStyleNormal
        wrdDoc.Paragraphs(i).Format.LeftIndent = InchesToPoints(0.2)
        If correctA = o Then
            wrdDoc.Paragraphs(i).Range.Characters(1).Font.Underline = True
            wrdDoc.Paragraphs(i).Range.Characters(2).Font.Underline = True
        Else
            wrdDoc.Paragraphs(i).Range.Font.Underline = False
        End If
        o = o + 1
    Next i
    ' Sua lai stt cau tra loi
    
End Sub

Sub Test()
Dim cNum As Integer
    cNum = 65
    
    Dim s As String
    s = "A. "
    If s = Chr(cNum) & ". " Then
        Debug.Print "True"
    End If
End Sub