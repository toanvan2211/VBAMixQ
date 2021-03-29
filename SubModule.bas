Attribute VB_Name = "NewMacros"
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
    'Gan vi tri moi cho cac cau tra loi
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

Sub TronDanhSach()
    ' Khai bao bien co ban
    Dim mRange As Range
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    wrdDoc.Range.ListFormat.ConvertNumbersToText
    Dim i As Integer, rndNumber As Integer
    i = 0
    'Khai bao collection cau hoi
    Dim collIndexQ As Collection 'Dung de luu vi tri doan bat dau cua cau hoi
    Set collIndexQ = New Collection
    
    'Tao bien chua vi tri ket thuc cua cau hoi cuoi cung
    Dim lastEndIndexQ As Integer
    
    ' Duyet cac doan trong document
    For Each Paragraph In wrdDoc.Paragraphs
        i = i + 1
        'Kiem tra xem doan van co phai cau hoi hay khong
        If Paragraph.Range.Words(1) = "Câu " Then
            If wrdDoc.Paragraphs(i + 2).Range.Words(1) = "B" Then
                collIndexQ.Add i
            End If
        End If
        If collIndexQ.Count <> 0 Then 'k hop ly - new cau hoi dai hon 1 doan thi sai logic
            If Paragraph.Range.Words(1) = "B" And collIndexQ(collIndexQ.Count) = i - 2 Then
                lastEndIndexQ = i
            ElseIf Paragraph.Range.Words(1) = "C" And collIndexQ(collIndexQ.Count) = i - 3 Then
                lastEndIndexQ = i
            ElseIf Paragraph.Range.Words(1) = "D" And collIndexQ(collIndexQ.Count) = i - 4 Then
                lastEndIndexQ = i
            ElseIf Paragraph.Range.Words(1) = "E" And collIndexQ(collIndexQ.Count) = i - 5 Then
                lastEndIndexQ = i
            ElseIf Paragraph.Range.Words(1) = "F" And collIndexQ(collIndexQ.Count) = i - 6 Then
                lastEndIndexQ = i
            End If
        End If
    Next
    
    'Tron cau tra loi cua tung cau hoi
    For i = 1 To collIndexQ.Count Step 1
        If i <> collIndexQ.Count Then
            TronCauTraLoi_Q collIndexQ(i), collIndexQ(i + 1) - 1
        Else
            TronCauTraLoi_Q collIndexQ(i), lastEndIndexQ
        End If
    Next i
    
    'Khai bao bien chua so luong cau hoi
    Dim qCount As Integer
    qCount = collIndexQ.Count
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
    
    'Duyet cau hoi vao tao vi tri ngau nhien moi cho tung cau hoi
    For i = 1 To qCount Step 1
       rndNumber = Int(1 + rnd * (4 - 1 + 1))
       'Xac dinh range cua cau hoi
       If i < qCount Then
            Set mRange = wrdDoc.Range( _
            Start:=wrdDoc.Paragraphs(collIndexQ(i)).Range.Start, _
            End:=wrdDoc.Paragraphs(collIndexQ(i + 1) - 1).Range.End)
        Else
            Set mRange = wrdDoc.Range( _
            Start:=wrdDoc.Paragraphs(collIndexQ(i)).Range.Start, _
            End:=wrdDoc.Paragraphs(lastEndIndexQ).Range.End)
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

    'Dua con tro chuot ve cuoi cung cua cau hoi
    mRange.Select
    Selection.MoveRight
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdParagraph, Count:=1

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
    'Xoa phan cau hoi cu
    Set mRange = wrdDoc.Range( _
        Start:=wrdDoc.Paragraphs(collIndexQ(1)).Range.Start, _
        End:=wrdDoc.Paragraphs(lastEndIndexQ).Range.End)
    mRange.Select
    Selection.Delete
    'Xoa dong thua
    wrdDoc.Paragraphs(lastEndIndexQ).Range.Select
    Selection.MoveDown
    Selection.Range.Delete
End Sub
Function TronCauTraLoi_Q(idParaStart As Integer, idParaEnd As Integer)
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
    
End Function
