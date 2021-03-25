Attribute VB_Name = "ModuleFunction"
Function Mix(arr()) As Variant
    Dim rndN, i As Integer
    'Tao mang chua gia tri vi tri moi
    Dim newIndexArr()
    ReDim newIndexArr(UBound(arr, 1))
    'Duyet mang & gan gia tri moi
    For i = 0 To UBound(arr, 1)
        rndN = Int(1 + rnd * (4 - 1 + 1))
        newIndexArr(i) = rndN
    Next i
    Mix = newIndexArr()
End Function
'Function duyet document tim cau hoi
Function FindQuestion(ByRef collCorrectAns As Collection) As Collection
    Dim wrdDoc As Document
    Set wrdDoc = Application.ActiveDocument
    wrdDoc.Range.ListFormat.ConvertNumbersToText
    Dim i As Integer
    Dim collQ As Collection
    Set collQ = New Collection
    
    Dim collC As Collection
    Set collC = New Collection
    
    Dim lastIndexQ As Integer
    ' Duyet cac doan trong document
    For Each Paragraph In wrdDoc.Paragraphs
        i = i + 1
        'Kiem tra xem doan van co phai cau hoi hay khong
        If Paragraph.Range.Words(1) = "Câu " Then
            Dim t, c As Integer
            t = 1
            c = 0 'So cau tra loi tim thay
            lastIndexQ = i - 1
            'Duyet cac doan cho den cau hoi tiep theo
            Do While i + t <= wrdDoc.Paragraphs.Count 'Duyet den cau hoi tiep theo
                If wrdDoc.Paragraphs(i + t).Range.Words(1) = "Câu " Then
                    Exit Do
                End If
                
                
                Dim tabA As Variant
                'Duyet cac cau tra loi cung hang
                tabA = Split(wrdDoc.Paragraphs(i + t).Range, vbTab)
                For Each a In tabA
                    If Left(a, 1) = Chr(65 + c) Then
                        c = c + 1
                        lastIndexQ = i + t
                    End If
                Next
                t = t + 1
            Loop
            
            If c >= 2 Then
                collQ.Add i
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
    collQ.Add lastIndexQ
    Set FindQuestion = collQ
End Function
Sub Test()
    Dim collCorrectAns As Collection
    Set collCorrectAns = New Collection
    
    Dim coll As Collection
    Set coll = FindQuestion(collCorrectAns)
    
    Dim i As Integer
    For i = 1 To collCorrectAns.Count
        Debug.Print coll(i) & "-" & collCorrectAns(i)
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
