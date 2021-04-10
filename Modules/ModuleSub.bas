Attribute VB_Name = "ModuleSub"
' Tron tat ca tren tai lieu hien tai
Sub MixBothThisDocument()

    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, coll, lIndex, 2)
    Set coll = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, coll, lIndex, 1)
    
End Sub
Sub MixQThisDocument()

    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 1)
    
End Sub

Sub MixAThisDocument()

    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 2)
    
End Sub
Sub MixBothTheSelection()

    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = Selection.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 2)
    Set CollQ = FindQuestion(lIndex, rangeFind) 'Tim cach toi uu
    Call Mix(rangeFind, CollQ, lIndex, 1)
    
End Sub
Sub MixATheSelection()

    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = Selection.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 2)
    
End Sub
Sub MixQTheSelection()

    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = Selection.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 1)
    
End Sub

'Format
Sub FormatTabStop()
    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    
    Dim r As Range
    Set r = ActiveDocument.Range( _
        Start:=rangeFind.Paragraphs(CollQ(1).ParaStartIndex).Range.Start, _
        End:=rangeFind.Paragraphs(lIndex).Range.End)
    r.Paragraphs.TabStops.ClearAll
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(0.5)
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(4.77)
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(9.07)
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(13.36)
    r.ParagraphFormat.Alignment = wdAlignParagraphJustify
    
End Sub

Sub MarkRedCA()
    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    For Each Item In CollQ
        Item.MarkRedCA
    Next
End Sub
Sub UnMarkRedCA()
    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    For Each Item In CollQ
        Item.UnMarkRedCA
    Next
End Sub
Sub MarkUnderlineCA()
    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    For Each Item In CollQ
        Item.MarkUnderlineCA
    Next
End Sub
Sub UnMarkUnderlineCA()
    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    For Each Item In CollQ
        Item.UnMarkUnderlineCA
    Next
End Sub
Sub MarkQuestionOrder()

    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion(lIndex, rangeFind)
    Dim i As Integer
    i = 0
    For Each Item In coll
        i = i + 1
        Item.RangeQ.Words(2) = i
    Next
    
End Sub
' Xuat de khong kem dap an
Sub MixToNewDocument(mixCount As Integer)

    Dim i, countDoc As Integer
    
    Dim path, nameThisFile As String
    nameThisFile = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1)
    
    Dim CollQ As Collection
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Dim lIndex As Integer
    Set CollQ = FindQuestion(lIndex, rangeFind)
    
    For i = 1 To mixCount
        countDoc = -1
        Do
            countDoc = countDoc + 1
            path = ThisDocument.path & Application.PathSeparator & nameThisFile & "_De" & countDoc + i & ".docx"
        Loop While Dir(path, vbNormal) <> ""
        
        
        Debug.Print path 'debug
        Dim doc As Document
        Set doc = New Document

        ThisDocument.Content.Copy
        doc.Content.Paste
        
        Call Mix(rangeFind, CollQ, lIndex, 2)
        
        Dim newRangeFind As Range
        Set newRangeFind = doc.Range
        Dim newCollQ As Collection
        
        Set newCollQ = FindQuestion(lIndex, newRangeFind) 'Tim cach toi uu
        Call Mix(newRangeFind, newCollQ, lIndex, 1)

        doc.SaveAs FileName:=path
        doc.Close
    Next


End Sub
'Xuat dap an
Sub ExportListAns()
    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion(lIndex, rangeFind)
    
    Dim intNoOfRows, intNoOfColumns As Integer
    Dim objDoc As Document
    Dim objRange
    Dim objTable
    
    
    intNoOfRows = Fix(coll.Count / 4)
    If coll.Count Mod 4 <> 0 Then
        intNoOfRows = intNoOfRows + 1
    End If
    
    
    intNoOfColumns = 4
    
    Set objDoc = Documents.Add

    Set objRange = objDoc.Range

    objDoc.Tables.Add objRange, intNoOfRows, intNoOfColumns

    Set objTable = objDoc.Tables(1)

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
    Dim path, nameThisFile As String
    nameThisFile = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1)
    path = ThisDocument.path & Application.PathSeparator & nameThisFile & "_DapAn.docx"
    objDoc.SaveAs FileName:=path, FileFormat:=wdFormatXMLDocument, AddtoRecentFiles:=False
    objDoc.Close
End Sub
' Xuat de kem dap an
