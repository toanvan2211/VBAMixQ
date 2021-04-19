Attribute VB_Name = "ModuleSub"
' Tron tat ca tren tai lieu hien tai
Sub MixBothThisDocument()

    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, coll, lIndex, 2)
    Call Mix(rangeFind, coll, lIndex, 1)
    
End Sub
Sub MixQThisDocument()

    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, collQ, lIndex, 1)
    
End Sub

Sub MixAThisDocument()

    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, collQ, lIndex, 2)
    
End Sub
Sub MixBothTheSelection()

    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = Selection.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, collQ, lIndex, 2)
    Call Mix(rangeFind, collQ, lIndex, 1)
    
End Sub
Sub MixATheSelection()

    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = Selection.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, collQ, lIndex, 2)
    
End Sub
Sub MixQTheSelection()

    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = Selection.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, collQ, lIndex, 1)
    
End Sub

'Format
Sub FormatTabStop()
    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    
    Dim r As Range
    Set r = ActiveDocument.Range( _
        Start:=rangeFind.Paragraphs(collQ(1).ParaStartIndex).Range.Start, _
        End:=rangeFind.Paragraphs(lIndex).Range.End)
    r.Paragraphs.TabStops.ClearAll
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(0.5)
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(4.77)
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(9.07)
    r.Paragraphs.TabStops.Add Position:=CentimetersToPoints(13.36)
    r.ParagraphFormat.Alignment = wdAlignParagraphJustify
    
End Sub

Sub MarkRedCA()
    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    For Each Item In collQ
        Item.MarkRedCA
    Next
End Sub
Sub UnMarkRedCA()
    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    For Each Item In collQ
        Item.UnMarkRedCA
    Next
End Sub
Sub MarkUnderlineCA()
    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    For Each Item In collQ
        Item.MarkUnderlineCA
    Next
End Sub
Sub UnMarkUnderlineCA()
    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    For Each Item In collQ
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
' Xuat de
Sub MixToNewDocument(mixCount As Integer, AttachAns As Boolean)

    Dim i, countDoc As Integer
    
    Dim path, nameThisFile As String
    nameThisFile = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1)
        
    Dim docTemplate As Document
    Set docTemplate = New Document
    docTemplate.ActiveWindow.Visible = False
    
    docTemplate.Activate
    
    ThisDocument.Content.Copy
    docTemplate.Content.Paste
    
    Dim collQ As Collection
    Dim rangeFind As Range
    Set rangeFind = docTemplate.Range
    Dim lIndex As Integer
    Set collQ = FindQuestion(lIndex, rangeFind)
        
    For i = 1 To mixCount
        countDoc = -1
        Do
            countDoc = countDoc + 1
            path = ThisDocument.path & Application.PathSeparator & nameThisFile & "_De" & countDoc + i & ".docx"
        Loop While Dir(path, vbNormal) <> ""
        
        Call Mix(rangeFind, collQ, lIndex, 2)
        Call Mix(rangeFind, collQ, lIndex, 1)
        
        Set collQ = FindQuestion(lIndex, rangeFind)
        
        Dim doc As Document
        Set doc = New Document
        doc.ActiveWindow.Visible = False
        
        docTemplate.Content.Copy
        doc.Content.Paste
        
        'Tuy chon xuat dap an
        If AttachAns = True Then
            Dim objTable
            objTable = CreateTableOfAns(collQ, doc)
        End If
        doc.SaveAs FileName:=path
        doc.Close SaveChanges:=wdDoNotSaveChanges
    Next

    docTemplate.Close SaveChanges:=wdDoNotSaveChanges
    ThisDocument.Activate
    
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
    
    
    Set objDoc = Documents.Add
    
    objTable = CreateTableOfAns(coll, objDoc)

    Dim path, nameThisFile As String
    nameThisFile = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1)
    path = ThisDocument.path & Application.PathSeparator & nameThisFile & "_DapAn.docx"
    'Save replace
    objDoc.SaveAs FileName:=path, FileFormat:=wdFormatXMLDocument, AddtoRecentFiles:=False
    objDoc.Close
End Sub

'Them dap an vao cuoi doc
Sub InsertListAns()
    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion(lIndex, rangeFind)
    
    Dim intNoOfRows, intNoOfColumns As Integer
    Dim objRange
    Dim objTable
    
    objTable = CreateTableOfAns(coll, ActiveDocument)
    
End Sub
'Them dau cham cuoi cau tra loi
Sub AddAnEndToTheAns()
    Dim coll As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set coll = FindQuestion(lIndex, rangeFind)
    For Each Item In coll
        Call Item.DotMarkEndAns
    Next
End Sub
