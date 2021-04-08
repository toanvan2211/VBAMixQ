Attribute VB_Name = "ModuleSub"
' Tron tat ca tren tai lieu hien tai
Sub MixBothThisDocument()

    Dim CollQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 2)
    Set CollQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, CollQ, lIndex, 1)
    
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

Sub MixToNewDocument(mixCount As Integer)
    Dim i As Integer
    
    Dim oFSO As Object
    Dim oFolder As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim path, nameThisFile As String
    nameThisFile = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1)
    path = ThisDocument.path & Application.PathSeparator
    
    Set oFolder = oFSO.GetFolder(path)
    Dim countDoc As Integer
    For Each oFile In oFolder.Files
        Dim os
        os = Left(oFile.Name, InStr(oFile.Name, ".") - 1)
        If Left(os, Len(os) - 1) = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1) & "_De" Then
            countDoc = CInt(Right(os, 1))
        End If
    Next
    
    For i = 1 To mixCount
        
        path = ThisDocument.path & Application.PathSeparator & nameThisFile & "_De" & countDoc + i & ".docx"
        Debug.Print path 'debug
        Dim doc As Document
        Set doc = New Document
        
        ThisDocument.Content.Copy
        doc.Content.Paste
        
        Dim CollQ As Collection
        Dim rangeFind As Range
        Set rangeFind = doc.Range
        
        Dim lIndex As Integer
        Set CollQ = FindQuestion(lIndex, rangeFind)
        Call Mix(rangeFind, CollQ, lIndex, 2)
        Set CollQ = FindQuestion(lIndex, rangeFind) 'Tim cach toi uu
        Call Mix(rangeFind, CollQ, lIndex, 1)
        
        doc.SaveAs FileName:=path
        doc.Close
    Next
End Sub

Sub Test()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim i As Integer
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim path, nameThisFile As String
    nameThisFile = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1)
    path = ThisDocument.path & Application.PathSeparator
    
    Set oFolder = oFSO.GetFolder(path)
    Dim countDoc As Integer
    For Each oFile In oFolder.Files
        Dim os
        os = Left(oFile.Name, InStr(oFile.Name, ".") - 1)
        If Left(os, Len(os) - 1) = Left(ThisDocument.Name, InStr(ThisDocument.Name, ".") - 1) & "_De" Then
            countDoc = CInt(Right(os, 1))
        End If
    Next
    
End Sub
