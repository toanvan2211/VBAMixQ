Attribute VB_Name = "ModuleSub"
' Tron tat ca tren tai lieu hien tai
Sub MixBothThisDocument()

    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)
    Call Mix(rangeFind, collQ, lIndex, 2)
    'Call Mix(rangeFind, collQ, lIndex, 1)
    
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
    ActiveDocument.Paragraphs.TabStops.ClearAll
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(0.5)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(4.77)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(9.07)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(13.36)
End Sub

Sub Test()
    Dim collQ As Collection
    Dim lIndex As Integer
    Dim rangeFind As Range
    Set rangeFind = ActiveDocument.Range
    Set collQ = FindQuestion(lIndex, rangeFind)

    Call Mix(rangeFind, collQ, lIndex, 1)

    
End Sub

