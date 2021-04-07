Attribute VB_Name = "ModuleSub"
' Tron tat ca tren tai lieu hien tai
Sub MixThisDocument()

    Dim collQ As Collection
    Dim lIndex As Integer
    Set collQ = FindQuestion(lIndex)
    Call Mix(collQ, lIndex, 2)
    Call Mix(collQ, lIndex, 1)
    
End Sub

'Format
Sub FormatTabStop()
    ActiveDocument.Paragraphs.TabStops.ClearAll
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(0.5)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(4.77)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(9.07)
    ActiveDocument.Paragraphs.TabStops.Add Position:=CentimetersToPoints(13.36)
End Sub

