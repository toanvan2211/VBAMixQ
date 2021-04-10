Attribute VB_Name = "ModuleCallForm"
Sub CallFormMixThisDoc()
    frmChooseTypeMix.Tag = "Doc"
    frmChooseTypeMix.Show
End Sub

Sub CallFormMixTheSelection()
    frmChooseTypeMix.Tag = "Select"
    frmChooseTypeMix.Show
End Sub

Sub ChangeFontSize()
       
    frmChangeFontSize.Show
    
End Sub

Sub MixToNewDoc()
    frmMixDoc.Show
End Sub
