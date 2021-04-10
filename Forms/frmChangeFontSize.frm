VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangeFontSize 
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "frmChangeFontSize.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangeFontSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    btnOk.Caption = UniConvert("Xasc nhaajn", "Telex")
    
End Sub

Private Sub btnOk_Click()

    
    
    If IsNumeric(tbSize.Text) Then
    
        Dim CollQ As Collection
        Dim lIndex As Integer
        Dim rangeFind As Range
        Set rangeFind = ActiveDocument.Range
        Set CollQ = FindQuestion(lIndex, rangeFind)
        
        Dim r As Range
        Set r = ActiveDocument.Range( _
            Start:=rangeFind.Paragraphs(CollQ(1).ParaStartIndex).Range.Start, _
            End:=rangeFind.Paragraphs(lIndex).Range.End)
        r.Font.Size = tbSize.Text
        Unload frmChangeFontSize
    Else
        Application.Assistant.DoAlert UniConvert("Looxi", "Telex"), _
            UniConvert("Haxy nhaajp ddusng duwx lieeju", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
    End If
    
    
    
End Sub
