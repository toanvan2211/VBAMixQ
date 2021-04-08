VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMixDoc 
   Caption         =   "Input"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3795
   OleObjectBlob   =   "frmMixDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMixDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOk_Click()

    If IsNumeric(tbMixCount.Text) Then
    
        MixToNewDocument (tbMixCount.Text)
        
        Application.Assistant.DoAlert UniConvert("Thoong basn", "Telex"), _
            UniConvert("Quas trifnh troojn ddax hoafn taast", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
        Unload frmMixDoc
    Else
        
        Application.Assistant.DoAlert UniConvert("Looxi", "Telex"), _
            UniConvert("Haxy nhaajp ddusng duwx lieeju", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
        
    End If

End Sub
Private Sub UserForm_Initialize()

    btnOk.Caption = UniConvert("Xasc nhaajn", "Telex")
    lbMixCount.Caption = UniConvert("Soos ddeef troojn: ", "Telex")
    
End Sub
