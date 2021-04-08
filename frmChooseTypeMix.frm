VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChooseTypeMix 
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   OleObjectBlob   =   "frmChooseTypeMix.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChooseTypeMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnCancel_Click()
    Unload frmChooseTypeMix
End Sub

Private Sub BtnRun_Click()
    If OpBtnMixQ = True Then
        'Goi function
        If Me.Tag = "Doc" Then
            MixQThisDocument
        ElseIf Me.Tag = "Select" Then
            MixQTheSelection
        End If
        'Thong bao
        Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), _
            UniConvert("DDax troojn hoafn taast", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
    ElseIf OpBtnMixA = True Then
        'Goi function
        If Me.Tag = "Doc" Then
            MixAThisDocument
        ElseIf Me.Tag = "Select" Then
            MixATheSelection
        End If
        'Thong bao
        Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), _
            UniConvert("DDax troojn hoafn taast", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
    ElseIf OpBtnMixBoth = True Then
        'Goi function
        If Me.Tag = "Doc" Then
            MixBothThisDocument
        ElseIf Me.Tag = "Select" Then
            MixBothTheSelection
        End If
        'Thong bao
        Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), _
            UniConvert("DDax troojn hoafn taast", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
    Else
        Application.Assistant.DoAlert UniConvert("Looxi", "Telex"), _
            UniConvert("Haxy chojn kieeru troojn!", "Telex"), msoAlertButtonOK _
            , msoAlertIconWarning, 0, 0, 0
    End If
    
End Sub

Private Sub UserForm_Initialize()
    'In ra caption tieng viet
    'Cb
    OpBtnMixQ.Caption = UniConvert("Troojn caau hori", "Telex")
    OpBtnMixA.Caption = UniConvert("Troojn caau trar lowfi", "Telex")
    OpBtnMixBoth.Caption = UniConvert("Troojn car hai", "Telex")
    'Btn
    BtnRun.Caption = UniConvert("Troojn", "Telex")
    BtnCancel.Caption = UniConvert("Huyr", "Telex")
    
End Sub


