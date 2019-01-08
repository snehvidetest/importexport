VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm011 
   Caption         =   "Forældelseskontrol"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm011.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Private Sub Ja_Click()

End Sub

Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()
    If OptionButton1.Value = False And OptionButton2.Value = False Then
        dFunc.msgError = "Vælg venligst et svar for at forsætte"
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
    End If

    Worksheets("SpmSvar").Range("C21:C21").Value = Controls("Label1").Caption
    If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("D21:D21").Value = "Ja"
    End If
    If OptionButton2.Value = True Then
    Worksheets("SpmSvar").Range("D21:D21").Value = "Nej"
    End If
    
    Me.Hide
    
    ' Worksheets("Konfiguration").Activate
    ' Activate sheet
    
    If OptionButton1 = True Then
        ' G2 Aktiveres
        Worksheets("Regler").Range("G43:G47").Value = "JA"
        ' Stoler på RIM
        Worksheets("Population").Range("B16:B16").Value = "JA"
        ' Gruppe 2 aktiveres
        Worksheets("Gruppering").Range("C3:C3").Value = "JA"
        Worksheets("Regler").Range("G40:G40").Value = "JA"
        SFunc.ShowFunc ("frm014")
        'frm014.Show
    
    Else
        SFunc.ShowFunc ("frm012")
        'frm012.Show
    End If

    
ending:
End Sub



Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub
Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm010")
'frm007.Show
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 9b
If Worksheets("SpmSvar").Range("D21:D21").Value = "Ja" Then
    OptionButton1.Value = True
ElseIf Worksheets("SpmSvar").Range("D21:D21").Value = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If



End Sub
