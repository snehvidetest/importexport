VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm008 
   Caption         =   "Forældelseskontrol"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm008.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()
Worksheets("SpmSvar").Range("C18:C18").Value = Controls("Label1").Caption
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "Vælg venligst et svar for at forsætte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("D18:D18").Value = "Ja"
    
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
    Worksheets("Gruppering").Range("C3:C3").Value = "JA"
    
    Worksheets("Population").Range("B16:B16").Value = "JA"
    Worksheets("Population").Range("B17:B17").Value = "NEJ"
ElseIf OptionButton1.Value = False Then
    Worksheets("SpmSvar").Range("D18:D18").Value = "Nej"
End If


Me.Hide

' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton1 = True Then
    SFunc.ShowFunc ("frm039")
    'frm039.Show

ElseIf OptionButton2 = True Then
    SFunc.ShowFunc ("frm009")
    'frm009.Show
End If


ending:
End Sub



Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Public Sub Tilbage_Click()
Me.Hide

SFunc.ShowFunc ("frm007")
'frm007.Show
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 9a
If Worksheets("SpmSvar").Range("D18:D18").Value = "Ja" Then
    OptionButton1.Value = True
    
ElseIf Worksheets("SpmSvar").Range("D18:D18").Value = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If


End Sub
