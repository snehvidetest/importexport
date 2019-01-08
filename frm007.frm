VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm007 
   Caption         =   "Forældelseskontrol"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm007.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub OKButton_Click()
Worksheets("SpmSvar").Range("C17:C17").Value = Controls("Label1").Caption

If OptionButton1.Value = True Then
Worksheets("SpmSvar").Range("D17:D17").Value = "Altid"
End If

If OptionButton2.Value = True Then
Worksheets("SpmSvar").Range("D17:D17").Value = "I visse tilfælde"
End If

If OptionButton3.Value = True Then
Worksheets("SpmSvar").Range("D17:D17").Value = "Aldrig"
End If





' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton1 = True Then
    Me.Hide
    SFunc.ShowFunc ("frm008")
    'frm008.Show

ElseIf OptionButton2 = True Then
    Me.Hide
    SFunc.ShowFunc ("frm011")
    'frm011.Show

ElseIf OptionButton3 = True Then
    Call dFunc.FOKO_Retracer
    Me.Hide
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
    SFunc.ShowFunc ("frm014")
    'frm014.Show
ElseIf OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False Then
    dFunc.msgError = "Vælg venligst et svar for at forsætte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

ending:
End Sub



Private Sub OptionButton3_Click()

End Sub

Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm006")
'frm006.Show
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 9
If Worksheets("SpmSvar").Range("D17:D17").Value = "Altid" Then
    OptionButton1.Value = True
ElseIf Worksheets("SpmSvar").Range("D17:D17").Value = "I visse tilfælde" Then
    OptionButton2.Value = True
ElseIf Worksheets("SpmSvar").Range("D17:D17").Value = "Aldrig" Then
    OptionButton3.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
    OptionButton3.Value = False
End If


End Sub
