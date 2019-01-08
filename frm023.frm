VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm023 
   Caption         =   "Frasortering"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm023.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm023"
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
    
Worksheets("SpmSvar").Range("C41:C41").Value = Controls("Label1").Caption
    
If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("D41:D41").Value = "Ja"
End If

If OptionButton2.Value = True Then
    Worksheets("SpmSvar").Range("D41:D41").Value = "Nej"
End If

' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton2.Value = True Then
    Worksheets("Regler").Range("J24:J24").Value = "-1825"
    Worksheets("Regler").Range("J25:J25").Value = "-1825"
    Worksheets("Regler").Range("J26:J26").Value = "-1825"
    Worksheets("Regler").Range("J27:J27").Value = "-1825"
    Worksheets("Regler").Range("J28:J28").Value = "-1825"
    
    Worksheets("Regler").Range("M24:M24").Value = "1"
    Worksheets("Regler").Range("M25:M25").Value = "1"
    Worksheets("Regler").Range("M26:M26").Value = "1"
    Worksheets("Regler").Range("M27:M27").Value = "1"
    Worksheets("Regler").Range("M28:M28").Value = "1"
End If

If OptionButton1.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm017")
    'frm017.Show
ElseIf frm005.OptionButton1.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm024")
    'frm024.Show
ElseIf frm027.OptionButton1.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm025")
    'frm025.Show
End If

ending:
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Public Sub Tilbage_Click()

    Me.Hide
    SFunc.ShowFunc ("frm022")
    'frm022.Show

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 14
If Worksheets("SpmSvar").Range("D41:D41").Value = "Ja" Then
    OptionButton1.Value = True
ElseIf Worksheets("SpmSvar").Range("D41:D41").Value = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If

End Sub
