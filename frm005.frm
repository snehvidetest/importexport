VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm005 
   Caption         =   "Populationsafgrænsning"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm005.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()
Worksheets("SpmSvar").Range("C13:C13").Value = Controls("Label1").Caption


If OptionButton1 = False And OptionButton2 = False Then
    dFunc.msgError = "Vælg venligst et svar"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox ("Vælg venligst et svar")
    GoTo ending
End If

If OptionButton1 Then
    Worksheets("SpmSvar").Range("D13:D13").Value = "Ja"
    
        ' KF0006 aktiveres - KF0007-KF0008 deaktiveres
    Worksheets("Regler").Range("G7:G7").Value = "JA"
    Worksheets("Regler").Range("G8:G8").Value = "NEJ"
    Worksheets("Regler").Range("G9:G9").Value = "NEJ"
Else
    Worksheets("SpmSvar").Range("D13:D13").Value = "Nej"
End If

If OptionButton1 Then
    Me.Hide
    SFunc.ShowFunc ("frm006")
    'frm006.Show
ElseIf OptionButton2 Then
    dFunc.msgError = "Hvis registreringspraksis er forskellig kan FlexFilteret ikke anvendes"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox ("Hvis registreringspraksis er forskellig kan FlexFilteret ikke anvendes")
    Me.Hide
    SFunc.ShowFunc ("frm002")
    'frm002.Show
End If

ending:

End Sub

Private Sub OptionButton2_Click()

End Sub

Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm002")
'frm002.Show
End Sub

Private Sub UserForm_Initialize()


OptionButton1.Value = False
OptionButton2.Value = False

'Fill JA/NEJ ComboBox
Image1.PictureSizeMode = fmPictureSizeModeStretch
' Activate sheet
' Worksheets("Population").Activate

If Worksheets("SpmSvar").Range("D13:D13").Value = "Ja" Then
    OptionButton1.Value = True
ElseIf Worksheets("SpmSvar").Range("D13:D13").Value = "Nej" Then
    OptionButton2.Value = True
End If

End Sub
