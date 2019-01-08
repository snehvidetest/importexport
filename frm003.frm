VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm003 
   Caption         =   "Populationsafgrænsning"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm003.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub OKButton_Click()

If OptionButton1 = False And OptionButton2 = False And OptionButton3 = False Then
    dFunc.msgError = "Vælg venligst et svar"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox ("Vælg venligst et svar")
    GoTo ending
End If

If OptionButton1 Then
    Me.Hide
    SFunc.ShowFunc ("frm004")
    'frm004.Show
ElseIf OptionButton2 Then
    Me.Hide
    SFunc.ShowFunc ("frm026")
    'frm026.Show
Else
    dFunc.msgError = "Populationen skal afgrænses på ny, hvis motorvejen skal kunne anvendes"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox "Populationen skal afgrænses på ny, hvis motorvejen skal kunne anvendes"
    Me.Hide
    SFunc.ShowFunc ("frm002")
    'frm002.Show
End If

Worksheets("SpmSvar").Range("C6:C6").Value = Label1.Caption
If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton1.Caption
ElseIf OptionButton2.Value = True Then
    Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton2.Caption
    frm002.txtModtStart.Value = "01-09-2013"
    Worksheets("SpmSvar").Range("D4:D4").Value = frm002.txtModtStart.Value
    Worksheets("Population").Range("B4:B4").Value = frm002.txtModtStart.Value
    frm002.txtModtSlut.Value = ""
    Worksheets("SpmSvar").Range("E4:E4").Value = frm002.txtModtSlut.Value
    Worksheets("Population").Range("B5:B5").Value = frm002.txtModtSlut.Value
    
    
    
ElseIf OptionButton3.Value = True Then
    Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton3.Caption
End If



ending:
End Sub



Private Sub OptionButton1_Click()

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
OptionButton3.Value = False

'Fill JA/NEJ ComboBox
Image1.PictureSizeMode = fmPictureSizeModeStretch
' Activate sheet
' Worksheets("Population").Activate

' Indlæs tidligere svar 4a
If Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton1.Caption Then
    OptionButton1.Value = True
ElseIf Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton2.Caption Then
    OptionButton2.Value = True
ElseIf Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton3.Caption Then
    OptionButton3.Value = True
End If

End Sub
