VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm021 
   Caption         =   "Forældelseskontrol"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm021.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CheckBox1_Click()

If CheckBox1.Value = True Then
    TextBox1.Enabled = False
    Label1.Enabled = False
    TextBox1.Value = ""
Else
    TextBox1.Enabled = True
    Label1.Enabled = True
    End If

End Sub

Public Sub OKButton_Click()
   
' Validering - Beløbet skal være en talværdi
    If CheckBox1.Value = False And IsNumeric(TextBox1.Value) = False Then
        dFunc.msgError = "Indsæt en gyldig værdi"
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Indsæt en gyldig værdi")
        Exit Sub
    End If

' Beløb skal være positivt
    If TextBox1.Value < 0 Then
        dFunc.msgError = "Beløb skal være positivt"
        SFunc.ShowFunc ("frmMsg")
        Exit Sub
    End If
  
' Validering - Beløbsfeltet skal være udfyldt, eller 'Ved ikke' valgt ovenfor
    If IsEmpty(TextBox1.Value) = True And CheckBox1.Value = False Then
        dFunc.msgError = "Udfyldt venligst et beløb, eller vælg 'Ved ikke', hvis det højeste beløb ikke vides."
        SFunc.ShowFunc ("fmrMsg")
        'MsgBox ("Udfyldt venligst et beløb, eller vælg 'Ved ikke', hvis det højeste beløb ikke vides")
        Exit Sub
    End If
            
    Worksheets("SpmSvar").Range("C55:C55").Value = Controls("Label2").Caption
            
    If CheckBox1.Value = True Then
        Worksheets("SpmSvar").Range("D55:D55").Value = "Ved ikke"
        Call Insert_to_sheet("Regler", "H73:H73", "")
        Call Insert_to_sheet("Regler", "H74:H74", "")
        Call Insert_to_sheet("Regler", "G73:G73", "NEJ")
        Call Insert_to_sheet("Regler", "G74:G74", "NEJ")
        Call Insert_to_sheet("Regler", "G75:G75", "NEJ")
        Call Insert_to_sheet("Regler", "G76:G76", "NEJ")
        Call Insert_to_sheet("Gruppering", "C6:C6", "NEJ")
        Call Insert_to_sheet("Gruppering", "C7:C7", "NEJ")
    Else
        Worksheets("SpmSvar").Range("D55:D55").Value = TextBox1.Value + " kr."
        Call Insert_to_sheet("Regler", "H73:H73", TextBox1)
        Call Insert_to_sheet("Regler", "H74:H74", TextBox1)
        Call Insert_to_sheet("Regler", "G73:G73", "JA")
        Call Insert_to_sheet("Regler", "G74:G74", "JA")
        Call Insert_to_sheet("Regler", "G75:G75", "NEJ")
        Call Insert_to_sheet("Regler", "G76:G76", "NEJ")
        Call Insert_to_sheet("Gruppering", "C6:C6", "JA")
        Call Insert_to_sheet("Gruppering", "C7:C7", "NEJ")
    End If
    
    Me.Hide
    SFunc.ShowFunc ("frm022")
    
    'frm022.Show
    
ending:
End Sub
Public Sub Tilbage_Click()
    Me.Hide
    
    If frm039.CheckBox4.Value = True Then
       SFunc.ShowFunc ("frm037")
       'frm037.Show
    Else
       SFunc.ShowFunc ("frm038")
       'frm038.Show
    End If
    
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 12
If Worksheets("SpmSvar").Range("D55:D55").Value = "Ved ikke" Then
    CheckBox1.Value = True
End If

Dim val1 As String
Dim val2 As String

val2 = Worksheets("SpmSvar").Range("D55:D55").Value
val2 = onlyDigits(val2)
TextBox1.Value = val2

End Sub
