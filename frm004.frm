VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm004 
   Caption         =   "UserForm1"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11280
   OleObjectBlob   =   "frm004.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub OKButton_Click()
' Validering af at der skal være en startdato for modtagelsesperioden
If TextBox1.Value = "" Then
    dFunc.msgError = "Startdatoen for perioden skal udfyldes."
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox ("Startdatoen for perioden skal udfyldes.")
    GoTo ending
End If

' Validering af datoformat for modtagelsesstartdatoen
If TextBox1.Value <> "" Then
    If FormatCheck(TextBox1, "date") = False Then
        TextBox1.SetFocus
        GoTo ending
    End If
End If

' Validering af dataformat for modtagelsesstartdatoen
If TextBox2.Value <> "" Then
    If FormatCheck(TextBox2, "date") = False Then
        TextBox2.SetFocus
        GoTo ending
    End If
End If

' Validering af om PeriodeStartdato er før PeriodeSlutdato
If TextBox2.Value <> "" Then
    If CDate(TextBox1.Value) > CDate(TextBox2.Value) Then
        dFunc.msgError = "Slutperioden kan ikke ligge før startperioden."
        SFunc.ShowFunc ("frmMsg")
        'frmMsg.Show
        'MsgBox ("Slutperioden kan ikke ligge før startperioden")
        GoTo ending
    End If
End If

If CDate(TextBox1.Value) < CDate("1 - 9 - 2013") Then
    SFunc.ShowFunc ("frm043")

'    MsgBox ("Fordringshaver har indtastet et begyndelsestidspunkt for modtagelsesperioden, der ligger før den 1. september 2013. Som udgangspunkt vil vi ikke konfigurere fordringer, der er modtaget før den 1. september 2013, da der er risiko for, at fordringer modtaget før den 1. september 2013 har mistet data i forbindelse med konverteringen til EFI/DMI. Såfremt der i populationsafgrænsningen vælges en modtagelsesperiode med start før den 1. september 2013, skal det afdækkes, om konverteringen af den afgrænsede population har medført ændringer i fordringernes data.")
    
    GoTo ending
End If

' Gemmer datoer i populations arket
Worksheets("Population").Range("B4:B4").Value = TextBox1.Value
Worksheets("Population").Range("B5:B5").Value = TextBox2.Value

' Gemmer datoer i SpmSvar arket
Worksheets("SpmSvar").Range("D4:D4").Value = TextBox1.Value
Worksheets("SpmSvar").Range("E4:E4").Value = TextBox2.Value

Me.Hide
SFunc.ShowFunc ("frm005")
'frm005.Show

ending:
End Sub

Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm003")
'frm003.Show
End Sub

Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()
Image1.PictureSizeMode = fmPictureSizeModeStretch
    ' Indlæs tidligere svar

If Worksheets("SpmSvar").Range("D4:D4").Value = "" Then
    TextBox1.Value = "01-09-2013"
Else
    TextBox1.Value = Worksheets("SpmSvar").Range("D4:D4").Value
End If

TextBox2.Value = Worksheets("SpmSvar").Range("E4:E4").Value

TextBox2.Value = Worksheets("SpmSvar").Range("E4:E4").Value

End Sub
