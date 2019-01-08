VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm017 
   Caption         =   "Frasortering"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm017.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub OKButton_Click()
    ' Gem svar
    Worksheets("SpmSvar").Range("C43:C43").Value = Controls("Label1").Caption
    Worksheets("SpmSvar").Range("D43:D43").Value = CheckBox1.Caption & " " & CheckBox1.Value
    Worksheets("SpmSvar").Range("E43:E43").Value = "SRB" & " " & CheckBox2.Value
    Worksheets("SpmSvar").Range("F43:F43").Value = CheckBox3.Caption & " " & CheckBox3.Value
    Worksheets("SpmSvar").Range("G43:G43").Value = "PeriodeStart" & " " & CheckBox4.Value
    Worksheets("SpmSvar").Range("H43:H43").Value = "PeriodeSlut" & " " & CheckBox5.Value

If CheckBox1.Value = False Then
    Worksheets("Regler").Range("J24:J24").Value = "-1825"
    Worksheets("Regler").Range("M24:M24").Value = "-1"
ElseIf CheckBox2.Value = False Then
    Worksheets("Regler").Range("J25:J25").Value = "-1825"
    Worksheets("Regler").Range("M25:M25").Value = "-1"
ElseIf CheckBox3.Value = False Then
    Worksheets("Regler").Range("J26:J26").Value = "-1825"
    Worksheets("Regler").Range("M26:M26").Value = "-1"
ElseIf CheckBox4.Value = False Then
    Worksheets("Regler").Range("J27:J27").Value = "-1825"
    Worksheets("Regler").Range("M27:M27").Value = "-1"
ElseIf CheckBox5.Value = False Then
    Worksheets("Regler").Range("J28:J28").Value = "-1825"
    Worksheets("Regler").Range("M28:M28").Value = "-1"
End If


If (CheckBox1.Value = True Or CheckBox2.Value = True Or CheckBox3.Value = True Or CheckBox4.Value = True Or CheckBox5.Value) Then
    Me.Hide
    SFunc.ShowFunc ("frm041")
    'frm041.Show
ElseIf frm005.OptionButton1.Value = True Then
    Me.Hide
    dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
    SFunc.ShowFunc ("frm024")
    'frm024.Show
ElseIf frm027.OptionButton1.Value = True Then
    Me.Hide
    dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
    SFunc.ShowFunc ("frm025")
    'frm025.Show
End If


End Sub

Public Sub Tilbage_Click()
    Me.Hide
    SFunc.ShowFunc ("frm023")
    'frm023.Show
End Sub

Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeStretch
    
    ' Indlæs tidligere svar
    If Not IsEmpty(Worksheets("SpmSvar").Range("D43:D43").Value) Then
    
        If vaArray = Split(Worksheets("SpmSvar").Range("D43:D43").Value, " ")(1) = True Then
            CheckBox1.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("E43:E43").Value, " ")(1) = True Then
            CheckBox2.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("F43:F43").Value, " ")(1) = True Then
            CheckBox3.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("G43:G43").Value, " ")(1) = True Then
            CheckBox4.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("H43:H43").Value, " ")(1) = True Then
            CheckBox5.Value = True
        End If
    End If
    
End Sub
