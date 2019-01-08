VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm042 
   ClientHeight    =   7840
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11256
   OleObjectBlob   =   "frm042.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm042"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub OKButton_Click()

If TextBox1.Enabled = True And TextBox1.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox2.Enabled = True And TextBox2.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox3.Enabled = True And TextBox3.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox4.Enabled = True And TextBox4.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox5.Enabled = True And TextBox5.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If ComboBox1.Enabled = True And ComboBox1.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox2.Enabled = True And ComboBox2.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox3.Enabled = True And ComboBox3.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox4.Enabled = True And ComboBox4.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If
    
If ComboBox5.Enabled = True And ComboBox5.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If
    
    Worksheets("SpmSvar").Range("C101:C101").Value = Controls("Label1").Caption
    
' indsæt forfaldsdato i Regler (MODR) og SpmSvar
    If ComboBox1.Value = "Dage" And IsNumeric(TextBox1.Value) Then
        Worksheets("Regler").Range("J29:J29").Value = "-1095"
        Worksheets("Regler").Range("M29:M29").Value = TextBox1.Value
        Worksheets("SpmSvar").Range("D106:D106").Value = TextBox1.Value
        Worksheets("SpmSvar").Range("E106:E106").Value = ComboBox1.Value
        
    ElseIf ComboBox1.Value = "Måneder" And IsNumeric(TextBox1.Value) Then
        Worksheets("Regler").Range("J29:J29").Value = "-1095"
        Worksheets("Regler").Range("N29:N29").Value = TextBox1.Value
        Worksheets("SpmSvar").Range("D106:D106").Value = TextBox1.Value
        Worksheets("SpmSvar").Range("E106:E106").Value = ComboBox1.Value
        
    ElseIf ComboBox1.Value = "År" And IsNumeric(TextBox1.Value) Then
        Worksheets("Regler").Range("J29:J29").Value = "-1095"
        Worksheets("Regler").Range("O29:O29").Value = TextBox1.Value
        Worksheets("SpmSvar").Range("D106:D106").Value = TextBox1.Value
        Worksheets("SpmSvar").Range("E106:E106").Value = ComboBox1.Value
    End If
    
' indsæt SRB i Regler (MODR) og SpmSvar
    If ComboBox2.Value = "Dage" And IsNumeric(TextBox2.Value) Then
        Worksheets("Regler").Range("J30:J30").Value = "-1095"
        Worksheets("Regler").Range("M30:M30").Value = TextBox2.Value
        Worksheets("SpmSvar").Range("D107:D107").Value = TextBox2.Value
        Worksheets("SpmSvar").Range("E107:E107").Value = ComboBox2.Value
    
    ElseIf ComboBox2.Value = "Måneder" And IsNumeric(TextBox2.Value) Then
       Worksheets("Regler").Range("J30:J30").Value = "-1095"
        Worksheets("Regler").Range("N30:N30").Value = TextBox2.Value
        Worksheets("SpmSvar").Range("D107:D107").Value = TextBox2.Value
        Worksheets("SpmSvar").Range("E107:E107").Value = ComboBox2.Value

    ElseIf ComboBox2.Value = "År" And IsNumeric(TextBox2.Value) Then
        Worksheets("Regler").Range("J30:J30").Value = "-1095"
        Worksheets("Regler").Range("O30:O30").Value = TextBox2.Value
        Worksheets("SpmSvar").Range("D107:D107").Value = TextBox2.Value
        Worksheets("SpmSvar").Range("E107:E107").Value = ComboBox2.Value
    Else
        dFunc.msgError = "Angiv venligst antal dage/måneder/år."
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Angiv venligst antal dage/måneder/år.")
        GoTo ending
    End If

' indsæt stiftelse i Regler (MODR) og SpmSvar
  If ComboBox3.Value = "Dage" And IsNumeric(TextBox3.Value) Then
        Worksheets("Regler").Range("J31:J31").Value = "-1095"
        Worksheets("Regler").Range("M31:M31").Value = TextBox3.Value
        Worksheets("SpmSvar").Range("D108:D108").Value = TextBox3.Value
        Worksheets("SpmSvar").Range("E108:E108").Value = ComboBox3.Value
        
    ElseIf ComboBox3.Value = "Måneder" And IsNumeric(TextBox3.Value) Then
        Worksheets("Regler").Range("J31:J31").Value = "-1095"
        Worksheets("Regler").Range("N31:N31").Value = TextBox3.Value
        Worksheets("SpmSvar").Range("D108:D108").Value = TextBox3.Value
        Worksheets("SpmSvar").Range("E108:E108").Value = ComboBox3.Value
        
    ElseIf ComboBox3.Value = "År" And IsNumeric(TextBox3.Value) Then
        Worksheets("Regler").Range("J31:J31").Value = "-1095"
        Worksheets("Regler").Range("O31:O31").Value = TextBox3.Value
        Worksheets("SpmSvar").Range("D108:D108").Value = TextBox3.Value
        Worksheets("SpmSvar").Range("E108:E108").Value = ComboBox3.Value
    End If

' indsæt periodestart i Regler (MODR) og SpmSvar
If ComboBox4.Value = "Dage" And IsNumeric(TextBox4.Value) Then
        Worksheets("Regler").Range("J32:J32").Value = "-1095"
        Worksheets("Regler").Range("M32:M32").Value = TextBox4.Value
        Worksheets("SpmSvar").Range("D109:D109").Value = TextBox4.Value
        Worksheets("SpmSvar").Range("E109:E109").Value = ComboBox4.Value
        
    ElseIf ComboBox4.Value = "Måneder" And IsNumeric(TextBox4.Value) Then
        Worksheets("Regler").Range("J32:J32").Value = "-1095"
        Worksheets("Regler").Range("N32:N32").Value = TextBox4.Value
        Worksheets("SpmSvar").Range("D109:D109").Value = TextBox4.Value
        Worksheets("SpmSvar").Range("E109:E109").Value = ComboBox4.Value
        
    ElseIf ComboBox4.Value = "År" And IsNumeric(TextBox4.Value) Then
        Worksheets("Regler").Range("J32:J32").Value = "-1095"
        Worksheets("Regler").Range("O32:O32").Value = TextBox4.Value
        Worksheets("SpmSvar").Range("D109:D109").Value = TextBox4.Value
        Worksheets("SpmSvar").Range("E109:E109").Value = ComboBox4.Value
    End If

' indsæt periodeslut i Regler (MODR) og SpmSvar
If ComboBox5.Value = "Dage" And IsNumeric(TextBox5.Value) Then
        Worksheets("Regler").Range("J32:J32").Value = "-1095"
        Worksheets("Regler").Range("M32:M32").Value = TextBox5.Value
        Worksheets("SpmSvar").Range("D110:D110").Value = TextBox5.Value
        Worksheets("SpmSvar").Range("E110:E110").Value = ComboBox5.Value
        
    ElseIf ComboBox5.Value = "Måneder" And IsNumeric(TextBox5.Value) Then
        Worksheets("Regler").Range("J32:J32").Value = "-1095"
        Worksheets("Regler").Range("N32:N32").Value = TextBox5.Value
        Worksheets("SpmSvar").Range("D110:D110").Value = TextBox5.Value
        Worksheets("SpmSvar").Range("E110:E110").Value = ComboBox5.Value
        
    ElseIf ComboBox5.Value = "År" And IsNumeric(TextBox5.Value) Then
        Worksheets("Regler").Range("J32:J32").Value = "-1095"
        Worksheets("Regler").Range("O32:O32").Value = TextBox5.Value
        Worksheets("SpmSvar").Range("D110:D110").Value = TextBox5.Value
        Worksheets("SpmSvar").Range("E110:E110").Value = ComboBox5.Value
    End If

    If ComboBox1.Value = "" Then
        Worksheets("Regler").Range("G29:G29").Value = "NEJ"
    End If

    If ComboBox2.Value = "" Then
        Worksheets("Regler").Range("G30:G30").Value = "NEJ"
    End If

    If ComboBox3.Value = "" Then
        Worksheets("Regler").Range("G31:G31").Value = "NEJ"
    End If

    If ComboBox4.Value = "" Then
        Worksheets("Regler").Range("G32:G32").Value = "NEJ"
    End If

    If ComboBox5.Value = "" Then
        Worksheets("Regler").Range("G33:G33").Value = "NEJ"
    End If
    
    If ComboBox1.Value <> "" Then
        Worksheets("Regler").Range("G29:G29").Value = "JA"
    End If

    If ComboBox2.Value <> "" Then
        Worksheets("Regler").Range("G30:G30").Value = "JA"
    End If

    If ComboBox3.Value <> "" Then
        Worksheets("Regler").Range("G31:G31").Value = "JA"
    End If

    If ComboBox4.Value <> "" Then
        Worksheets("Regler").Range("G32:G32").Value = "JA"
    End If

    If ComboBox5.Value <> "" Then
        Worksheets("Regler").Range("G33:G33").Value = "JA"
    End If

Me.Hide
SFunc.ShowFunc ("frm025")
'frm025.Show
ending:
End Sub

Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm024")
'frm024.Show

TextBox1.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox5.Value = ""

ComboBox1.Value = ""
ComboBox2.Value = ""
ComboBox3.Value = ""
ComboBox4.Value = ""
ComboBox5.Value = ""

Label8.ForeColor = RGB(0, 0, 0)
Label9.ForeColor = RGB(0, 0, 0)
Label10.ForeColor = RGB(0, 0, 0)
Label11.ForeColor = RGB(0, 0, 0)
Label12.ForeColor = RGB(0, 0, 0)


End Sub

Private Sub UserForm_Initialize()
Image1.PictureSizeMode = fmPictureSizeModeStretch
    
    
    With ComboBox1
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With

    With ComboBox2
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With

    With ComboBox3
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With
    
    With ComboBox4
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With

    With ComboBox5
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With


    If frm019.CheckBox1.Value = True Then
        TextBox1.Enabled = True
        ComboBox1.Enabled = True
    Else
        TextBox1.Enabled = False
        ComboBox1.Enabled = False
        TextBox1.Value = ""
        ComboBox1.Value = ""
    End If
    
    If frm019.CheckBox2.Value = True Then
        TextBox2.Enabled = True
        ComboBox2.Enabled = True
    Else
        TextBox2.Enabled = False
        ComboBox2.Enabled = False
        TextBox2.Value = ""
        ComboBox2.Value = ""
    End If
    
    If frm019.CheckBox3.Value = True Then
        TextBox3.Enabled = True
        ComboBox3.Enabled = True
    Else
        TextBox3.Enabled = False
        ComboBox3.Enabled = False
        TextBox3.Value = ""
        ComboBox3.Value = ""
    End If
    
    If frm019.CheckBox4.Value = True Then
        TextBox4.Enabled = True
        ComboBox4.Enabled = True
    Else
        TextBox4.Enabled = False
        ComboBox4.Enabled = False
        TextBox4.Value = ""
        ComboBox4.Value = ""
    End If

    If frm019.CheckBox5.Value = True Then
        TextBox5.Enabled = True
        ComboBox5.Enabled = True
    Else
        TextBox5.Enabled = False
        ComboBox5.Enabled = False
        TextBox5.Value = ""
        ComboBox5.Value = ""
    End If

' Indlæs forfaldsdato dato fra tidligere besvarelse
    If Worksheets("SpmSvar").Range("E106:E106").Value = "Dage" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D106:D106").Value
        ComboBox1.Value = "Dage"
        
    ElseIf Worksheets("SpmSvar").Range("E106:E106").Value = "Måneder" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D106:D106").Value
        ComboBox1.Value = "Måneder"
        
    ElseIf Worksheets("SpmSvar").Range("E106:E106").Value = "År" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D106:D106").Value
        ComboBox1.Value = "År"
    End If

' Indlæs SRB dato fra tidligere besvarelse
    If Worksheets("SpmSvar").Range("E107:E107").Value = "Dage" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D107:D107").Value
        ComboBox1.Value = "Dage"

    ElseIf Worksheets("SpmSvar").Range("E107:E107").Value = "Måneder" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D107:D107").Value
        ComboBox1.Value = "Måneder"

    ElseIf Worksheets("SpmSvar").Range("E107:E107").Value = "År" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D107:D107").Value
        ComboBox1.Value = "År"
    End If
    
' Indlæs stiftelsesdato fra tidligere besvarelse
    If Worksheets("SpmSvar").Range("E108:E108").Value = "Dage" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D108:D108").Value
        ComboBox1.Value = "Dage"
        
    ElseIf Worksheets("SpmSvar").Range("E108:E108").Value = "Måneder" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D108:D108").Value
        ComboBox1.Value = "Måneder"
        
    ElseIf Worksheets("SpmSvar").Range("E108:E108").Value = "År" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D108:D108").Value
        ComboBox1.Value = "År"
    End If

' Indlæs periodestartdato fra tidligere besvarelse
    If Worksheets("SpmSvar").Range("E109:E109").Value = "Dage" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D109:D109").Value
        ComboBox1.Value = "Dage"
        
    ElseIf Worksheets("SpmSvar").Range("E109:E109").Value = "Måneder" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D109:D109").Value
        ComboBox1.Value = "Måneder"
        
    ElseIf Worksheets("SpmSvar").Range("E109:E109").Value = "År" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D109:D109").Value
        ComboBox1.Value = "År"
    End If

' Indlæs periodeslutdato fra tidligere besvarelse
    If Worksheets("SpmSvar").Range("E110:E110").Value = "Dage" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D110:D110").Value
        ComboBox1.Value = "Dage"
        
    ElseIf Worksheets("SpmSvar").Range("E110:E110").Value = "Måneder" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D110:D110").Value
        ComboBox1.Value = "Måneder"
        
    ElseIf Worksheets("SpmSvar").Range("E110:E110").Value = "År" Then
        TextBox1.Value = Worksheets("SpmSvar").Range("D110:D110").Value
        ComboBox1.Value = "År"
    End If


End Sub

