VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm014 
   Caption         =   "Forældelseskontrol"
   ClientHeight    =   8730
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11592
   OleObjectBlob   =   "frm014.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    Forfaldsdato = False
    SRB = False
    PeriodeStartdato = False
    PeriodeSlutdato = False
    Stiftelsesdato = False
    Forfaldsdato.Enabled = False
    SRB.Enabled = False
    PeriodeStartdato.Enabled = False
    PeriodeSlutdato.Enabled = False
    Stiftelsesdato.Enabled = False

ElseIf CheckBox2.Value = False Then
    Forfaldsdato.Enabled = True
    SRB.Enabled = True
    PeriodeStartdato.Enabled = True
    PeriodeSlutdato.Enabled = True
    Stiftelsesdato.Enabled = True
End If
End Sub

Public Sub OKButton_Click()
If Forfaldsdato.Value = False And SRB.Value = False And PeriodeStartdato.Value = False And PeriodeSlutdato.Value = False And Stiftelsesdato.Value = False And CheckBox2.Value = False Then
    dFunc.msgError = "Mindst ét af stamdatafelterne eller 'Ingen' skal vælges for at forsætte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
    'MsgBox ("Mindst ét af stamdatafelterne eller 'Ingen' skal vælges for at forsætte")
End If

If Forfaldsdato.Value = True Then
    Worksheets("Regler").Range("G48:G48").Value = "JA"
    Worksheets("Regler").Range("G49:G49").Value = "JA"
    Worksheets("Regler").Range("G50:G50").Value = "JA"
    Worksheets("Regler").Range("G51:G51").Value = "JA"
    Worksheets("Regler").Range("G68:G68").Value = "JA"
ElseIf Forfaldsdato.Value = False Then
    Worksheets("Regler").Range("G48:G48").Value = "NEJ"
    Worksheets("Regler").Range("G49:G49").Value = "NEJ"
    Worksheets("Regler").Range("G50:G50").Value = "NEJ"
    Worksheets("Regler").Range("G51:G51").Value = "NEJ"
    Worksheets("Regler").Range("G68:G68").Value = "NEJ"
End If

If Stiftelsesdato.Value = True Then
    Worksheets("Regler").Range("G52:G52").Value = "JA"
    Worksheets("Regler").Range("G53:G53").Value = "JA"
    Worksheets("Regler").Range("G54:G54").Value = "JA"
    Worksheets("Regler").Range("G55:G55").Value = "JA"
    Worksheets("Regler").Range("G69:G69").Value = "JA"
ElseIf Stiftelsesdato.Value = False Then
    Worksheets("Regler").Range("G52:G52").Value = "NEJ"
    Worksheets("Regler").Range("G53:G53").Value = "NEJ"
    Worksheets("Regler").Range("G54:G54").Value = "NEJ"
    Worksheets("Regler").Range("G55:G55").Value = "NEJ"
    Worksheets("Regler").Range("G69:G69").Value = "NEJ"
End If

If PeriodeStartdato.Value = True Then
    Worksheets("Regler").Range("G56:G56").Value = "JA"
    Worksheets("Regler").Range("G57:G57").Value = "JA"
    Worksheets("Regler").Range("G58:G58").Value = "JA"
    Worksheets("Regler").Range("G59:G59").Value = "JA"
    Worksheets("Regler").Range("G70:G70").Value = "JA"
ElseIf PeriodeStartdato.Value = False Then
    Worksheets("Regler").Range("G56:G56").Value = "NEJ"
    Worksheets("Regler").Range("G57:G57").Value = "NEJ"
    Worksheets("Regler").Range("G58:G58").Value = "NEJ"
    Worksheets("Regler").Range("G59:G59").Value = "NEJ"
    Worksheets("Regler").Range("G70:G70").Value = "NEJ"
End If

If PeriodeSlutdato.Value = True Then
    Worksheets("Regler").Range("G60:G60").Value = "JA"
    Worksheets("Regler").Range("G61:G61").Value = "JA"
    Worksheets("Regler").Range("G62:G62").Value = "JA"
    Worksheets("Regler").Range("G63:G63").Value = "JA"
    Worksheets("Regler").Range("G71:G71").Value = "JA"
ElseIf PeriodeSlutdato.Value = False Then
    Worksheets("Regler").Range("G60:G60").Value = "NEJ"
    Worksheets("Regler").Range("G61:G61").Value = "NEJ"
    Worksheets("Regler").Range("G62:G63").Value = "NEJ"
    Worksheets("Regler").Range("G63:G63").Value = "NEJ"
    Worksheets("Regler").Range("G71:G71").Value = "NEJ"
End If

If SRB.Value = True Then
    Worksheets("Regler").Range("G64:G64").Value = "JA"
    Worksheets("Regler").Range("G65:G65").Value = "JA"
    Worksheets("Regler").Range("G66:G66").Value = "JA"
    Worksheets("Regler").Range("G67:G67").Value = "JA"
    Worksheets("Regler").Range("G72:G72").Value = "JA"
ElseIf SRB.Value = False Then
    Worksheets("Regler").Range("G64:G64").Value = "NEJ"
    Worksheets("Regler").Range("G65:G65").Value = "NEJ"
    Worksheets("Regler").Range("G66:G66").Value = "NEJ"
    Worksheets("Regler").Range("G67:G67").Value = "NEJ"
    Worksheets("Regler").Range("G72:G72").Value = "NEJ"
End If

Worksheets("SpmSvar").Range("C24:C24").Value = Controls("Label5").Caption
    Worksheets("SpmSvar").Range("D24:D24").Value = "Forfaldsdato" & " " & Forfaldsdato.Value
    Worksheets("SpmSvar").Range("E24:E24").Value = "SRB" & " " & SRB.Value
    Worksheets("SpmSvar").Range("F24:F24").Value = "Stiftelsesdato" & " " & Stiftelsesdato.Value
    Worksheets("SpmSvar").Range("G24:G24").Value = "PeriodeStart" & " " & PeriodeStartdato.Value
    Worksheets("SpmSvar").Range("H24:H24").Value = "PeriodeSlut" & " " & PeriodeSlutdato.Value
    Worksheets("SpmSvar").Range("I24:I24").Value = "Ingen" & " " & CheckBox2.Value

' If CheckBox2.Value = False Then
'    Worksheets("Regler").Range("G41:G41").Value = "JA"
' End If

If CheckBox2.Value = True Then
    Worksheets("Population").Range("B17:B17").Value = "NEJ"
ElseIf CheckBox2.Value = False Then
    Worksheets("Population").Range("B17:B17").Value = "JA"
End If

If Forfaldsdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm028")
    'frm028.Show
    GoTo ending
ElseIf SRB.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm032")
    'frm032.Show
    GoTo ending
ElseIf Stiftelsesdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm029")
    'frm029.Show
    GoTo ending
ElseIf PeriodeStartdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm030")
    'frm030.Show
    GoTo ending
ElseIf PeriodeSlutdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm031")
    'frm031.Show
ElseIf CheckBox2.Value = True And frm007.OptionButton3.Value = True Then
    'MsgBox ("RIM kan ikke beregne et tidligst muligt forældelsestidspunkt for fordringer omfattet af den afgrænsede population.")
    dFunc.msgError = "RIM kan ikke beregne et tidligst muligt forældelsestidspunkt for fordringer omfattet af den afgrænsede population."
    SFunc.ShowFunc ("frmMsg")
    Me.Hide
    SFunc.ShowFunc ("frm002")
    'frm002.Show
ElseIf CheckBox2.Value = True And frm007.OptionButton2.Value = True And (frm012.OptionButton1.Value = True Or frm011.OptionButton1.Value = True) Then
    dFunc.msgError = "RIM kan ikke beregne et tidligst muligt forældelsestidspunkt for fordringer omfattet af den afgrænsede population."
    SFunc.ShowFunc ("frmMsg")
    Me.Hide
    SFunc.ShowFunc ("frm039")
ElseIf CheckBox2.Value = True And frm007.OptionButton1.Value = True Then
    dFunc.msgError = "RIM kan ikke beregne et tidligst muligt forældelsestidspunkt for fordringer omfattet af den afgrænsede population."
    SFunc.ShowFunc ("frmMsg")
    Me.Hide
    SFunc.ShowFunc ("frm039")

End If

If CheckBox2.Value = True Then
    Call Insert_to_sheet("Gruppering", "C2:C2", "NEJ")
End If

ending:
End Sub


Public Sub Tilbage_Click()
    Me.Hide
    SFunc.ShowFunc ("frm007")
    'frm007.Show
End Sub

Private Sub UserForm_Initialize()
    
Image1.PictureSizeMode = fmPictureSizeModeStretch

    If Not IsEmpty(Worksheets("SpmSvar").Range("D24:D24").Value) Then
    
        If Split(Worksheets("SpmSvar").Range("D24:D24").Value, " ")(1) = True Then
            Forfaldsdato.Value = True
        End If
        
        If Split(Worksheets("SpmSvar").Range("E24:E24").Value, " ")(1) = True Then
            SRB.Value = True
        End If
        
        If Split(Worksheets("SpmSvar").Range("F24:F24").Value, " ")(1) = True Then
            Stiftelsesdato.Value = True
        End If
        
        If Split(Worksheets("SpmSvar").Range("G24:G24").Value, " ")(1) = True Then
            PeriodeStartdato.Value = True
        End If
        
        If Split(Worksheets("SpmSvar").Range("H24:H24").Value, " ")(1) = True Then
            PeriodeSlutdato.Value = True
        End If
        
        If Split(Worksheets("SpmSvar").Range("I24:I24").Value, " ")(1) = True Then
            CheckBox2.Value = True
        End If
    End If

End Sub

'Private Sub UserForm_Initialize()
'    Forfaldsdato.Value = Worksheets("SpmSvar").range("D24:D24").Value
'    SRB.Value = Worksheets("SpmSvar").range("E24:E24").Value
'    Stiftelsesdato.Value = Worksheets("SpmSvar").range("F24:F24").Value
'    PeriodeStartdato.Value = Worksheets("SpmSvar").range("G24:G24").Value
'    PeriodeSlutdato.Value = Worksheets("SpmSvar").range("H24:H24").Value
'    CheckBox2.Value = Worksheets("SpmSvar").range("H24:H24").Value'
' End Sub
