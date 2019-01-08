VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm010 
   Caption         =   "Forældelseskontrol"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm010.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()
' Validering - Tomme felter
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "Vælg venligst et svar for at forsætte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
ElseIf OptionButton1.Value = True And TextBox1.Value = "" Then
    dFunc.msgError = "Indsæt venligst antal dage for at forsætte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If
    
' Validering - Max og min værdier
If TextBox1.Value > 1000 And OptionButton1.Value = False Then
    dFunc.msgError = "Antal dage kan ikke være mere end 1000"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If
    
If TextBox1.Value < -1000 And OptionButton1.Value = False Then
    dFunc.msgError = "Værdien er ugyldig"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If
    
    
    
    
    ' Worksheets("Regler").Activate
    If OptionButton1 = True Then
        If IsNumeric(TextBox1.Value) = True Then
            ' FOKO som i Retracer
            Worksheets("Regler").Range("J43:J47").Value = TextBox1.Value
            Worksheets("Regler").Range("G43:G47").Value = "JA"
            
            ' Grupper aktiveres og deaktiveres
            Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
            Worksheets("Gruppering").Range("C3:C3").Value = "JA"
                    
            ' RIM
            Worksheets("Population").Range("B16:B16").Value = "JA"
            Worksheets("Population").Range("B17:B17").Value = "NEJ"
            Me.Hide
            SFunc.ShowFunc ("frm039")
            'frm039.Show
            
            Worksheets("SpmSvar").Range("C20:C20").Value = Controls("Label1").Caption
    
        Else
            dFunc.msgError = "Antal dage er indtastet forkert"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Antal dage er indtastet forkert")
        End If
    
    ElseIf OptionButton2 = True Then
        Call dFunc.FOKO_Retracer
        Me.Hide
        SFunc.ShowFunc ("frm014")
        'frm014.Show
        
    Else
        dFunc.msgError = "Vælg venligst én af de ovenstående muligheder"
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Vælg venligst én af de ovenstående muligheder")
    End If
    
    If OptionButton2.Value = True Then
        Worksheets("SpmSvar").Range("D20:D20").Value = "Ved ikke"
    End If

ending:
End Sub



Private Sub OptionButton1_Click()

TextBox1.Enabled = True
Label2.Enabled = True
Label3.Enabled = False


End Sub

Private Sub OptionButton2_Click()

TextBox1.Enabled = False
Label2.Enabled = False
Label3.Enabled = True


End Sub




Private Sub TextBox1_Change()

End Sub

Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm009")
'frm009.Show


End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 9a22
If IsNumeric(Worksheets("SpmSvar").Range("D20:D20").Value) Then
    Call OptionButton1_Click
    OptionButton1.Value = True
    TextBox1.Value = Worksheets("SpmSvar").Range("D20:D20").Value
ElseIf Worksheets("SpmSvar").Range("D20:D20").Value = "Ved ikke" Then
    Call OptionButton2_Click
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If


End Sub
