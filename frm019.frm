VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm019 
   Caption         =   "Frasortering"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm019.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub OKButton_Click()
    ' Gem svar i "SpmSvar"
    Worksheets("SpmSvar").Range("C45:C45").Value = Controls("Label1").Caption
    Worksheets("SpmSvar").Range("D45:D45").Value = CheckBox1.Caption & " " & CheckBox1.Value
    Worksheets("SpmSvar").Range("E45:E45").Value = "SRB" & " " & CheckBox2.Value
    Worksheets("SpmSvar").Range("F45:F45").Value = CheckBox3.Caption & " " & CheckBox3.Value
    Worksheets("SpmSvar").Range("G45:G45").Value = "PeriodeStart" & " " & CheckBox4.Value
    Worksheets("SpmSvar").Range("H45:H45").Value = "PeriodeSlut" & " " & CheckBox5.Value
    
If CheckBox1.Value = False Then
    Worksheets("Regler").Range("J29:J29").Value = "-1825"
    Worksheets("Regler").Range("M29:M29").Value = "-1"
ElseIf CheckBox2.Value = False Then
    Worksheets("Regler").Range("J30:J30").Value = "-1825"
    Worksheets("Regler").Range("M30:M30").Value = "-1"
ElseIf CheckBox3.Value = False Then
    Worksheets("Regler").Range("J31:J31").Value = "-1825"
    Worksheets("Regler").Range("M31:M31").Value = "-1"
ElseIf CheckBox4.Value = False Then
    Worksheets("Regler").Range("J32:J32").Value = "-1825"
    Worksheets("Regler").Range("M32:M32").Value = "-1"
ElseIf CheckBox5.Value = False Then
    Worksheets("Regler").Range("J33:J33").Value = "-1825"
    Worksheets("Regler").Range("M33:M33").Value = "-1"
End If
    
    
    ' Vælg form
    If (CheckBox1.Value = True Or CheckBox2.Value = True Or CheckBox3.Value = True Or CheckBox4.Value = True Or CheckBox5.Value = True) Then
        Me.Hide
        SFunc.ShowFunc ("frm042")
        'frm042.Show
    Else
        dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der oprettes til modregning inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der oprettes til modregning inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
        Me.Hide
        SFunc.ShowFunc ("frm025")
        'frm025.Show
    End If



End Sub



Public Sub Tilbage_Click()

    Me.Hide
    SFunc.TestMode ("frm024")
    'frm024.Show

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

    ' Indlæs tidligere svar
    
    If Not IsEmpty(Worksheets("SpmSvar").Range("D45:D45").Value) Then
    
        If vaArray = Split(Worksheets("SpmSvar").Range("D45:D45").Value, " ")(1) = True Then
            CheckBox1.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("E45:E45").Value, " ")(1) = True Then
            CheckBox2.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("F45:F45").Value, " ")(1) = True Then
            CheckBox3.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("G45:G45").Value, " ")(1) = True Then
            CheckBox4.Value = True
        End If
        
        If vaArray = Split(Worksheets("SpmSvar").Range("H45:H45").Value, " ")(1) = True Then
            CheckBox5.Value = True
        End If
    End If
End Sub
