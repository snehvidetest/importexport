VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm026 
   Caption         =   "Populationsafgrænsning"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm026.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm026"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Forfaldsdato_Click()
If Forfaldsdato.Value = True Then
    txtFFStart.Enabled = True
    txtFFSlut.Enabled = True

Else
    txtFFStart.Enabled = False
    txtFFSlut.Enabled = False

End If
End Sub


Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub



Private Sub Label13_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub SRB_Click()
If SRB.Value = True Then
    txtSRBstart.Enabled = True
    txtSRBslut.Enabled = True

Else
    txtSRBstart.Enabled = False
    txtSRBslut.Enabled = False

End If
End Sub

Private Sub Stiftelsesdato_Click()
If Stiftelsesdato.Value = True Then
    txtSTIstart.Enabled = True
    txtSTIslut.Enabled = True

Else
    txtSTIstart.Enabled = False
    txtSTIslut.Enabled = False

End If
End Sub

Private Sub PeriodeStartdato_Click()
If PeriodeStartdato.Value = True Then
    txtPSTstart.Enabled = True
    txtPSTslut.Enabled = True

Else
    txtPSTstart.Enabled = False
    txtPSTslut.Enabled = False

End If
End Sub

Private Sub PeriodeSlutdato_Click()
If PeriodeSlutdato.Value = True Then
    txtPSLstart.Enabled = True
    txtPSLslut.Enabled = True

Else
    txtPSLstart.Enabled = False
    txtPSLslut.Enabled = False

End If
End Sub

Public Sub OKButton_Click()
    
    
    If Forfaldsdato.Value = True And txtFFSlut = "" Then txtFFSlut = Replace(Date, ".", "-")
    If SRB.Value = True And txtSRBslut = "" Then txtSRBslut = Replace(Date, ".", "-")
    If Stiftelsesdato.Value = True And txtSTIslut = "" Then txtSTIslut = Replace(Date, ".", "-")
    If PeriodeStartdato.Value = True And txtPSTslut = "" Then txtPSTslut = Replace(Date, ".", "-")
    If PeriodeSlutdato.Value = True And txtPSLslut = "" Then txtPSLslut = Replace(Date, ".", "-")
    
    ' Vælg mindst een dato
    If Not (Forfaldsdato And SRB And Stiftelsesdato And PeriodeStartdato And PeriodeSlutdato) Then
        dFunc.msgError = "Vælg som minimum et stamdatofelt for at gå videre"
        SFunc.ShowFunc ("frmMsg")
        Exit Sub
    End If
    
    If Forfaldsdato.Value = True Then
        If DateDiff("d", txtFFStart, txtFFSlut) < 0 Then
            dFunc.msgError = "Startdatoen (indtastet: " & txtFFStart & ") skal ligge før slutdatoen (indtastet: " & txtFFSlut & ")"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Startdatoen (indtastet: " & txtFFStart & ") skal ligge før slutdatoen (indtastet: " & txtFFSlut & ")")
            GoTo ending
        End If
        If (IsNull(txtFFStart) And IsNull(txtFFSlut)) Then
            dFunc.msgError = "Til og/eller fra datoerne skal være udfyldt for forfaldsdatoen"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Til og/eller fra datoerne skal være udfyldt for forfaldsdatoen")
            GoTo ending
        End If
        If Not IsNull(txtFFStart) Then
            If FormatCheck(txtFFStart, "date") = False Then
                GoTo ending
            End If
            
        End If
        If Not IsNull(txtFFSlut) Then
            If FormatCheck(txtFFSlut, "date") = False Then
                GoTo ending
            End If
        End If
    End If
    
    If SRB.Value = True Then
        If DateDiff("d", txtSRBstart, txtSRBslut) < 0 Then
            dFunc.msgError = "Startdatoen (indtastet: " & txtSRBstart & ") skal ligge før slutdatoen (indtastet: " & txtSRBslut & ")"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Startdatoen (indtastet: " & txtSRBstart & ") skal ligge før slutdatoen (indtastet: " & txtSRBslut & ")")
            GoTo ending
        End If
        If (IsNull(txtSRBstart) And IsNull(txtSRBslut)) Then
            dFunc.msgError = "Til og/eller fra datoerne skal være udfyldt for SRB datoen"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Til og/eller fra datoerne skal være udfyldt for SRB datoen")
            GoTo ending
        End If
        If Not IsNull(txtSRBstart) Then
            If FormatCheck(txtSRBstart, "date") = False Then
                GoTo ending
            End If
            
        End If
        If Not IsNull(txtSRBslut) Then
            If FormatCheck(txtSRBslut, "date") = False Then
                GoTo ending
            End If
        End If
        
    End If
    
    If Stiftelsesdato.Value = True Then
        
        If DateDiff("d", txtSTIstart, txtSTIslut) < 0 Then
            dFunc.msgError = "Startdatoen (indtastet: " & txtSTIstart & ") skal ligge før slutdatoen (indtastet: " & txtSTIslut & ")"
            SFunc.ShowFunc ("frmMsg")
           ' MsgBox ("Startdatoen (indtastet: " & txtSTIstart & ") skal ligge før slutdatoen (indtastet: " & txtSTIslut & ")")
            GoTo ending
        End If
        
        If (IsNull(txtSTIstart) And IsNull(txtSTIslut)) Then
            dFunc.msgError = "Til og/eller fra datoerne skal være udfyldt for Stiftelsesdatoen"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Til og/eller fra datoerne skal være udfyldt for Stiftelsesdatoen")
            GoTo ending
        End If
        If Not IsNull(txtSTIstart) Then
            If FormatCheck(txtSTIstart, "date") = False Then
                GoTo ending
            End If
            
        End If
        If Not IsNull(txtSTIslut) Then
            If FormatCheck(txtSTIslut, "date") = False Then
                GoTo ending
            End If
        End If
        
    End If
    
    
    If PeriodeStartdato.Value = True Then
        
        If DateDiff("d", txtPSTstart, txtPSTslut) < 0 Then
            dFunc.msgError = "Startdatoen (indtastet: " & txtPSTstart & ") skal ligge før slutdatoen (indtastet: " & txtPSTslut & ")"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Startdatoen (indtastet: " & txtPSTstart & ") skal ligge før slutdatoen (indtastet: " & txtPSTslut & ")")
            GoTo ending
        End If
        
        If (IsNull(txtPSTstart) And IsNull(txtPSTslut)) Then
            dFunc.msgError = "Til og/eller fra datoerne skal være udfyldt for Periode startdatoen"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Til og/eller fra datoerne skal være udfyldt for Periode startdatoen")
            GoTo ending
        End If
        If Not IsNull(txtPSTstart) Then
            If FormatCheck(txtPSTstart, "date") = False Then
                GoTo ending
            End If
            
        End If
        If Not IsNull(txtPSTslut) Then
            If FormatCheck(txtPSTslut, "date") = False Then
                GoTo ending
            End If
        End If
        
    End If
    
    If PeriodeSlutdato.Value = True Then
        
        If DateDiff("d", txtPSLstart, txtPSLslut) < 0 Then
            dFunc.msgError = "Startdatoen (indtastet: " & txtPSLstart & ") skal ligge før slutdatoen (indtastet: " & txtPSLslut & ")"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Startdatoen (indtastet: " & txtPSLstart & ") skal ligge før slutdatoen (indtastet: " & txtPSLslut & ")")
            GoTo ending
        End If
        
        If (IsNull(txtPSLstart) And IsNull(txtPSLslut)) Then
            dFunc.msgError = "Til og/eller fra datoerne skal være udfyldt for Periode startdatoen"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("Til og/eller fra datoerne skal være udfyldt for Periode startdatoen")
            GoTo ending
        End If
        If Not IsNull(txtPSLstart) Then
            If FormatCheck(txtPSLstart, "date") = False Then
                GoTo ending
            End If
            
        End If
        If Not IsNull(txtPSLslut) Then
            If FormatCheck(txtPSLslut, "date") = False Then
                GoTo ending
            End If
        End If
        
    End If
    
    
    Worksheets("SpmSvar").Range("C7:C7").Value = Controls("Label1").Caption
    
    If Forfaldsdato.Value = True Then
        Worksheets("SpmSvar").Range("D8:D8").Value = "Forfaldsdato"
        Worksheets("SpmSvar").Range("E8:E8").Value = txtFFStart
        Worksheets("SpmSvar").Range("F8:F8").Value = txtFFSlut
    Else
        Worksheets("SpmSvar").Range("D8:D8").Value = ""
        Worksheets("SpmSvar").Range("E8:E8").Value = ""
        Worksheets("SpmSvar").Range("F8:F8").Value = ""
    End If
    
    If SRB.Value = True Then
        Worksheets("SpmSvar").Range("D9:D9").Value = "SRB Dato"
        Worksheets("SpmSvar").Range("E9:E9").Value = txtSRBstart
        Worksheets("SpmSvar").Range("F9:F9").Value = txtSRBslut
    Else
        Worksheets("SpmSvar").Range("D9:D9").Value = ""
        Worksheets("SpmSvar").Range("E9:E9").Value = ""
        Worksheets("SpmSvar").Range("F9:F9").Value = ""
    End If
    
    If Stiftelsesdato.Value = True Then
        Worksheets("SpmSvar").Range("D10:D10").Value = "Stiftelsesdato"
        Worksheets("SpmSvar").Range("E10:E10").Value = txtSTIstart
        Worksheets("SpmSvar").Range("F10:F10").Value = txtSTIslut
    Else
        Worksheets("SpmSvar").Range("D10:D10").Value = ""
        Worksheets("SpmSvar").Range("E10:E10").Value = ""
        Worksheets("SpmSvar").Range("F10:F10").Value = ""
    End If
    
    If PeriodeStartdato.Value = True Then
        Worksheets("SpmSvar").Range("D11:D11").Value = "PeriodeStartdato"
        Worksheets("SpmSvar").Range("E11:E11").Value = txtPSTstart
        Worksheets("SpmSvar").Range("F11:F11").Value = txtPSTslut
    Else
        Worksheets("SpmSvar").Range("D11:D11").Value = ""
        Worksheets("SpmSvar").Range("E11:E11").Value = ""
        Worksheets("SpmSvar").Range("F11:F11").Value = ""
    End If
    
    If PeriodeSlutdato.Value = True Then
        Worksheets("SpmSvar").Range("D12:D12").Value = "PeriodeSlutdato"
        Worksheets("SpmSvar").Range("E12:E12").Value = txtPSLstart
        Worksheets("SpmSvar").Range("F12:F12").Value = txtPSLslut
    Else
        Worksheets("SpmSvar").Range("D12:D12").Value = ""
        Worksheets("SpmSvar").Range("E12:E12").Value = ""
        Worksheets("SpmSvar").Range("F12:F12").Value = ""
    End If
    
    
    If Forfaldsdato.Value = True Then
        Worksheets("Population").Range("B6:B6").Value = txtFFStart.Value
        Worksheets("Population").Range("B7:B7").Value = txtFFSlut.Value
    Else
        Worksheets("Population").Range("B6:B6").Value = ""
        Worksheets("Population").Range("B7:B7").Value = ""
    End If
    
    If SRB.Value = True Then
        Worksheets("Population").Range("B8:B8").Value = txtSRBstart.Value
        Worksheets("Population").Range("B9:B9").Value = txtSRBslut.Value
    Else
        Worksheets("Population").Range("B8:B8").Value = ""
        Worksheets("Population").Range("B9:B9").Value = ""
        
    End If
    
    If Stiftelsesdato.Value = True Then
        Worksheets("Population").Range("B10:B10").Value = txtSTIstart.Value
        Worksheets("Population").Range("B11:B11").Value = txtSTIslut.Value
    Else
        Worksheets("Population").Range("B10:B10").Value = ""
        Worksheets("Population").Range("B11:B11").Value = ""
        
    End If
    
    If PeriodeStartdato.Value = True Then
        Worksheets("Population").Range("B12:B12").Value = txtPSTstart.Value
        Worksheets("Population").Range("B13:B13").Value = txtPSTslut.Value
    Else
        Worksheets("Population").Range("B12:B12").Value = ""
        Worksheets("Population").Range("B13:B13").Value = ""
        
    End If
    
    If PeriodeSlutdato.Value = True Then
        Worksheets("Population").Range("B14:B14").Value = txtPSLstart.Value
        Worksheets("Population").Range("B15:B15").Value = txtPSLslut.Value
    Else
        Worksheets("Population").Range("B14:B14").Value = ""
        Worksheets("Population").Range("B15:B15").Value = ""
        
    End If
    
    If Forfaldsdato Or SRB Or Stiftelsesdato Or PeriodeStartdato Or PeriodeSlutdato Then
        'Worksheets("Population").range("B4:B4").Value = "01-09-2013" ' Workshop 25.09.2018
        Me.Hide
        SFunc.ShowFunc ("frm005")
        'frm005.Show
    Else
        dFunc.msgError = "Vælg som minimum et stamdatafelt for at gå videre"
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Vælg som minimum et stamdatafelt for at gå videre")
    
    End If
    
    Exit Sub
    
ending:
End Sub


Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm003")
'frm003.Show
End Sub

Private Sub txtSRBstart_Change()

End Sub

'Worksheets("Population").Range("B2:B2").Value = txtFordringsId.Value
'Worksheets("Population").Range("B3:B3").Value = cboFordringstype.Value
'Worksheets("Population").Range("B4:B4").Value = txtModtStart.Value
'Worksheets("Population").Range("B5:B5").Value = txtModtSlut.Value


Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

If Worksheets("SpmSvar").Range("D8:D8").Value = "Forfaldsdato" Then
    Call Forfaldsdato_Click
    Forfaldsdato.Value = True
    txtFFStart = Worksheets("SpmSvar").Range("E8:E8").Value
    txtFFSlut = Worksheets("SpmSvar").Range("F8:F8").Value
End If

If Worksheets("SpmSvar").Range("D9:D9").Value = "SRB Dato" Then
    Call SRB_Click
    SRB.Value = True
    txtSRBstart = Worksheets("SpmSvar").Range("E9:E9").Value
    txtSRBslut = Worksheets("SpmSvar").Range("F9:F9").Value
End If

If Worksheets("SpmSvar").Range("D10:D10").Value = "Stiftelsesdato" Then
    Call Stiftelsesdato_Click
    Stiftelsesdato.Value = True
    txtSTIstart = Worksheets("SpmSvar").Range("E10:E10").Value
    txtSTIslut = Worksheets("SpmSvar").Range("F10:F10").Value
End If

If Worksheets("SpmSvar").Range("D11:D11").Value = "PeriodeStartdato" Then
    Call PeriodeStartdato_Click
    PeriodeStartdato.Value = True
    txtPSTstart = Worksheets("SpmSvar").Range("E11:E11").Value
    txtPSTslut = Worksheets("SpmSvar").Range("F11:F11").Value
End If

If Worksheets("SpmSvar").Range("D12:D12").Value = "PeriodeSlutdato" Then
    Call PeriodeSlutdato_Click
    PeriodeSlutdato.Value = True
    txtPSLstart = Worksheets("SpmSvar").Range("E12:E12").Value
    txtPSLslut = Worksheets("SpmSvar").Range("F12:F12").Value
End If

End Sub
