VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm034 
   Caption         =   "Frasortering"
   ClientHeight    =   7850
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11328
   OleObjectBlob   =   "frm034.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub OKButton_Click()
    
    ' Validering for numeriske værdier
    
    Dim cControl As Control
        
    For Each cControl In Me.Controls
        
        control_type = UCase(Left(cControl.Name, 4))
            
        If control_type = "TEXT" Then
           If cControl.Text = "" Then
              cControl.SetFocus
              dFunc.msgError = "Felt skal udfyldes med tal."
              SFunc.ShowFunc ("frmMsg")
              GoTo ending
           End If
           
           If cControl.Text <> "" Then
              If IsNumeric(cControl.Text) = False Then
                 cControl.SetFocus
                 dFunc.msgError = "Felt skal udfyldes med tal."
                 SFunc.ShowFunc ("frmMsg")
                 GoTo ending
              End If
           End If
        End If
        
    Next cControl
    
    ' Validering for forkert anvendelse af før/efter
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "før" Then
        dFunc.msgError = "Forkert anvendelse af før/efter"
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
    End If
    
    ' Validering for 'efter'
    
    If ComboBox2.Value = "efter" Then
        If Int(TextBox1.Value) > Int(TextBox2.Value) Then
            dFunc.msgError = "Værdien i 'Fra' skal være mindre end værdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    Dim antal As Integer
           
    Dim x1 As Variant
    Dim x2 As Variant
    
    ' Reset values
    
    Call Insert_to_sheet("Regler", "J14:O14", "")
    
    'Relationen mellem "forfaldsdato" og "sidste rettidige betalingsdato"
    
    x1 = TextBox1.Value
    x2 = TextBox2.Value
    
    ' 'Før' fra foranstilles med minus
    If ComboBox2.Value = "før" Then
        x1 = "-" + x1
    End If
    
    ' 'Før' fra foranstilles med minus
    If ComboBox4.Value = "før" Then
        x2 = "-" + x2
    End If
    
    ' Validering for 'før'
    
    If ComboBox2.Value = "før" Then
        If Int(x1) > Int(x2) Then
            dFunc.msgError = "Værdien i 'Fra' skal være mindre end værdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Validering af 'Sidste rettidige betalingsdato' kan ligge samme dag som eller op til 119 dage efter 'Forfaldsdato'.
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "før" Then
        If (Int(TextBox1.Value) - Int(TextBox2.Value) > 119) Then
            dFunc.msgError = "Antal dage mellem 'Sidste rettidige betalingsdato' og 'Forfaldsdato' kan maksimalt være 119 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) + Int(TextBox1.Value) > 119) Then
            dFunc.msgError = "Antal dage mellem 'Sidste rettidige betalingsdato' og 'Forfaldsdatoen' kan maksimalt være 119 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 119) Then
            dFunc.msgError = "Antal dage mellem 'Sidste rettidige betalingsdato' og 'Forfaldsdatoen' kan maksimalt være 119 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
            
    ' Indsæt værdier i regeler
    Call Insert_to_sheet("Regler", "J14:J14", x1)
    Call Insert_to_sheet("Regler", "M14:M14", x2)
        
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G14:G14", "JA")
          
    ' Relationen mellem "forfaldsdato" og "sidste rettidige betalingsdato"
    a = "SRB"
    b = "Forfaldsdato"
    VisuTitle = a & " i forhold til " & b
    
    ' Skriv svar ned i 'SpmSvar'
    Worksheets("SpmSvar").Range("C60:C60").Value = VisuTitle
    Worksheets("SpmSvar").Range("D60:D60").Value = TextBox1.Value
    Worksheets("SpmSvar").Range("E60:E60").Value = "dage"
    Worksheets("SpmSvar").Range("F60:F60").Value = ComboBox2.Value
    Worksheets("SpmSvar").Range("G60:G60").Value = TextBox2.Value
    Worksheets("SpmSvar").Range("H60:H60").Value = "dage"
    Worksheets("SpmSvar").Range("I60:I60").Value = ComboBox4.Value
    Worksheets("SpmSvar").Range("C59:C59").Value = Label48.Caption
    
    ' Hvis fordringshaver svarer, at "sidste rettidige betalingsdato" kan ligge før "forfaldsdatoen"
    ' skal der komme en advarsel om, at dette ikke er "normalt".
    If ComboBox2.Value = "før" Or ComboBox4.Value = "før" Then
        SFunc.ShowFunc ("frm044")
        GoTo ending
    End If
    
    Me.Hide
    SFunc.ShowFunc ("frm035")
       
ending:
End Sub

Private Sub TextBox1_Change()
    
    Count = TextBox1.Value
    DMY = "dage"
    FE = ComboBox2.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
    
    ' Antal dage
    x = Count
        
    Worksheets("SpmSvar").Range("L2") = x
    
    Worksheets("SpmSvar").Range("K2") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("SpmSvar").Range("L2") = x
        
    ElseIf FE = "før" Then
    
        Worksheets("SpmSvar").Range("L2") = -x
        
    End If
    
    Call DrawChart
    
    
End Sub

Private Sub TextBox2_Change()
    
    Count = TextBox2.Value
    DMY = "dage"
    FE = ComboBox4.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
       
    x = Count
    
    Worksheets("SpmSvar").Range("L4") = x
    
    Worksheets("SpmSvar").Range("K4") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("SpmSvar").Range("L4") = x
        
    ElseIf FE = "før" Then
    
        Worksheets("SpmSvar").Range("L4") = -x
        
    End If
    
    Call DrawChart

End Sub

Private Sub DrawChart()

    Dim Fname As String

    Call SaveChart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    Me.Image2.Picture = LoadPicture(Fname)
    Call DeleteFile
    
End Sub

Private Sub SaveChart()

    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = Sheets("SpmSvar").ChartObjects(1).Chart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    MyChart.Export Filename:=Fname, FilterName:="GIF"
    
End Sub

Sub DeleteFile()

    Dim Fname As String
    On Error Resume Next
    Fname = ThisWorkbook.Path & "\temp1.gif"
    Kill Fname
    On Error GoTo 0
    
End Sub
Public Sub Tilbage_Click()
    Me.Hide
    SFunc.ShowFunc ("frm039")
End Sub
Private Sub ComboBox2_Change()
    Call TextBox1_Change
End Sub
Private Sub ComboBox4_Change()
    Call TextBox2_Change
End Sub
Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeStretch

    a = "SRB"
    b = "Forfaldsdato"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("SpmSvar").Range("K3") = b
    Worksheets("SpmSvar").Range("J1") = VisuTitle
    Worksheets("SpmSvar").Range("K2") = "20 dage efter"
    Worksheets("SpmSvar").Range("L2") = 20
    Worksheets("SpmSvar").Range("K4") = "30 dage efter"
    Worksheets("SpmSvar").Range("L4") = 30
    
    ' Initier comboboxe
    With ComboBox2
        .AddItem "før"
        .AddItem "efter"
    End With
          
    With ComboBox4
        .AddItem "før"
        .AddItem "efter"
    End With

    
    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("I1").Select
 
    Call DrawChart
    
    ' Indlæs tidligere svar fra 'SpmSvar'
    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    TextBox1.Value = Worksheets("SpmSvar").Range("D60:D60").Value
    
    If Not IsEmpty(Worksheets("SpmSvar").Range("F60:F60").Value) Then
        ComboBox2.Value = Worksheets("SpmSvar").Range("F60:F60").Value
    End If
    
    TextBox2.Value = Worksheets("SpmSvar").Range("G60:G60").Value
    
    If Not IsEmpty(Worksheets("SpmSvar").Range("I60:I60").Value) Then
        ComboBox4.Value = Worksheets("SpmSvar").Range("I60:I60").Value
    End If
End Sub

