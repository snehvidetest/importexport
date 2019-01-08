VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm037 
   Caption         =   "Frasortering"
   ClientHeight    =   7850
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11328
   OleObjectBlob   =   "frm037.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm037"
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
        'MsgBox ("Forkert anvendelse af før/efter")
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
    
    Call Insert_to_sheet("Regler", "J21:O21", "")
    
    'Relationen mellem "stiftelsesdato" og "periode start"
    
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
            'MsgBox ("Værdien i 'Fra' skal være mindre end værdien i 'Til'.")
            GoTo ending
        End If
    End If
    
    ' Validering af 'Stiftelsesdato' kan ligge samme dag som eller op til 365 dage efter 'Periode start'.
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "før" Then
        If (Int(TextBox1.Value) - Int(TextBox2.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Periode start' kan maksimalt være 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) + Int(TextBox1.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Periode start' kan maksimalt være 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Periode start' kan maksimalt være 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Indsæt værdier i regler
    Call Insert_to_sheet("Regler", "J21:J21", x1)
    Call Insert_to_sheet("Regler", "M21:M21", x2)
   
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G21:G21", "JA")
    
    ' Skriv svar ned i 'SpmSvar'
    
    ' Relationen mellem "stiftelsesdato" og "periode start"
    a = "Stiftelsesdato"
    b = "Periode start"
    VisuTitle = a & " i forhold til " & b
    Worksheets("SpmSvar").Range("C63:C63").Value = VisuTitle
    Worksheets("SpmSvar").Range("D63:D63").Value = TextBox1.Value
    Worksheets("SpmSvar").Range("E63:E63").Value = "dage"
    Worksheets("SpmSvar").Range("F63:F63").Value = ComboBox2.Value
    Worksheets("SpmSvar").Range("G63:G63").Value = TextBox2.Value
    Worksheets("SpmSvar").Range("H63:H63").Value = "dage"
    Worksheets("SpmSvar").Range("I63:I63").Value = ComboBox4.Value
    
    ' Hvis fordringshaver svarer, at "stiftelsesdato" kan ligge før "periode start"
    ' skal der komme en advarsel om, at dette ikke er "normalt".
        
    If ComboBox2.Value = "før" Or ComboBox4.Value = "før" Then
        SFunc.ShowFunc ("frm047")
        GoTo ending
    End If
     
    Me.Hide
    SFunc.ShowFunc ("frm021")
    
ending:
End Sub

Private Sub TextBox1_Change()
    
    Count = TextBox1.Value
    DMY = "dage"
    FE = ComboBox2.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
    
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
    SFunc.ShowFunc ("frm036")
End Sub

Private Sub ComboBox2_Change()
    Call TextBox1_Change
End Sub

Private Sub ComboBox4_Change()
    Call TextBox2_Change
End Sub
Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeStretch

    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("I1").Select
    
    a = "Stiftelsesdato"
    b = "Periode start"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("SpmSvar").Range("K3") = b
    Worksheets("SpmSvar").Range("J1") = VisuTitle
    Worksheets("SpmSvar").Range("K2") = "20 dage efter"
    Worksheets("SpmSvar").Range("L2") = 20
    Worksheets("SpmSvar").Range("K4") = "30 dage efter"
    Worksheets("SpmSvar").Range("L4") = 30
    
    Call DrawChart
    
    With ComboBox2
        .AddItem "før"
        .AddItem "efter"
    End With
    
    With ComboBox4
        .AddItem "før"
        .AddItem "efter"
    End With

    
    ' Indlæs tidligere svar fra 'SpmSvar'

    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    TextBox1.Value = Worksheets("SpmSvar").Range("D63:D63").Value
    If Not IsEmpty(Worksheets("SpmSvar").Range("F63:F63").Value) Then ComboBox2.Value = Worksheets("SpmSvar").Range("F63:F63").Value
    TextBox2.Value = Worksheets("SpmSvar").Range("G63:G63").Value
    If Not IsEmpty(Worksheets("SpmSvar").Range("I63:I63").Value) Then ComboBox4.Value = Worksheets("SpmSvar").Range("I63:I63").Value

End Sub

