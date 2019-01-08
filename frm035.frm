VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm035 
   Caption         =   "Frasortering"
   ClientHeight    =   7850
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11328
   OleObjectBlob   =   "frm035.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub OKButton_Click()
       
    ' Validering for numeriske v�rdier
    
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
    
    ' Validering for forkert anvendelse af f�r/efter
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "f�r" Then
        dFunc.msgError = "Forkert anvendelse af f�r/efter"
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
    End If
            
    ' Validering for 'efter'
    
    If ComboBox2.Value = "efter" Then
        If Int(TextBox1.Value) > Int(TextBox2.Value) Then
            dFunc.msgError = "V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'. (2)"
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    'Dim antal As Integer
    
    Dim x1 As Variant
    Dim x2 As Variant
    
    ' Reset values
    
    Call Insert_to_sheet("Regler", "J15:O15", "")
    
    'Relationen mellem forfaldsdato" og "stiftelsesdato"
    
    x1 = TextBox1.Value
    x2 = TextBox2.Value
    
     ' 'F�r' fra foranstilles med minus
    If ComboBox2.Value = "f�r" Then
        x1 = "-" + x1
    End If
    
    ' 'F�r' fra foranstilles med minus
    If ComboBox4.Value = "f�r" Then
        x2 = "-" + x2
    End If
    
    ' Validering for 'f�r'
    
    If ComboBox2.Value = "f�r" Then
        If Int(x1) > Int(x2) Then
            dFunc.msgError = "V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Validering af 'Stiftelsesdato' kan ligge samme dag som eller op til 365 dage efter 'Forfaldsdato'.
    
    If ComboBox2.Value = "f�r" And ComboBox4.Value = "f�r" Then
        If (Int(TextBox1.Value) - Int(TextBox2.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Forfaldsdato' kan maksimalt v�re 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "f�r" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) + Int(TextBox1.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Forfaldsdato' kan maksimalt v�re 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Forfaldsdato' kan maksimalt v�re 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Inds�t v�rdier i regeler
    Call Insert_to_sheet("Regler", "J15:J15", x1)
    Call Insert_to_sheet("Regler", "M15:M15", x2)
    
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G15:G15", "JA")
    
    ' Skriv svar ned i 'SpmSvar'
    
    ' Relationen mellem forfaldsdato" og "stiftelsesdato"
    a = "Stiftelsesdato"
    b = "Forfaldsdato"
    VisuTitle = a & " i forhold til " & b
    Worksheets("SpmSvar").Range("C61:C61").Value = VisuTitle
    Worksheets("SpmSvar").Range("D61:D61").Value = TextBox1.Value
    Worksheets("SpmSvar").Range("E61:E61").Value = "dage"
    Worksheets("SpmSvar").Range("F61:F61").Value = ComboBox2.Value
    Worksheets("SpmSvar").Range("G61:G61").Value = TextBox2.Value
    Worksheets("SpmSvar").Range("H61:H61").Value = "dage"
    Worksheets("SpmSvar").Range("I61:I61").Value = ComboBox4.Value
    
    ' Hvis fordringshaver svarer, at "stiftelsesdato" kan ligge efter "forfaldsdatoen"
    ' skal der komme en advarsel om, at dette ikke er "normalt".
    
    
    If ComboBox2.Value = "efter" Or ComboBox4.Value = "efter" Then
       SFunc.ShowFunc ("frm045")
       GoTo ending
    End If
    
    Me.Hide
    SFunc.ShowFunc ("frm036")
       
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
        
    ElseIf FE = "f�r" Then
    
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
        
    ElseIf FE = "f�r" Then
    
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
    SFunc.ShowFunc ("frm034")
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
    b = "Forfaldsdato"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("SpmSvar").Range("K3") = b
    Worksheets("SpmSvar").Range("J1") = VisuTitle
    Worksheets("SpmSvar").Range("K2") = "20 dage efter"
    Worksheets("SpmSvar").Range("L2") = 20
    Worksheets("SpmSvar").Range("K4") = "30 dage efter"
    Worksheets("SpmSvar").Range("L4") = 30
    
    Call DrawChart
    
    With ComboBox2
        .AddItem "f�r"
        .AddItem "efter"
    End With
    
    With ComboBox4
        .AddItem "f�r"
        .AddItem "efter"
    End With

    
    ' Indl�s tidligere svar fra 'SpmSvar'

    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    TextBox1.Value = Worksheets("SpmSvar").Range("D61:D61").Value
    If Not IsEmpty(Worksheets("SpmSvar").Range("F61:F61").Value) Then ComboBox2.Value = Worksheets("SpmSvar").Range("F61:F61").Value
    TextBox2.Value = Worksheets("SpmSvar").Range("G61:G61").Value
    If Not IsEmpty(Worksheets("SpmSvar").Range("I61:I61").Value) Then ComboBox4.Value = Worksheets("SpmSvar").Range("I61:I61").Value

End Sub

