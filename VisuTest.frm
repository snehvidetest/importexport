VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VisuTest 
   Caption         =   "UserForm1"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   22116
   OleObjectBlob   =   "VisuTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VisuTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
    Call DrawChart
End Sub

Private Sub OptionButton1_Click()
    Worksheets("Grafik01").Range("H19:H19").Value = "Periode start"
End Sub

Private Sub OptionButton2_Click()
    Worksheets("Grafik01").Range("H19:H19").Value = "Stiftelsesdato"
End Sub

Private Sub TextBox1_Change()
    Worksheets("Grafik01").Range("F19:F19").Value = val(TextBox1)
End Sub


Private Sub UserForm_Click()

End Sub


Private Sub DrawChart()

    Dim Fname As String

    Call SaveChart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    Me.Image1.Picture = LoadPicture(Fname)
    Call DeleteFile
    
End Sub

Private Sub SaveChart()

    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = Sheets("Grafik01").ChartObjects(1).Chart
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
