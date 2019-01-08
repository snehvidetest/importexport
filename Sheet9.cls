VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public spmChangedCells As Scripting.Dictionary

Private Sub Worksheet_Change(ByVal Target As Range)
    If (Sheet1.recordChangingCells = True) Then
        If (spmChangedCells.exists(Target.Address(0, 0)) = False) Then
            spmChangedCells.Add Target.Address(0, 0), Target
        End If
    End If
End Sub
Private Sub GeneratePDF_Click()
    Application.ScreenUpdating = False
    Worksheets("PDF").Activate
    
    ' Declare Arrays
    Dim spm As Variant
    Dim pdf As Variant
    Dim check As Variant
    check = ""
    
    ' Clear All
    ActiveSheet.UsedRange.ClearContents
    
    ' Create PDF sheet fram SpmSvar sheet
    Call Spm2PDF
    
    ' Insert Intro at the top
    Call CreateIntro("Opsummering af spørgeskema")
    
    ' Format, style and clean up sheet
    Call DeleteEmptyRows
    Call StyleSheet
    'Call HorizLines("A1:A200")
    
    ' Set reformat to True only if reformatting is needed. This may take some time.
    Dim reformat As Boolean
    reformat = False
    If reformat Then Call FormatCells
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub Spm2PDF()

    ' Spm 1
    spm = Array("C2", "D2")
    pdf = Array("A1", "B1")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 2
    spm = Array("C3", "D3")
    pdf = Array("A2", "B2")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 3
    spm = Array("C4", "D4", "E4")
    pdf = Array("A3", "B3", "D3")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    ActiveSheet.Range("C3") = "-"
    If IsEmpty(ActiveSheet.Range("D3")) Then ActiveSheet.Range("D3") = "dags dato"
    
    
    ' Spm 4
    spm = Array("C5", "D5")
    pdf = Array("A4", "B4")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 4a
    spm = Array("C6", "D6")
    pdf = Array("A5", "B5")
    check = "D6" ' only do something if spm->D6 is not empty
    Call CopyVal(spm, pdf, check)

    
    
    ' Spm 4a1
    spm = Array("C7", "D8", "E8", "F8", "D9", "E9", "F9", "D10", "E10", "F10", "D11", "E11", "F11", "D12", "E12", "F12")
    pdf = Array("A6", "B7", "C7", "D7", "B8", "C8", "D8", "B9", "C9", "D9", "B10", "C10", "D10", "B11", "C11", "D11")
    check = "" ' not done
    Call CopyVal(spm, pdf, check)
    If Not IsEmpty(ActiveSheet.Range("A6")) Then
        ActiveSheet.Range("B6") = "Stamdatofelt"
        ActiveSheet.Range("C6") = "Fra"
        ActiveSheet.Range("D6") = "Til"
    End If
    
    ' Spm 5
    spm = Array("C13", "D13")
    pdf = Array("A12", "B12")
    check = ""
    Call CopyVal(spm, pdf, check)
    
        
    ' Spm 6
    spm = Array("C14", "D14")
    pdf = Array("A13", "B13")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 7
    spm = Array("C15", "D15")
    pdf = Array("A14", "B14")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 8
    spm = Array("C16", "D16")
    pdf = Array("A15", "B15")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    ' Spm 9
    spm = Array("C17", "D17")
    pdf = Array("A16", "B16")
    check = ""
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 9a
    spm = Array("C18", "D18")
    pdf = Array("A17", "B17")
    check = "D18"
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 9a2
    spm = Array("C19", "D19")
    pdf = Array("A18", "B18")
    check = "D19"
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 9a22
    spm = Array("C20", "D20")
    pdf = Array("A19", "B19")
    check = "D20"
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 9b
    spm = Array("C21", "D21")
    pdf = Array("A20", "B20")
    check = "D21"
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 9b2
    spm = Array("C22", "D22")
    pdf = Array("A21", "B21")
    check = "D22"
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 9b22
    spm = Array("C23", "D23")
    pdf = Array("A22", "B22")
    check = "D23"
    Call CopyVal(spm, pdf, check)
    
    
    ' Spm 10
    spm = Array("C24", "D24", "E24", "F24", "G24", "H24", "I24")
    pdf = Array("A22")
    check = "D24"
    Call CopyVal(Array(spm(0)), pdf, check) ' only copy question
    If Not IsEmpty(ActiveSheet.Range("A22")) Then
        result = ""
        For Each cell In spm
            If Split(Worksheets("SpmSvar").Range(cell).Value, " ")(1) = True Then
                result = result & Split(Worksheets("SpmSvar").Range(cell).Value, " ")(0) & ", "
            End If
        Next
        result = Left(result, Len(result) - 2)
        ActiveSheet.Range("B22") = result
       
        ' See spmsvar line 71-101
        spm = Array("C71", "D71", "C72", "D72", "C73", "D73")
        pdf = Array("A23", "B23", "A24", "B24", "A25", "B25")
        check = "D71" ' This may need double checking
        Call CopyVal(spm, pdf, check)
        
        spm = Array("C76", "D76", "C77", "D77", "C78", "D78")
        pdf = Array("A26", "B26", "A27", "B27", "A28", "B28")
        check = "D76" ' This may need double checking
        Call CopyVal(spm, pdf, check)
            
        spm = Array("C81", "D81", "C82", "D82", "C83", "D83")
        pdf = Array("A29", "B29", "A30", "B30", "A31", "B31")
        check = "D81" ' This may need double checking
        Call CopyVal(spm, pdf, check)
                
        spm = Array("C86", "D86", "C87", "D87", "C88", "D88")
        pdf = Array("A32", "B32", "A33", "B33", "A34", "B34")
        check = "D86" ' This may need double checking
        Call CopyVal(spm, pdf, check)
    
        spm = Array("C92", "D92", "C93", "D93", "C94", "D94")
        pdf = Array("A35", "B35", "A36", "B36", "A37", "B37")
        check = "D92" ' This may need double checking
        Call CopyVal(spm, pdf, check)
        
        For Each cell In Array("B24", "B25", "B27", "B28", "B30", "B31", "B33", "B34", "B36", "B37")
            If Not IsEmpty(ActiveSheet.Range(cell)) Then
                ActiveSheet.Range(cell) = CStr(ActiveSheet.Range(cell)) & " dage"
            End If
        Next
    
    End If
        
    
    ' Spm 11
    spm = Array("C27", "D27", "C28", "D28", "E28")
    pdf = Array("A38", "B38", "A39", "B39", "D39")
    check = "D27"
    Call CopyVal(spm, pdf, check)
    If Not IsEmpty(ActiveSheet.Range("A38")) Then
        ActiveSheet.Range("C39") = "-" ' for date interval
    End If
    
    'Spm 11a-c
    spm = Array("C59", _
                "C60", "D60", "E60", "F60", "G60", "H60", "I60", _
                "C61", "D61", "E61", "F61", "G61", "H61", "I61", _
                "C62", "D62", "E62", "F62", "G62", "H62", "I62")
    pdf = Array("A40", _
                "A41", "B41", "C41", "D41", "F41", "G41", "H41", _
                "A42", "B42", "C42", "D42", "F42", "G42", "H42", _
                "A43", "B43", "C43", "D43", "F43", "G43", "H43")
    check = ""
    Call CopyVal(spm, pdf, check)
    If Not IsEmpty(ActiveSheet.Range("A40")) Then
        ActiveSheet.Range("E41:E43") = "-" ' for date interval
    End If
    
    'spm 11d
    spm = Array("C63", "D63", "E63", "F63", "G63", "H63", "I63")
    pdf = Array("A44", "B44", "C44", "D44", "F44", "G44", "H44")
    check = "C63"
    Call CopyVal(spm, pdf, check)
    If Not IsEmpty(ActiveSheet.Range("A44")) Then
        ActiveSheet.Range("E44") = "-" ' for date interval
    End If


    'spm 11e
    spm = Array("C64", "D64", "E64", "F64", "G64", "H64", "I64")
    pdf = Array("A45", "B45", "C45", "D45", "F45", "G45", "H45")
    check = "C64"
    Call CopyVal(spm, pdf, check)
    If Not IsEmpty(ActiveSheet.Range("A45")) Then
        ActiveSheet.Range("E55") = "-" ' for date interval
    End If
    
    
    ' Spm 12
    spm = Array("C55", "D55")
    pdf = Array("A47", "B47")
    check = "D55"
    Call CopyVal(spm, pdf, check)
    
    ' Spm 13
    spm = Array("C30", _
                "C31", "D31", "E31", "F31", _
                "D32", "E32", "F32", _
                "D33", "E33", "F33", _
                "D34", "E34", "F34", _
                "D35", "E35", "F35", _
                "C36", "D36", "E36", "F36", _
                "D37", "E37", "F37", _
                "D38", "E38", "F38", _
                "D39", "E39", "F39", _
                "D40", "E40", "F40" _
                )
    
    ' move everything 1-20 row(s) down!
    pdf = Array("A49", _
                "A50", "B50", "C50", "D50", _
                "B51", "C51", "D51", _
                "B52", "C52", "D52", _
                "B53", "C53", "D53", _
                "B54", "C54", "D54", _
                "A55", "B55", "C55", "D55", _
                "B56", "C56", "D56", _
                "B57", "C57", "D57", _
                "B58", "C58", "D58", _
                "B59", "C59", "D59" _
                )
    check = ""
    Call CopyVal(spm, pdf, check)
    
    ' Spm 13
    
    ' Spm 14
    
    ' Spm 15
    
    
End Sub

Private Sub CopyVal(spm As Variant, pdf As Variant, check As Variant)
    
    ' Copy value from SpmSvar to PDF-sheet
    
    ' Only copy if question was answered
    If check = "" Then
        ' do nothing
    ElseIf Not IsEmpty(Worksheets("SpmSvar").Range(check).Value) Then
        ' do nothing
    Else
        Exit Sub ' break
    End If


    cnt = UBound(spm)

    For i = 0 To cnt
    
        Worksheets("PDF").Range(pdf(i)).Value = Worksheets("SpmSvar").Range(spm(i)).Value
        
    Next i
    
End Sub


Private Sub FormatCells()
    ' Format cell content and width/height
    
    Dim mCell As Range
    
    For Each mCell In Worksheets("PDF").UsedRange.Cells
        mCell.EntireColumn.AutoFit
      
        If mCell.EntireColumn.ColumnWidth > 50 Then mCell.EntireColumn.ColumnWidth = 50
        If mCell.EntireColumn.ColumnWidth < 8 Then mCell.EntireColumn.ColumnWidth = 8
        
        mCell.WrapText = True
        mCell.EntireRow.AutoFit
    Next mCell
    
End Sub


Private Sub DeleteEmptyRows()
    Dim r As Range, rows As Long, i As Long
    Set r = ActiveSheet.Range("A1:I200")
    rows = r.rows.Count
    For i = rows To 1 Step (-1)
        If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
    Next
End Sub


Private Sub CreateIntro(introString As Variant)
    
    'Insert row in the top
    ActiveSheet.Range("A1").EntireRow.Insert
    ActiveSheet.Range("A1").Value = introString & vbNewLine
    
End Sub


Private Sub StyleSheet()
    ' Used for styling
    With ActiveSheet
        .Columns("B:Z").HorizontalAlignment = xlCenter
        .Columns("B:Z").VerticalAlignment = xlBottom
    End With

End Sub


Private Sub HorizLines(rng As Variant)
    
    ' Not finished!
    
    For Each cell In ActiveSheet.Range(rng)
        If Not IsEmpty(cell) Then
            With cell.Borders(xlTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End If
    Next
    
End Sub


Private Sub Debugger()
    
    For Each cell In ActiveSheet.Range("A1:A10")
        cell.Value = "Test"
    Next
    
End Sub


