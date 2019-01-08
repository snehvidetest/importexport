Attribute VB_Name = "Frm001_test"
Private result As String
Private formID As Integer
Private formName As String
Private changed As Boolean
Private unChanged As Boolean
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
Sub RunTests()

    'Which form are we testing?
    formName = "frm001"
    formID = 1
    
    'Get parameters relevant for testcase including their respective columns
    Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)

    'Get the total number of testcases associated with the form
    Dim nrTC As Integer, i As Integer
    nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)
    
    'Open progressFile For Output As #1
    
    'Run all testcases incl. printing of results to the testcase workbook
    For i = 1 To nrTC
        Set parameters = New Scripting.Dictionary
        Testcase i
    Next i

    
    
    
End Sub


'The following code is the skeleton for form 1 testcases.
Private Function Testcase(tc As Integer)
    Dim review As Boolean, tcid As String

    'Reset spørgeskema workbook
    Global_Test_Func.resetSheets ThisWorkbook

    'Create the TCID
    tcid = Global_Test_Func.GetTCID(tc, formID)
    If logging Then
        Write #1, tcid
    End If
    
    Set parameters = New Scripting.Dictionary 'Resets the testcase parameters
    Set parameters = Global_Test_Func.getData(tcid, parametersAndCols)

    'Clear all fields related to spørskema
    'ClearAllFields ThisWorkbook

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
    
    setRandomConfig
    Select Case parameters("testSubject")
        Case "printsToSpmSheet"
            If (parameters("testParameter") = "defaultConfig") Then
                TestDefaultConfig ("SpmSvar")
            Else
                TestExistingConfig ("SpmSvar")
            End If
        Case "printsToPopSheet"
            If (parameters("testParameter") = "defaultConfig") Then
                 TestDefaultConfig ("Population")
            Else
                 TestExistingConfig ("Population")
            End If
        Case "printsToRulSheet"
            If (parameters("testParameter") = "defaultConfig") Then
                 TestDefaultConfig ("Regler")
            Else
                TestExistingConfig ("Regler")
            End If
        Case "printsToGroSheet"
            If (parameters("testParameter") = "defaultConfig") Then
                 TestDefaultConfig ("Gruppering")
            Else
                 TestExistingConfig ("Gruppering")
            End If
        Case "errorMessage"
            frm001.CommandButton1_Click
            result = Global_Test_Func.errorMessage()
        Case "nextStep"
            frm001.OKButton_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
        Case "backButton"
            frm001.CommandButton1_Click
            frmMsg.CommandButton1_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
    End Select
    
    'Compare actual and expected
    If result = parameters("expected") Then
        review = True
    Else:
        review = False
    End If

    Call KillForms
    
     'Print results
    Global_Test_Func.PrintTestResults tcid, result, review
    

End Function

Private Function setRandomConfig()
    Worksheets("SpmSvar").Range("D6").Value = frm003.OptionButton2.Caption
    Worksheets("SpmSvar").Range("E4").Value = ""
    Worksheets("Population").Range("B3").Value = "VAASKAT"
    Worksheets("Population").Range("B5").Value = ""
    Worksheets("Regler").Range("G48").Value = "JA"
    Worksheets("Regler").Range("G49").Value = "JA"
    Worksheets("Regler").Range("G50").Value = "JA"
    Worksheets("Gruppering").Range("C6").Value = "JA"
    Worksheets("Gruppering").Range("C7").Value = "JA"

End Function


Private Function TestExistingConfig(sheet As String)
    frm001.OKButton_Click
    
    Select Case sheet
        Case "SpmSvar"
            ChangedFunc sheet, "D6", frm003.OptionButton2.Caption
            ChangedFunc sheet, "E4", ""
        Case "Population"
            ChangedFunc sheet, "B3", "VAASKAT"
            ChangedFunc sheet, "B5", ""
        Case "Regler"
            ChangedFunc sheet, "G48", "JA"
            ChangedFunc sheet, "G49", "JA"
            ChangedFunc sheet, "G50", "JA"
        Case "Gruppering"
            ChangedFunc sheet, "C6", "JA"
            ChangedFunc sheet, "C7", "JA"
    End Select

    If changed = True Then
            result = "False"
        Else
            result = "True"
    End If
    
    changed = False
    
End Function

Private Function ChangedFunc(sheet As String, cell As String, addedString As String)
    If (Worksheets(sheet).Range(cell).Text <> addedString) Then
        changed = True
    End If
    
End Function


Private Function TestDefaultConfig(sheet As String)
    frm001.CommandButton1_Click
    frmMsg.CommandButton1_Click
    Select Case sheet
        Case "SpmSvar"
            ChangedFunc sheet, "D6", ""
            ChangedFunc sheet, "E4", ""
            
        Case "Population"
            ChangedFunc sheet, "B4", ""
            ChangedFunc sheet, "B5", ""
        Case "Regler"
            ChangedFunc sheet, "G48", "NEJ"
            ChangedFunc sheet, "G49", "NEJ"
            ChangedFunc sheet, "G50", "NEJ"
        Case "Gruppering"
            ChangedFunc sheet, "C6", "NEJ"
            ChangedFunc sheet, "C7", "NEJ"
    End Select
    
    If changed = True Then
            result = "False"
    Else
            result = "True"
    End If
        changed = False
End Function


Private Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm001") Then
        Unload frm001
    End If
    If Global_Test_Func.IsLoaded("frm002") Then
        Unload frm002
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function



