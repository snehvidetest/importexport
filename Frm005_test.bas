Attribute VB_Name = "Frm005_test"
Private result As String
Private formID As Integer
Private formName As String
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
Private spmCells() As Variant
Private popCells() As Variant
Private rulCells() As Variant
Private groCells() As Variant

Sub RunTests()

    formName = "frm005"
    formID = 5
    
    Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)

    
    Dim nrTC As Integer, i As Integer
    nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)
    
    For i = 1 To nrTC
        Set parameters = New Scripting.Dictionary
        Testcase i
    Next i

End Sub



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
    
    Select Case parameters("testSubject")
        Case "printsToRulSheet"
            SetFields
            frm005.OKButton_Click 'Click on Videre button
            CheckFields "Regler"
            
        Case "printsToSpmSheet"
            SetFields
            frm005.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm005.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm005.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm005.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar", "D13"
            
        Case "noExtraPrints"
            SFunc.ShowFunc ("frm002") 'Prevent form 2 initialisation affecting results
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm005.Tilbage_Click 'Click back button
            Else
                frm005.OKButton_Click 'Click on Videre button
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
            
        Case Else
            MsgBox "Error in 'testsubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    End Select
    
    If result = parameters("expected") Then
        review = True
    Else:
        review = False
    End If

    KillForms
    
     'Print results
    Global_Test_Func.PrintTestResults tcid, result, review

End Function


Private Function SetFields()
    
    'The folowing code inserts the inputs into the actual form
    frm005.OptionButton1.Value = parameters("optionButton1")
    frm005.OptionButton2.Value = parameters("optionButton2")
  
End Function

Private Function CheckFields(sheet As String)
    'Check results
    If (sheet = "SpmSvar") Then
        result = ThisWorkbook.Sheets(sheet).Range("D13").Text
    Else
        Select Case parameters("rule")
            Case "R0006"
                result = ThisWorkbook.Sheets(sheet).Range("G7").Text
            Case "R0007"
                result = ThisWorkbook.Sheets(sheet).Range("G8").Text
            Case "R0008"
                result = ThisWorkbook.Sheets(sheet).Range("G9").Text
        End Select
    End If
End Function

Private Function DataIsSaved(sheet As String, cell As String)

    If parameters("expected") = True Then
         Select Case parameters("testParameter")
            Case "optionButton1"
                ThisWorkbook.Sheets(sheet).Range(cell).Value = "Ja"
                ShowFunc (formName)
                result = CStr(frm005.OptionButton1.Value)
            Case "optionButton2"
                ThisWorkbook.Sheets(sheet).Range(cell).Value = "Nej"
                ShowFunc (formName)
                result = CStr(frm005.OptionButton2.Value)
        End Select
    Else
        Select Case parameters("testParameter")
            Case "optionButton1"
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
                ShowFunc (formName)
                result = CStr(frm005.OptionButton1.Value)
            Case "optionButton2"
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
                ShowFunc (formName)
                result = CStr(frm005.OptionButton2.Value)
        End Select
    End If
End Function

Private Function CheckNoExtraPrints()
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
            popCells = Array()
            rulCells = Array()
            groCells = Array()
            spmCells = Array()
        Case "noChangeWhenBackButton"
            popCells = Array()
            rulCells = Array()
            groCells = Array()
            spmCells = Array()
        Case "config1"
            popCells = Array()
            rulCells = Array("G7", "G8", "G9")
            groCells = Array()
            spmCells = Array("D13")
        Case "config2"
            popCells = Array()
            rulCells = Array("")
            groCells = Array()
            spmCells = Array("D13")
    End Select
    
    'returns a string which shows either true or has the input of the cells that changed that shouldn't have been changed
    result = Global_Test_Func.CheckPrintsInAllSheets(spmCells, popCells, rulCells, groCells)
    
     'Cleans up all arrays and dictionaries
    Erase popCells, rulCells, groCells, spmCells
    Sheet9.spmChangedCells.RemoveAll
    Sheet5.groChangedCells.RemoveAll
    Sheet3.rulChangedCells.RemoveAll
    Sheet1.popChangedCells.RemoveAll
End Function

Private Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm002") Then
        Unload frm002
    End If
    If Global_Test_Func.IsLoaded("frm005") Then
        Unload frm005
    End If
    If Global_Test_Func.IsLoaded("frm006") Then
        Unload frm006
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function




