Attribute VB_Name = "Frm031_test"
Private result As String
Private formID As Integer
Private formName As String
Private stopFormTest As Boolean
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
Private spmCells() As Variant
Private popCells() As Variant
Private rulCells() As Variant
Private groCells() As Variant


Sub RunTests()

    formName = "frm031"
    formID = 31
    
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
    

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
        
    Select Case parameters("testSubject")
        Case "printsToSpmSheet"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
            Select Case parameters("testParameter")
                Case "optionButton1"
                    CheckFields "SpmSvar", "D86"
                Case "optionButton2"
                    CheckFields "SpmSvar", "D86"
                Case "textbox1"
                    CheckFields "SpmSvar", "D87"
                Case "textbox2"
                    CheckFields "SpmSvar", "D88"
                Case "checkbox1"
                    CheckFields "SpmSvar", "D87"
                Case "checkbox2"
                    CheckFields "SpmSvar", "D88"
            End Select
        Case "printsToPopSheet"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
            CheckFields "Population", "B17"
            
        Case "printsToRulSheet"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
            If (parameters("testParameter") = "ruleActivation") Then
                Select Case parameters("rule")
                    Case "R0059"
                        CheckFields "Regler", "G60"
                    Case "R0060"
                        CheckFields "Regler", "G61"
                    Case "R0061"
                        CheckFields "Regler", "G62"
                    Case "R0062"
                        CheckFields "Regler", "G63"
                    Case "R0070"
                        CheckFields "Regler", "G71"
                End Select
            Else
                Select Case parameters("rule")
                    Case "R0059"
                        CheckFields "Regler", "J60"
                    Case "R0060"
                        CheckFields "Regler", "J61"
                    Case "R0061"
                        CheckFields "Regler", "J62"
                    Case "R0062"
                        CheckFields "Regler", "J63"
                    Case "R0070"
                        CheckFields "Regler", "J71"
                End Select
            End If
            
        Case "printsToGroSheet"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
            CheckFields "Gruppering", "C2"
            
        Case "errorMessage"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
            
            If (clickOnErrorMessage = True) Then
                frmMsg.CommandButton1_Click
            End If
            
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm031.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            frm031.OKButton_Click 'Click on Videre button
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
        Case "checkCaption"
            SetFields
            frm031.OKButton_Click 'Click on Videre button
        Case Else
            MsgBox "Error in 'testsubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    End Select
    
    'Comparison
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
   
    frm031.OptionButton1.Value = parameters("optionButton1")
    frm031.OptionButton2.Value = parameters("optionButton2")
    frm031.TextBox1.Value = parameters("textbox1")
    frm031.TextBox2.Value = parameters("textbox2")
    frm031.CheckBox1.Value = parameters("checkbox1")
    frm031.CheckBox2.Value = parameters("checkbox2")
    
    If (parameters("checkbox3") = True) Then
        frm031.CheckBox3.Value = True
        frm031.CheckBox3_Click
    End If
    
    Select Case parameters("spm9bSvar")
        Case "Ja"
            frm008.OptionButton1.Value = True
            frm008.OptionButton2.Value = False
        Case "Nej"
            frm008.OptionButton1.Value = False
            frm008.OptionButton2.Value = True
    End Select
    
    Select Case parameters("spm9b2Svar")
        Case "Ja"
            frm009.OptionButton1.Value = True
            frm009.OptionButton2.Value = False
        Case "Nej"
            frm009.OptionButton1.Value = False
            frm009.OptionButton2.Value = True
    End Select
    
    Select Case parameters("spm9b22Svar")
        Case "Antal dage angivet"
            frm010.OptionButton1.Value = True
            frm010.OptionButton2.Value = False
        Case "Ved ikke"
            frm010.OptionButton1.Value = False
            frm010.OptionButton2.Value = True
    End Select
    
End Function
Private Function CheckFields(sheet As String, cell As String)
    'Check results
    result = ThisWorkbook.Sheets(sheet).Range(cell).Text

End Function
Private Function DataIsSaved(sheet As String)
   
   If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "optionButton1"
               ThisWorkbook.Sheets(sheet).Range("D86").Value = "Før det valgte stamdatafelt"
               ShowFunc (formName)
               result = CStr(frm031.OptionButton1.Value)
           Case "optionButton2"
               ThisWorkbook.Sheets(sheet).Range("D86").Value = "Samme dag eller senere end det valgte stamdatafelt"
               result = CStr(frm031.OptionButton2.Value)
           Case "textbox1"
               ThisWorkbook.Sheets(sheet).Range("D87").Value = "10"
               ShowFunc (formName)
               result = CStr(frm031.TextBox1.Value)
            Case "textbox2"
               ThisWorkbook.Sheets(sheet).Range("D88").Value = "10"
               ShowFunc (formName)
               result = CStr(frm031.TextBox2.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "optionButton1"
               ThisWorkbook.Sheets(sheet).Range("D86").Value = ""
               ShowFunc (formName)
               result = CStr(frm031.OptionButton1.Value)
           Case "optionButton2"
               ThisWorkbook.Sheets(sheet).Range("D86").Value = ""
               result = CStr(frm031.OptionButton2.Value)
           Case "textbox1"
               ThisWorkbook.Sheets(sheet).Range("D87").Value = ""
               ShowFunc (formName)
               result = CStr(frm031.TextBox1.Value)
            Case "textbox2"
               ThisWorkbook.Sheets(sheet).Range("D88").Value = ""
               ShowFunc (formName)
               result = CStr(frm031.TextBox2.Value)
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
        Case "config1"
            popCells = Array("B17")
            rulCells = Array("J60", "J61", "J62", "J63", "J71", "G60", "G61", "G62", "G63", "G71")
            groCells = Array("C2")
            spmCells = Array("C86", "C87", "C88", "D86", "D87", "D88")
        Case "config2"
            popCells = Array("B17")
            rulCells = Array("G60", "G61", "G62", "G63", "G71")
            groCells = Array("C2")
            spmCells = Array("C86", "C87", "C88", "D86", "D87", "D88")
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
    If Global_Test_Func.IsLoaded("frm014") Then
        Unload frm014
    End If
    If Global_Test_Func.IsLoaded("frm028") Then
        Unload frm028
    End If
    If Global_Test_Func.IsLoaded("frm009") Then
        Unload frm009
    End If
    If Global_Test_Func.IsLoaded("frm010") Then
        Unload frm010
    End If
    If Global_Test_Func.IsLoaded("frm008") Then
        Unload frm008
    End If
    If Global_Test_Func.IsLoaded("frm029") Then
        Unload frm029
    End If
    If Global_Test_Func.IsLoaded("frm030") Then
        Unload frm030
    End If
    If Global_Test_Func.IsLoaded("frm031") Then
        Unload frm031
    End If
    If Global_Test_Func.IsLoaded("frm032") Then
        Unload frm032
    End If
    If Global_Test_Func.IsLoaded("frm039") Then
        Unload frm039
    End If
    If Global_Test_Func.IsLoaded("frm040") Then
        Unload frm040
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function




