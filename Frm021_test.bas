Attribute VB_Name = "Frm021_test"
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

    formName = "frm021"
    formID = 21
    
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
            frm021.OKButton_Click 'Click on Videre button
            Select Case parameters("testParameter")
                Case "textbox1"
                    CheckFields "SpmSvar", "D55"
                Case "checkbox1"
                    CheckFields "SpmSvar", "D55"
            End Select
            
        Case "printsToRulSheet"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            If (parameters("testParameter") = "ruleActivation") Then
                Select Case parameters("rule")
                    Case "R0072"
                        CheckFields "Regler", "G73"
                    Case "R0073"
                        CheckFields "Regler", "G74"
                    Case "R0074"
                        CheckFields "Regler", "G76"
                    Case "R0103"
                        CheckFields "Regler", "G75"
                 End Select
            ElseIf (parameters("testParameter") = "amount") Then
                Select Case parameters("rule")
                    Case "R0072"
                        CheckFields "Regler", "H73"
                    Case "R0073"
                        CheckFields "Regler", "H74"
                 End Select
            End If
            
        Case "printsToGroSheet"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            Select Case parameters("group")
                Case "G0005"
                    CheckFields "Gruppering", "C6"
                Case "G0006"
                    CheckFields "Gruppering", "C7"
            End Select
            
        Case "errorMessage"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            If (parameters("testParameter") = "frm037") Then
                frm039.CheckBox4.Value = True
            Else
                frm039.CheckBox4.Value = False
            End If
            frm021.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            frm021.OKButton_Click 'Click on Videre button
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
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
   
    frm021.TextBox1.Value = parameters("textbox1")
    frm021.CheckBox1.Value = parameters("checkbox1")
    
End Function
Private Function CheckFields(sheet As String, cell As String)
    'Check results
    result = ThisWorkbook.Sheets(sheet).Range(cell).Text

End Function
Private Function DataIsSaved(sheet As String)
   
   If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "textbox1"
               ThisWorkbook.Sheets(sheet).Range("D55").Value = "10"
               ShowFunc (formName)
               result = CStr(frm021.TextBox1.Value)
           Case "checkbox1"
               ThisWorkbook.Sheets(sheet).Range("D55").Value = "Ved ikke"
               result = CStr(frm021.CheckBox1.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "textbox1"
               ThisWorkbook.Sheets(sheet).Range("D55").Value = ""
               ShowFunc (formName)
               result = CStr(frm021.TextBox1.Value)
           Case "checkbox1"
               ThisWorkbook.Sheets(sheet).Range("D55").Value = ""
               result = CStr(frm021.CheckBox1.Value)
               Debug.Print (result)
               Debug.Print ("hej")
               Debug.Print (result)
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
            popCells = Array()
            rulCells = Array("G73", "G74", "G75", "G76", "H73", "H74")
            groCells = Array("C6", "C7")
            spmCells = Array("D55", "C55")
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
    If Global_Test_Func.IsLoaded("frm021") Then
        Unload frm021
    End If
    If Global_Test_Func.IsLoaded("frm022") Then
        Unload frm022
    End If
    If Global_Test_Func.IsLoaded("frm037") Then
        Unload frm037
    End If
    If Global_Test_Func.IsLoaded("frm038") Then
        Unload frm038
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function

