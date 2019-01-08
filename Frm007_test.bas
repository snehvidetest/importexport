Attribute VB_Name = "Frm007_test"
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

    formName = "frm007"
    formID = 7
    
    Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)
    
    Dim nrTC As Integer, i As Integer
    nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)
    
'    For i = 1 To nrTC
'        Set parameters = New Scripting.Dictionary
'        Testcase i
'    Next i
    
    Testcase 37
    
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
    ClearAllFields ThisWorkbook

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
        
    Select Case parameters("testSubject")
    
        Case "printsToSpmSheet"
            SetFields
            frm007.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar", "D17"
            
        Case "printsToRulSheet"
            SetFields
            frm007.OKButton_Click 'Click on Videre button
            Select Case parameters("rule")
                Case "R0039"
                        CheckFields "Regler", "G40"
                Case "R0042"
                    Select Case parameters("testParameter")
                        Case "ruleActivation"
                            CheckFields "Regler", "G43"
                        Case "ruleDurXDays"
                            CheckFields "Regler", "J43"
                        Case "ruleDurXYears"
                            CheckFields "Regler", "K43"
                        Case "ruleDurXMonths"
                            CheckFields "Regler", "L43"
                    End Select
                Case "R0043"
                    Select Case parameters("testParameter")
                        Case "ruleActivation"
                            CheckFields "Regler", "G44"
                        Case "ruleDurXDays"
                            CheckFields "Regler", "J44"
                        Case "ruleDurXYears"
                            CheckFields "Regler", "K44"
                        Case "ruleDurXMonths"
                            CheckFields "Regler", "L44"
                    End Select
                Case "R0044"
                    Select Case parameters("testParameter")
                        Case "ruleActivation"
                            CheckFields "Regler", "G45"
                        Case "ruleDurXDays"
                            CheckFields "Regler", "J45"
                        Case "ruleDurXYears"
                            CheckFields "Regler", "K45"
                        Case "ruleDurXMonths" '
                            CheckFields "Regler", "L45"
                    End Select
                Case "R0045"
                    Select Case parameters("testParameter")
                        Case "ruleActivation"
                            CheckFields "Regler", "G46"
                        Case "ruleDurXDays"
                            CheckFields "Regler", "J46"
                        Case "ruleDurXYears"
                            CheckFields "Regler", "K46"
                        Case "ruleDurXMonths"
                            CheckFields "Regler", "L46"
                    End Select
                Case "R0046"
                    Select Case parameters("testParameter")
                        Case "ruleActivation"
                            CheckFields "Regler", "G47"
                        Case "ruleDurXDays"
                            CheckFields "Regler", "J47"
                        Case "ruleDurXYears"
                            CheckFields "Regler", "K47"
                        Case "ruleDurXMonths"
                            CheckFields "Regler", "L47"
                    End Select
                    
                End Select
                
        Case "printsToPopSheet"
            SetFields
            frm007.OKButton_Click 'Click on Videre button
            CheckFields "Population", "B16"
            
        Case "errorMessage"
            SetFields
            frm007.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm007.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm007.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar", "D17"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm007.Tilbage_Click 'Click back button
            Else
                frm007.OKButton_Click 'Click on Videre button
            End If
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

    Call KillForms

    'Print results
    Global_Test_Func.PrintTestResults tcid, result, review
    
    
End Function
Private Function SetFields()
   'The folowing code inserts the inputs into the actual form
   
    frm007.OptionButton1.Value = parameters("optionButton1")
    frm007.OptionButton2.Value = parameters("optionButton2")
    frm007.OptionButton3.Value = parameters("optionButton3")

End Function

Private Function CheckFields(sheet As String, cell As String)
    'Check results
    result = ThisWorkbook.Sheets(sheet).Range(cell).Text
End Function

Private Function DataIsSaved(sheet As String, cell As String)
   
   If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "optionButton1"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = "Altid"
               ShowFunc (formName)
               result = CStr(frm007.OptionButton1.Value)
           Case "optionButton2"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = "I visse tilfælde"
               ShowFunc (formName)
               result = CStr(frm007.OptionButton2.Value)
           Case "optionButton3"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = "Aldrig"
               ShowFunc (formName)
               result = CStr(frm007.OptionButton3.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "optionButton1"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
               ShowFunc (formName)
               result = CStr(frm007.OptionButton1.Value)
           Case "optionButton2"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
               ShowFunc (formName)
               result = CStr(frm007.OptionButton2.Value)
           Case "optionButton3"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
               ShowFunc (formName)
               result = CStr(frm007.OptionButton3.Value)
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
        Case "config2"
            popCells = Array("B16")
            rulCells = Array("G43:G47", "J43:J47")
            groCells = Array()
            spmCells = Array("D17", "C17")
        Case "config1"
            popCells = Array()
            rulCells = Array()
            groCells = Array()
            spmCells = Array("D17", "C17")
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
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm007") Then
        Unload frm007
    End If
    If Global_Test_Func.IsLoaded("frm008") Then
        Unload frm008
    End If
    If Global_Test_Func.IsLoaded("frm006") Then
        Unload frm006
    End If
    If Global_Test_Func.IsLoaded("frm011") Then
        Unload frm011
    End If
     If Global_Test_Func.IsLoaded("frm014") Then
        Unload frm014
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function


