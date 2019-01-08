Attribute VB_Name = "Frm013_test"
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

    formName = "frm013"
    formID = 13
    
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
    resetSheets ThisWorkbook
    
    'Create the TCID
    tcid = GetTCID(tc, formID)
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
    
        Case "printsToRulSheet"
            SetFields
            frm013.OKButton_Click 'Click on Videre button
            CheckFrmFields "Regler"
            
        Case "printsToPopSheet"
            SetFields
            frm013.OKButton_Click 'Click on Videre button
            CheckFrmFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm013.OKButton_Click 'Click on Videre button
            CheckFrmFields "SpmSvar"
                    
        Case "errorMessage"
            SetFields
            frm013.OKButton_Click 'Click on Videre button
            result = errorMessage
            
        Case "nextStep"
            SetFields
            frm013.OKButton_Click 'Click on Videre button
            result = NextStep(parameters("expected"))
            
        Case "backButton"
            frm013.Tilbage_Click
            result = NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar", "D23"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm013.Tilbage_Click 'Click back button
            Else
                frm013.OKButton_Click 'Click on Videre button
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

    Call KillForms
    
     'Print results
    PrintTestResults tcid, result, review

End Function


Private Function SetFields()
    
    ThisWorkbook.Sheets("SpmSvar").Range("D24:H24").Value = "" 'Prevents crashing when frm010 initialises frm014
    
    'The folowing code inserts the inputs into the actual form
    frm013.OptionButton1.Value = parameters("optionButton1")
    frm013.TextBox1.Value = parameters("antalDage")
    frm013.OptionButton2.Value = parameters("optionButton2")
      
End Function

Private Function CheckFrmFields(sheet As String)
    
    'Check results
    If (sheet = "SpmSvar") Then
        result = ThisWorkbook.Sheets(sheet).Range("D23").Text
        
    ElseIf (sheet = "Population") Then
        Select Case parameters("testParameter")
            Case "trustRIM"
                result = ThisWorkbook.Sheets(sheet).Range("B16").Text
            Case "rimFOKO"
                result = ThisWorkbook.Sheets(sheet).Range("B17").Text
        End Select
        
    ElseIf (sheet = "Gruppering") Then
        Select Case parameters("group")
            Case "G0001"
                result = ThisWorkbook.Sheets(sheet).Range("C2").Text
            Case "G0002"
                result = ThisWorkbook.Sheets(sheet).Range("C3").Text
        End Select
    ElseIf (sheet = "Regler") And parameters("testParameter") = "ruleActivation" Then
        Select Case parameters("rule")
            Case "R0042"
                result = ThisWorkbook.Sheets(sheet).Range("G43").Text
            Case "R0043"
                result = ThisWorkbook.Sheets(sheet).Range("G44").Text
            Case "R0044"
                result = ThisWorkbook.Sheets(sheet).Range("G45").Text
            Case "R0045"
                result = ThisWorkbook.Sheets(sheet).Range("G46").Text
            Case "R0046"
                result = ThisWorkbook.Sheets(sheet).Range("G47").Text
        End Select
    ElseIf (sheet = "Regler") And parameters("testParameter") = "ruleDurXDays" Then
        Select Case parameters("rule")
            Case "R0042"
                result = ThisWorkbook.Sheets(sheet).Range("J43").Text
            Case "R0043"
                result = ThisWorkbook.Sheets(sheet).Range("J44").Text
            Case "R0044"
                result = ThisWorkbook.Sheets(sheet).Range("J45").Text
            Case "R0045"
                result = ThisWorkbook.Sheets(sheet).Range("J46").Text
            Case "R0046"
                result = ThisWorkbook.Sheets(sheet).Range("J47").Text
        End Select
    End If
End Function


Private Function DataIsSaved(sheet As String, cell As String)
    
    
    Select Case parameters("testParameter")
        Case "optionButton1"
            
            If parameters("optionButton1") = "True" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = parameters("antalDage")
            ElseIf parameters("optionButton1") = "False" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm013.OptionButton2.Value)
        
        Case "antalDage"
            
            If parameters("optionButton1") = "True" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = parameters("antalDage")
            ElseIf parameters("optionButton1") = "False" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm013.TextBox1.Value)
            
        Case "optionButton2"
            If parameters("optionButton2") = "True" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = "Ved ikke"
            ElseIf parameters("optionButton2") = "False" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm013.OptionButton2.Value)
            
    End Select
            
            
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
            popCells = Array("B16")
            rulCells = Array("J43:J47")
            groCells = Array()
            spmCells = Array("D23", "C23")
        Case "config2"
            popCells = Array("B16")
            rulCells = Array("G43:G47", "J43:J47")
            groCells = Array()
            spmCells = Array("D23", "C23")
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
    If IsLoaded("frm008") Then
        Unload frm008
    End If
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm013") Then
        Unload frm013
    End If
    If IsLoaded("frm012") Then
        Unload frm012
    End If
    If IsLoaded("frm014") Then
        Unload frm014
    End If
    If IsLoaded("frm039") Then
        Unload frm039
    End If
End Function










