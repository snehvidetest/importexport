Attribute VB_Name = "Frm034_test"
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

    formName = "frm034"
    formID = 34
    
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
    
        Case "printsToSpmSheet"
            SetFields
            frm034.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "printsToRulSheet"
            SetFields
            frm034.OKButton_Click 'Click on Videre button
            CheckFields "Regler"
            
        Case "errorMessage"
            SetFields
            frm034.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm034.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm034.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm034.Tilbage_Click 'Click back button
            Else
                frm034.OKButton_Click 'Click on Videre button
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
    frm034.TextBox1.Value = parameters("textbox1")
    frm034.TextBox2.Value = parameters("textbox2")
    frm034.ComboBox2.Value = parameters("combobox2")
    frm034.ComboBox4.Value = parameters("combobox4")
    
End Function

Private Function CheckFields(sheet As String)
    Select Case parameters("testParameter")
        Case "textbox1"
            result = ThisWorkbook.Sheets(sheet).Range("D60").Text
        Case "textbox2"
            result = ThisWorkbook.Sheets(sheet).Range("G60").Text
        Case "combobox2"
            result = ThisWorkbook.Sheets(sheet).Range("F60").Text
        Case "combobox4"
            result = ThisWorkbook.Sheets(sheet).Range("I60").Text
        Case "ruleActivation"
            result = ThisWorkbook.Sheets(sheet).Range("G14").Text
        Case "ruleXDays"
            result = ThisWorkbook.Sheets(sheet).Range("J14").Text
        Case "ruleYDays"
            result = ThisWorkbook.Sheets(sheet).Range("M14").Text
    End Select
    
    
End Function


Function DataIsSaved(sheet As String)
    If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "textbox1"
               ThisWorkbook.Sheets(sheet).Range("D60").Value = "10"
               ShowFunc (formName)
               result = CStr(frm034.TextBox1.Value)
           Case "textbox2"
               ThisWorkbook.Sheets(sheet).Range("G60").Value = "100"
               result = CStr(frm034.TextBox2.Value)
           Case "combobox2"
               ThisWorkbook.Sheets(sheet).Range("F60").Value = "efter"
               ShowFunc (formName)
               result = CStr(frm034.ComboBox2.Value)
            Case "ombobox4"
               ThisWorkbook.Sheets(sheet).Range("I60").Value = "efter"
               ShowFunc (formName)
               result = CStr(frm034.ComboBox4.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "textbox1"
               ThisWorkbook.Sheets(sheet).Range("D60").Value = ""
               ShowFunc (formName)
               result = CStr(frm034.TextBox1.Value)
           Case "textbox2"
               ThisWorkbook.Sheets(sheet).Range("G60").Value = ""
               result = CStr(frm034.TextBox2.Value)
           Case "combobox2"
               ThisWorkbook.Sheets(sheet).Range("F60").Value = ""
               ShowFunc (formName)
               result = CStr(frm034.ComboBox2.Value)
            Case "ombobox4"
               ThisWorkbook.Sheets(sheet).Range("I60").Value = ""
               ShowFunc (formName)
               result = CStr(frm034.ComboBox4.Value)
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
            rulCells = Array("G14", "J14", "M14")
            groCells = Array()
            spmCells = Array("D60", "F60", "G60", "I60", "C60", "C59")
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
Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If IsLoaded("frm034") Then
        Unload frm034
    End If
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm039") Then
        Unload frm039
    End If
    If IsLoaded("frm035") Then
        Unload frm035
    End If
    If IsLoaded("frm044") Then
        Unload frm044
    End If
End Function



