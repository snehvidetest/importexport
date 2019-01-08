Attribute VB_Name = "Frm008_test"
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

    formName = "frm008"
    formID = 8
    
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
    
        Case "printsToGroSheet"
            SetFields
            frm008.OKButton_Click 'Click on Videre button
            CheckFields "Gruppering"
            
        Case "printsToPopSheet"
            SetFields
            frm008.OKButton_Click 'Click on Videre button
            CheckFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm008.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm008.OKButton_Click 'Click on Videre button
            result = errorMessage
            
        Case "nextStep"
            SetFields
            frm008.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm008.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar", "D18"
        
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm008.Tilbage_Click 'Click back button
            Else
                frm008.OKButton_Click 'Click on Videre button
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
    frm008.OptionButton1.Value = parameters("optionButton1")
    frm008.OptionButton2.Value = parameters("optionButton2")
  
End Function

Private Function CheckFields(sheet As String)
    'Check results
    If (sheet = "SpmSvar") Then
        result = ThisWorkbook.Sheets(sheet).Range("D18").Text
        
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
    End If
End Function


Private Function DataIsSaved(sheet As String, cell As String)
    
    Select Case parameters("testParameter")
        Case "optionButton1"
            
            If parameters("optionButton1") = "True" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = "Ja"
            ElseIf parameters("optionButton1") = "False" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm008.OptionButton1.Value)
            
        Case "optionButton2"
            If parameters("optionButton2") = "True" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = "Nej"
            ElseIf parameters("optionButton2") = "False" Then
                ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm008.OptionButton2.Value)
            
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
        Case "config2"
            popCells = Array()
            rulCells = Array()
            groCells = Array()
            spmCells = Array("D18", "C18")
        Case "config1"
            popCells = Array("B16", "B17")
            rulCells = Array()
            groCells = Array("C2", "C3")
            spmCells = Array("D18", "C18")
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
    If Global_Test_Func.IsLoaded("frm008") Then
        Unload frm008
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If Global_Test_Func.IsLoaded("frm007") Then
        Unload frm007
    End If
    If Global_Test_Func.IsLoaded("frm009") Then
        Unload frm009
    End If
    If Global_Test_Func.IsLoaded("frm039") Then
        Unload frm039
    End If
End Function






