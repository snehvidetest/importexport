Attribute VB_Name = "Frm026_test"
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

    formName = "frm026"
    formID = 26
    
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
        
        Case "printsToPopSheet"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            CheckFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
        
            
        Case "nextStep"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))

            
        Case "backButton"
            frm026.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm026.Tilbage_Click 'Click back button
            Else
                frm026.OKButton_Click 'Click on Videre button
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
    frm026.Forfaldsdato.Value = parameters("forfaldsdato")
    frm026.txtFFStart.Value = parameters("forfaldsdatoFrom")
    frm026.txtFFSlut.Value = parameters("forfaldsdatoTo")
    
    frm026.SRB.Value = parameters("srb")
    frm026.txtSRBstart.Value = parameters("srbFrom")
    frm026.txtSRBslut.Value = parameters("srbTo")
    
    frm026.Stiftelsesdato.Value = parameters("stiftelsesdato")
    frm026.txtSTIstart.Value = parameters("stiftelsesdatoFrom")
    frm026.txtSTIslut.Value = parameters("stiftelsesdatoTo")
    
    frm026.PeriodeStartdato.Value = parameters("periodeStart")
    frm026.txtPSTstart.Value = parameters("periodeStartFrom")
    frm026.txtPSTslut.Value = parameters("periodeStartTo")
    
    frm026.PeriodeSlutdato.Value = parameters("periodeSlut")
    frm026.txtPSLstart.Value = parameters("periodeSlutFrom")
    frm026.txtPSLslut.Value = parameters("periodeSlutTo")
    
End Function

Private Function CheckFields(sheet As String)

    Select Case sheet
        Case "SpmSvar"
            Select Case parameters("testParameter")
                Case "forfaldsdato"
                    result = ThisWorkbook.Sheets(sheet).Range("D8").Text
                Case "forfaldsdatoFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("E8").Text
                Case "forfaldsdatoTo"
                    result = ThisWorkbook.Sheets(sheet).Range("F8").Text
                Case "srb"
                    result = ThisWorkbook.Sheets(sheet).Range("D9").Text
                Case "srbFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("E9").Text
                Case "srbTo"
                    result = ThisWorkbook.Sheets(sheet).Range("F9").Text
                Case "stiftelsesdato"
                    result = ThisWorkbook.Sheets(sheet).Range("D10").Text
                Case "stiftelsesdatoFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("E10").Text
                Case "stiftelsesdatoTo"
                    result = ThisWorkbook.Sheets(sheet).Range("F10").Text
                Case "periodeStart"
                    result = ThisWorkbook.Sheets(sheet).Range("D11").Text
                Case "periodeStartFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("E11").Text
                Case "periodeStartTo"
                    result = ThisWorkbook.Sheets(sheet).Range("F11").Text
                Case "periodeSlut"
                    result = ThisWorkbook.Sheets(sheet).Range("D12").Text
                Case "periodeSlutFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("E12").Text
                Case "periodeSlutTo"
                    result = ThisWorkbook.Sheets(sheet).Range("F12").Text
            End Select
            
        Case "Population"
            Select Case parameters("testParameter")
                Case "forfaldsdatoFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B6").Text
                Case "forfaldsdatoTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B7").Text
                Case "srbFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B8").Text
                Case "srbTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B9").Text
                Case "stiftelsesdatoFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B10").Text
                Case "stiftelsesdatoTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B11").Text
                Case "periodeStartFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B12").Text
                Case "periodeStartTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B13").Text
                Case "periodeSlutFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B14").Text
                Case "periodeSlutTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B15").Text
            End Select
            
        End Select
        
End Function


Function DataIsSaved(sheet As String)

    If parameters("forfaldsdato") = True Then
        ThisWorkbook.Sheets(sheet).Range("D8").Value = "Forfaldsdato"
        ThisWorkbook.Sheets(sheet).Range("E8").Value = parameters("forfaldsdatoFrom")
        ThisWorkbook.Sheets(sheet).Range("F8").Value = parameters("forfaldsdatoTo")
    End If
        
    If parameters("srb") = True Then
        ThisWorkbook.Sheets(sheet).Range("D9").Value = "SRB Dato"
        ThisWorkbook.Sheets(sheet).Range("E9").Value = parameters("srbFrom")
        ThisWorkbook.Sheets(sheet).Range("F9").Value = parameters("srbTo")
    End If
    
    If parameters("stiftelsesdato") = True Then
        ThisWorkbook.Sheets(sheet).Range("D10").Value = "Stiftelsesdato"
        ThisWorkbook.Sheets(sheet).Range("E10").Value = parameters("stiftelsesdatoFrom")
        ThisWorkbook.Sheets(sheet).Range("F10").Value = parameters("stiftelsesdatoTo")
    End If
    
    If parameters("periodeStart") = True Then
        ThisWorkbook.Sheets(sheet).Range("D11").Value = "PeriodeStartdato"
        ThisWorkbook.Sheets(sheet).Range("E11").Value = parameters("periodeStartFrom")
        ThisWorkbook.Sheets(sheet).Range("F11").Value = parameters("periodeStartTo")
    End If
    
    If parameters("periodeSlut") = True Then
        ThisWorkbook.Sheets(sheet).Range("D12").Value = "PeriodeSlutdato"
        ThisWorkbook.Sheets(sheet).Range("E12").Value = parameters("periodeSlutFrom")
        ThisWorkbook.Sheets(sheet).Range("F12").Value = parameters("periodeSlutTo")
    End If
            
    ShowFunc (formName)
    
    Select Case parameters("testParameter")
        Case "forfaldsdato"
            result = CStr(frm026.Forfaldsdato.Value)
        Case "forfaldsdatoFrom"
            result = CStr(frm026.txtFFStart.Value)
        Case "forfaldsdatoTo"
            result = CStr(frm026.txtFFSlut.Value)
        Case "srb"
            result = CStr(frm026.SRB.Value)
        Case "srbFrom"
            result = CStr(frm026.txtSRBstart.Value)
        Case "srbTo"
            result = CStr(frm026.txtSRBslut.Value)
        Case "stiftelsesdato"
            result = CStr(frm026.Stiftelsesdato.Value)
        Case "stiftelsesdatoFrom"
            result = CStr(frm026.txtSTIstart.Value)
        Case "stiftelsesdatoTo"
            result = CStr(frm026.txtSTIslut.Value)
        Case "periodeStart"
            result = CStr(frm026.PeriodeStartdato.Value)
        Case "periodeStartFrom"
            result = CStr(frm026.txtPSTstart.Value)
        Case "periodeStartTo"
            result = CStr(frm026.txtPSTslut.Value)
        Case "periodeSlut"
            result = CStr(frm026.PeriodeSlutdato.Value)
        Case "periodeSlutFrom"
            result = CStr(frm026.txtPSLstart.Value)
        Case "periodeSlutTo"
            result = CStr(frm026.txtPSLslut.Value)
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
            popCells = Array("B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15")
            rulCells = Array()
            groCells = Array()
            spmCells = Array("C7", "D8", "D9", "D10", "D11", "D12", "E8", "E9", "E10", "E11", "E12", "F8", "F9", "F10", "F11", "F12")
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
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm026") Then
        Unload frm026
    End If
    If IsLoaded("frm003") Then
        Unload frm003
    End If
    If IsLoaded("frm005") Then
        Unload frm005
    End If
End Function






