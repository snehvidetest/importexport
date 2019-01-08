Attribute VB_Name = "Frm014_test"
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

    formName = "frm014"
    formID = 14
    
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
            frm014.OKButton_Click 'Click on Videre button
            CheckFields "Regler"
        
        Case "printsToGroSheet"
            SetFields
            frm014.OKButton_Click 'Click on Videre button
            CheckFields "Gruppering"
            
        Case "printsToPopSheet"
            SetFields
            frm014.OKButton_Click 'Click on Videre button
            CheckFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm014.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm014.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "checkCaption"
            SetFields
            frm014.OKButton_Click 'Click on Videre button
            Select Case parameters("testParameter")
                Case "ingen"
                    If IsLoaded("frmMsg") Then
                        result = dFunc.msgError
                    End If
            End Select
            
        Case "nextStep"
            Select Case parameters("testParameter")
                Case ""
                    SetFields
                    frm014.OKButton_Click 'Click on Videre button
                    result = Global_Test_Func.NextStep(parameters("expected"))
                Case "nextForm"
                    SetFields
                    frm014.OKButton_Click 'Click on Videre button
                    If IsLoaded("frmMsg") Then
                        frmMsg.CommandButton1_Click
                        result = Global_Test_Func.NextStep(parameters("expected"))
                    Else
                        result = "MessageBox didn't show so it wasn't possible to complete test"
                    End If
                Case "message"
                    SetFields
                    frm014.OKButton_Click 'Click on Videre button
                    result = Global_Test_Func.NextStep(parameters("expected"))
            End Select
            
        Case "backButton"
            
            frm014.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm014.Tilbage_Click 'Click back button
            Else
                frm014.OKButton_Click 'Click on Videre button
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
    
    ThisWorkbook.Sheets("SpmSvar").Range("D24:H24").Value = "" 'Prevents crashing when frm010 initialises frm014
    
    'The folowing code inserts the inputs into the actual form
    frm014.Forfaldsdato.Value = parameters("forfaldsdato")
    frm014.SRB.Value = parameters("srb")
    frm014.Stiftelsesdato.Value = parameters("stiftelsesdato")
    frm014.PeriodeStartdato.Value = parameters("periodeStartDato")
    frm014.PeriodeSlutdato.Value = parameters("periodeSlutDato")
    frm014.CheckBox2.Value = parameters("ingen")
    
    'Insert necessary previous question answers
    Select Case parameters("spm9Svar")
        Case "Altid"
            frm007.OptionButton1.Value = True
            frm007.OptionButton2.Value = False
            frm007.OptionButton3.Value = False
        Case "I visse tilfælde"
            frm007.OptionButton1.Value = False
            frm007.OptionButton2.Value = True
            frm007.OptionButton3.Value = False
        Case "Aldrig"
            frm007.OptionButton1.Value = False
            frm007.OptionButton2.Value = False
            frm007.OptionButton3.Value = True
    End Select
    
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

Private Function CheckFields(sheet As String)

    Select Case sheet
        Case "SpmSvar"
            Select Case parameters("testParameter")
                Case "forfaldsdato"
                    result = ThisWorkbook.Sheets(sheet).Range("D24").Text
                Case "srb"
                    result = ThisWorkbook.Sheets(sheet).Range("E24").Text
                Case "stiftelsesdato"
                    result = ThisWorkbook.Sheets(sheet).Range("F24").Text
                Case "periodeStartDato"
                    result = ThisWorkbook.Sheets(sheet).Range("G24").Text
                Case "periodeSlutDato"
                    result = ThisWorkbook.Sheets(sheet).Range("H24").Text
                Case "ingen"
                    result = ThisWorkbook.Sheets(sheet).Range("I24").Text
            End Select
            
        Case "Gruppering"
            Select Case parameters("group")
                Case "G0001"
                    result = ThisWorkbook.Sheets(sheet).Range("C2").Text
                Case "G0002"
                    result = ThisWorkbook.Sheets(sheet).Range("C3").Text
            End Select
            
        Case "Population"
            Select Case parameters("testParameter")
                Case "trustRIM"
                    result = ThisWorkbook.Sheets(sheet).Range("B16").Text
                Case "rimFOKO"
                    result = ThisWorkbook.Sheets(sheet).Range("B17").Text
            End Select
        
        Case "Regler"
            Select Case parameters("rule")
                Case "R0047"
                    result = ThisWorkbook.Sheets(sheet).Range("G48").Text
                Case "R0048"
                    result = ThisWorkbook.Sheets(sheet).Range("G49").Text
                Case "R0049"
                    result = ThisWorkbook.Sheets(sheet).Range("G50").Text
                Case "R0050"
                    result = ThisWorkbook.Sheets(sheet).Range("G51").Text
                Case "R0051"
                    result = ThisWorkbook.Sheets(sheet).Range("G52").Text
                Case "R0052"
                    result = ThisWorkbook.Sheets(sheet).Range("G53").Text
                Case "R0053"
                    result = ThisWorkbook.Sheets(sheet).Range("G54").Text
                Case "R0054"
                    result = ThisWorkbook.Sheets(sheet).Range("G55").Text
                Case "R0055"
                    result = ThisWorkbook.Sheets(sheet).Range("G56").Text
                Case "R0056"
                    result = ThisWorkbook.Sheets(sheet).Range("G57").Text
                Case "R0057"
                    result = ThisWorkbook.Sheets(sheet).Range("G58").Text
                Case "R0058"
                    result = ThisWorkbook.Sheets(sheet).Range("G59").Text
                Case "R0059"
                    result = ThisWorkbook.Sheets(sheet).Range("G60").Text
                Case "R0060"
                    result = ThisWorkbook.Sheets(sheet).Range("G61").Text
                Case "R0061"
                    result = ThisWorkbook.Sheets(sheet).Range("G62").Text
                Case "R0062"
                    result = ThisWorkbook.Sheets(sheet).Range("G63").Text
                Case "R0063"
                    result = ThisWorkbook.Sheets(sheet).Range("G64").Text
                Case "R0064"
                    result = ThisWorkbook.Sheets(sheet).Range("G65").Text
                Case "R0065"
                    result = ThisWorkbook.Sheets(sheet).Range("G66").Text
                Case "R0066"
                    result = ThisWorkbook.Sheets(sheet).Range("G67").Text
                Case "R0067"
                    result = ThisWorkbook.Sheets(sheet).Range("G68").Text
                Case "R0068"
                    result = ThisWorkbook.Sheets(sheet).Range("G69").Text
                Case "R0069"
                    result = ThisWorkbook.Sheets(sheet).Range("G70").Text
                Case "R0070"
                    result = ThisWorkbook.Sheets(sheet).Range("G71").Text
                Case "R0071"
                    result = ThisWorkbook.Sheets(sheet).Range("G72").Text
            End Select
    End Select

End Function


Function DataIsSaved(sheet As String)

    If parameters("forfaldsdato") <> "" Then
        ThisWorkbook.Sheets(sheet).Range("D24").Value = "Forfaldsdato " & parameters("forfaldsdato")
        ThisWorkbook.Sheets(sheet).Range("E24").Value = "SRB " & parameters("srb")
        ThisWorkbook.Sheets(sheet).Range("F24").Value = "Stiftelsesdato " & parameters("stiftelsesdato")
        ThisWorkbook.Sheets(sheet).Range("G24").Value = "PeriodeStart " & parameters("periodeStartDato")
        ThisWorkbook.Sheets(sheet).Range("H24").Value = "PeriodeSlut " & parameters("periodeSlutDato")
        ThisWorkbook.Sheets(sheet).Range("I24").Value = "Ingen " & parameters("ingen")
    Else
        ThisWorkbook.Sheets(sheet).Range("D24").Value = ""
        ThisWorkbook.Sheets(sheet).Range("E24").Value = ""
        ThisWorkbook.Sheets(sheet).Range("F24").Value = ""
        ThisWorkbook.Sheets(sheet).Range("G24").Value = ""
        ThisWorkbook.Sheets(sheet).Range("H24").Value = ""
        ThisWorkbook.Sheets(sheet).Range("I24").Value = ""
    End If
    ShowFunc (formName)
    
    Select Case parameters("testParameter")
            Case "forfaldsdato"
                result = CStr(frm014.Forfaldsdato.Value)
            Case "srb"
                result = CStr(frm014.SRB.Value)
            Case "stiftelsesdato"
                result = CStr(frm014.Stiftelsesdato.Value)
            Case "periodeStartDato"
                result = CStr(frm014.PeriodeStartdato.Value)
            Case "periodeSlutDato"
                result = CStr(frm014.PeriodeSlutdato.Value)
            Case "ingen"
                result = CStr(frm014.CheckBox2.Value)
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
            popCells = Array("B17")
            rulCells = Array("G48", "G49", "G50", "G51", "G52", "G53", "G54", "G55", "G56", "G57", "G58", "G59", "G60", "G61", "G62", "G63", "G64", "G65", "G66", "G67", "G68", "G69", "G70", "G71", "G72")
            groCells = Array()
            spmCells = Array("D24", "E24", "F24", "G24", "H24", "I24", "C24")
        Case "config2"
            popCells = Array("B17")
            rulCells = Array("G48", "G49", "G50", "G51", "G52", "G53", "G54", "G55", "G56", "G57", "G58", "G59", "G60", "G61", "G62", "G63", "G64", "G65", "G66", "G67", "G68", "G69", "G70", "G71", "G72")
            groCells = Array("C2")
            spmCells = Array("D24", "E24", "F24", "G24", "H24", "I24")
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
    If IsLoaded("frm002") Then
        Unload frm002
    End If
    If IsLoaded("frm007") Then
        Unload frm007
    End If
    If IsLoaded("frm014") Then
        Unload frm014
    End If
    If IsLoaded("frm028") Then
        Unload frm028
    End If
    If IsLoaded("frm029") Then
        Unload frm029
    End If
    If IsLoaded("frm030") Then
        Unload frm030
    End If
    If IsLoaded("frm031") Then
        Unload frm031
    End If
    If IsLoaded("frm032") Then
        Unload frm032
    End If
    If IsLoaded("frm039") Then
        Unload frm039
    End If
End Function




