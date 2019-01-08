Attribute VB_Name = "Frm004_test"
'****Test script for frm004****
Private formID As Integer
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
Private spmCells() As Variant
Private popCells() As Variant
Private rulCells() As Variant
Private groCells() As Variant
Private result As String


Sub RunTests()

'****GUIDE****:

'The main testcase template for form 4 is called:
'frm004_testcase ( tc )
'
'Input: tc
'   An integer which identifies the test case number. This is not the tcid!
'
'The function read data from the testcases excel workbook, hereunder
'
'   - testSubject: This string defines what the testcase is testing. Possible values are:
'       - printsToPopSheet: Checks the form input is written correctly to the population sheet
'       - printsToSpmSheet: Checks the form input is written correctly to the SpmSvar sheet
'       - errorMessage: Check error message is correct
'       - tidligereBesvarelse: Checks that a form can correctly load a previous response to that form
'       - nextStep: Checks that next form(s) is(are) called correctly
'       - backButton: Checks the back button functions
'
'   - testParameter (where relevant): If the testcase relates to a certain parameter, this variable identifies it. Possible values:
'       - "modtagelseStart"
'       - "modtagelseSlut"
'
'   - expected: The relevant value we expect to find
'
'   - The relevant test paramters. For this form they are:
'       - modtagelseStart (String): The modtagelses start date
'       - modtagelseSlut (String): The modtagelses end date
'

'Which form are we testing?
formID = 4


'Get parameters relevant for testcase including their respective columns
Set parametersAndCols = New Scripting.Dictionary
Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)

'Get the total number of testcases associated with the form
Dim nrTC As Integer, i As Integer
nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)

'Run all testcases incl. printing of results to the testcase workbook
For i = 1 To nrTC
    Testcase i
Next i


End Sub


'The following code is the skeleton for form 2 testcases.
Private Function Testcase(tc As Integer)

    Dim review As Boolean, tcid As String
    ThisWorkbook.Activate
    
    'Reset spørgeskema workbook
    Global_Test_Func.resetSheets ThisWorkbook
    
    'Create the TCID
    tcid = Global_Test_Func.GetTCID(tc, formID)
    If logging Then
        Write #1, tcid
    End If
    
    'Get testcase data
    Set parameters = New Scripting.Dictionary 'Resets the testcase parameters
    Set parameters = Global_Test_Func.getData(tcid, parametersAndCols)
    
    'Check if testcase should be run
    If parameters("run") = 0 Then
        Exit Function
    End If
    
    Select Case parameters("testSubject")
        Case "printsToPopSheet"
        
            'Enter data into form
            SetFields
            'frm004.Show 'Check it is correct (but it will clear fields too afterwards)

            'Execute/Click button
            frm004.OKButton_Click
    
            Select Case parameters("testParameter")
                Case "modtagelseStart"
                    result = ThisWorkbook.Sheets("Population").Range("B4").Text
                Case "modtagelseSlut"
                    result = ThisWorkbook.Sheets("Population").Range("B5").Text
                Case Else
                    MsgBox "Error in 'testParameter' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
            End Select
    
    Case "printsToSpmSheet"
        
        'Enter data into form
        SetFields
        'frm004.Show 'Check it is correct (but it will clear fields too afterwards)
        
        'Execute/Click button
        frm004.OKButton_Click
        
        Select Case parameters("testParameter")
            Case "modtagelseStart"
                result = ThisWorkbook.Sheets("SpmSvar").Range("D4").Text
            Case "modtagelseSlut"
                result = ThisWorkbook.Sheets("SpmSvar").Range("E4").Text
            Case Else
                MsgBox "Error in 'testParameter' input: tcid " & tcid
        End Select

    
    Case "errorMessage"
        
            'Enter data into form
            SetFields
            'frm004.Show 'Check it is correct
            
            'Execute/Click button
            frm004.OKButton_Click
            
            If parameters("testParameter") <> "before01092013" Then
                'Get the error message
                result = Global_Test_Func.errorMessage()
            Else
                If IsLoaded("frm043") Then
                    result = frm043.Label1.Caption
                Else
                    result = "Message did not pop up"
                End If
            End If
            
    Case "nextStep"
        
            'Enter data into form
            SetFields
            'frm004.Show 'Check it is correct
            
            'Execute/Click button
            frm004.OKButton_Click
            
            'Check if the expected form opened
            If Global_Test_Func.IsLoaded(parameters("expected")) Then
                result = parameters("expected")
            Else
                result = "Incorrect"
            End If
            
        Case "backButton"
        
            'Enter data into form
            SetFields
            'frm004.Show 'Check it is correct
            
            'Execute/Click button
            frm004.Tilbage_Click
            
            'Check if the expected form opened
            If Global_Test_Func.IsLoaded(parameters("expected")) Then
                result = parameters("expected")
            Else
                result = "Incorrect"
            End If
            
        Case "tidligereBesvarelse"
            
            'Pre-populate SpmSvar sheet
            prePopulateFields
            
            'Initialise form
            ShowFunc ("frm004")
        
            Select Case parameters("testParameter")
                Case "modtagelseStart"
                    result = frm004.TextBox1.Value
                Case "modtagelseSlut"
                    result = frm004.TextBox2.Value
                Case Else
                    MsgBox "Error in 'testsubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
            End Select
            
        Case "noExtraPrints"
            
            SetFields
            Sheet1.recordChangingCells = True
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm004.Tilbage_Click 'Click back button
            Else
                frm004.OKButton_Click 'Click on Videre button
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
            
        Case Else
        
            MsgBox "Error in 'testSubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    
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
    
    'Clear relevant fields
    ThisWorkbook.Activate
    'ThisWorkbook.Sheets("SpmSvar").Range("D2", "E200").Value = ""
    'ThisWorkbook.Sheets("Population").Range("B2", "B200").Value = ""
    
    ShowFunc ("frm004")
    
    'Set values in form
    frm004.TextBox1.Value = parameters("modtagelseStart")
    frm004.TextBox2.Value = parameters("modtagelseSlut")
    
End Function

Private Function prePopulateFields()
    'The folowing code inserts the inputs spmSvar sheet
    
    Dim ws As Worksheet
    
    'Clear relevant fields
    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Sheets("SpmSvar")
    'ws.Range("D2", "E200").Value = ""
    
    'Set values in SpmSvar sheet
    ws.Range("D4").Value = "'" & parameters("modtagelseStart")
    ws.Range("E4").Value = "'" & parameters("modtagelseSlut")
    
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
            popCells = Array("B4", "B5")
            rulCells = Array()
            groCells = Array()
            spmCells = Array("D4", "E4")
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
    If Global_Test_Func.IsLoaded("frm003") Then
        Unload frm003
    End If
    If Global_Test_Func.IsLoaded("frm005") Then
        Unload frm005
    End If
    If Global_Test_Func.IsLoaded("frm004") Then
        Unload frm004
    End If
    If Global_Test_Func.IsLoaded("frm043") Then
        Unload frm043
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If Global_Test_Func.IsLoaded("frm002") Then
        Unload frm002
    End If
End Function




