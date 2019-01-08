Attribute VB_Name = "Global_Test_Func"

'ClearAllFields in This workbook
Function ClearAllFields(wb As Workbook)
    wb.Sheets("SpmSvar").Range("D2:H150").Value = ""
    wb.Sheets("Population").Range("B2", "B200").Value = ""
End Function

'Clear result is Test sheet
Function ClearTestResults()
    For i = 3 To 5000
        If IsNumeric(testWS.Cells(i, 1).Value) Then
            testWS.Cells(i, 10).Value = ""
            testWS.Cells(i, 11).Value = ""
        End If
    Next i
End Function

'Reset sheets
Function resetSheets(wb As Workbook)

'wb.Sheets("SpmSvar").Range("D2:I111") = "Test"
'wb.Sheets("Population").Range("B2:B18") = "Test"
'wb.Sheets("Regler").Range("G2:R105") = "Test"
'wb.Sheets("Gruppering").Range("C2:C18") = "Test"
wb.Sheets("SpmSvar").Range("D2:I111") = ""
wb.Sheets("Population").Range("B2:B18") = ""
wb.Sheets("Regler").Range("G2:R105") = ""
wb.Sheets("Gruppering").Range("C2:C18") = ""

End Function


'Gets data for the specific testcase and puts it in a dictionary which i returned. Key: Parameter Val: Data related to that parameter
Function getData(tcid As String, parametersAndCols As Scripting.Dictionary) As Scripting.Dictionary

    Dim ws As Worksheet, test As String, tcRow As Long, parameters As New Scripting.Dictionary
    testWS.Activate
    
    'Get the row for the specific testcase
    tcRow = GetDataRow(tcid)

    'Set retrive the testcase parameter using the dictionary containing alle parameter names and their respective columns
    For Each parameter In parametersAndCols.Keys
        parameters.Add parameter, testWS.Cells(tcRow, parametersAndCols(parameter)).Text
    Next parameter
    
    'Replace any "Empty"/"blank" values with "" instead
    For Each varItem In parameters
        If parameters(varItem) = "blank" Or parameters(varItem) = "empty" Then
            parameters(varItem) = ""
        End If
    Next varItem
    
    Set getData = parameters
    
End Function

'Gets the parameters specific to the form an their respective columvalue and put it in a dictionary which it returns
Function getParamtersAndTheirCols(formID As Integer) As Scripting.Dictionary
    Dim parameterRow As Long
    Dim parametersAndCols As New Scripting.Dictionary

    'Retrive the row of the parameters
    parameterRow = GetDataRow(formID & ".01") - 1
    
    'Retrive the parameter names and store their columns
    parametersAndCols.Add testWS.Cells(parameterRow, 6).Text, 6 'TestSubject
    parametersAndCols.Add testWS.Cells(parameterRow, 7).Text, 7 'TestParameter'
    parametersAndCols.Add testWS.Cells(parameterRow, 9).Text, 9 'Expected'
    For i = 13 To 100   'Other parameters
        'start at column "W"/23
        If testWS.Cells(parameterRow, i).Text = "run" Then
            parametersAndCols.Add testWS.Cells(parameterRow, i).Text, i
            Exit For
        Else
            parametersAndCols.Add testWS.Cells(parameterRow, i).Text, i 'Store parameter and location
        End If
    Next
    
    Set getParamtersAndTheirCols = parametersAndCols
    
End Function


Function PrintTestResults(tcid As String, result As String, review As Boolean)
    'The following code prints the results to the testcase result sheet by matching TCID
    
    Dim wks As Worksheet
    Set wks = Workbooks(Main_Test_Form.workbookName).Worksheets(Main_Test_Form.worksheetName)
    wks.Activate
    
    Dim rng As Range, tc_row As Long
    
    'If expected and/or result are empty, then write "Empty" in the testcase results instead
   
    If result = "" Then
        result = "Empty"
    End If
    Set rng = wks.Range("C1", "C5000")
    tc_row = rng.Find(tcid, LookIn:=xlValues).row
    rng.Cells(tc_row, 8).Value = "'" & result
    rng.Cells(tc_row, 9).Value = "'" & review

End Function

Function GetTCID(tc As Integer, formID As Integer) As String
    'The following code creates the TCID fra the testcase number
    If tc < 10 Then
        GetTCID = formID & ".0" & tc
    Else:
        GetTCID = formID & "." & tc
    End If
End Function
   
   
Function GetDataRow(tcid As String) As Long
'The following code returns the Row linked to the input testcase id
    Dim rng As Range
    
    Set rng = testWS.Range("C1", "C5000")
    GetDataRow = rng.Find(tcid, LookIn:=xlValues).row
    
End Function

Function GetDataCol(row As Long, parameter As String) As Long
'The following code returns the colum linked to the input testcase id and the choosen parameter
    Dim rng As Range
    
    With testWS
        Set rng = .Range(.Cells(row, "A"), .Cells(row, "ZZ"))
        GetDataCol = rng.Find(parameter, LookIn:=xlValues).Column
    End With
    
End Function
   
Public Function IsLoaded(formName As String) As Boolean
'The following code returns true or false weateher the form is loaded or not
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.Name = formName Then
        IsLoaded = True
        Exit Function
    End If
Next frm
IsLoaded = False
End Function

Function NextStep(formName As String) As String
'The following code returns the string if it has been loaded otherwise it returns "incorrect"
    If IsLoaded(formName) Then
        NextStep = formName
    Else
        NextStep = "Incorrect"
    End If

End Function

Function errorMessage() As String
'The following code returns the errormessage if it has been loaded
    If IsLoaded("frmMsg") Then
        errorMessage = frmMsg.lblMsg.Caption
    Else
        errorMessage = "Messege did not pop up"
    End If
End Function


Function CleanUp(wbFrom As Workbook, wbTo As Workbook)

    Dim i, str As String, wksFrom As Worksheet, vbCom As Object
    
    Set wksFrom = wbFrom.Worksheets("Copy")
    
    wbTo.Activate
    Set vbCom = Application.VBE.ActiveVBProject.VBComponents
    
        For i = 0 To WorksheetFunction.CountA(wksFrom.Range("A:A")) - 1
            str = wksFrom.Range("B1").Offset(i, 0).Text
            vbCom.Remove VBComponent:= _
            vbCom.Item(str)
        Next i
    
End Function

Function CheckPrintsInAllSheets(spmCells() As Variant, popCells() As Variant, rulCells() As Variant, groCells() As Variant) As String
    
    resultSpm = CheckPrintsInSpecificSheet(Sheet9.spmChangedCells, spmCells())
    resultPop = CheckPrintsInSpecificSheet(Sheet1.popChangedCells, popCells())
    resultRul = CheckPrintsInSpecificSheet(Sheet3.rulChangedCells, rulCells())
    resultGro = CheckPrintsInSpecificSheet(Sheet5.groChangedCells, groCells())
    
    If (resultSpm = True And resultPop = True And resultRul = True And resultGro = True) Then
         CheckPrintsInAllSheets = True
    Else
        CheckPrintsInAllSheets = "SpmSvar: " + resultSpm & vbLf & "Population: " + resultPop & vbLf & "Regler: " + resultRul & vbLf & "Gruppering: " + resultGro
    End If
    
End Function

Function CheckPrintsInSpecificSheet(actChangedCells As Scripting.Dictionary, expChangedCells() As Variant) As String
    
    For Each element In expChangedCells()
        For Each varItem In actChangedCells
            If (element = varItem) Then
                actChangedCells.Remove varItem
            End If
        Next varItem
    Next element

    If (actChangedCells.Count = 0) Then
        CheckPrintsInSpecificSheet = True
    Else
        temp = actChangedCells.Keys()
        CheckPrintsInSpecificSheet = Join(temp, ", ")
    End If
End Function



