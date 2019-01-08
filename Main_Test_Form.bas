Attribute VB_Name = "Main_Test_Form"
Public workbookName As String
Public worksheetName As String
Public logging As Boolean
Public progressFile As String

Public testWB As Workbook
Public testWS As Worksheet

Sub Main_test_forms()

'~~~~~~~~~~~~~~~~I've added this comment to demonstrate Githubs version control!'~~~~~~~~~~~~~~~~

'----------------Now we try with the importexport add-in!--------------


'    Application.ScreenUpdating = False
    Dim wb As Workbook, wks As Worksheet
    
'~~~~~~~~~~~~~~~~Also here!~~~~~~~~~~~~~~~~
    
    workbookName = "FF_Spgskema_TC_main.xlsm"               'Name of test workbook
    worksheetName = "FF - Test Design"                      'Name of sheet with testcases
    Set testWB = Workbooks(workbookName)
    Set testWS = testWB.Worksheets(worksheetName)
    
    logging = False                                      'Logging of testcase process
    If logging Then
        progressFile = Application.ActiveWorkbook.Path & "\progress.csv"
        Open progressFile For Output As #1
    End If
    
    Set Sheet9.spmChangedCells = New Scripting.Dictionary   'Dictionaries to log changes in SpmSvar sheet
    Set Sheet1.popChangedCells = New Scripting.Dictionary   'Dictionaries to log changes in Population sheet
    Set Sheet5.groChangedCells = New Scripting.Dictionary   'Dictionaries to log changes in Gruppering sheet
    Set Sheet3.rulChangedCells = New Scripting.Dictionary   'Dictionaries to log changes in Regler sheet
    
    
    'Clear cells in testcase results sheet
    Global_Test_Func.ClearTestResults
    
    
    
    'Start testing!
    TestMode = True
'    Frm001_test.RunTests
'    Debug.Print ("1")
    Frm002_test.RunTests
    Debug.Print ("2")
'    Frm003_test.RunTests
'    Debug.Print ("3")
'    Frm004_test.RunTests
'    Debug.Print ("4")
'    Frm005_test.RunTests
'    Debug.Print ("5")
'    Frm006_test.RunTests
'    Debug.Print ("6")
'    Frm007_test.RunTests
'    Debug.Print ("7")
'    Frm008_test.RunTests
'    Debug.Print ("8")
'    Frm009_test.RunTests
'    Debug.Print ("9")
'    Frm010_test.RunTests
'    Debug.Print ("10")
'    Frm011_test.RunTests
'    Debug.Print ("11")
'    Frm012_test.RunTests
'    Debug.Print ("12")
'    Frm013_test.RunTests
'    Debug.Print ("13")
'    Frm014_test.RunTests
'    Debug.Print ("14")
'    Frm021_test.RunTests
'    Debug.Print ("21")
'    Frm026_test.RunTests
'    Debug.Print ("26")
'    Frm028_test.RunTests
'    Debug.Print ("28")
'    Frm029_test.RunTests
'    Debug.Print ("29")
'    Frm030_test.RunTests
'    Debug.Print ("30")
'    Frm031_test.RunTests
'    Debug.Print ("31")
'    Frm032_test.RunTests
'    Debug.Print ("32")
'    Frm033_test.RunTests
'    Debug.Print ("33")
'    Frm034_test.RunTests
'    Debug.Print ("34")
'    Frm035_test.RunTests
'    Debug.Print ("35")
'    Frm036_test.RunTests
'    Debug.Print ("36")
'    Frm037_test.RunTests
'    Debug.Print ("37")
'    Frm038_test.RunTests
'    Debug.Print ("38")
'    Frm039_test.RunTests
'    Debug.Print ("39")
'    Frm043_test.RunTests
'    Debug.Print ("43")
'    Frm044_test.RunTests
'    Debug.Print ("44")
'    Frm045_test.RunTests
'    Debug.Print ("45")
'    Frm046_test.RunTests
'    Debug.Print ("46")
'    Frm047_test.RunTests
'    Debug.Print ("47")
    TestMode = False
End Sub








'*********** Use the following to export modules from a project**********'´
Sub ExportModules()
    
    Dim wks As Worksheet, i As Integer, row As Integer, exists As Boolean, wbExport As Workbook, tempFolderPath As String, FSO As Object
    
    '****************EDIT THESE VARIABLES****************'
    tempFolderPath = "C:\FF_sporegeskema_test_temp"         '<-- The folder to export files to
    Set wbExport = Workbooks("FlexFilter_master_R3_v2.xlsm")   '<-- The workbook to export from
    Set testWB = Workbooks("FF_Spgskema_TC_main.xlsm")      '<-- The testworkbook where the names of exported files are logged in a sheet called "Copy".
    '****************************************************'
    
    'Deletes any previous copies and create export folder if it doesn't exist
    Set FSO = CreateObject("scripting.filesystemobject")
    If FSO.FolderExists(tempFolderPath) = True Then
        On Error Resume Next
        Kill tempFolderPath & "\*.*"    'Delete files
    Else
        MkDir tempFolderPath            'Create folder
    End If
    
    'Check if copy sheet in the test workbook exists, add it if it doesn't
    testWB.Activate
    For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Copy" Then
        exists = True
    End If
    Next i
    If Not exists Then
        Worksheets.Add.Name = "Copy"
    End If
    
    Set wks = testWB.Worksheets("Copy")
    wks.Range("A:A").Value = ""
    
    'Export all files from
    row = 0
    With wbExport.VBProject.VBComponents
        For i = 1 To .Count
        If .Item(i).Type = 1 Then
            wks.Range("A1").Offset(row, 0).Value = .Item(i).Name            'Save name of file
            .Item(i).Export tempFolderPath & "\" & .Item(i).Name & ".bas"   'Save file in folder
            row = row + 1
        End If
        Next i
    End With
    
    MsgBox ("Please open the Copy-sheet in the testworkbook and delete modules you don't want to import to the new project")
    
End Sub

'*********** Use the following to import modules to a project (Use ExportModules() first!)**********'´
Sub ImportModules()
    Dim x, i, wbFrom As Workbook, wbTo As Workbook, wksFrom As Worksheet, tempFolderPath As String
    
    '****************EDIT THESE VARIABLES****************'
    tempFolderPath = "C:\FF_sporegeskema_test_temp"         '<-- The folder to export files to
    Set wbFrom = Workbooks("FlexFilter_master_R3.xlsm")     '<-- The workbook to export from
    Set wbTo = Workbooks("FlexFilter_master_R3_v2.xlsm")     '<-- The workbook to import to
    Set testWB = Workbooks("FF_Spgskema_TC_main.xlsm")      '<-- The testworkbook where the names of exported files are logged in a sheet called "Copy".
    '****************************************************'
    
    Set wksFrom = testWB.Worksheets("Copy")
    With wbTo.VBProject.VBComponents
        For i = 0 To WorksheetFunction.CountA(wksFrom.Range("A:A")) - 1
            .Import tempFolderPath & "\" & wksFrom.Range("A1").Offset(i, 0).Text & ".bas"
        Next i
    End With
End Sub


Function InsertProc()
 Dim mdl As Module, strText As String
 
 'On Error GoTo Error_InsertProc
 ' Open module.
 ThisWorkbook.Activate
 'DoCmd.OpenModule ("Module1")
 ' Return reference to Module object.
 Set mdl = Modules("Module1")
 ' Initialize string variable.
 strText = "Sub DisplayMessage()" & vbCrLf _
 & vbTab & "MsgBox ""Wild!""" & vbCrLf _
 & "End Sub"
 ' Insert text into module.
 mdl.InsertText strText
 InsertProc = True
 
Exit_InsertProc:
 Exit Function
 
Error_InsertProc:
 MsgBox Err & ": " & Err.Description
 InsertProc = False
 Resume Exit_InsertProc
End Function
