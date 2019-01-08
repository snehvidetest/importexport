Attribute VB_Name = "Func"
Function check_numeric(x, msg)
    If IsNumeric(x) = False Then
         MsgBox (msg)
    End If
End Function

Function Insert_to_sheet(sheet0, range0, value0)
    Worksheets(sheet0).Range(range0).Value = value0
    If value0 = "JA" Then
        Worksheets(sheet0).Range(range0).Interior.Color = RGB(198, 239, 206) 'Background color
        Worksheets(sheet0).Range(range0).Font.Color = RGB(0, 97, 0)          'Text color
    ElseIf value0 = "NEJ" Then
        Worksheets(sheet0).Range(range0).Interior.Color = RGB(255, 199, 206) 'Background color
        Worksheets(sheet0).Range(range0).Font.Color = RGB(156, 0, 6)         'Text color
    End If


End Function


Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer

    retval = ""
                        
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
                       '
    onlyDigits = retval
End Function

Function maxValueMonths(x As Variant) As String
    maxValueMonths = ""
    If x > 12 Then
    maxValueMonths = "Indsæt en gyldig måned i året"
    End If
End Function

Function maxValueDays(x As Variant) As String
    maxValueDays = ""
    If x > 31 Then
    maxValueDays = "Indsæt en gyldig dag i måneden"
    End If
End Function
Function check_day_month(x As String, msg As String, check As String) As Boolean
    
    Dim a As Long
    check_day_month = False
    
    If x = "" Then
        a = 0
    End If
    
    If x <> "" Then
        If IsNumeric(x) = False Then
            check_day_month = True
            dFunc.msgError = msg + " (1 og 2)"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox (msg + " (1 og 2)")
            GoTo Tilbage
        End If
      
        If check = "1" Then
            a = CLng(x)
            If (a <= 0 Or a > 31) Then
                check_day_month = True
                dFunc.msgError = msg + " (1)"
                SFunc.ShowFunc ("frmMsg")
                'MsgBox (msg + " (1)")
                GoTo Tilbage
            End If
        End If
    
    
        If check = "2" Then
            a = CLng(x)
            If (a <= 0 Or a > 12) Then
                check_day_month = True
                dFunc.msgError = msg + " (2)"
                SFunc.ShowFunc ("frmMsg")
                'MsgBox (msg + " (2)")
            End If
        End If
    End If
    
Tilbage:

End Function

Function check_month(a1 As String, msg) As Boolean

    Dim a As Long
    
    check_month = False
            
    a = CLng(a1)
       
    If (a <= 0 Or a > 12) Then
        check_month = True
        dFunc.msgError = msg
        SFunc.ShowFunc ("frmMsg")
        'MsgBox (msg)
        GoTo ending
    End If
            
ending:

End Function

Function test(x) As Integer
    
    test = x + 2
    
End Function

Sub test2()
a = onlyDigits("hej123he4j")
MsgBox (a)
End Sub
