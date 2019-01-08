Attribute VB_Name = "ValidCheck"
Function FormatCheck(val, expected_format)

    Dim w_month As Variant
    w_month = Mid(val, 4, 2)
    
    If IsNumeric(w_month) Then
        'w_month = Mid(val, 4, 2)
    Else
        res = False
        dFunc.msgError = "Dato ikke udfyldt korrekt.."
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
        'frmMsg.Show
    End If
    
    If expected_format = "date" Then
        res = IsDate(val)
                
        If w_month > 12 Then
            res = False
            dFunc.msgError = CStr(val) + " er ikke en gyldig dato."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
        
        If Mid(val, 3, 1) = "/" Or Mid(val, 6, 1) = "/" Then
            res = False
            dFunc.msgError = "Formattet skal vÊre DD-MM-≈≈≈≈."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
        
        If Mid(val, 3, 1) = " " Or Mid(val, 6, 1) = " " Then
            res = False
            dFunc.msgError = "Formattet skal vÊre DD-MM-≈≈≈≈."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    
        If res = False Then
            dFunc.msgError = CStr(val) + " er ikke en dato. Formattet skal vÊre DD-MM-≈≈≈≈."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If expected_format = "num" Then
        res = IsNumeric(val)
        If res = False Then
            dFunc.msgError = CStr(val) + " er ikke en dato. Formattet skal vÊre DD-MM-≈≈≈≈."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
            'frmMsg.Show
            'MsgBox (CStr(val) + " er ikke en dato. Formattet skal vÊre DD-MM-≈≈≈≈")
        End If
    End If

    If expected_format = "str" Then
        res = True
    End If
    
ending:

    FormatCheck = res

End Function


Sub test()

    C = FormatCheck("21 - 07 - 2018", "date")
    
    MsgBox (C)

End Sub

