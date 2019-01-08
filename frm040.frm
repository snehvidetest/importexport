VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm040 
   Caption         =   "Advarsel"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "frm040.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()
Me.Hide


If frm028.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    If frm028.ActiveControl = True Then
        frm028.Hide
        GoTo ending
    End If
End If


If frm029.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    If frm029.ActiveControl = True Then
        frm029.Hide
        GoTo ending
    End If
End If


If frm030.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    If frm030.ActiveControl = True Then
        frm030.Hide
        GoTo ending
    End If
End If


If frm031.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    If frm031.ActiveControl = True Then
        frm031.Hide
        GoTo ending
    End If
End If


If frm032.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    If frm032.ActiveControl = True Then
        frm032.Hide
        GoTo ending
    End If
End If

ending:
SFunc.ShowFunc ("frm002")
'frm002.Show

End Sub

Public Sub CommandButton2_Click()
Me.Hide

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
