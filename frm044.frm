VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm044 
   Caption         =   "Advarsel"
   ClientHeight    =   4060
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "frm044.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm044"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()
Me.Hide
End Sub

Public Sub CommandButton2_Click()
Me.Hide
frm034.Hide
SFunc.ShowFunc ("frm035")
'frm035.Show

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
