VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsgYesNo 
   Caption         =   "Spørgsmål?"
   ClientHeight    =   2320
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4728
   OleObjectBlob   =   "frmMsgYesNo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMsgYesNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub cmdNo_Click()

dFunc.msgYesNo = "NEJ"
Me.Hide

End Sub
Public Sub cmdYes_Click()

dFunc.msgYesNo = "JA"
Me.Hide

End Sub
Private Sub UserForm_Initialize()

lblMsg.Caption = dFunc.msgYesNoTxt
dFunc.msgYesNo = ""

End Sub
