VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm001 
   Caption         =   "Population"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11004
   OleObjectBlob   =   "frm001.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()

    dFunc.msgYesNoTxt = "Er du sikker? Dette vil slette den tidligere besvarelse, hvis en sådan eksisterer."
    SFunc.ShowFunc ("frmMsgYesNo")
        
    If dFunc.msgYesNo = "NEJ" Then
       'bliv på siden
    Else
       'start forfra
        Worksheets("SpmSvar").Range("D2:H150").Value = ""
        frm002.lblFtypeTxt.Caption = ""
        frm002.lblFhaverTxt.Caption = ""
        frm002.UserForm_Initialize
        Me.Hide
        dFunc.msgYesNoTxt = ""
        SFunc.ShowFunc ("frm002")
        'frm002.Show
    End If
    'Call YesNoMessageBox
    
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
End Sub
Public Sub OKButton_Click()
    Me.Hide
    ShowFunc ("frm002")
    'frm002.Show
End Sub

Sub YesNoMessageBox()
 
Dim Answer As String
Dim MyNote As String
 
    'Place your text here
    MyNote = "Er du sikker? Dette vil slette den tidligere besvarelse, hvis en sådan eksisterer."
 
    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbOKCancel, "Ny Besvarelse")
 
    If Answer = vbOK Then
        Worksheets("SpmSvar").Range("D2:H150").Value = ""
        frm002.UserForm_Initialize
        Me.Hide
        frm002.Show
    End If
 
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Image1.PictureSizeMode = fmPictureSizeModeStretch
    Worksheets("SpmSvar").Activate
End Sub
