VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public rulChangedCells As Scripting.Dictionary

Private Sub Worksheet_Change(ByVal Target As Range)
    If (Sheet1.recordChangingCells = True) Then
        If (rulChangedCells.exists(Target.Address(0, 0)) = False) Then
            rulChangedCells.Add Target.Address(0, 0), Target
        End If
    End If
End Sub
