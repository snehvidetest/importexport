VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public popChangedCells As Scripting.Dictionary
Public recordChangingCells As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    If (recordChangingCells = True) Then
        If (popChangedCells.exists(Target.Address(0, 0)) = False) Then
            popChangedCells.Add Target.Address(0, 0), Target
        End If
    End If
End Sub

