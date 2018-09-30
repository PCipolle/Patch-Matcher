VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "Pocket Type"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTop_Click()
Dim flag As Integer

flag = 1

frmSelect.Hide

    patchMatcher flag
    
End Sub

Private Sub cmdBottom_Click()
Dim flag As Integer
flag = 2

frmSelect.Hide

    patchMatcher flag

End Sub

Private Sub cmdWood_Click()

Dim flag As Integer
flag = 3

frmSelect.Hide

    patchMatcher flag
    
End Sub

Private Sub cmdBronze_Click()
Dim flag As Integer

'MsgBox ("This feature is not yet available")
'Exit Sub

flag = 4

frmSelect.Hide

    patchMatcher flag
    
End Sub




