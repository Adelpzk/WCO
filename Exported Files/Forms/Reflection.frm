VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Reflection 
   Caption         =   "Reflection on the Past Week"
   ClientHeight    =   4896
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5808
   OleObjectBlob   =   "Reflection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Reflection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()

End Sub
Private Sub btnConfirm_Click()
Set wB = ThisWorkbook
Set Ws = wB.Worksheets("Main Page")


'Very simple
If (txtReflection.Value <> "") Then
Range("A2:D17") = txtReflection.Value
End If


Unload Reflection

End Sub




Private Sub UserForm_Activate()
txtReflection = Sheets("Main Page").Range("A2").Value
End Sub

