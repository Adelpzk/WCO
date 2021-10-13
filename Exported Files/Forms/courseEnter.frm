VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} courseEnter 
   Caption         =   "Enter Your Courses"
   ClientHeight    =   5556
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10320
   OleObjectBlob   =   "courseEnter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "courseEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub btnConfirm_Click()
Set wB = ThisWorkbook
Set Ws = wB.Worksheets("Classes_Page")
'Placing the Enteries in the right places in our workbook
If (txtCourse1.Value <> "") Then
Sheets("Classes_Page").Range("A2:B4") = txtCourse1.Value
Sheets("Grade Report").Range("A1:D1") = txtCourse1.Value
Ws.Cells(1000, "A") = txtCourse1.Value
End If
If (txtCourse2.Value <> "") Then
Sheets("Classes_Page").Range("A5:B7") = txtCourse2.Value
Sheets("Grade Report").Range("E1:H1") = txtCourse2.Value
Ws.Cells(1001, "A") = txtCourse2.Value
End If
If (txtCourse3.Value <> "") Then
Sheets("Classes_Page").Range("A8:B10") = txtCourse3.Value
Sheets("Grade Report").Range("I1:L1") = txtCourse3.Value
Ws.Cells(1002, "A") = txtCourse3.Value
End If
If (txtCourse4.Value <> "") Then
Sheets("Classes_Page").Range("A11:B13") = txtCourse4.Value
Sheets("Grade Report").Range("M1:P1") = txtCourse4.Value
Ws.Cells(1003, "A") = txtCourse4.Value
End If
If (txtCourse5.Value <> "") Then
Sheets("Classes_Page").Range("A12:B16") = txtCourse5.Value
Sheets("Grade Report").Range("Q1:T1") = txtCourse5.Value
Ws.Cells(1004, "A") = txtCourse5.Value
End If

Unload courseEnter
End Sub


Private Sub CommandButton1_Click()
Unload courseEnter
End Sub

