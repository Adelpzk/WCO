VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompleteTaskEdit 
   Caption         =   "Completed Task"
   ClientHeight    =   2256
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5868
   OleObjectBlob   =   "CompleteTaskEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompleteTaskEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRow1 As Integer
Dim boxCourse As Variant
Dim assignmentName As Variant
Dim iRow2 As Integer
Dim d As Date
Dim DueDate As Variant





Private Sub CommandButton1_Click()
assignmentName = txtName
boxCourse = boxCoursetitle.Value
DueDate = txtDuedate.Value
d = Date
'Error Handlings

If Application.WorksheetFunction.CountIf(Sheets("Completed Tasks").Range("C2:G30"), assignmentName) > 0 Then
MsgBox ("The Assessment has Already Marked as Complete"), vbCritical
Unload CompleteTaskEdit
Exit Sub
End If
 
 On Error Resume Next
Me.txtDuedate = CDate(Me.txtDuedate)
If Not IsDate(txtDuedate.Text) Then
MsgBox ("please enter a date"), vbInformation
Exit Sub
Unload CompleteTaskEdit
Unload CompletedTask
End If

If DueDate < d Then
MsgBox ("Please Enter a Valid Date"), vbInformation
Exit Sub
End If

If boxCoursetitle = "" Then
MsgBox ("Please Enter the Course Title"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If txtDuedate = "" Then
MsgBox ("Please  add the Due Date"), vbInformation
Exit Sub
Unload addNewdeliverable
End If

If txtName = "" Then
MsgBox ("Please  add the Task Name"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If txtDuedate = "" Then
MsgBox ("Please  add the Due Date"), vbInformation
Exit Sub
Unload addNewdeliverable
End If

'Putting the inputs in the right cells

iRow1 = Sheets("Completed Tasks").Range("A1000").Value + 1
Sheets("Completed Tasks").Range("CompleteTask").Offset(iRow1, -7).Value = txtName
Sheets("Classes_Page").Range("Condition").Offset(iRow1, 0) = txtName
Sheets("Main Page").Range("ConditionM").Offset(iRow1, 0) = txtName
Sheets("Completed Tasks").Range("CompleteTask").Offset(iRow1, -2).Value = txtDuedate
Sheets("Completed Tasks").Range("CompleteTask").Offset(iRow1, -9).Value = boxCoursetitle


If boxCourse <> "" And txtName <> "" And txtDuedate <> "" Then
MsgBox assignmentName & " " & ("Was Marked as Completed Succussfully")
Unload addNewdeliverable
End If



Unload CompleteTaskEdit
Unload CompletedTask


End Sub


Private Sub Label2_Click()

End Sub
