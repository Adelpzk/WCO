VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addNewdeliverable 
   Caption         =   "Add a New Deliverable"
   ClientHeight    =   5580
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6480
   OleObjectBlob   =   "addNewdeliverable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addNewdeliverable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim boxCourse As Variant
Dim iRow1 As Integer
Dim iRow2 As Integer
Dim iRow3 As Integer
Dim iRow4 As Integer
Dim iRow5 As Integer
Dim iRowM As Integer
Dim iRow11 As Integer
Dim Name As Variant
Dim assignmentName As Variant
Dim d As Date
Dim DueDate As Variant
Dim EstDate As Variant






Private Sub btnConfirm_Click()
Set wB = ThisWorkbook
Set Ws = wB.Worksheets("Classes_Page")
boxCourse = boxCoursetitle.Value
assignmentName = txtName
DueDate = txtDuedate.Value
EstDate = txtEst.Value
d = Date



'Checks so the same Assessment doesn't repeat
If Application.WorksheetFunction.CountIf(Sheets("Classes_Page").Range("C2:E16"), assignmentName) > 0 Then
MsgBox ("The Assessment Already Exists"), vbCritical
Exit Sub
End If

On Error Resume Next
Me.txtDuedate = CDate(Me.txtDuedate)
Me.txtEst = CDate(Me.txtEst)

If Not IsDate(txtDuedate.Text) Then
MsgBox ("please enter a date"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If Not IsDate(txtEst.Text) Then
MsgBox ("please enter a date"), vbInformation
Exit Sub
Unload addNewdeliverable
End If

If DueDate < d Then
MsgBox ("Please Enter a Valid Date"), vbInformation
Exit Sub
End If

If EstDate < d Then
MsgBox ("Please Enter a Valid Date"), vbInformation
Exit Sub
End If

If DueDate < EstDate Then
MsgBox (" Try your Best to Get it Done Before the DueDate"), vbInformation
End If

' Not letting each Course Title Have more than 3 assignments due a week since that's the functionality of the tool.
iRow1 = Range("A1010").Value + 1
If boxCourse = Ws.Cells(1000, "A") Then
If iRow1 > 3 Then
MsgBox ("You're Out of Space , Wait for the Next Week"), vbCritical
Unload addNewdeliverable
Exit Sub
End If
End If
iRow2 = Range("A1011").Value + 1
If boxCourse = Ws.Cells(1001, "A") Then
If iRow2 > 3 Then
MsgBox ("You're Out of Space , Wait for the Next Week"), vbCritical
Unload addNewdeliverable
Exit Sub
End If
End If


iRow3 = Range("A1012").Value + 1
If boxCourse = Ws.Cells(1002, "A") Then
If iRow3 > 3 Then
MsgBox ("You're Out of Space , Wait for the Next Week"), vbCritical
Unload addNewdeliverable
Exit Sub
End If
End If


iRow4 = Range("A1013").Value + 1
If boxCourse = Ws.Cells(1003, "A") Then
If iRow4 > 3 Then
MsgBox ("You're Out of Space , Wait for the Next Week"), vbCritical
Unload addNewdeliverable
Exit Sub
End If
End If

iRow5 = Range("A1014").Value + 1
If boxCourse = Ws.Cells(1004, "A") Then
If iRow5 > 3 Then
MsgBox ("You're Out of Space , Wait for the Next Week"), vbCritical
Unload addNewdeliverable
Exit Sub
End If
End If



' Message boxes to guide the user
If boxCourse = "Choose Your Course Title" Then
MsgBox ("Please Choose your Course Title"), vbInformation
Exit Sub
End If


If txtName = "" Then
MsgBox ("Please add the Task Name"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If txtDuedate = "" Then
MsgBox ("Please add the Due Date"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If txtDes = "" Then
MsgBox ("Please add the Description"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If txtEst = "" Then
MsgBox ("Please add the Estimated Time to Finsih the Assessment"), vbInformation
Exit Sub
Unload addNewdeliverable
End If





'Placing enteries in the right place
If boxCourse = Ws.Cells(1000, "A") Then
Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -3).Value = txtEst

ElseIf boxCourse = Ws.Cells(1001, "A") Then
Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -3).Value = txtEst


ElseIf boxCourse = Ws.Cells(1002, "A") Then
Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -3).Value = txtEst



ElseIf boxCourse = Ws.Cells(1003, "A") Then
Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -3).Value = txtEst


ElseIf boxCourse = Ws.Cells(1004, "A") Then
Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -3).Value = txtEst
End If

iRowM = Sheets("Main Page").Range("A1000").Value + 1
Sheets("Main Page").Range("MainPage").Offset(iRowM, -9).Value = txtName
Sheets("Main Page").Range("MainPage").Offset(iRowM, -3).Value = txtDuedate
Sheets("Main Page").Range("MainPage").Offset(iRowM, -11).Value = boxCoursetitle


If boxCourse <> "Choose Your Course Title" And txtName <> "" And txtDuedate <> "" And txtDes <> "" And txtEst <> "" Then
MsgBox assignmentName & " " & ("Was Added Succussfully")
Unload addNewdeliverable
End If









End Sub

 
Private Sub CommandButton1_Click()

Unload addNewdeliverable

End Sub


