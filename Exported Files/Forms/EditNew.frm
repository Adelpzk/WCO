VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditNew 
   Caption         =   "Edit your Existing Task"
   ClientHeight    =   5652
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6468
   OleObjectBlob   =   "EditNew.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CommandButton1_Click()
Dim boxCourse As Variant
Dim assignmentName As Variant
Dim iRowMain As Integer
Dim d As Date
Dim DueDate As Variant
Dim EstDate As Variant


Set wB = ThisWorkbook
Set Ws = wB.Worksheets("Classes_Page")

DueDate = txtDuedate.Value
EstDate = txtEst.Value
d = Date

'Erro Handling

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
MsgBox ("Please Enter a Valid Date")
Exit Sub
End If

If EstDate < d Then
MsgBox ("Please Enter a Valid Date")
Exit Sub
End If

If DueDate < EstDate Then
MsgBox (" Try your Best to Get it Done Before the DueDate"), vbInformation
End If


boxCourse = boxCoursetitle
assignmentName = txtName

 
' Error Handling
If boxCourse = "" Then
MsgBox ("Please Choose your Course Title"), vbInformation
Exit Sub
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


If txtDes = "" Then
MsgBox ("Please  add the Description"), vbInformation
Exit Sub
Unload addNewdeliverable
End If


If txtEst = "" Then
MsgBox ("Please  add the Estimated Time to Finsih the Assessment"), vbInformation
Exit Sub
Unload addNewdeliverable
End If





If editTask.boxEdit.Value = Sheets("Classes_Page").Range("C2").Value Then
Sheets("Classes_Page").Range("courseTitel1").Offset(1, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitel1").Offset(1, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitel1").Offset(1, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitel1").Offset(1, -3).Value = txtEst

ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C3").Value Then
Sheets("Classes_Page").Range("courseTitel1").Offset(2, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitel1").Offset(2, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitel1").Offset(2, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitel1").Offset(2, -3).Value = txtEst

ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C4").Value Then
Sheets("Classes_Page").Range("courseTitel1").Offset(3, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitel1").Offset(3, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitel1").Offset(3, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitel1").Offset(3, -3).Value = txtEst
End If


If editTask.boxEdit.Value = Sheets("Classes_Page").Range("C5").Value Then
Sheets("Classes_Page").Range("courseTitle2").Offset(1, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle2").Offset(1, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle2").Offset(1, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle2").Offset(1, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C6").Value Then
Sheets("Classes_Page").Range("courseTitle2").Offset(2, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle2").Offset(2, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle2").Offset(2, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle2").Offset(2, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C7").Value Then
Sheets("Classes_Page").Range("courseTitle2").Offset(3, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle2").Offset(3, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle2").Offset(3, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle2").Offset(3, -3).Value = txtEst
End If









If editTask.boxEdit.Value = Sheets("Classes_Page").Range("C8").Value Then
Sheets("Classes_Page").Range("courseTitle3").Offset(1, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle3").Offset(1, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle3").Offset(1, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle3").Offset(1, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C9").Value Then
Sheets("Classes_Page").Range("courseTitle3").Offset(2, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle3").Offset(2, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle3").Offset(2, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle3").Offset(2, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C10").Value Then
Sheets("Classes_Page").Range("courseTitle3").Offset(3, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle3").Offset(3, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle3").Offset(3, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle3").Offset(3, -3).Value = txtEst
End If







If editTask.boxEdit.Value = Sheets("Classes_Page").Range("C11").Value Then
Sheets("Classes_Page").Range("courseTitle4").Offset(1, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle4").Offset(1, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle4").Offset(1, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle4").Offset(1, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C12").Value Then
Sheets("Classes_Page").Range("courseTitle4").Offset(2, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle4").Offset(2, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle4").Offset(2, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle4").Offset(2, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C13").Value Then
Sheets("Classes_Page").Range("courseTitle4").Offset(3, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle4").Offset(3, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle4").Offset(3, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle4").Offset(3, -3).Value = txtEst
End If






If editTask.boxEdit.Value = Sheets("Classes_Page").Range("C14").Value Then
Sheets("Classes_Page").Range("courseTitle5").Offset(1, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle5").Offset(1, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle5").Offset(1, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle5").Offset(1, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C15").Value Then
Sheets("Classes_Page").Range("courseTitle5").Offset(2, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle5").Offset(2, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle5").Offset(2, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle5").Offset(2, -3).Value = txtEst
ElseIf editTask.boxEdit.Value = Sheets("Classes_Page").Range("C16").Value Then
Sheets("Classes_Page").Range("courseTitle5").Offset(3, -15).Value = txtName
Sheets("Classes_Page").Range("courseTitle5").Offset(3, -12).Value = txtDuedate
Sheets("Classes_Page").Range("courseTitle5").Offset(3, -10).Value = txtDes
Sheets("Classes_Page").Range("courseTitle5").Offset(3, -3).Value = txtEst
End If



'Editing on the Main Page as well

iRowMain = Sheets("Main Page").Range("A201").Value
Sheets("Main Page").Range("MainPage").Offset(iRowMain, -9).Value = txtName
Sheets("Main Page").Range("MainPage").Offset(iRowMain, -3).Value = txtDuedate
Sheets("Main Page").Range("MainPage").Offset(iRowMain, -11).Value = boxCoursetitle

If boxCourse <> "" And txtName <> "" And txtDuedate <> "" And txtDes <> "" And txtEst <> "" Then
MsgBox assignmentName & " " & ("Was Added Succussfully")
Unload addNewdeliverable
End If

Unload EditNew
Unload editTask



End Sub
