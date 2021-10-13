VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gradeReport 
   Caption         =   "Add a New Report"
   ClientHeight    =   3060
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5688
   OleObjectBlob   =   "gradeReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gradeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRow1 As Integer
Dim iRow2 As Integer
Dim iRow3 As Integer
Dim iRow4 As Integer
Dim iRow5 As Integer
Dim iRow11 As Integer
Dim CourseTitle As Variant
Dim assignmentName As String



Private Sub CommandButton1_Click()
Set wB = ThisWorkbook
Set Ws = wB.Worksheets("Grade Report")
iRow1 = Sheets("Grade Report").Range("A200").Value + 2
iRow2 = Sheets("Grade Report").Range("A201").Value + 2
iRow3 = Sheets("Grade Report").Range("A202").Value + 2
iRow4 = Sheets("Grade Report").Range("A203").Value + 2
iRow5 = Sheets("Grade Report").Range("A204").Value + 2
CourseTitle = tit_Course.Value
assignmentName = txt_Name.Value

'Error Handlings


If CourseTitle = "Choose Your Course Title" Then
    MsgBox ("Please Choose your Course Title"), vbInformation
Exit Sub
End If

If txt_Name.Value = "" Then
    MsgBox ("Please Enter the Name of the Assignment")
Exit Sub
End If

If txt_Grade.Value = "" Then
    MsgBox ("Please Enter the Grade Associated with the Assignment")
Exit Sub
End If

If Not IsNumeric(txt_Grade.Value) Then
    MsgBox ("Only Numbers are Allowed for Grades")
Exit Sub
End If

If CourseTitle = Sheets("Classes_Page").Range("A1000") Then
Sheets("Grade Report").Range("Help").Offset(iRow1, -20).Value = txt_Name
Sheets("Grade Report").Range("Help").Offset(iRow1, -17).Value = txt_Grade
End If


If CourseTitle = Sheets("Classes_Page").Range("A1001") Then
Sheets("Grade Report").Range("Help").Offset(iRow2, -16).Value = txt_Name
Sheets("Grade Report").Range("Help").Offset(iRow2, -13).Value = txt_Grade
End If

If CourseTitle = Sheets("Classes_Page").Range("A1002") Then
Sheets("Grade Report").Range("Help").Offset(iRow3, -12).Value = txt_Name
Sheets("Grade Report").Range("Help").Offset(iRow3, -9).Value = txt_Grade
End If

If CourseTitle = Sheets("Classes_Page").Range("A1003") Then
Sheets("Grade Report").Range("Help").Offset(iRow4, -8).Value = txt_Name
Sheets("Grade Report").Range("Help").Offset(iRow4, -5).Value = txt_Grade
End If

If CourseTitle = Sheets("Classes_Page").Range("A1004") Then
Sheets("Grade Report").Range("Help").Offset(iRow5, -4).Value = txt_Name
Sheets("Grade Report").Range("Help").Offset(iRow5, -1).Value = txt_Grade
End If
Unload gradeReport

If CourseTitle <> "" And txt_Name <> "" And txt_Grade <> "" Then
MsgBox assignmentName & " " & ("Was Added Succussfully")
Unload addNewdeliverable
End If

End Sub




