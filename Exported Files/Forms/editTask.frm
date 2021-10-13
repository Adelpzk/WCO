VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} editTask 
   Caption         =   "Edit Task"
   ClientHeight    =   1944
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "editTask.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "editTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub boxEdit_Change()

End Sub

Sub Confirm1_Click()

Dim Editbox As Variant
Dim Range11 As Integer
Dim Range12 As Integer
Dim Range13 As Integer
Dim Range14 As Integer
Dim Range15 As Integer
Dim iRow1 As Integer
Dim iRow2 As Integer
Dim iRow3 As Integer
Dim iRow4 As Integer
Dim iRow5 As Integer
Dim TargetRow As Integer

Editbox = boxEdit.Value

If Editbox = "" Then
MsgBox ("Please Choose the Task You'd Like to Edit"), vbInformation
Exit Sub
End If

' Putting in the previous values from the cell in the textboxes for the user.

Range11 = Sheets("Classes_Page").Range("A1010").Value
Range12 = Sheets("Classes_Page").Range("A1011").Value
Range13 = Sheets("Classes_Page").Range("A1012").Value
Range14 = Sheets("Classes_Page").Range("A1013").Value
Range15 = Sheets("Classes_Page").Range("A1014").Value

If Range11 <> 0 Then
If Editbox = Sheets("Classes_Page").Range("C2").Value Then
iRow1 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range1"), 0)
EditNew.txtName = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1000").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C3").Value Then
iRow1 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range1"), 0)
Sheets("Classes_page").Range("A802").Value = iRow1
EditNew.txtName = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1000").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C4").Value Then
iRow1 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range1"), 0)
Sheets("Classes_page").Range("A803").Value = iRow1
EditNew.txtName = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitel1").Offset(iRow1, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1000").Value

End If
End If


If Range12 <> 0 Then
If Editbox = Sheets("Classes_Page").Range("C5").Value Then
iRow2 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range2"), 0)
Sheets("Classes_page").Range("A804").Value = iRow2
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1001").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C6").Value Then
iRow2 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range2"), 0)
Sheets("Classes_page").Range("A805").Value = iRow2
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1001").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C7").Value Then
iRow2 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range2"), 0)
Sheets("Classes_page").Range("A806").Value = iRow2
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle2").Offset(iRow2, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1001").Value

End If
End If


If Range13 <> 0 Then
If Editbox = Sheets("Classes_Page").Range("C8").Value Then
iRow3 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range3"), 0)
Sheets("Classes_page").Range("A807").Value = iRow3
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1002").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C9").Value Then
iRow3 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range3"), 0)
Sheets("Classes_page").Range("A808").Value = iRow3
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1002").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C10").Value Then
iRow3 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range3"), 0)
Sheets("Classes_page").Range("A809").Value = iRow3
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle3").Offset(iRow3, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1002").Value

End If
End If


If Range14 <> 0 Then
If Editbox = Sheets("Classes_Page").Range("C11").Value Then
iRow4 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range4"), 0)
Sheets("Classes_page").Range("A810").Value = iRow4
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1003").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C12").Value Then
iRow4 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range4"), 0)
Sheets("Classes_page").Range("A811").Value = iRow4
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1003").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C13").Value Then
iRow4 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range4"), 0)
Sheets("Classes_page").Range("A812").Value = iRow4
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle4").Offset(iRow4, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1003").Value

End If
End If


If Range15 <> 0 Then
If Editbox = Sheets("Classes_Page").Range("C14").Value Then
iRow5 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range5"), 0)
Sheets("Classes_page").Range("A813").Value = iRow5
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1004").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C15").Value Then
iRow5 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range5"), 0)
Sheets("Classes_page").Range("A814").Value = iRow5
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1004").Value

ElseIf Editbox = Sheets("Classes_Page").Range("C16").Value Then
iRow5 = Application.WorksheetFunction.Match(boxEdit, Sheets("Classes_Page").Range("Range5"), 0)
Sheets("Classes_page").Range("A815").Value = iRow5
EditNew.txtName = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -15).Value
EditNew.txtDuedate = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -12).Value
EditNew.txtDes = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -10).Value
EditNew.txtEst = Sheets("Classes_Page").Range("courseTitle5").Offset(iRow5, -3).Value
EditNew.boxCoursetitle = Sheets("Classes_Page").Range("A1004").Value

End If
End If


TargetRow = Application.WorksheetFunction.Match(boxEdit, Sheets("Main Page").Range("Dyn_Name"), 0)
Sheets("Main Page").Range("A201").Value = TargetRow


 'The Addnewdeliverable UF pops up

EditNew.Show




End Sub




Sub UserForm_Activate()
Dim Cell As Variant
'Showing all the assignment with no space between them in the ComboBox
   For Each Cell In Range("C2:C16")
        If Cell.Value <> "" Then
            boxEdit.AddItem Cell.Value
        End If
    Next Cell
End Sub


