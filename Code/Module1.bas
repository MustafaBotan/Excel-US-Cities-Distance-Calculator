Attribute VB_Name = "Module1"
Option Explicit

Sub RunForm()
Application.ScreenUpdating = False
Worksheets("Sheet1").Visible = True

PopulateStates
UserForm1.Show

Worksheets("Sheet1").Visible = False
Application.ScreenUpdating = True
End Sub

Sub PopulateStates()
Dim tWB As Workbook
Dim ncategories As Integer, i As Integer
Set tWB = ThisWorkbook
tWB.Activate

Sheets("Sheet1").Select
Range("A1").Select
ncategories = WorksheetFunction.CountA(Columns("E:E"))
For i = 1 To ncategories
    UserForm1.state1select.AddItem Range("E1:E" & ncategories).Cells(i, 1)
    UserForm1.state2select.AddItem Range("E1:E" & ncategories).Cells(i, 1)
Next i
UserForm1.state1select.Text = Range("E1:E51").Cells(1, 1)
UserForm1.state2select.Text = Range("E1:E51").Cells(1, 1)
End Sub
