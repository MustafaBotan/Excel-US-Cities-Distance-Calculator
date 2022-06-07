VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Geography Toolbox"
   ClientHeight    =   4608
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   14832
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub GoButton1_Click()
Dim d As Double
Dim i, j As Integer, Lat1 As Double, Lat2 As Double, Long1 As Double, Long2 As Double

If City1Label = "" Or City2Label = "" Then
    GoTo IsBlank
End If

Range("A1").Select
For i = 1 To 855
    If Range("A1:A855").Cells(i, 1) = UserForm1.State1Label Then
        Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            If Range("A" & j + i) = City1Label Then
                Range("A" & j + i).Select
                Lat1 = ActiveCell.Offset(0, 1)
                Long1 = ActiveCell.Offset(0, 2)
            End If
            j = j + 1
        Loop
        j = 0
    End If
Next i
For i = 1 To 855
    If Range("A1:A855").Cells(i, 1) = UserForm1.State2Label Then
        Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            If Range("A" & j + i) = City2Label Then
                Range("A" & j + i).Select
                Lat2 = ActiveCell.Offset(0, 1)
                Long2 = ActiveCell.Offset(0, 2)
            End If
            j = j + 1
        Loop
        j = 0
    End If
Next i

d = NumbersFunction(Lat1, Lat2, Long1, Long2)

MsgBox (UserForm1.City1Label & " and " & UserForm1.City2Label & " are " & FormatNumber(d, 0) & " miles apart as the crow flies.")
Exit Sub
IsBlank:
MsgBox ("Please enter two cities")
End Sub
Function NumbersFunction(Lat1 As Double, Lat2 As Double, Long1 As Double, Long2 As Double) As Double
Dim a, b, c, d, Rad, pi As Double
pi = WorksheetFunction.pi()
Rad = 3960

a = Cos(Lat1 * pi / 180) * Cos(Lat2 * pi / 180) * Cos(Long1 * pi / 180) * Cos(Long2 * pi / 180)
b = Cos(Lat1 * pi / 180) * Sin(Long1 * pi / 180) * Cos(Lat2 * pi / 180) * Sin(Long2 * pi / 180)
c = Sin(Lat1 * pi / 180) * Sin(Lat2 * pi / 180)
NumbersFunction = WorksheetFunction.Acos(a + b + c) * Rad
End Function
Private Sub GoButton2_Click()
Dim d As Double
Dim i As Integer, Lat1 As Double, Lat2 As Double, Long1 As Double, Long2 As Double
Range("A1").Select
For i = 1 To 855
    If Range("A1:A855").Cells(i, 1) = UserForm1.state1select.Text Then
        Range("A" & i).Select
        Lat1 = ActiveCell.Offset(UserForm1.city1select.ListIndex + 1, 1)
        Long1 = ActiveCell.Offset(UserForm1.city1select.ListIndex + 1, 2)
    End If
Next i
For i = 1 To 855
    If Range("A1:A855").Cells(i, 1) = UserForm1.state2select.Text Then
        Range("A" & i).Select
        Lat2 = ActiveCell.Offset(UserForm1.city2select.ListIndex + 1, 1)
        Long2 = ActiveCell.Offset(UserForm1.city2select.ListIndex + 1, 2)
    End If
Next i

d = NumbersFunction(Lat1, Lat2, Long1, Long2)

If city1select.Text = city2select.Text And state1select.Text = state2select.Text Then
MsgBox ("Please select different cities")
Else
MsgBox (UserForm1.city1select.Text & " and " & UserForm1.city2select.Text & " are " & FormatNumber(d, 0) & " miles apart as the crow flies.")
End If
End Sub

Private Sub state1select_Change()
Dim i As Integer, j As Integer
UserForm1.city1select.Clear
Range("A1").Select
For i = 1 To 855
    j = 1
    If Range("A1:A855").Cells(i, 1) = UserForm1.state1select.Text Then
        Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            UserForm1.city1select.AddItem ActiveCell.Offset(j, 0)
            j = j + 1
        Loop
        UserForm1.city1select.Text = ActiveCell.Offset(1, 0)
        Exit For
    End If
Next i
End Sub

Private Sub state2select_Change()
Dim i As Integer, j As Integer
UserForm1.city2select.Clear
Range("A1").Select
For i = 1 To 855
    j = 1
    If Range("A1:A855").Cells(i, 1) = UserForm1.state2select.Text Then
        Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            UserForm1.city2select.AddItem ActiveCell.Offset(j, 0)
            j = j + 1
        Loop
        UserForm1.city2select.Text = ActiveCell.Offset(1, 0)
        Exit For
    End If
Next i
End Sub

Private Sub SearchButton1_Click()
Dim i As Integer, j As Integer, k As Integer
Dim Cities() As String, States() As String
Dim Ans As Integer, commaposition

If city1input.Text = "" Then
GoTo BlankInput2
End If

If IsNumeric(city1input.Text) = True Then
GoTo numInput2:
End If

For i = 1 To 855
    commaposition = InStr(UserForm1.city1input.Text, ",")
    If commaposition = 0 Then
        If UCase(Left(Range("A1:A855").Cells(i, 1), Len(UserForm1.city1input.Text))) = UCase(UserForm1.city1input.Text) _
        And Not UCase(Range("A1:A855").Cells(i, 1)) = Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
            ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not UCase(ActiveCell.Offset(-j, 0)) = ActiveCell.Offset(-j, 0) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j, 0)
                    Exit Do
                End If
            Loop
        End If
    Else
        If UCase(Left(Range("A1:A855").Cells(i, 1), commaposition - 1)) = UCase(Left(UserForm1.city1input.Text, commaposition - 1)) _
        And Not UCase(Range("A1:A855").Cells(i, 1)) = Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
        ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not IsEmpty(ActiveCell.Offset(-j, 0)) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j + 1, 0)
                    Exit Do
                End If
            Loop
        End If
    End If
Next i

If k > 1 Then
    UserForm2.ComboBox1.Clear
    For j = 1 To k
        UserForm2.ComboBox1.AddItem Cities(j) & ", " & States(j)
    Next j
    UserForm2.ComboBox1.Text = Cities(1) & ", " & States(1)
    UserForm2.Show
        Dim cityState() As String
            cityState() = Split(UserForm2.ComboBox1.Text, ", ")
            On Error GoTo line1
            UserForm1.City1Label = cityState(0)
            UserForm1.State1Label = cityState(1)
                'City1Label = Cities(UserForm2.ComboBox1.ListIndex + 1)
                'State1Label = States(UserForm2.ComboBox1.ListIndex + 1)
line1:
On Error GoTo 0
Else
On Error GoTo not_found2
    Ans = MsgBox("Did you mean " & Cities(1) & ", " & States(1) & "?", vbYesNo)
On Error GoTo 0
    If Ans = 7 Then
        MsgBox "Sorry, that's the only location we could find meeting your search criterion."
        Exit Sub
    End If
    City1Label = Cities(1)
    State1Label = States(1)
End If
Exit Sub

'error handling
not_found2:
MsgBox ("No match found")
Exit Sub
BlankInput2:
MsgBox ("Please enter a city")
Exit Sub
numInput2:
MsgBox ("Must be a string")
Exit Sub
End Sub

Private Sub SearchButton2_Click()
Dim i As Integer, j As Integer, k As Integer
Dim Cities() As String, States() As String
Dim Ans As Integer, commaposition

If city2input.Text = "" Then
GoTo BlankInput:
End If

If IsNumeric(city2input.Text) = True Then
GoTo numInput:
End If

For i = 1 To 855
    commaposition = InStr(UserForm1.city2input.Text, ",")
    If commaposition = 0 Then
        If UCase(Left(Range("A1:A855").Cells(i, 1), Len(UserForm1.city2input.Text))) = UCase(UserForm1.city2input.Text) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
            ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not UCase(ActiveCell.Offset(-j, 0)) = ActiveCell.Offset(-j, 0) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j, 0)
                    Exit Do
                End If
            Loop
        End If
    Else
        If UCase(Left(Range("A1:A855").Cells(i, 1), commaposition - 1)) = UCase(Left(UserForm1.city2input.Text, commaposition - 1)) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
        ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not IsEmpty(ActiveCell.Offset(-j, 0)) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j + 1, 0)
                    Exit Do
                End If
            Loop
        End If
    End If
Next i

If k > 1 Then
    UserForm2.ComboBox1.Clear
    For j = 1 To k
        UserForm2.ComboBox1.AddItem Cities(j) & ", " & States(j)
    Next j
    UserForm2.ComboBox1.Text = Cities(1) & ", " & States(1)
    UserForm2.Show
    Dim cityState() As String
cityState() = Split(UserForm2.ComboBox1.Text, ", ")
On Error GoTo line1
UserForm1.City2Label = cityState(0)
UserForm1.State2Label = cityState(1)
    'City2Label = Cities(UserForm2.ComboBox1.ListIndex + 1)
    'State2Label = States(UserForm2.ComboBox1.ListIndex + 1)
line1:
On Error GoTo 0
Else
On Error GoTo not_found
    Ans = MsgBox("Did you mean " & Cities(1) & ", " & States(1) & "?", vbYesNo)
On Error GoTo 0
    If Ans = 7 Then
        MsgBox "Sorry, that's the only location we could find meeting your search criterion."
        Exit Sub
    End If
    City2Label = Cities(1)
    State2Label = States(1)
End If
Exit Sub
'error handling
not_found:
MsgBox ("No match found")
On Error GoTo 0
Exit Sub
BlankInput:
MsgBox ("Please enter a city")
Exit Sub
numInput:
MsgBox ("Must be a string")
Exit Sub
End Sub

Private Sub QuitButton_Click()
Unload Me
End Sub

Private Sub ResetButton_Click()
UserForm1.City1Label = ""
UserForm1.State1Label = ""
UserForm1.City2Label = ""
UserForm1.State2Label = ""

state1select.Text = "ALABAMA"
state2select.Text = "ALABAMA"
End Sub


