Attribute VB_Name = "Feedback"
Dim coursearray(1 To 50, 1 To 50) As Variant
Dim noanswercount As Integer
Dim pname As String
Sub isitallblank()
If WorksheetFunction.Sum(Range("d17:d30")) = 0 And WorksheetFunction.Sum(Range("i17:i30")) = 0 Then
    coursearray(49, 1) = 1 'this is the no response given field
End If
End Sub
'THIS IS THE MAIN MACRO................................................................................
Sub copycoursefb()
'50 50 of array is the no. of attendees
Erase coursearray 'erase array
Call copyfbtoarray
Call openfbwb
'find first empty cell
Call findemptyinfbanalysis
ActiveCell.Offset(1, 0).Select
Call checkifnewcourse
'now go to the insertion point for information
Call copyarraytofb
ActiveCell.Offset(0, -27).Select
'now we are in the n column to enter No response given for courses
coursearray(49, 2) = ActiveCell.Value
ActiveCell.Value = coursearray(49, 2) + coursearray(49, 1)
Call insert15
Call secondfeedback
'Now delete place-holding text
Call delplaceholder
'remove duplicate date and course fields
Call autoclose
Workbooks("feedback analysis 2017.xlsx").Sheets("feedback data").Select
Call autoclose
Workbooks("feedback.xlsx").Sheets("sheet1").Select
Call autoclose
Application.Quit
End Sub

Sub copyfbtoarray() 'Copies fb form to array
Dim cc As Integer
Dim cc2 As Integer
coursearray(1, 1) = Range("d7").Value 'What was good about your training today?
'coursearray(2, 1) = Range("d8").Value appears to be no purpose for this
coursearray(3, 1) = Range("f9").Value 'Overall rating 1-4
coursearray(4, 1) = Range("d12").Value 'Suggested improvements
coursearray(5, 1) = Range("f13").Value 'Was the course confirmation pack received?
coursearray(48, 1) = Range("i9").Value 'Today’s date
    For cc = 0 To 13
        coursearray(6 + cc, 1) = Range("d17").Offset(cc, 0).Value 'add the values of d17 to d30 of course evaluation to array
    Next cc
    For cc2 = 0 To 12
        coursearray(20 + cc2, 1) = Range("i17").Offset(cc2, 0).Value 'add the values of i17 to i24 of course evaluation to array
    Next cc2
'if no courses given then add 1 to 49,1 of array
Call isitallblank
coursearray(50, 1) = Range("f35").Value 'Are there any courses you would like to see that are not included above?
End Sub
Sub openfbwb() 'open destination workbook
    Workbooks.Open Filename:="e:\Feedback analysis 2017.xlsx"
End Sub
Sub findemptyinfbanalysis()
'open feedback workbook
'Workbooks("Feedback analysis 2017.xlsx").Activate
'find an empty cell in column a
Range("a3").Select
    While IsEmpty(ActiveCell) = False
        ActiveCell.Offset(1, 0).Select
    Wend
End Sub

Sub checkifnewcourse() 'if it’s a new course we want the head to move down one line, otherwise stay on the same line
Dim iamdate As Date
iamdate = ActiveCell.Offset(0, 1).Value
    'checks whether course name and date match. If they don’t we move down a line and add details
    If ActiveCell.Value = "Introduction to using SchüCal (Commercial)" Then
    'check that dates tally
    Dim iamdate2 As Date
    iamdate2 = coursearray(48, 1)
    If iamdate = iamdate2 Or iamdate = iamdate2 + 1 Or iamdate = iamdate2 - 1 Then
    Exit Sub
    End If
    Else
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "Introduction to using SchüCal (Commercial)"
    ActiveCell.Offset(0, 1).Value = coursearray(48, 1)
    ActiveCell.Offset(0, 2).Value = "CZ"
    End If
    End Sub

Sub copyarraytofb()
Dim dd As Integer
ActiveCell.Offset(0, 14).Select
    For dd = 0 To 25
        coursearray(dd + 6, 2) = ActiveCell.Value
        ActiveCell.Value = coursearray(6 + dd, 1) + coursearray(dd + 6, 2)
        ActiveCell.Offset(0, 1).Select
    Next dd
End Sub

Sub insert15()
'insert 1-5 rating
ActiveCell.Offset(0, -7).Select
noanswercount = ActiveCell.Offset(0, 1).Value
'1-2 to 5-2 of course array are for incrementing overall course rating values
'6-2 to 35-2 are for incrementing courses requested
    If coursearray(3, 1) = 4 Then
        coursearray(1, 2) = ActiveCell.Value
        ActiveCell.Value = coursearray(1, 2) + 1
        GoTo line99
    End If
    If coursearray(3, 1) = 3 Then
        coursearray(1, 2) = ActiveCell.Offset(0, -1).Value
        ActiveCell.Offset(0, -1).Value = coursearray(1, 2) + 1
        GoTo line99
    End If
    If coursearray(3, 1) = 2 Then
        coursearray(1, 2) = ActiveCell.Offset(0, -2).Value
        ActiveCell.Offset(0, -2).Value = coursearray(1, 2) + 1
        GoTo line99
    End If
    If coursearray(3, 1) = 1 Then
        coursearray(1, 2) = ActiveCell.Offset(0, -3).Value
        ActiveCell.Offset(0, -3).Value = coursearray(1, 2) + 1
        GoTo line99
    Else: ActiveCell.Offset(0, 1).Value = noanswercount + 1
    End If
'now insert course pack
line99:
Dim confpack As Integer
If coursearray(5, 1) = "Yes" Then
    confpack = ActiveCell.Offset(0, 3).Value
    confpack = confpack + 1
    ActiveCell.Offset(0, 3).Value = confpack
End If
If coursearray(5, 1) = "No" Then
    confpack = ActiveCell.Offset(0, 4).Value
    confpack = confpack + 1
    ActiveCell.Offset(0, 4).Value = confpack
End If
If coursearray(5, 1) = "" Then
    confpack = ActiveCell.Offset(0, 5).Value
    confpack = confpack + 1
    ActiveCell.Offset(0, 5).Value = confpack
End If
'How many attendees were there?
coursearray(50, 50) = Range("m" & ActiveCell.Row).Value
End Sub

Sub secondfeedback()
Workbooks.Open Filename:="e:\Feedback analysis spreadsheet 2017.xls"
Range("f1").Select
    'find bottom cell based on f
Dim LngLastRow As Long
LngLastRow = Range("f1").SpecialCells(xlCellTypeLastCell).Row
ActiveCell.Offset(LngLastRow, 0).Select
    While IsEmpty(ActiveCell) = False
        ActiveCell.Offset(1, 0).Select
    Wend
ActiveCell.Offset(0, 1).Select
While IsEmpty(ActiveCell) = False
        ActiveCell.Offset(1, 0).Select
    Wend
'enter date field
ActiveCell.Offset(0, -6).Select
'only enter if date and person is not same as above
If (coursearray(48, 1) <> ActiveCell.Offset(-1, 0)) And (ActiveCell.Offset(-1, 1).Value <> "CZ") _
And (ActiveCell.Offset(-2, 1).Value <> "CZ") And (ActiveCell.Offset(-3, 1).Value <> "CZ") _
And (ActiveCell.Offset(-4, 1).Value <> "CZ") And (ActiveCell.Offset(-5, 1).Value <> "CZ") _
And (ActiveCell.Offset(-6, 1).Value <> "CZ") And (ActiveCell.Offset(-7, 1).Value <> "CZ") _
And (ActiveCell.Offset(-8, 1).Value <> "CZ") Then
    ActiveCell.Value = coursearray(48, 1)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "CZ"
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Offset(0, -1).Value = "Introduction to using SchüCal (Commercial)"
End If
'insert att. count
'make f the active column
Range("f" & ActiveCell.Row).Select
ActiveCell.Value = coursearray(1, 1)
ActiveCell.Offset(0, 1).Value = coursearray(4, 1)
Range("c" & ActiveCell.Row).Select
'if there is no course in the active cell, then look upwards
    While IsEmpty(ActiveCell) = True
    ActiveCell.Offset(-1, 0).Select
    Wend
ActiveCell.Offset(0, 1).Value = coursearray(50, 50)
End Sub

Sub delplaceholder()
Workbooks("feedback analysis spreadsheet 2017.xls").Sheets("2017").Activate
    Cells.Replace What:="write here (don’t worry if text extends beyond page)", _
        Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Sub deletedupdateandcourse()
ActiveCell.Offset(0, -3).Range("A1").Select
    While ActiveCell.Value = ActiveCell.Offset(-1, 0)
        Range(Selection, "A" & Selection.Row).Clear
        'Get out if blank
        ActiveCell.Offset(-1, 0).Select
        If ActiveCell.Value = "" Then
        Exit Sub
        End If
        Wend
    ActiveCell.Offset(1, 1).Select
    While ActiveCell.Value = ActiveCell.Offset(-1, 0)
    ActiveCell.Clear
    Wend
End Sub

Sub autoclose()
ActiveWorkbook.Close savechanges:=True
End Sub



