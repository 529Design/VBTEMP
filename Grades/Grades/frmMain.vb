'Programmed by: Jeffrey Marron
'CRN 1475
Option Explicit On
Option Strict On
Imports System.IO


Public Module frmMainModule
    Public listOfStudentRecords() As studentRecord = {New studentRecord, New studentRecord, New studentRecord, New studentRecord, New studentRecord,
    New studentRecord, New studentRecord, New studentRecord, New studentRecord, New studentRecord, New studentRecord, New studentRecord,
    New studentRecord, New studentRecord, New studentRecord, New studentRecord, New studentRecord, New studentRecord}
    Public ClassAverage As Double = 0
    'Public listOfStudentRecords() As studentRecord = New studentRecord(18) {} //Tried this but it wouldn't work
    'https://msdn.microsoft.com/en-us/library/487y7874(v=vs.100).aspx
    'syntax from link above MSDN doesn't work
End Module

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CreateMyBorderlesWindow()
        FileParser()
        lstNames.SelectedIndex = 0
    End Sub


    Public Sub CreateMyBorderlesWindow()
        'FormBorderStyle = FormBorderStyle.None
        MaximizeBox = False
        MinimizeBox = False
        StartPosition = FormStartPosition.CenterScreen
        ' Remove the control box so the form will only display client area.
        ControlBox = False
    End Sub 'CreateMyBorderlesWindow


    Public Sub FileParser()
        'Dim Temp As New studentRecord
        Dim i As Integer = 0
        Dim FilePath As String = "grades.csv"
        Dim sr As StreamReader = New StreamReader(FilePath)

        Try
            Do While sr.EndOfStream() > True
                Dim columns() As String = sr.ReadLine().Split(","c)
                listOfStudentRecords(i).FirstName = columns(0)
                listOfStudentRecords(i).LastName = columns(1)
                listOfStudentRecords(i).Assignment1 = Convert.ToDouble(columns(2))
                listOfStudentRecords(i).Assignment2 = Convert.ToDouble(columns(3))
                listOfStudentRecords(i).Assignment3 = Convert.ToDouble(columns(4))
                listOfStudentRecords(i).Project = Convert.ToDouble(columns(5))
                listOfStudentRecords(i).Midterm = Convert.ToDouble(columns(6))
                listOfStudentRecords(i).FinalGrade = Convert.ToDouble(columns(7))
                'Calculate Average
                listOfStudentRecords(i).Average = Math.Truncate((Convert.ToDouble(((listOfStudentRecords(i).Assignment1 * 2) + (listOfStudentRecords(i).Assignment2 * 2) +
                    (listOfStudentRecords(i).Assignment3 * 2) + (listOfStudentRecords(i).Project * 4) + (listOfStudentRecords(i).Midterm * 5) +
                    (listOfStudentRecords(i).FinalGrade * 5)) / 20)) + 0.5)
                listOfStudentRecords(i).LetterGrade = CalcLetterGrade(listOfStudentRecords(i).Average)

                lstNames.Items.Add((listOfStudentRecords(i).LastName) + ", " + (listOfStudentRecords(i).FirstName))
                ClassAverage = listOfStudentRecords(i).Average + ClassAverage
                i = i + 1
            Loop
            sr.Close()
        Catch e As Exception
            MsgBox("Failed to read input line")
        End Try
        ClassAverage = Math.Truncate(ClassAverage / i)
    End Sub

    Public Function CalcLetterGrade(t As Double) As String
        Dim Temp As String = ""
        Dim number As Double = t
        Select Case number
            Case 0 To 59
                Temp = "E"
            Case 60 To 66
                Temp = "D"
            Case 67 To 69
                Temp = "D+"
            Case 70 To 72
                Temp = "C-"
            Case 73 To 76
                Temp = "C"
            Case 77 To 79
                Temp = "C+"
            Case 80 To 82
                Temp = "B-"
            Case 83 To 86
                Temp = "B"
            Case 87 To 89
                Temp = "B+"
            Case 90 To 92
                Temp = "A-"
            Case 93 To 100
                Temp = "A"
            Case Else
                Debug.WriteLine("Invalid Grade")
        End Select

        Return Temp
    End Function

    Private Sub lstNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstNames.SelectedIndexChanged
        Dim i As Integer = Me.lstNames.SelectedIndex
        lblAsgn1Out.Text = Convert.ToString(listOfStudentRecords(i).Assignment1)
        lblAsgn2Out.Text = Convert.ToString(listOfStudentRecords(i).Assignment2)
        lblAsgn3Out.Text = Convert.ToString(listOfStudentRecords(i).Assignment3)
        lblProjOut.Text = Convert.ToString(listOfStudentRecords(i).Project)
        lblMidOut.Text = Convert.ToString(listOfStudentRecords(i).Midterm)
        lblFinOut.Text = Convert.ToString(listOfStudentRecords(i).FinalGrade)
        lblAvgOut.Text = Convert.ToString(listOfStudentRecords(i).Average)
        lblGradeOut.Text = (listOfStudentRecords(i).LetterGrade)
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim result = MessageBox.Show(" Are you sure you want to exit", "Leaving?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If result = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub
End Class


Public Class studentRecord
    Public Assignment1 As Double, Assignment2 As Double, Assignment3 As Double,
        Project As Double, Midterm As Double = 0, FinalGrade As Double = 0, Average As Double
    Public LetterGrade As String, FirstName As String, LastName As String
End Class