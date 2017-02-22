<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.mnuTitle = New System.Windows.Forms.MenuStrip()
        Me.DirectoryEntry1 = New System.DirectoryServices.DirectoryEntry()
        Me.lstNames = New System.Windows.Forms.ListBox()
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblNames = New System.Windows.Forms.Label()
        Me.lblAssignmentGrades = New System.Windows.Forms.Label()
        Me.lblProjectGrades = New System.Windows.Forms.Label()
        Me.lblMidtermGrade = New System.Windows.Forms.Label()
        Me.lvlAverage = New System.Windows.Forms.Label()
        Me.lblGrade = New System.Windows.Forms.Label()
        Me.lblFinalGrade = New System.Windows.Forms.Label()
        Me.lblAsgn1Out = New System.Windows.Forms.Label()
        Me.lblAsgn2Out = New System.Windows.Forms.Label()
        Me.lblAsgn3Out = New System.Windows.Forms.Label()
        Me.lblProjOut = New System.Windows.Forms.Label()
        Me.lblMidOut = New System.Windows.Forms.Label()
        Me.lblAvgOut = New System.Windows.Forms.Label()
        Me.lblFinOut = New System.Windows.Forms.Label()
        Me.lblGradeOut = New System.Windows.Forms.Label()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'mnuTitle
        '
        Me.mnuTitle.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.mnuTitle.Location = New System.Drawing.Point(0, 0)
        Me.mnuTitle.Name = "mnuTitle"
        Me.mnuTitle.Size = New System.Drawing.Size(450, 24)
        Me.mnuTitle.TabIndex = 0
        Me.mnuTitle.Text = "MenuStrip1"
        '
        'lstNames
        '
        Me.lstNames.FormattingEnabled = True
        Me.lstNames.Location = New System.Drawing.Point(8, 94)
        Me.lstNames.Name = "lstNames"
        Me.lstNames.Size = New System.Drawing.Size(120, 199)
        Me.lstNames.TabIndex = 3
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.lblTitle.Location = New System.Drawing.Point(5, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(107, 13)
        Me.lblTitle.TabIndex = 4
        Me.lblTitle.Text = "Jeff's Grade Reporter"
        '
        'lblNames
        '
        Me.lblNames.AutoSize = True
        Me.lblNames.Location = New System.Drawing.Point(12, 63)
        Me.lblNames.Name = "lblNames"
        Me.lblNames.Size = New System.Drawing.Size(43, 13)
        Me.lblNames.TabIndex = 5
        Me.lblNames.Text = "Names:"
        '
        'lblAssignmentGrades
        '
        Me.lblAssignmentGrades.AutoSize = True
        Me.lblAssignmentGrades.Location = New System.Drawing.Point(160, 94)
        Me.lblAssignmentGrades.Name = "lblAssignmentGrades"
        Me.lblAssignmentGrades.Size = New System.Drawing.Size(61, 26)
        Me.lblAssignmentGrades.TabIndex = 6
        Me.lblAssignmentGrades.Text = "Assignment" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Grades:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'lblProjectGrades
        '
        Me.lblProjectGrades.AutoSize = True
        Me.lblProjectGrades.Location = New System.Drawing.Point(160, 155)
        Me.lblProjectGrades.Name = "lblProjectGrades"
        Me.lblProjectGrades.Size = New System.Drawing.Size(44, 26)
        Me.lblProjectGrades.TabIndex = 7
        Me.lblProjectGrades.Text = "Project" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Grades:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'lblMidtermGrade
        '
        Me.lblMidtermGrade.AutoSize = True
        Me.lblMidtermGrade.Location = New System.Drawing.Point(160, 220)
        Me.lblMidtermGrade.Name = "lblMidtermGrade"
        Me.lblMidtermGrade.Size = New System.Drawing.Size(44, 26)
        Me.lblMidtermGrade.TabIndex = 8
        Me.lblMidtermGrade.Text = "Midterm" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Grade:"
        '
        'lvlAverage
        '
        Me.lvlAverage.AutoSize = True
        Me.lvlAverage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvlAverage.Location = New System.Drawing.Point(160, 280)
        Me.lvlAverage.Name = "lvlAverage"
        Me.lvlAverage.Size = New System.Drawing.Size(58, 13)
        Me.lvlAverage.TabIndex = 9
        Me.lvlAverage.Text = "Average:"
        '
        'lblGrade
        '
        Me.lblGrade.AutoSize = True
        Me.lblGrade.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGrade.Location = New System.Drawing.Point(320, 280)
        Me.lblGrade.Name = "lblGrade"
        Me.lblGrade.Size = New System.Drawing.Size(45, 13)
        Me.lblGrade.TabIndex = 10
        Me.lblGrade.Text = "Grade:"
        '
        'lblFinalGrade
        '
        Me.lblFinalGrade.AutoSize = True
        Me.lblFinalGrade.Location = New System.Drawing.Point(320, 220)
        Me.lblFinalGrade.Name = "lblFinalGrade"
        Me.lblFinalGrade.Size = New System.Drawing.Size(39, 26)
        Me.lblFinalGrade.TabIndex = 11
        Me.lblFinalGrade.Text = "Final" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Grade:"
        '
        'lblAsgn1Out
        '
        Me.lblAsgn1Out.AutoSize = True
        Me.lblAsgn1Out.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAsgn1Out.Location = New System.Drawing.Point(227, 94)
        Me.lblAsgn1Out.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblAsgn1Out.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblAsgn1Out.Name = "lblAsgn1Out"
        Me.lblAsgn1Out.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblAsgn1Out.Size = New System.Drawing.Size(42, 21)
        Me.lblAsgn1Out.TabIndex = 12
        Me.lblAsgn1Out.Text = "         "
        '
        'lblAsgn2Out
        '
        Me.lblAsgn2Out.AutoSize = True
        Me.lblAsgn2Out.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAsgn2Out.Location = New System.Drawing.Point(298, 94)
        Me.lblAsgn2Out.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblAsgn2Out.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblAsgn2Out.Name = "lblAsgn2Out"
        Me.lblAsgn2Out.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblAsgn2Out.Size = New System.Drawing.Size(42, 21)
        Me.lblAsgn2Out.TabIndex = 13
        Me.lblAsgn2Out.Text = "         "
        '
        'lblAsgn3Out
        '
        Me.lblAsgn3Out.AutoSize = True
        Me.lblAsgn3Out.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAsgn3Out.Location = New System.Drawing.Point(371, 94)
        Me.lblAsgn3Out.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblAsgn3Out.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblAsgn3Out.Name = "lblAsgn3Out"
        Me.lblAsgn3Out.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblAsgn3Out.Size = New System.Drawing.Size(42, 21)
        Me.lblAsgn3Out.TabIndex = 14
        Me.lblAsgn3Out.Text = "         "
        '
        'lblProjOut
        '
        Me.lblProjOut.AutoSize = True
        Me.lblProjOut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblProjOut.Location = New System.Drawing.Point(227, 155)
        Me.lblProjOut.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblProjOut.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblProjOut.Name = "lblProjOut"
        Me.lblProjOut.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblProjOut.Size = New System.Drawing.Size(42, 21)
        Me.lblProjOut.TabIndex = 15
        Me.lblProjOut.Text = "         "
        '
        'lblMidOut
        '
        Me.lblMidOut.AutoSize = True
        Me.lblMidOut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMidOut.Location = New System.Drawing.Point(227, 220)
        Me.lblMidOut.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblMidOut.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblMidOut.Name = "lblMidOut"
        Me.lblMidOut.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblMidOut.Size = New System.Drawing.Size(42, 21)
        Me.lblMidOut.TabIndex = 16
        Me.lblMidOut.Text = "         "
        '
        'lblAvgOut
        '
        Me.lblAvgOut.AutoSize = True
        Me.lblAvgOut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAvgOut.Location = New System.Drawing.Point(227, 280)
        Me.lblAvgOut.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblAvgOut.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblAvgOut.Name = "lblAvgOut"
        Me.lblAvgOut.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblAvgOut.Size = New System.Drawing.Size(42, 21)
        Me.lblAvgOut.TabIndex = 17
        Me.lblAvgOut.Text = "         "
        '
        'lblFinOut
        '
        Me.lblFinOut.AutoSize = True
        Me.lblFinOut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFinOut.Location = New System.Drawing.Point(371, 220)
        Me.lblFinOut.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblFinOut.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblFinOut.Name = "lblFinOut"
        Me.lblFinOut.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblFinOut.Size = New System.Drawing.Size(42, 21)
        Me.lblFinOut.TabIndex = 18
        Me.lblFinOut.Text = "         "
        '
        'lblGradeOut
        '
        Me.lblGradeOut.AutoSize = True
        Me.lblGradeOut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGradeOut.Location = New System.Drawing.Point(371, 280)
        Me.lblGradeOut.MaximumSize = New System.Drawing.Size(42, 21)
        Me.lblGradeOut.MinimumSize = New System.Drawing.Size(42, 21)
        Me.lblGradeOut.Name = "lblGradeOut"
        Me.lblGradeOut.Padding = New System.Windows.Forms.Padding(0, 0, 6, 6)
        Me.lblGradeOut.Size = New System.Drawing.Size(42, 21)
        Me.lblGradeOut.TabIndex = 19
        Me.lblGradeOut.Text = "         "
        '
        'btnReport
        '
        Me.btnReport.Location = New System.Drawing.Point(89, 359)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Size = New System.Drawing.Size(115, 23)
        Me.btnReport.TabIndex = 20
        Me.btnReport.Text = "Summary Report"
        Me.btnReport.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(250, 359)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(115, 23)
        Me.btnExit.TabIndex = 21
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(450, 419)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnReport)
        Me.Controls.Add(Me.lblGradeOut)
        Me.Controls.Add(Me.lblFinOut)
        Me.Controls.Add(Me.lblAvgOut)
        Me.Controls.Add(Me.lblMidOut)
        Me.Controls.Add(Me.lblProjOut)
        Me.Controls.Add(Me.lblAsgn3Out)
        Me.Controls.Add(Me.lblAsgn2Out)
        Me.Controls.Add(Me.lblAsgn1Out)
        Me.Controls.Add(Me.lblFinalGrade)
        Me.Controls.Add(Me.lblGrade)
        Me.Controls.Add(Me.lvlAverage)
        Me.Controls.Add(Me.lblMidtermGrade)
        Me.Controls.Add(Me.lblProjectGrades)
        Me.Controls.Add(Me.lblAssignmentGrades)
        Me.Controls.Add(Me.lblNames)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.lstNames)
        Me.Controls.Add(Me.mnuTitle)
        Me.MainMenuStrip = Me.mnuTitle
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents mnuTitle As MenuStrip
    Friend WithEvents DirectoryEntry1 As DirectoryServices.DirectoryEntry
    Friend WithEvents lstNames As ListBox
    Friend WithEvents SerialPort1 As IO.Ports.SerialPort
    Friend WithEvents lblTitle As Label
    Friend WithEvents lblNames As Label
    Friend WithEvents lblAssignmentGrades As Label
    Friend WithEvents lblProjectGrades As Label
    Friend WithEvents lblMidtermGrade As Label
    Friend WithEvents lvlAverage As Label
    Friend WithEvents lblGrade As Label
    Friend WithEvents lblFinalGrade As Label
    Friend WithEvents lblAsgn1Out As Label
    Friend WithEvents lblAsgn2Out As Label
    Friend WithEvents lblAsgn3Out As Label
    Friend WithEvents lblProjOut As Label
    Friend WithEvents lblMidOut As Label
    Friend WithEvents lblAvgOut As Label
    Friend WithEvents lblFinOut As Label
    Friend WithEvents lblGradeOut As Label
    Friend WithEvents btnReport As Button
    Friend WithEvents btnExit As Button
End Class
