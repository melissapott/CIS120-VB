Option Strict On
Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

    'Weekly Payroll Calculator
    'Author:        Melissa Pott
    'Class:         CIS120
    'Assignment:    Lab #6
    'Instructor:    Loay Alnaji
    'This program will read values from two text files containing employee records and
    'time clock records, and output each employee's weekly pay and payroll summary to
    'a third text file

    Dim EmployeeReader As StreamReader
    Dim TimeClockReader As StreamReader
    Dim PayrollOutputWriter As StreamWriter
    Dim EmployeeFile, TimeClockFile As String
    Dim EmployeeSelected, ClockSelected, OutputSelected As Boolean


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCalculate As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLocateEmployee As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLocateClock As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLocateOutput As System.Windows.Forms.MenuItem
    Friend WithEvents lblEmployeeFile As System.Windows.Forms.Label
    Friend WithEvents lblTimeClockFile As System.Windows.Forms.Label
    Friend WithEvents lblPayrollOutputFile As System.Windows.Forms.Label
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuInstructions As System.Windows.Forms.MenuItem
    Friend WithEvents btnCalculate As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuCalculate = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.mnuLocateEmployee = New System.Windows.Forms.MenuItem
        Me.mnuLocateClock = New System.Windows.Forms.MenuItem
        Me.mnuLocateOutput = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.mnuAbout = New System.Windows.Forms.MenuItem
        Me.mnuInstructions = New System.Windows.Forms.MenuItem
        Me.lblEmployeeFile = New System.Windows.Forms.Label
        Me.lblTimeClockFile = New System.Windows.Forms.Label
        Me.lblPayrollOutputFile = New System.Windows.Forms.Label
        Me.btnCalculate = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem4, Me.MenuItem8})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCalculate, Me.mnuExit})
        Me.MenuItem1.Text = "File"
        '
        'mnuCalculate
        '
        Me.mnuCalculate.Index = 0
        Me.mnuCalculate.Text = "Calculate"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 1
        Me.mnuExit.Text = "Exit"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuLocateEmployee, Me.mnuLocateClock, Me.mnuLocateOutput})
        Me.MenuItem4.Text = "Locate"
        '
        'mnuLocateEmployee
        '
        Me.mnuLocateEmployee.Index = 0
        Me.mnuLocateEmployee.Text = "Employee File"
        '
        'mnuLocateClock
        '
        Me.mnuLocateClock.Index = 1
        Me.mnuLocateClock.Text = "Time Clock File"
        '
        'mnuLocateOutput
        '
        Me.mnuLocateOutput.Index = 2
        Me.mnuLocateOutput.Text = "Weekly Payroll File"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAbout, Me.mnuInstructions})
        Me.MenuItem8.Text = "Help"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 0
        Me.mnuAbout.Text = "About"
        '
        'mnuInstructions
        '
        Me.mnuInstructions.Index = 1
        Me.mnuInstructions.Text = "Instructions"
        '
        'lblEmployeeFile
        '
        Me.lblEmployeeFile.Location = New System.Drawing.Point(16, 32)
        Me.lblEmployeeFile.Name = "lblEmployeeFile"
        Me.lblEmployeeFile.Size = New System.Drawing.Size(488, 40)
        Me.lblEmployeeFile.TabIndex = 0
        Me.lblEmployeeFile.Text = "Employee File:  "
        '
        'lblTimeClockFile
        '
        Me.lblTimeClockFile.Location = New System.Drawing.Point(16, 88)
        Me.lblTimeClockFile.Name = "lblTimeClockFile"
        Me.lblTimeClockFile.Size = New System.Drawing.Size(488, 40)
        Me.lblTimeClockFile.TabIndex = 1
        Me.lblTimeClockFile.Text = "Time Clock File:  "
        '
        'lblPayrollOutputFile
        '
        Me.lblPayrollOutputFile.Location = New System.Drawing.Point(16, 144)
        Me.lblPayrollOutputFile.Name = "lblPayrollOutputFile"
        Me.lblPayrollOutputFile.Size = New System.Drawing.Size(488, 40)
        Me.lblPayrollOutputFile.TabIndex = 2
        Me.lblPayrollOutputFile.Text = "Weekly Payroll File:  "
        '
        'btnCalculate
        '
        Me.btnCalculate.Location = New System.Drawing.Point(408, 216)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.TabIndex = 3
        Me.btnCalculate.Text = "Calculate"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(568, 266)
        Me.Controls.Add(Me.btnCalculate)
        Me.Controls.Add(Me.lblPayrollOutputFile)
        Me.Controls.Add(Me.lblTimeClockFile)
        Me.Controls.Add(Me.lblEmployeeFile)
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "Weekly Payroll Calculator"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        'Exit the application in response to click on Exit menu option
        Application.Exit()

    End Sub

    Private Sub mnuLocateEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLocateEmployee.Click
        'This routine displays file dialog to allow user to select the input file containing employee records

        Dim dlgOpenEmployeeFile As New OpenFileDialog

        'set initial directory and default file types
        dlgOpenEmployeeFile.InitialDirectory = "C:\"
        dlgOpenEmployeeFile.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        'assign the file name selected to the file name variable and set the file selected flag to true
        If dlgOpenEmployeeFile.ShowDialog = DialogResult.OK Then
            lblEmployeeFile.Text = "Employee File:  " & dlgOpenEmployeeFile.FileName
            EmployeeFile = dlgOpenEmployeeFile.FileName
            EmployeeSelected = True
        End If

    End Sub

    Private Sub mnuLocateClock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLocateClock.Click
        'This routine displays the file dialog to allow user to select the input file containing timeclock records

        Dim dlgOpenClockFile As New OpenFileDialog

        'set initial directory and default file types
        dlgOpenClockFile.InitialDirectory = "C:\"
        dlgOpenClockFile.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        'assign the file name selected to the file name variable and set the file selected flag to true
        If dlgOpenClockFile.ShowDialog = DialogResult.OK Then
            lblTimeClockFile.Text = "Time Clock File:  " & dlgOpenClockFile.FileName
            TimeClockFile = dlgOpenClockFile.FileName
            ClockSelected = True
        End If

    End Sub

    Private Sub mnuLocateOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLocateOutput.Click
        'This routine displays the dialog to allow the user to select the output file name and location
        Dim dlgSavePayrollOutput As New SaveFileDialog

        'set initial directory and default file types
        dlgSavePayrollOutput.InitialDirectory = "C:\"
        dlgSavePayrollOutput.Filter = "Text Files (*.txt)|*.txt|All Files|*3*"

        'open the output file for writing and set the file selected flag to true
        If dlgSavePayrollOutput.ShowDialog() = DialogResult.OK Then
            PayrollOutputWriter = File.CreateText(dlgSavePayrollOutput.FileName)
            lblPayrollOutputFile.Text = "Payroll Output File:  " & dlgSavePayrollOutput.FileName
            OutputSelected = True
        End If
    End Sub

    Private Function paycheck(ByVal intEmployeeNumber As Integer, ByVal dblPayRate As Double) As Double
        'This function reads through the timeclock file looking for records associated with the employee
        'number that was passed in.  When it finds a record for that employee, it will subtract the start
        'time from the stop time to find the hours worked, multiply that by the pay rate (which was also
        'passed in) and accumulate that amount into the dblPay variable.  The function returns the final
        'value of the dblPay variable, which is the amount of that employee's weekly pay amount.

        Dim intClockEmployee As Integer
        Dim StartTime, StopTime As Date
        Dim dblHoursWorked, dblPay As Double

        Try
            dblPay = 0
            TimeClockReader = File.OpenText(TimeClockFile)
            Do Until TimeClockReader.Peek = -1
                intClockEmployee = CInt(TimeClockReader.ReadLine)
                If intClockEmployee = -1 Then
                    Exit Do
                End If

                StartTime = CDate(TimeClockReader.ReadLine)
                StopTime = CDate(TimeClockReader.ReadLine)

                If intEmployeeNumber = intClockEmployee Then
                    dblHoursWorked = (StopTime.Subtract(StartTime)).TotalHours
                    dblPay = dblPay + (dblHoursWorked * dblPayRate)
                End If
            Loop
        Catch
            MsgBox("There was a problem reading from the file.  Please check file locations and try again")
        End Try

        Return dblPay

    End Function

    Private Sub mnuCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCalculate.Click
        Calculate()
    End Sub

    Private Sub Calculate()
        'This routine reads through the employee file, and passes the employee number
        'and pay rate for each employee to the paycheck function

        Dim intEmployeeNumber As Integer
        Dim strEmployeeName As String
        Dim dblPayRate, dblPayroll, dblEmployeePaycheck As Double
        Try
            If EmployeeSelected And ClockSelected And OutputSelected Then 'make sure all required files have been selected first

                EmployeeReader = File.OpenText(EmployeeFile)

                Do Until EmployeeReader.Peek = -1
                    intEmployeeNumber = CInt(EmployeeReader.ReadLine)
                    If intEmployeeNumber = -1 Then
                        Exit Do 'we've reached the sentinal value - we're finished processing
                    End If
                    strEmployeeName = EmployeeReader.ReadLine
                    dblPayRate = CDbl(EmployeeReader.ReadLine)

                    'send the employee number and payrate to the paycheck  function
                    dblEmployeePaycheck = paycheck(intEmployeeNumber, dblPayRate)

                    'add the employee's pay amount (returned from the paycheck function above)
                    'to the total company payroll
                    dblPayroll = dblPayroll + dblEmployeePaycheck

                    'write the employee's name and paycheck amount to the output file
                    PayrollOutputWriter.WriteLine(strEmployeeName & vbTab & FormatCurrency(dblEmployeePaycheck))

                Loop

                'All the employees have been processed, so write the summary line to the output file
                PayrollOutputWriter.WriteLine("Total Payroll:" & vbTab & FormatCurrency(dblPayroll))
                MsgBox("Payroll Calculation Complete")

                'clean up all the open files
                EmployeeReader.Close()
                TimeClockReader.Close()
                PayrollOutputWriter.Close()
            Else
                MsgBox("You must locate all files before calculating")
            End If
        Catch
            MsgBox("There was a problem reading from the file.  Please check file locations and try again")
        End Try
    End Sub

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        'Display the "About" form
        Dim frmAbout As New AboutForm
        frmAbout.ShowDialog()
        frmAbout.Dispose()
    End Sub

    Private Sub MnuInstructions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInstructions.Click
        'Display the "Instructions" form
        Dim frmInstructions As New InstructionsForm
        frmInstructions.ShowDialog()
        frmInstructions.Dispose()
    End Sub

    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click
        Calculate()
    End Sub
End Class
