Public Class Form1
    Inherits System.Windows.Forms.Form

    'Lifetime Earnings Redux
    'Author:        Melissa Pott
    'Class:         CIS120
    'Assignment:    Lab #4
    'Instructor:    Loay Alnaji
    'Based on user input of current age, current salary, annual raise percent and retirement age,
    'this program calculates the new salary for each year and accumulates the total to produce an
    'output of the user's total earnings from the start point until retirement.

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
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtAge As System.Windows.Forms.TextBox
    Friend WithEvents txtCurrentSalary As System.Windows.Forms.TextBox
    Friend WithEvents txtAnnualRaise As System.Windows.Forms.TextBox
    Friend WithEvents txtRetirementAge As System.Windows.Forms.TextBox
    Friend WithEvents btnCalculate As System.Windows.Forms.Button
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents lblAge As System.Windows.Forms.Label
    Friend WithEvents lblCurrentSalary As System.Windows.Forms.Label
    Friend WithEvents lblAnnualRaise As System.Windows.Forms.Label
    Friend WithEvents lblRetirementAge As System.Windows.Forms.Label
    Friend WithEvents lblOutput As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtAge = New System.Windows.Forms.TextBox
        Me.txtCurrentSalary = New System.Windows.Forms.TextBox
        Me.txtAnnualRaise = New System.Windows.Forms.TextBox
        Me.txtRetirementAge = New System.Windows.Forms.TextBox
        Me.btnCalculate = New System.Windows.Forms.Button
        Me.lblName = New System.Windows.Forms.Label
        Me.lblAge = New System.Windows.Forms.Label
        Me.lblCurrentSalary = New System.Windows.Forms.Label
        Me.lblAnnualRaise = New System.Windows.Forms.Label
        Me.lblRetirementAge = New System.Windows.Forms.Label
        Me.lblOutput = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(8, 40)
        Me.txtName.Name = "txtName"
        Me.txtName.TabIndex = 0
        Me.txtName.Text = ""
        '
        'txtAge
        '
        Me.txtAge.Location = New System.Drawing.Point(128, 40)
        Me.txtAge.Name = "txtAge"
        Me.txtAge.TabIndex = 1
        Me.txtAge.Text = ""
        '
        'txtCurrentSalary
        '
        Me.txtCurrentSalary.Location = New System.Drawing.Point(248, 40)
        Me.txtCurrentSalary.Name = "txtCurrentSalary"
        Me.txtCurrentSalary.TabIndex = 2
        Me.txtCurrentSalary.Text = ""
        '
        'txtAnnualRaise
        '
        Me.txtAnnualRaise.Location = New System.Drawing.Point(368, 40)
        Me.txtAnnualRaise.Name = "txtAnnualRaise"
        Me.txtAnnualRaise.TabIndex = 3
        Me.txtAnnualRaise.Text = ""
        '
        'txtRetirementAge
        '
        Me.txtRetirementAge.Location = New System.Drawing.Point(488, 40)
        Me.txtRetirementAge.Name = "txtRetirementAge"
        Me.txtRetirementAge.TabIndex = 4
        Me.txtRetirementAge.Text = ""
        '
        'btnCalculate
        '
        Me.btnCalculate.Location = New System.Drawing.Point(152, 80)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(296, 32)
        Me.btnCalculate.TabIndex = 5
        Me.btnCalculate.Text = "Calculate Lifetime Earnings"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(8, 24)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(100, 16)
        Me.lblName.TabIndex = 7
        Me.lblName.Text = "Name"
        '
        'lblAge
        '
        Me.lblAge.Location = New System.Drawing.Point(128, 24)
        Me.lblAge.Name = "lblAge"
        Me.lblAge.Size = New System.Drawing.Size(100, 16)
        Me.lblAge.TabIndex = 8
        Me.lblAge.Text = "Age"
        '
        'lblCurrentSalary
        '
        Me.lblCurrentSalary.Location = New System.Drawing.Point(248, 24)
        Me.lblCurrentSalary.Name = "lblCurrentSalary"
        Me.lblCurrentSalary.Size = New System.Drawing.Size(100, 16)
        Me.lblCurrentSalary.TabIndex = 9
        Me.lblCurrentSalary.Text = "Current Salary"
        '
        'lblAnnualRaise
        '
        Me.lblAnnualRaise.Location = New System.Drawing.Point(368, 24)
        Me.lblAnnualRaise.Name = "lblAnnualRaise"
        Me.lblAnnualRaise.Size = New System.Drawing.Size(100, 16)
        Me.lblAnnualRaise.TabIndex = 10
        Me.lblAnnualRaise.Text = "Annual Raise %"
        '
        'lblRetirementAge
        '
        Me.lblRetirementAge.Location = New System.Drawing.Point(488, 24)
        Me.lblRetirementAge.Name = "lblRetirementAge"
        Me.lblRetirementAge.Size = New System.Drawing.Size(100, 16)
        Me.lblRetirementAge.TabIndex = 11
        Me.lblRetirementAge.Text = "Retirement Age"
        '
        'lblOutput
        '
        Me.lblOutput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOutput.Location = New System.Drawing.Point(40, 136)
        Me.lblOutput.Name = "lblOutput"
        Me.lblOutput.Size = New System.Drawing.Size(512, 40)
        Me.lblOutput.TabIndex = 12
        Me.lblOutput.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 198)
        Me.Controls.Add(Me.lblOutput)
        Me.Controls.Add(Me.lblRetirementAge)
        Me.Controls.Add(Me.lblAnnualRaise)
        Me.Controls.Add(Me.lblCurrentSalary)
        Me.Controls.Add(Me.lblAge)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.btnCalculate)
        Me.Controls.Add(Me.txtRetirementAge)
        Me.Controls.Add(Me.txtAnnualRaise)
        Me.Controls.Add(Me.txtCurrentSalary)
        Me.Controls.Add(Me.txtAge)
        Me.Controls.Add(Me.txtName)
        Me.Name = "Form1"
        Me.Text = "Lifetime Earnings Redux"
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click
        'Declare the variables we are going to use
        Dim dblEarnings As Double
        Dim dblNewSalary As Double
        Dim intAgeCount As Integer

        'Initiate the variables using the starting values entered by the user
        dblEarnings = CDbl(txtCurrentSalary.Text)
        dblNewSalary = CDbl(txtCurrentSalary.Text)

        'loop through each year between start and retirement age, calculating the annual raise
        'and accumulating the total earnings.  Add one to the start of the loop because we have
        'already captured the first year earnings when we initialized the variables
        For intAgeCount = CInt(txtAge.Text) + 1 To CInt(txtRetirementAge.Text)
            dblNewSalary = dblNewSalary + (dblNewSalary * CDbl(txtAnnualRaise.Text) / 100)
            dblEarnings = dblEarnings + dblNewSalary
        Next

        'Output the result
        lblOutput.Text = txtName.Text & " will earn " & FormatCurrency(dblEarnings) & " before retiring."

    End Sub

    Private Sub txtAge_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAge.Leave
        'Make sure that the user enters a valid number for Age
        If Not (IsNumeric(txtAge.Text)) Then
            MsgBox("Please enter your age in years as a whole number")
            txtAge.Text = ""
            txtAge.Focus()
        End If
    End Sub

    Private Sub txtCurrentSalary_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCurrentSalary.Leave
        'Make sure the user enters a valid number for salary
        If Not (IsNumeric(txtCurrentSalary.Text)) Then
            MsgBox("Please enter your current salary")
            txtCurrentSalary.Text = ""
            txtCurrentSalary.Focus()
        Else
            txtCurrentSalary.Text = FormatCurrency(txtCurrentSalary.Text)
        End If
    End Sub

    Private Sub txtAnnualRaise_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAnnualRaise.Leave
        'Make sure the user enters a valid number for Annual Raise Percent
        If Not (IsNumeric(txtAnnualRaise.Text)) Then
            MsgBox("Please enter your annual raise amount" & vbCrLf & "  *note: 5 = 5%, .05 = .05%")
            txtAnnualRaise.Text = ""
            txtAnnualRaise.Focus()
        End If


    End Sub

    Private Sub txtRetirementAge_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetirementAge.Leave
        'Make sure the user enters a valid number for retirement age
        If Not (IsNumeric(txtRetirementAge.Text)) Then
            MsgBox("Please enter your retirement age as a whole number")
            txtRetirementAge.Text = ""
            txtRetirementAge.Focus()
        End If

    End Sub
End Class
