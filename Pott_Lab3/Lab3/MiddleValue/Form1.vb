Public Class Form1
    Inherits System.Windows.Forms.Form
    'Middle Value
    'Author:    Melissa Pott
    'Class:     CIS120
    'Assignment:    Lab #3
    'Instructor:    Loay Alnaji
    ' This program will determine the middle of three user-entered values

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
    Friend WithEvents txtFirstValue As System.Windows.Forms.TextBox
    Friend WithEvents txtSecondValue As System.Windows.Forms.TextBox
    Friend WithEvents txtThirdValue As System.Windows.Forms.TextBox
    Friend WithEvents btnCalcMiddle As System.Windows.Forms.Button
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents txtExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtFirstValue = New System.Windows.Forms.TextBox
        Me.txtSecondValue = New System.Windows.Forms.TextBox
        Me.txtThirdValue = New System.Windows.Forms.TextBox
        Me.btnCalcMiddle = New System.Windows.Forms.Button
        Me.txtResult = New System.Windows.Forms.TextBox
        Me.txtExit = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtFirstValue
        '
        Me.txtFirstValue.Location = New System.Drawing.Point(16, 24)
        Me.txtFirstValue.Name = "txtFirstValue"
        Me.txtFirstValue.Size = New System.Drawing.Size(64, 20)
        Me.txtFirstValue.TabIndex = 0
        Me.txtFirstValue.Text = ""
        '
        'txtSecondValue
        '
        Me.txtSecondValue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSecondValue.Location = New System.Drawing.Point(104, 24)
        Me.txtSecondValue.Name = "txtSecondValue"
        Me.txtSecondValue.Size = New System.Drawing.Size(64, 20)
        Me.txtSecondValue.TabIndex = 1
        Me.txtSecondValue.Text = ""
        '
        'txtThirdValue
        '
        Me.txtThirdValue.Location = New System.Drawing.Point(192, 24)
        Me.txtThirdValue.Name = "txtThirdValue"
        Me.txtThirdValue.Size = New System.Drawing.Size(64, 20)
        Me.txtThirdValue.TabIndex = 2
        Me.txtThirdValue.Text = ""
        '
        'btnCalcMiddle
        '
        Me.btnCalcMiddle.Location = New System.Drawing.Point(32, 72)
        Me.btnCalcMiddle.Name = "btnCalcMiddle"
        Me.btnCalcMiddle.Size = New System.Drawing.Size(208, 24)
        Me.btnCalcMiddle.TabIndex = 3
        Me.btnCalcMiddle.Text = "Determine Middle Value"
        '
        'txtResult
        '
        Me.txtResult.Location = New System.Drawing.Point(96, 160)
        Me.txtResult.Name = "txtResult"
        Me.txtResult.TabIndex = 4
        Me.txtResult.Text = ""
        '
        'txtExit
        '
        Me.txtExit.Location = New System.Drawing.Point(192, 224)
        Me.txtExit.Name = "txtExit"
        Me.txtExit.TabIndex = 5
        Me.txtExit.Text = "Exit"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(96, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Middle Value is:"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtExit)
        Me.Controls.Add(Me.txtResult)
        Me.Controls.Add(Me.txtThirdValue)
        Me.Controls.Add(Me.txtSecondValue)
        Me.Controls.Add(Me.txtFirstValue)
        Me.Controls.Add(Me.btnCalcMiddle)
        Me.Name = "Form1"
        Me.Text = "Middle Value"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtFirstValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFirstValue.TextChanged
        txtResult.Text = ""
    End Sub

    Private Sub txtSecondValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSecondValue.TextChanged
        txtResult.Text = ""
    End Sub

    Private Sub txtThirdValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtThirdValue.TextChanged
        txtResult.Text = ""
    End Sub

    Private Sub btnCalcMiddle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalcMiddle.Click
        Dim FirstValue, SecondValue, ThirdValue As Double

        'check to make sure the user entered numbers
        If Not (IsNumeric(txtFirstValue.Text)) Or Not (IsNumeric(txtSecondValue.Text)) Or Not (IsNumeric(txtThirdValue.Text)) Then
            MsgBox("Please Enter Numeric Values")
            txtResult.Text = ""
            Exit Sub
        Else 'convert everything to a double at once so we don't have to do it for each comparison
            FirstValue = CDbl(txtFirstValue.Text)
            SecondValue = CDbl(txtSecondValue.Text)
            ThirdValue = CDbl(txtThirdValue.Text)
        End If

        'check to make sure that the same number wasn't entered more than once
        If (Val(txtFirstValue.Text) = Val(txtSecondValue.Text)) Or (Val(txtFirstValue.Text) = Val(txtThirdValue.Text)) Or (Val(txtSecondValue.Text) = Val(txtThirdValue.Text)) Then
            txtResult.Text = "Duplicate Values"
            Exit Sub
        End If

        'begin the comparison to see which one is the middle value
        If (FirstValue < SecondValue And SecondValue < ThirdValue) Or (FirstValue > SecondValue And SecondValue > ThirdValue) Then
            txtResult.Text = SecondValue
        ElseIf (FirstValue < SecondValue And FirstValue > ThirdValue) Or (FirstValue > SecondValue And FirstValue < ThirdValue) Then
            txtResult.Text = FirstValue
        Else
            txtResult.Text = ThirdValue
        End If

    End Sub

    Private Sub txtExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExit.Click
        End
    End Sub
End Class
