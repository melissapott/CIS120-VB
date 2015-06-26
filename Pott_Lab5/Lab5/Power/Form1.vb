Public Class Form1
    Inherits System.Windows.Forms.Form

    'Middle Value
    'Author:    Melissa Pott
    'Class:     CIS120
    'Assignment:    Lab #5
    'Instructor:    Loay Alnaji
    ' This program will calculate the value of the first user entered number to the power of the second user entered number

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
    Friend WithEvents txtArgument1 As System.Windows.Forms.TextBox
    Friend WithEvents txtArgument2 As System.Windows.Forms.TextBox
    Friend WithEvents btnCallPower As System.Windows.Forms.Button
    Friend WithEvents lblArgument1 As System.Windows.Forms.Label
    Friend WithEvents lblArgument2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtArgument1 = New System.Windows.Forms.TextBox
        Me.txtArgument2 = New System.Windows.Forms.TextBox
        Me.btnCallPower = New System.Windows.Forms.Button
        Me.lblArgument1 = New System.Windows.Forms.Label
        Me.lblArgument2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtArgument1
        '
        Me.txtArgument1.Location = New System.Drawing.Point(48, 64)
        Me.txtArgument1.Name = "txtArgument1"
        Me.txtArgument1.Size = New System.Drawing.Size(64, 20)
        Me.txtArgument1.TabIndex = 0
        Me.txtArgument1.Text = ""
        '
        'txtArgument2
        '
        Me.txtArgument2.Location = New System.Drawing.Point(152, 64)
        Me.txtArgument2.Name = "txtArgument2"
        Me.txtArgument2.Size = New System.Drawing.Size(64, 20)
        Me.txtArgument2.TabIndex = 1
        Me.txtArgument2.Text = ""
        '
        'btnCallPower
        '
        Me.btnCallPower.Location = New System.Drawing.Point(64, 104)
        Me.btnCallPower.Name = "btnCallPower"
        Me.btnCallPower.Size = New System.Drawing.Size(128, 23)
        Me.btnCallPower.TabIndex = 2
        Me.btnCallPower.Text = "Call Power"
        '
        'lblArgument1
        '
        Me.lblArgument1.Location = New System.Drawing.Point(48, 40)
        Me.lblArgument1.Name = "lblArgument1"
        Me.lblArgument1.Size = New System.Drawing.Size(64, 23)
        Me.lblArgument1.TabIndex = 3
        Me.lblArgument1.Text = "Argument1"
        '
        'lblArgument2
        '
        Me.lblArgument2.Location = New System.Drawing.Point(152, 40)
        Me.lblArgument2.Name = "lblArgument2"
        Me.lblArgument2.Size = New System.Drawing.Size(64, 23)
        Me.lblArgument2.TabIndex = 4
        Me.lblArgument2.Text = "Argument2"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(256, 182)
        Me.Controls.Add(Me.lblArgument2)
        Me.Controls.Add(Me.lblArgument1)
        Me.Controls.Add(Me.btnCallPower)
        Me.Controls.Add(Me.txtArgument2)
        Me.Controls.Add(Me.txtArgument1)
        Me.Name = "Form1"
        Me.Text = "Power"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCallPower_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCallPower.Click
        Dim dblArgument1, dblArgument2 As Double

        If IsNumeric(txtArgument1.Text) Then 'make sure that the user entered a number for the 1st value
            dblArgument1 = txtArgument1.Text

            If IsNumeric(txtArgument2.Text) Then 'make sure the user entered a number for the 2nd value
                dblArgument2 = txtArgument2.Text
                Call Power(dblArgument1, dblArgument2) 'we have 2 numbers, so we can call the Power sub
            Else
                MsgBox("Please enter a number for Argument 2")
            End If
        Else
            MsgBox("Please enter a number for Argument 1")
        End If
    End Sub
    Private Sub Power(ByVal dblArgument1, ByVal dblArgument2)
        MsgBox(dblArgument1 ^ dblArgument2) 'calculate the value output the result in a message box
    End Sub
End Class
