Public Class InstructionsForm
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 104)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "This application will produce an output file consisting of pay amounts for each e" & _
        "mployee and a weekly payroll summary amount.  The program requires an input file" & _
        " for the employee records, consisting of employee number, employee name, and pay" & _
        " rate and an input file for the time clock records consisting of  employee numbe" & _
        "r, time in and time out."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 136)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 72)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Select the employee  file, time clock file, and the desired output file from the " & _
        "'Locate' menu.  Then click 'Calculate' on the File menu.  The payroll will be ca" & _
        "lculated and saved to the output file selected."
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(112, 224)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        '
        'InstructionsForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "InstructionsForm"
        Me.Text = "Instructions"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
