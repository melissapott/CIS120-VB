'Author:        Melissa Pott
'Date:          March 5, 2005
'Class:         CIS120
'Assignment:    Week 1 Lab
'Title:         Excuse-o-Rama

'This program will change the text in lblExcuses when each of 4 buttons is clicked


Public Class Form1
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
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents btnExcuse1 As System.Windows.Forms.Button
    Friend WithEvents btnExcuse2 As System.Windows.Forms.Button
    Friend WithEvents btnExcuse3 As System.Windows.Forms.Button
    Friend WithEvents btnExcuse4 As System.Windows.Forms.Button
    Friend WithEvents lblExcuses As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label
        Me.btnExcuse1 = New System.Windows.Forms.Button
        Me.btnExcuse2 = New System.Windows.Forms.Button
        Me.btnExcuse3 = New System.Windows.Forms.Button
        Me.btnExcuse4 = New System.Windows.Forms.Button
        Me.lblExcuses = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Red
        Me.lblTitle.Location = New System.Drawing.Point(16, 16)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(248, 32)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Why did I not submit my homework on time?"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnExcuse1
        '
        Me.btnExcuse1.Location = New System.Drawing.Point(48, 96)
        Me.btnExcuse1.Name = "btnExcuse1"
        Me.btnExcuse1.TabIndex = 1
        Me.btnExcuse1.Text = "Excuse 1"
        '
        'btnExcuse2
        '
        Me.btnExcuse2.Location = New System.Drawing.Point(160, 96)
        Me.btnExcuse2.Name = "btnExcuse2"
        Me.btnExcuse2.TabIndex = 2
        Me.btnExcuse2.Text = "Excuse 2"
        '
        'btnExcuse3
        '
        Me.btnExcuse3.Location = New System.Drawing.Point(48, 136)
        Me.btnExcuse3.Name = "btnExcuse3"
        Me.btnExcuse3.TabIndex = 3
        Me.btnExcuse3.Text = "Excuse 3"
        '
        'btnExcuse4
        '
        Me.btnExcuse4.Location = New System.Drawing.Point(160, 136)
        Me.btnExcuse4.Name = "btnExcuse4"
        Me.btnExcuse4.TabIndex = 4
        Me.btnExcuse4.Text = "Excuse 4"
        '
        'lblExcuses
        '
        Me.lblExcuses.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblExcuses.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExcuses.Location = New System.Drawing.Point(24, 200)
        Me.lblExcuses.Name = "lblExcuses"
        Me.lblExcuses.Size = New System.Drawing.Size(248, 40)
        Me.lblExcuses.TabIndex = 5
        Me.lblExcuses.Text = "Click a button for an excuse...."
        Me.lblExcuses.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.lblExcuses)
        Me.Controls.Add(Me.btnExcuse4)
        Me.Controls.Add(Me.btnExcuse3)
        Me.Controls.Add(Me.btnExcuse2)
        Me.Controls.Add(Me.btnExcuse1)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "Form1"
        Me.Text = "Excuse-o-Rama"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnExcuse1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcuse1.Click
        lblExcuses.Text = "It was so good the Smithsonian put it on display."
    End Sub

    Private Sub btnExcuse2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcuse2.Click
        lblExcuses.Text = "It was trampled by a herd of rare urban rhinos."
    End Sub

    Private Sub btnExcuse3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcuse3.Click
        lblExcuses.Text = "Aliens abducted it for use in strange experiments."
    End Sub

    Private Sub btnExcuse4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcuse4.Click
        lblExcuses.Text = "I accidently sold it on EBay."
    End Sub
End Class
