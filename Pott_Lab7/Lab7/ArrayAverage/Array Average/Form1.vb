Public Class Form1
    Inherits System.Windows.Forms.Form

    'Array Average
    'Author:        Melissa Pott
    'Class:         CIS120
    'Assignment:    Lab #7
    'Instructor:    Loay Alnaji
    'This program will read generate and display 10 random numbers, determine the average
    'of those 10 numbers, and display any numbers higher than the average

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
    Friend WithEvents lstResults As System.Windows.Forms.ListBox
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lstResults = New System.Windows.Forms.ListBox
        Me.btnGenerate = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lstResults
        '
        Me.lstResults.Location = New System.Drawing.Point(168, 40)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(232, 407)
        Me.lstResults.TabIndex = 0
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(24, 32)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(96, 40)
        Me.btnGenerate.TabIndex = 1
        Me.btnGenerate.Text = "Generate Values"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(432, 526)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.lstResults)
        Me.Name = "Form1"
        Me.Text = "Array Average"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
        Dim intRandNums(9) As Integer
        Dim i, intTotal As Integer
        Dim dblAverage As Double


        lstResults.Items.Clear() 'clear out any old data

        'generate random numbers and store them in the array
        For i = 0 To 9
            intRandNums(i) = 1000 * Rnd()
        Next

        Array.Sort(intRandNums) 'sort the array

        'add each of the numers to the list box and accumulate the total
        For i = 0 To 9
            lstResults.Items.Add(intRandNums(i))
            intTotal = intTotal + intRandNums(i)
        Next

        dblAverage = intTotal / 10 'calculate the average
        lstResults.Items.Add("") 'just a blank line
        lstResults.Items.Add("Average is " & dblAverage) 'display the average
        lstResults.Items.Add("") 'just another blank line
        lstResults.Items.Add("Greater than average:")

        'iterate through the array again, and find the items above the average
        For i = 0 To 9
            If intRandNums(i) > dblAverage Then
                lstResults.Items.Add(intRandNums(i))
            End If
        Next
    End Sub
End Class
