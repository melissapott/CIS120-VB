Public Class Form1
    Inherits System.Windows.Forms.Form
    'Sales Receipt
    'Author:        Melissa Pott
    'Class:         CIS120
    'Assignment:    Lab #2
    'Instructor:    Loay Alnaji

    'This program will accept user inputs for the amount of an item, add the item to the receipt,
    'and calculate the subtotal for all items, the sales tax based on 6.5%, and the grand total

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
    Friend WithEvents btnAddItem As System.Windows.Forms.Button
    Friend WithEvents lstReceipt As System.Windows.Forms.ListBox
    Friend WithEvents lblSalesReceipt As System.Windows.Forms.Label
    Friend WithEvents lblItemPrice As System.Windows.Forms.Label
    Friend WithEvents lblSubTotal As System.Windows.Forms.Label
    Friend WithEvents lblSalesTax As System.Windows.Forms.Label
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents txtItemInput As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents txtSubtotalOut As System.Windows.Forms.TextBox
    Friend WithEvents txtSalesTaxOut As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalOut As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtItemInput = New System.Windows.Forms.TextBox
        Me.lstReceipt = New System.Windows.Forms.ListBox
        Me.btnAddItem = New System.Windows.Forms.Button
        Me.lblItemPrice = New System.Windows.Forms.Label
        Me.lblSalesReceipt = New System.Windows.Forms.Label
        Me.lblSubTotal = New System.Windows.Forms.Label
        Me.lblSalesTax = New System.Windows.Forms.Label
        Me.lblTotal = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.txtSubtotalOut = New System.Windows.Forms.TextBox
        Me.txtSalesTaxOut = New System.Windows.Forms.TextBox
        Me.txtTotalOut = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'txtItemInput
        '
        Me.txtItemInput.Location = New System.Drawing.Point(16, 32)
        Me.txtItemInput.Name = "txtItemInput"
        Me.txtItemInput.Size = New System.Drawing.Size(88, 20)
        Me.txtItemInput.TabIndex = 0
        Me.txtItemInput.Text = ""
        '
        'lstReceipt
        '
        Me.lstReceipt.Location = New System.Drawing.Point(128, 32)
        Me.lstReceipt.Name = "lstReceipt"
        Me.lstReceipt.Size = New System.Drawing.Size(176, 108)
        Me.lstReceipt.TabIndex = 2
        '
        'btnAddItem
        '
        Me.btnAddItem.Location = New System.Drawing.Point(16, 56)
        Me.btnAddItem.Name = "btnAddItem"
        Me.btnAddItem.Size = New System.Drawing.Size(88, 23)
        Me.btnAddItem.TabIndex = 1
        Me.btnAddItem.Text = "Add to Receipt"
        '
        'lblItemPrice
        '
        Me.lblItemPrice.Location = New System.Drawing.Point(8, 8)
        Me.lblItemPrice.Name = "lblItemPrice"
        Me.lblItemPrice.Size = New System.Drawing.Size(100, 16)
        Me.lblItemPrice.TabIndex = 4
        Me.lblItemPrice.Text = "Item Price"
        Me.lblItemPrice.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSalesReceipt
        '
        Me.lblSalesReceipt.Location = New System.Drawing.Point(136, 8)
        Me.lblSalesReceipt.Name = "lblSalesReceipt"
        Me.lblSalesReceipt.Size = New System.Drawing.Size(100, 16)
        Me.lblSalesReceipt.TabIndex = 5
        Me.lblSalesReceipt.Text = "Sales Receipt"
        Me.lblSalesReceipt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSubTotal
        '
        Me.lblSubTotal.Location = New System.Drawing.Point(240, 152)
        Me.lblSubTotal.Name = "lblSubTotal"
        Me.lblSubTotal.Size = New System.Drawing.Size(100, 16)
        Me.lblSubTotal.TabIndex = 6
        Me.lblSubTotal.Text = "Sub Total"
        '
        'lblSalesTax
        '
        Me.lblSalesTax.Location = New System.Drawing.Point(240, 176)
        Me.lblSalesTax.Name = "lblSalesTax"
        Me.lblSalesTax.Size = New System.Drawing.Size(100, 16)
        Me.lblSalesTax.TabIndex = 8
        Me.lblSalesTax.Text = "Sales Tax"
        '
        'lblTotal
        '
        Me.lblTotal.Location = New System.Drawing.Point(240, 200)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(100, 16)
        Me.lblTotal.TabIndex = 10
        Me.lblTotal.Text = "Total"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(288, 248)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "Exit"
        '
        'txtSubtotalOut
        '
        Me.txtSubtotalOut.BackColor = System.Drawing.SystemColors.Control
        Me.txtSubtotalOut.Location = New System.Drawing.Point(128, 152)
        Me.txtSubtotalOut.Name = "txtSubtotalOut"
        Me.txtSubtotalOut.TabIndex = 3
        Me.txtSubtotalOut.Text = "0"
        '
        'txtSalesTaxOut
        '
        Me.txtSalesTaxOut.BackColor = System.Drawing.SystemColors.Control
        Me.txtSalesTaxOut.Location = New System.Drawing.Point(128, 176)
        Me.txtSalesTaxOut.Name = "txtSalesTaxOut"
        Me.txtSalesTaxOut.TabIndex = 4
        Me.txtSalesTaxOut.Text = "0"
        '
        'txtTotalOut
        '
        Me.txtTotalOut.BackColor = System.Drawing.SystemColors.Control
        Me.txtTotalOut.Location = New System.Drawing.Point(128, 200)
        Me.txtTotalOut.Name = "txtTotalOut"
        Me.txtTotalOut.TabIndex = 5
        Me.txtTotalOut.Text = "0"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(416, 302)
        Me.Controls.Add(Me.txtTotalOut)
        Me.Controls.Add(Me.txtSalesTaxOut)
        Me.Controls.Add(Me.txtSubtotalOut)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.lblSalesTax)
        Me.Controls.Add(Me.lblSubTotal)
        Me.Controls.Add(Me.lblSalesReceipt)
        Me.Controls.Add(Me.lblItemPrice)
        Me.Controls.Add(Me.btnAddItem)
        Me.Controls.Add(Me.lstReceipt)
        Me.Controls.Add(Me.txtItemInput)
        Me.Name = "Form1"
        Me.Text = "Sales Receipt"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddItem.Click

        Const SALES_TAX_RATE = 0.065 'set 6.5% sales tax
        Dim Item, Subtotal, SalesTax, Total As Double
        Dim ItemStr As String

        'check to see if a valid number has been entered
        If IsNumeric(txtItemInput.Text) Then
            Item = txtItemInput.Text
        Else
            MsgBox("Please enter a numeric value")
            'reset the input box for another try
            txtItemInput.Text = ""
            txtItemInput.Focus()
            Exit Sub
        End If

        'do some display formatting work
        ItemStr = FormatCurrency(Item) 'we need to have a string value to use the currency format
        lstReceipt.Items.Add(ItemStr)
        txtItemInput.Text = "" 'reset the input box

        'do some math and display the results as currency in the output text boxes
        Subtotal = txtSubtotalOut.Text + Item
        txtSubtotalOut.Text = FormatCurrency(Subtotal)
        SalesTax = Subtotal * SALES_TAX_RATE
        txtSalesTaxOut.Text = FormatCurrency(SalesTax)
        Total = Subtotal + SalesTax
        txtTotalOut.Text = FormatCurrency(Total)

        txtItemInput.Focus() 'get ready for the next entry


    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

End Class
