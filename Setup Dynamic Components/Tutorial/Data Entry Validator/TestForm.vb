Public Class TestForm
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
    Public WithEvents ExitButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents FirstButton As System.Windows.Forms.Button
    Friend WithEvents PreviousButton As System.Windows.Forms.Button
    Friend WithEvents NextButton As System.Windows.Forms.Button
    Friend WithEvents LastButton As System.Windows.Forms.Button
    Public WithEvents CustomerName As System.Windows.Forms.TextBox
    Public WithEvents CustomerID As System.Windows.Forms.TextBox
    Public WithEvents Address As System.Windows.Forms.TextBox
    Public WithEvents MaxDebit As System.Windows.Forms.TextBox
    Public WithEvents LastDeal As System.Windows.Forms.TextBox
    Public WithEvents Phone As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(TestForm))
        Me.CustomerName = New System.Windows.Forms.TextBox()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.CustomerID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Address = New System.Windows.Forms.TextBox()
        Me.Phone = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.MaxDebit = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.LastDeal = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.FirstButton = New System.Windows.Forms.Button()
        Me.PreviousButton = New System.Windows.Forms.Button()
        Me.NextButton = New System.Windows.Forms.Button()
        Me.LastButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'CustomerName
        '
        Me.CustomerName.AcceptsReturn = True
        Me.CustomerName.AutoSize = False
        Me.CustomerName.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerName.Location = New System.Drawing.Point(123, 40)
        Me.CustomerName.MaxLength = 50
        Me.CustomerName.Name = "CustomerName"
        Me.CustomerName.Size = New System.Drawing.Size(263, 25)
        Me.CustomerName.TabIndex = 1
        Me.CustomerName.Text = ""
        '
        'ExitButton
        '
        Me.ExitButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExitButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ExitButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExitButton.Image = CType(resources.GetObject("ExitButton.Image"), System.Drawing.Bitmap)
        Me.ExitButton.Location = New System.Drawing.Point(592, 248)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ExitButton.Size = New System.Drawing.Size(72, 48)
        Me.ExitButton.TabIndex = 14
        Me.ExitButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CustomerID
        '
        Me.CustomerID.AcceptsReturn = True
        Me.CustomerID.AutoSize = False
        Me.CustomerID.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerID.Location = New System.Drawing.Point(124, 6)
        Me.CustomerID.MaxLength = 4
        Me.CustomerID.Name = "CustomerID"
        Me.CustomerID.Size = New System.Drawing.Size(84, 26)
        Me.CustomerID.TabIndex = 0
        Me.CustomerID.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 24)
        Me.Label1.TabIndex = 60
        Me.Label1.Text = "Customer ID"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label2.Location = New System.Drawing.Point(8, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 24)
        Me.Label2.TabIndex = 61
        Me.Label2.Text = "Customer Name"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label3.Location = New System.Drawing.Point(8, 73)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 24)
        Me.Label3.TabIndex = 62
        Me.Label3.Text = "Address"
        '
        'Address
        '
        Me.Address.AcceptsReturn = True
        Me.Address.AutoSize = False
        Me.Address.BackColor = System.Drawing.SystemColors.Window
        Me.Address.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Address.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Address.Location = New System.Drawing.Point(123, 73)
        Me.Address.MaxLength = 50
        Me.Address.Name = "Address"
        Me.Address.Size = New System.Drawing.Size(263, 25)
        Me.Address.TabIndex = 2
        Me.Address.Text = ""
        '
        'Phone
        '
        Me.Phone.AcceptsReturn = True
        Me.Phone.AutoSize = False
        Me.Phone.BackColor = System.Drawing.SystemColors.Window
        Me.Phone.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Phone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Phone.Location = New System.Drawing.Point(122, 107)
        Me.Phone.MaxLength = 12
        Me.Phone.Name = "Phone"
        Me.Phone.Size = New System.Drawing.Size(134, 25)
        Me.Phone.TabIndex = 3
        Me.Phone.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label4.Location = New System.Drawing.Point(6, 107)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 24)
        Me.Label4.TabIndex = 64
        Me.Label4.Text = "Phone"
        '
        'MaxDebit
        '
        Me.MaxDebit.AcceptsReturn = True
        Me.MaxDebit.AutoSize = False
        Me.MaxDebit.BackColor = System.Drawing.SystemColors.Window
        Me.MaxDebit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.MaxDebit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.MaxDebit.Location = New System.Drawing.Point(123, 141)
        Me.MaxDebit.MaxLength = 12
        Me.MaxDebit.Name = "MaxDebit"
        Me.MaxDebit.Size = New System.Drawing.Size(133, 25)
        Me.MaxDebit.TabIndex = 4
        Me.MaxDebit.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label5.Location = New System.Drawing.Point(8, 144)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 24)
        Me.Label5.TabIndex = 66
        Me.Label5.Text = "Max Debit"
        '
        'LastDeal
        '
        Me.LastDeal.AcceptsReturn = True
        Me.LastDeal.AutoSize = False
        Me.LastDeal.BackColor = System.Drawing.SystemColors.Window
        Me.LastDeal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.LastDeal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LastDeal.Location = New System.Drawing.Point(123, 172)
        Me.LastDeal.MaxLength = 10
        Me.LastDeal.Name = "LastDeal"
        Me.LastDeal.Size = New System.Drawing.Size(133, 25)
        Me.LastDeal.TabIndex = 5
        Me.LastDeal.Text = ""
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label6.Location = New System.Drawing.Point(8, 176)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 24)
        Me.Label6.TabIndex = 68
        Me.Label6.Text = "Last Deal"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label8.Location = New System.Drawing.Point(400, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(248, 24)
        Me.Label8.TabIndex = 72
        Me.Label8.Text = "must be numeric characters"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label9.Location = New System.Drawing.Point(402, 40)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(248, 24)
        Me.Label9.TabIndex = 73
        Me.Label9.Text = "must be alphabetic characters"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label10.Location = New System.Drawing.Point(400, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(288, 24)
        Me.Label10.TabIndex = 74
        Me.Label10.Text = "may be alphabetic or numeic characters"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label11.Location = New System.Drawing.Point(400, 144)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(288, 24)
        Me.Label11.TabIndex = 75
        Me.Label11.Text = "formated to number with 2 decimal digit"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label12.Location = New System.Drawing.Point(400, 107)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(288, 24)
        Me.Label12.TabIndex = 76
        Me.Label12.Text = "may be alphabetic or numeic characters"
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label13.Location = New System.Drawing.Point(400, 176)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(248, 24)
        Me.Label13.TabIndex = 77
        Me.Label13.Text = "must be accepted date"
        '
        'FirstButton
        '
        Me.FirstButton.Location = New System.Drawing.Point(24, 249)
        Me.FirstButton.Name = "FirstButton"
        Me.FirstButton.Size = New System.Drawing.Size(72, 48)
        Me.FirstButton.TabIndex = 7
        Me.FirstButton.Text = "First"
        '
        'PreviousButton
        '
        Me.PreviousButton.Location = New System.Drawing.Point(112, 249)
        Me.PreviousButton.Name = "PreviousButton"
        Me.PreviousButton.Size = New System.Drawing.Size(72, 48)
        Me.PreviousButton.TabIndex = 8
        Me.PreviousButton.Text = "Previous"
        '
        'NextButton
        '
        Me.NextButton.Location = New System.Drawing.Point(208, 249)
        Me.NextButton.Name = "NextButton"
        Me.NextButton.Size = New System.Drawing.Size(72, 48)
        Me.NextButton.TabIndex = 9
        Me.NextButton.Text = "Next"
        '
        'LastButton
        '
        Me.LastButton.Location = New System.Drawing.Point(296, 249)
        Me.LastButton.Name = "LastButton"
        Me.LastButton.Size = New System.Drawing.Size(72, 48)
        Me.LastButton.TabIndex = 10
        Me.LastButton.Text = "last"
        '
        'TestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(682, 304)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.LastButton, Me.NextButton, Me.PreviousButton, Me.FirstButton, Me.Label13, Me.Label12, Me.Label11, Me.Label10, Me.Label9, Me.Label8, Me.LastDeal, Me.Label6, Me.MaxDebit, Me.Label5, Me.Phone, Me.Label4, Me.Address, Me.Label3, Me.Label2, Me.Label1, Me.CustomerName, Me.ExitButton, Me.CustomerID})
        Me.ForeColor = System.Drawing.Color.Blue
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TestForm"
        Me.Tag = "Orders Form"
        Me.Text = "TestForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim DV As New DynamicComponents.DataEntryValidator()
    Dim CN As New ADODB.Connection()
    Dim oCust As New ADODB.Recordset()
    Dim oAccess As New Access.Application()
    Dim DAO_DBEngine As New DAO.DBEngine()

    Private Sub TestForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'establish DSN
        oAccess.DBEngine.RegisterDatabase("DCDM_Nwind", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & VB6.GetPath & "\Nwind.mdb")
        CN.Open("DSN=DCDM_NWind")
        oCust.Open("Customers", CN, oCust.CursorType.adOpenKeyset, oCust.LockType.adLockOptimistic)
        PopulateDate()
        DV.InitForm(Me)                              'must be your first assignment , an error occurs if not
        DV.NumericFields("CustomerId")               'CustomerId must be numeric characters(0123456789)
        DV.AlphabeticFields("CustomerName")          'CustomerName must be alphabetic characters (abcdefghijklmnopqrstuvwzyzABCDEFGHIJKLMNOPQRSTUVWXYZ) 
        DV.FirstCharOfWordsFields("CustomerName")    'First charecter of every word will be in uooer case
        DV.AlphaNumericFields("Address")             'Address must be numeric or alphabetic characters (0123456789abcdefghijklmnopqrstuvwzyzABCDEFGHIJKLMNOPQRSTUVWXYZ)  
        DV.FirstCharOnlyFields("Address")            'First charecter of first word only will be in uooer case 
        DV.AlphaNumericFields("phone")                    'Phone must be numeric characters(0123456789)
        DV.DecimalFields("MaxDebit")                 'MaxDebit must be decimal characters(0123456789 & .) 
        DV.DecimalPlaces(2)                          'MaxDebit will be formatted with 2 decimal digits
        DV.DateFields("LastDeal")                    'LastDeal must be accepted date(0123456789-\/)
    End Sub


    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        oAccess.Quit()
        Me.Close()

    End Sub

    Private Sub PopulateDate()
        Me.CustomerID.Text = oCust.Fields("CustomerID").Value
        Me.CustomerName.Text = oCust.Fields("CustomerName").Value
        Me.Address.Text = oCust.Fields("Address").Value
        Me.Phone.Text = oCust.Fields("Phone").Value
        Me.MaxDebit.Text = oCust.Fields("MaxDebit").Value
        Me.LastDeal.Text = oCust.Fields("LastDeal").Value

    End Sub

    Private Sub FirstButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstButton.Click
        oCust.MoveFirst()
        PopulateDate()
    End Sub

    Private Sub PreviousButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PreviousButton.Click
        On Error Resume Next

        oCust.MovePrevious()
        PopulateDate()
    End Sub

    Private Sub NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextButton.Click
        On Error Resume Next

        oCust.MoveNext()
        PopulateDate()
    End Sub

    Private Sub LastButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LastButton.Click
        oCust.MoveLast()
        PopulateDate()
    End Sub
End Class
