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
    Friend WithEvents FirstButton As System.Windows.Forms.Button
    Friend WithEvents PreviousButton As System.Windows.Forms.Button
    Friend WithEvents NextButton As System.Windows.Forms.Button
    Friend WithEvents LastButton As System.Windows.Forms.Button
    Friend WithEvents AxDataGrid1 As AxMSDataGridLib.AxDataGrid
    Public WithEvents Freight As System.Windows.Forms.TextBox
    Public WithEvents Freight_Label As System.Windows.Forms.Label
    Public WithEvents Orderdate_Label As System.Windows.Forms.Label
    Public WithEvents OrderDate As System.Windows.Forms.TextBox
    Public WithEvents ShipVia As System.Windows.Forms.TextBox
    Public WithEvents CustomerID As System.Windows.Forms.TextBox
    Public WithEvents CustomerID_Label As System.Windows.Forms.Label
    Public WithEvents ShipVia_Label As System.Windows.Forms.Label
    Public WithEvents OrderId_Label As System.Windows.Forms.Label
    Public WithEvents OrderID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(TestForm))
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.FirstButton = New System.Windows.Forms.Button()
        Me.PreviousButton = New System.Windows.Forms.Button()
        Me.NextButton = New System.Windows.Forms.Button()
        Me.LastButton = New System.Windows.Forms.Button()
        Me.AxDataGrid1 = New AxMSDataGridLib.AxDataGrid()
        Me.Freight = New System.Windows.Forms.TextBox()
        Me.Freight_Label = New System.Windows.Forms.Label()
        Me.Orderdate_Label = New System.Windows.Forms.Label()
        Me.OrderDate = New System.Windows.Forms.TextBox()
        Me.ShipVia = New System.Windows.Forms.TextBox()
        Me.CustomerID = New System.Windows.Forms.TextBox()
        Me.CustomerID_Label = New System.Windows.Forms.Label()
        Me.ShipVia_Label = New System.Windows.Forms.Label()
        Me.OrderId_Label = New System.Windows.Forms.Label()
        Me.OrderID = New System.Windows.Forms.TextBox()
        CType(Me.AxDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ExitButton
        '
        Me.ExitButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExitButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ExitButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExitButton.Image = CType(resources.GetObject("ExitButton.Image"), System.Drawing.Bitmap)
        Me.ExitButton.Location = New System.Drawing.Point(432, 308)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ExitButton.Size = New System.Drawing.Size(72, 48)
        Me.ExitButton.TabIndex = 14
        Me.ExitButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FirstButton
        '
        Me.FirstButton.Image = CType(resources.GetObject("FirstButton.Image"), System.Drawing.Bitmap)
        Me.FirstButton.Location = New System.Drawing.Point(16, 308)
        Me.FirstButton.Name = "FirstButton"
        Me.FirstButton.Size = New System.Drawing.Size(72, 48)
        Me.FirstButton.TabIndex = 7
        '
        'PreviousButton
        '
        Me.PreviousButton.Image = CType(resources.GetObject("PreviousButton.Image"), System.Drawing.Bitmap)
        Me.PreviousButton.Location = New System.Drawing.Point(96, 308)
        Me.PreviousButton.Name = "PreviousButton"
        Me.PreviousButton.Size = New System.Drawing.Size(72, 48)
        Me.PreviousButton.TabIndex = 8
        '
        'NextButton
        '
        Me.NextButton.Image = CType(resources.GetObject("NextButton.Image"), System.Drawing.Bitmap)
        Me.NextButton.Location = New System.Drawing.Point(168, 308)
        Me.NextButton.Name = "NextButton"
        Me.NextButton.Size = New System.Drawing.Size(72, 48)
        Me.NextButton.TabIndex = 9
        '
        'LastButton
        '
        Me.LastButton.Image = CType(resources.GetObject("LastButton.Image"), System.Drawing.Bitmap)
        Me.LastButton.Location = New System.Drawing.Point(248, 308)
        Me.LastButton.Name = "LastButton"
        Me.LastButton.Size = New System.Drawing.Size(72, 48)
        Me.LastButton.TabIndex = 10
        '
        'AxDataGrid1
        '
        Me.AxDataGrid1.DataSource = Nothing
        Me.AxDataGrid1.Location = New System.Drawing.Point(19, 178)
        Me.AxDataGrid1.Name = "AxDataGrid1"
        Me.AxDataGrid1.OcxState = CType(resources.GetObject("AxDataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDataGrid1.Size = New System.Drawing.Size(488, 120)
        Me.AxDataGrid1.TabIndex = 68
        Me.AxDataGrid1.Tag = "Orders Form"
        '
        'Freight
        '
        Me.Freight.AcceptsReturn = True
        Me.Freight.AutoSize = False
        Me.Freight.BackColor = System.Drawing.SystemColors.Window
        Me.Freight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Freight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Freight.Location = New System.Drawing.Point(128, 112)
        Me.Freight.MaxLength = 10
        Me.Freight.Name = "Freight"
        Me.Freight.Size = New System.Drawing.Size(84, 25)
        Me.Freight.TabIndex = 67
        Me.Freight.Text = ""
        '
        'Freight_Label
        '
        Me.Freight_Label.BackColor = System.Drawing.SystemColors.Control
        Me.Freight_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.Freight_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Freight_Label.ForeColor = System.Drawing.Color.Blue
        Me.Freight_Label.Location = New System.Drawing.Point(16, 112)
        Me.Freight_Label.Name = "Freight_Label"
        Me.Freight_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Freight_Label.Size = New System.Drawing.Size(112, 25)
        Me.Freight_Label.TabIndex = 75
        Me.Freight_Label.Text = "Freight"
        Me.Freight_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Orderdate_Label
        '
        Me.Orderdate_Label.BackColor = System.Drawing.SystemColors.Control
        Me.Orderdate_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.Orderdate_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Orderdate_Label.ForeColor = System.Drawing.Color.Blue
        Me.Orderdate_Label.Location = New System.Drawing.Point(16, 144)
        Me.Orderdate_Label.Name = "Orderdate_Label"
        Me.Orderdate_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Orderdate_Label.Size = New System.Drawing.Size(112, 25)
        Me.Orderdate_Label.TabIndex = 73
        Me.Orderdate_Label.Text = "Order Date"
        Me.Orderdate_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderDate
        '
        Me.OrderDate.AcceptsReturn = True
        Me.OrderDate.AutoSize = False
        Me.OrderDate.BackColor = System.Drawing.SystemColors.Window
        Me.OrderDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.OrderDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.OrderDate.Location = New System.Drawing.Point(128, 144)
        Me.OrderDate.MaxLength = 10
        Me.OrderDate.Name = "OrderDate"
        Me.OrderDate.Size = New System.Drawing.Size(88, 25)
        Me.OrderDate.TabIndex = 64
        Me.OrderDate.Text = ""
        '
        'ShipVia
        '
        Me.ShipVia.AcceptsReturn = True
        Me.ShipVia.AutoSize = False
        Me.ShipVia.BackColor = System.Drawing.SystemColors.Window
        Me.ShipVia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.ShipVia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ShipVia.Location = New System.Drawing.Point(128, 80)
        Me.ShipVia.MaxLength = 2
        Me.ShipVia.Name = "ShipVia"
        Me.ShipVia.Size = New System.Drawing.Size(84, 25)
        Me.ShipVia.TabIndex = 66
        Me.ShipVia.Text = ""
        '
        'CustomerID
        '
        Me.CustomerID.AcceptsReturn = True
        Me.CustomerID.AutoSize = False
        Me.CustomerID.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerID.Location = New System.Drawing.Point(128, 48)
        Me.CustomerID.MaxLength = 3
        Me.CustomerID.Name = "CustomerID"
        Me.CustomerID.Size = New System.Drawing.Size(84, 25)
        Me.CustomerID.TabIndex = 65
        Me.CustomerID.Text = ""
        '
        'CustomerID_Label
        '
        Me.CustomerID_Label.BackColor = System.Drawing.SystemColors.Control
        Me.CustomerID_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.CustomerID_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.CustomerID_Label.ForeColor = System.Drawing.Color.Blue
        Me.CustomerID_Label.Location = New System.Drawing.Point(16, 48)
        Me.CustomerID_Label.Name = "CustomerID_Label"
        Me.CustomerID_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CustomerID_Label.Size = New System.Drawing.Size(112, 25)
        Me.CustomerID_Label.TabIndex = 72
        Me.CustomerID_Label.Text = "Customer Id"
        Me.CustomerID_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ShipVia_Label
        '
        Me.ShipVia_Label.BackColor = System.Drawing.SystemColors.Control
        Me.ShipVia_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.ShipVia_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ShipVia_Label.ForeColor = System.Drawing.Color.Blue
        Me.ShipVia_Label.Location = New System.Drawing.Point(16, 80)
        Me.ShipVia_Label.Name = "ShipVia_Label"
        Me.ShipVia_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShipVia_Label.Size = New System.Drawing.Size(112, 25)
        Me.ShipVia_Label.TabIndex = 71
        Me.ShipVia_Label.Text = "Ship Via"
        Me.ShipVia_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderId_Label
        '
        Me.OrderId_Label.BackColor = System.Drawing.SystemColors.Control
        Me.OrderId_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.OrderId_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.OrderId_Label.ForeColor = System.Drawing.Color.Blue
        Me.OrderId_Label.Location = New System.Drawing.Point(16, 16)
        Me.OrderId_Label.Name = "OrderId_Label"
        Me.OrderId_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OrderId_Label.Size = New System.Drawing.Size(112, 25)
        Me.OrderId_Label.TabIndex = 70
        Me.OrderId_Label.Text = "Order Id"
        Me.OrderId_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderID
        '
        Me.OrderID.AcceptsReturn = True
        Me.OrderID.AutoSize = False
        Me.OrderID.BackColor = System.Drawing.SystemColors.Window
        Me.OrderID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.OrderID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.OrderID.Location = New System.Drawing.Point(128, 16)
        Me.OrderID.MaxLength = 5
        Me.OrderID.Name = "OrderID"
        Me.OrderID.Size = New System.Drawing.Size(84, 26)
        Me.OrderID.TabIndex = 63
        Me.OrderID.Text = ""
        '
        'TestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(530, 368)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxDataGrid1, Me.Freight, Me.Freight_Label, Me.Orderdate_Label, Me.OrderDate, Me.ShipVia, Me.CustomerID, Me.CustomerID_Label, Me.ShipVia_Label, Me.OrderId_Label, Me.OrderID, Me.LastButton, Me.NextButton, Me.PreviousButton, Me.FirstButton, Me.ExitButton})
        Me.ForeColor = System.Drawing.Color.Blue
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TestForm"
        Me.Tag = "Orders Form"
        Me.Text = "TestForm"
        CType(Me.AxDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim HA As New DynamicComponents.HelpAuthority()
    Dim CN As New ADODB.Connection()
    Dim oOrders As New ADODB.Recordset()
    Dim oOrderDetails As New ADODB.Recordset()
    Dim oAccess As New Access.Application()
    Dim DAO_DBEngine As New DAO.DBEngine()

    'Press F12 to get help to any control on your form

    Private Sub TestForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'establish DSN
        oAccess.DBEngine.RegisterDatabase("DCDM_Nwind", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & VB6.GetPath & "\Nwind.mdb")
        CN.Open("DSN=DCDM_NWind")
        oOrders.Open("Orders", CN, oOrders.CursorType.adOpenKeyset, oOrders.LockType.adLockOptimistic)
        oOrderDetails.Open("OrderDetails", CN, oOrderDetails.CursorType.adOpenKeyset, oOrderDetails.LockType.adLockOptimistic)
        PopulateDate()
        Me.AxDataGrid1.DataSource = oOrderDetails
        HA.PrepareHelp(CN, Me, Me.AxDataGrid1)
    End Sub


    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        oAccess.Quit()
        Me.Close()

    End Sub

    Private Sub PopulateDate()
        Me.OrderID.Text = oOrders.Fields("OrderID").Value
        Me.OrderDate.Text = oOrders.Fields("OrderDate").Value
        Me.CustomerID.Text = oOrders.Fields("CustomerID").Value
        Me.ShipVia.Text = oOrders.Fields("ShipVia").Value
        Me.Freight.Text = oOrders.Fields("Freight").Value
        oOrderDetails.Filter = "OrderId ='" + Me.OrderID.Text + "'"
    End Sub

    Private Sub FirstButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstButton.Click
        oOrders.MoveFirst()
        PopulateDate()
    End Sub

    Private Sub PreviousButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PreviousButton.Click
        On Error Resume Next

        oOrders.MovePrevious()
        PopulateDate()
    End Sub

    Private Sub NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextButton.Click
        On Error Resume Next

        oOrders.MoveNext()
        PopulateDate()
    End Sub

    Private Sub LastButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LastButton.Click
        oOrders.MoveLast()
        PopulateDate()
    End Sub
End Class
