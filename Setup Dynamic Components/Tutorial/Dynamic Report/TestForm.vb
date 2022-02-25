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
    Friend WithEvents ViewButton As System.Windows.Forms.Button
    Friend WithEvents ExitButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ViewButton = New System.Windows.Forms.Button()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ViewButton
        '
        Me.ViewButton.Location = New System.Drawing.Point(24, 80)
        Me.ViewButton.Name = "ViewButton"
        Me.ViewButton.Size = New System.Drawing.Size(128, 40)
        Me.ViewButton.TabIndex = 0
        Me.ViewButton.Text = "View Report"
        '
        'ExitButton
        '
        Me.ExitButton.Location = New System.Drawing.Point(240, 72)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(72, 48)
        Me.ExitButton.TabIndex = 1
        Me.ExitButton.Text = "Exit"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(288, 40)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Orders List Report"
        '
        'TestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(336, 150)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.ExitButton, Me.ViewButton})
        Me.Name = "TestForm"
        Me.Text = "Test Form"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private oAccess As New Access.Application()
    Private Sub ViewButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewButton.Click
        Dim oRep As New DynamicComponents.DynamicReport()
        Dim CN As New ADODB.Connection()
        Dim oMaster As New ADODB.Recordset()
        Dim DAO_DBEngine As New DAO.DBEngine()
        Dim SqlStatment As String


        'establish DSN
        oAccess.DBEngine.RegisterDatabase("DCDR_Nwind", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & VB6.GetPath & "\Nwind.mdb")
        CN.Open("DCDR_NWind")
        SqlStatment = "select OrderID,ProductID,ProductName,UnitPrice,Quantity "
        ' you must order your data by the field you willing to groub by
        SqlStatment += "from OrderDetails where OrderId < '10270' order by OrderID"
        oMaster.Open(SqlStatment, CN, oMaster.CursorType.adOpenKeyset, oMaster.LockType.adLockOptimistic)
        oRep.InitReport(False) ' this must be youe first assignment
        oRep.SetTitle("Good Morning World")
        oRep.LogoImage("Logo.bmp", VB6.GetPath + "\")
        oRep.SetReportHeader("This is Dynamic Report v1.0", "It is powered by EgyFirst inc.", "Dynamic Components is a trade mark since 2004")
        oRep.GroubBy("OrderID", True, True)
        oRep.SetCaption("Order ID", "Product ID", "Product Name", "Unit Price", "Quantity", "Discount")
        oRep.SumFields(oMaster, "Quantity")
        oRep.ReadTheme(DynamicComponents.DynamicReport.Theme_ID.Simple) ' if ignored it is by default classic theme
        oRep.SetReportFooter("This is your report footer Section", "you can add here as many lines as you want")
        oRep.PopulateReport(oMaster) '' this must be your last assignment
    End Sub

    Private Sub TestForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        On Error Resume Next
        oAccess.Quit()
        Me.Close()
    End Sub
End Class
