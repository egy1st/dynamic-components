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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Firebrick
        Me.Label1.Location = New System.Drawing.Point(16, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(360, 96)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Main Form"
        '
        'TestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(392, 266)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1})
        Me.Name = "TestForm"
        Me.Text = "TestForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    ' Not: first dialog appears is the Protection for DC_AppProtector itself
    ' second dialog appears is the your tutorial protection
    Private Sub TestForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim MyProtection As New DynamicComponents.AppProtector()
        Dim ProductName As String
        Dim CompanyInfo As String

        CompanyInfo = "Company Name: Your Company Name" + vbCrLf 'vbcrlf force new line
        CompanyInfo += "Home Page: http://www.your domain name.com" + vbCrLf
        CompanyInfo += "License: Free XX Days Trail Version"
        ProductName = "Your Product"
        MyProtection.SetInformation(ProductName, CompanyInfo, "http://www.mygoldensoft.com/Order Now.htm")
        MyProtection.SetAlgorithms(1234, 56, 78, "abcdefg")
        MyProtection.SetLicense(30)
        MyProtection.ShowAuthor()
    End Sub
End Class
