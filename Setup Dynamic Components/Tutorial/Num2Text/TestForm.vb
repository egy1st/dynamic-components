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
    Friend WithEvents Combo_Language As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents GenerateButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Combo_Language = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.GenerateButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Combo_Language
        '
        Me.Combo_Language.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Language.Items.AddRange(New Object() {"Arabic", "English", "French", "German"})
        Me.Combo_Language.Location = New System.Drawing.Point(152, 16)
        Me.Combo_Language.Name = "Combo_Language"
        Me.Combo_Language.Size = New System.Drawing.Size(160, 21)
        Me.Combo_Language.Sorted = True
        Me.Combo_Language.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Define Language"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Enter Numer"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(152, 48)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(208, 26)
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(16, 88)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(664, 20)
        Me.TextBox2.TabIndex = 7
        Me.TextBox2.Text = ""
        '
        'GenerateButton
        '
        Me.GenerateButton.Location = New System.Drawing.Point(280, 128)
        Me.GenerateButton.Name = "GenerateButton"
        Me.GenerateButton.Size = New System.Drawing.Size(144, 40)
        Me.GenerateButton.TabIndex = 6
        Me.GenerateButton.Text = "Translate Number"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(696, 190)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Combo_Language, Me.Label2, Me.Label1, Me.TextBox1, Me.TextBox2, Me.GenerateButton})
        Me.Name = "Form1"
        Me.Text = "Test Form"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim oTextNum As New DynamicComponents.Num2Text()

    Private Sub GenerateButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateButton.Click

        oTextNum.SetCurrency("Dollar", "Cent", "Dollars", "Cents")

        Select Case Me.Combo_Language.Text
            Case "English"
                Me.TextBox2.Text = oTextNum.TranslateNumber(Me.TextBox1.Text, DynamicComponents.Num2Text.Language_ID.English)

            Case "French"
                Me.TextBox2.Text = oTextNum.TranslateNumber(Me.TextBox1.Text, DynamicComponents.Num2Text.Language_ID.Frensh)

            Case "German"
                Me.TextBox2.Text = oTextNum.TranslateNumber(Me.TextBox1.Text, DynamicComponents.Num2Text.Language_ID.German)

            Case "Arabic"
                Me.TextBox2.Text = oTextNum.TranslateNumber(Me.TextBox1.Text, DynamicComponents.Num2Text.Language_ID.Arabic)
        End Select
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Combo_Language.Text = "English"
    End Sub
End Class
