<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class TextExtract
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PickFile = New System.Windows.Forms.OpenFileDialog
        Me.Open = New System.Windows.Forms.Button
        Me.ExitProgram = New System.Windows.Forms.Button
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel
        Me.Label1 = New System.Windows.Forms.Label
        Me.FileLength = New System.Windows.Forms.Label
        Me.FileText = New System.Windows.Forms.TextBox
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PickFile
        '
        Me.PickFile.DefaultExt = "doc"
        Me.PickFile.Title = "Pick a file"
        '
        'Open
        '
        Me.Open.Location = New System.Drawing.Point(3, 3)
        Me.Open.Name = "Open"
        Me.Open.Size = New System.Drawing.Size(69, 23)
        Me.Open.TabIndex = 0
        Me.Open.Text = "&Open..."
        Me.Open.UseVisualStyleBackColor = True
        '
        'ExitProgram
        '
        Me.ExitProgram.Location = New System.Drawing.Point(78, 3)
        Me.ExitProgram.Name = "ExitProgram"
        Me.ExitProgram.Size = New System.Drawing.Size(75, 23)
        Me.ExitProgram.TabIndex = 1
        Me.ExitProgram.Text = "E&xit"
        Me.ExitProgram.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Open)
        Me.FlowLayoutPanel1.Controls.Add(Me.ExitProgram)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 535)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(792, 31)
        Me.FlowLayoutPanel1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 14)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Length:"
        '
        'FileLength
        '
        Me.FileLength.AutoSize = True
        Me.FileLength.Location = New System.Drawing.Point(66, 15)
        Me.FileLength.Name = "FileLength"
        Me.FileLength.Size = New System.Drawing.Size(0, 14)
        Me.FileLength.TabIndex = 4
        '
        'FileText
        '
        Me.FileText.Location = New System.Drawing.Point(20, 43)
        Me.FileText.Multiline = True
        Me.FileText.Name = "FileText"
        Me.FileText.ReadOnly = True
        Me.FileText.Size = New System.Drawing.Size(748, 473)
        Me.FileText.TabIndex = 5
        Me.FileText.WordWrap = False
        '
        'TextExtract
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(792, 566)
        Me.Controls.Add(Me.FileText)
        Me.Controls.Add(Me.FileLength)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "TextExtract"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Text Extractor Test"
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PickFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Open As System.Windows.Forms.Button
    Friend WithEvents ExitProgram As System.Windows.Forms.Button
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FileLength As System.Windows.Forms.Label
    Friend WithEvents FileText As System.Windows.Forms.TextBox
#End Region
End Class