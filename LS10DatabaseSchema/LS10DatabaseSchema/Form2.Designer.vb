<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lstTables = New System.Windows.Forms.ListBox()
        Me.lstFields = New System.Windows.Forms.ListBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnShowFields = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstTables
        '
        Me.lstTables.FormattingEnabled = True
        Me.lstTables.Location = New System.Drawing.Point(12, 12)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(120, 225)
        Me.lstTables.TabIndex = 0
        '
        'lstFields
        '
        Me.lstFields.FormattingEnabled = True
        Me.lstFields.Location = New System.Drawing.Point(294, 12)
        Me.lstFields.Name = "lstFields"
        Me.lstFields.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstFields.Size = New System.Drawing.Size(120, 225)
        Me.lstFields.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(171, 187)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(171, 117)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnShowFields
        '
        Me.btnShowFields.Location = New System.Drawing.Point(171, 45)
        Me.btnShowFields.Name = "btnShowFields"
        Me.btnShowFields.Size = New System.Drawing.Size(75, 23)
        Me.btnShowFields.TabIndex = 4
        Me.btnShowFields.Text = "Show Fields"
        Me.btnShowFields.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(426, 261)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnShowFields)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lstFields)
        Me.Controls.Add(Me.lstTables)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lstTables As ListBox
    Friend WithEvents lstFields As ListBox
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnOK As Button
    Friend WithEvents btnShowFields As Button
End Class
