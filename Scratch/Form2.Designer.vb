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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NameTextBox = New System.Windows.Forms.TextBox()
        Me.TitleDropDown = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GenderDropDown = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ProfessionDropDown = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.YearsOfExperienceTextBox = New System.Windows.Forms.TextBox()
        Me.ResetButton = New System.Windows.Forms.Button()
        Me.SubmitButton = New System.Windows.Forms.Button()
        Me.SummaryTextBox = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CopyToListBoxButton = New System.Windows.Forms.Button()
        Me.SummaryListBox = New System.Windows.Forms.ListBox()
        Me.UpButton = New System.Windows.Forms.Button()
        Me.ClearButton = New System.Windows.Forms.Button()
        Me.RemoveButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Title"
        '
        'NameTextBox
        '
        Me.NameTextBox.Location = New System.Drawing.Point(12, 83)
        Me.NameTextBox.Name = "NameTextBox"
        Me.NameTextBox.Size = New System.Drawing.Size(236, 26)
        Me.NameTextBox.TabIndex = 1
        '
        'TitleDropDown
        '
        Me.TitleDropDown.FormattingEnabled = True
        Me.TitleDropDown.Location = New System.Drawing.Point(12, 31)
        Me.TitleDropDown.Name = "TitleDropDown"
        Me.TitleDropDown.Size = New System.Drawing.Size(121, 27)
        Me.TitleDropDown.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 19)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Name"
        '
        'GenderDropDown
        '
        Me.GenderDropDown.FormattingEnabled = True
        Me.GenderDropDown.Location = New System.Drawing.Point(12, 134)
        Me.GenderDropDown.Name = "GenderDropDown"
        Me.GenderDropDown.Size = New System.Drawing.Size(121, 27)
        Me.GenderDropDown.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 19)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Gender"
        '
        'ProfessionDropDown
        '
        Me.ProfessionDropDown.FormattingEnabled = True
        Me.ProfessionDropDown.Location = New System.Drawing.Point(12, 186)
        Me.ProfessionDropDown.Name = "ProfessionDropDown"
        Me.ProfessionDropDown.Size = New System.Drawing.Size(121, 27)
        Me.ProfessionDropDown.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 164)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(99, 19)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Profession"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 216)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(180, 19)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Years of Experience"
        '
        'YearsOfExperienceTextBox
        '
        Me.YearsOfExperienceTextBox.Location = New System.Drawing.Point(8, 238)
        Me.YearsOfExperienceTextBox.Name = "YearsOfExperienceTextBox"
        Me.YearsOfExperienceTextBox.Size = New System.Drawing.Size(125, 26)
        Me.YearsOfExperienceTextBox.TabIndex = 8
        '
        'ResetButton
        '
        Me.ResetButton.Location = New System.Drawing.Point(8, 289)
        Me.ResetButton.Name = "ResetButton"
        Me.ResetButton.Size = New System.Drawing.Size(75, 23)
        Me.ResetButton.TabIndex = 10
        Me.ResetButton.Text = "Reset"
        Me.ResetButton.UseVisualStyleBackColor = True
        '
        'SubmitButton
        '
        Me.SubmitButton.Location = New System.Drawing.Point(8, 318)
        Me.SubmitButton.Name = "SubmitButton"
        Me.SubmitButton.Size = New System.Drawing.Size(75, 23)
        Me.SubmitButton.TabIndex = 11
        Me.SubmitButton.Text = "Submit"
        Me.SubmitButton.UseVisualStyleBackColor = True
        '
        'SummaryTextBox
        '
        Me.SummaryTextBox.Location = New System.Drawing.Point(284, 31)
        Me.SummaryTextBox.Multiline = True
        Me.SummaryTextBox.Name = "SummaryTextBox"
        Me.SummaryTextBox.Size = New System.Drawing.Size(236, 233)
        Me.SummaryTextBox.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(280, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 19)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Summary"
        '
        'CopyToListBoxButton
        '
        Me.CopyToListBoxButton.Location = New System.Drawing.Point(284, 270)
        Me.CopyToListBoxButton.Name = "CopyToListBoxButton"
        Me.CopyToListBoxButton.Size = New System.Drawing.Size(204, 23)
        Me.CopyToListBoxButton.TabIndex = 14
        Me.CopyToListBoxButton.Text = "Copy Info To ListBox"
        Me.CopyToListBoxButton.UseVisualStyleBackColor = True
        '
        'SummaryListBox
        '
        Me.SummaryListBox.FormattingEnabled = True
        Me.SummaryListBox.ItemHeight = 19
        Me.SummaryListBox.Location = New System.Drawing.Point(557, 29)
        Me.SummaryListBox.Name = "SummaryListBox"
        Me.SummaryListBox.Size = New System.Drawing.Size(242, 232)
        Me.SummaryListBox.TabIndex = 15
        '
        'UpButton
        '
        Me.UpButton.Location = New System.Drawing.Point(805, 29)
        Me.UpButton.Name = "UpButton"
        Me.UpButton.Size = New System.Drawing.Size(75, 23)
        Me.UpButton.TabIndex = 16
        Me.UpButton.Text = "Up"
        Me.UpButton.UseVisualStyleBackColor = True
        '
        'ClearButton
        '
        Me.ClearButton.Location = New System.Drawing.Point(805, 116)
        Me.ClearButton.Name = "ClearButton"
        Me.ClearButton.Size = New System.Drawing.Size(75, 23)
        Me.ClearButton.TabIndex = 19
        Me.ClearButton.Text = "Clear"
        Me.ClearButton.UseVisualStyleBackColor = True
        '
        'RemoveButton
        '
        Me.RemoveButton.Location = New System.Drawing.Point(805, 87)
        Me.RemoveButton.Name = "RemoveButton"
        Me.RemoveButton.Size = New System.Drawing.Size(75, 23)
        Me.RemoveButton.TabIndex = 18
        Me.RemoveButton.Text = "Remove"
        Me.RemoveButton.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(918, 385)
        Me.Controls.Add(Me.ClearButton)
        Me.Controls.Add(Me.RemoveButton)
        Me.Controls.Add(Me.UpButton)
        Me.Controls.Add(Me.SummaryListBox)
        Me.Controls.Add(Me.CopyToListBoxButton)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.SummaryTextBox)
        Me.Controls.Add(Me.SubmitButton)
        Me.Controls.Add(Me.ResetButton)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.YearsOfExperienceTextBox)
        Me.Controls.Add(Me.ProfessionDropDown)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GenderDropDown)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TitleDropDown)
        Me.Controls.Add(Me.NameTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("JetBrains Mono", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents NameTextBox As TextBox
    Friend WithEvents TitleDropDown As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents GenderDropDown As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents ProfessionDropDown As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents YearsOfExperienceTextBox As TextBox
    Friend WithEvents ResetButton As Button
    Friend WithEvents SubmitButton As Button
    Friend WithEvents SummaryTextBox As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents CopyToListBoxButton As Button
    Friend WithEvents SummaryListBox As ListBox
    Friend WithEvents UpButton As Button
    Friend WithEvents ClearButton As Button
    Friend WithEvents RemoveButton As Button
End Class
