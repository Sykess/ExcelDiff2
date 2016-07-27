<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.FileToButton = New System.Windows.Forms.Button()
        Me.FileTwoLabel = New System.Windows.Forms.Label()
        Me.FileOneComboBox = New System.Windows.Forms.ComboBox()
        Me.FileOneLabel = New System.Windows.Forms.Label()
        Me.FileOneButton = New System.Windows.Forms.Button()
        Me.FileTwoComboBox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SaveComboBox = New System.Windows.Forms.ComboBox()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.DisplayCheckBox = New System.Windows.Forms.CheckBox()
        Me.SaveCheckBox = New System.Windows.Forms.CheckBox()
        Me.CompareButton = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TestTextToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.05031!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 71.54088!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.4088!))
        Me.TableLayoutPanel1.Controls.Add(Me.FileToButton, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.FileTwoLabel, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.FileOneComboBox, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FileOneLabel, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FileOneButton, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FileTwoComboBox, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.SaveComboBox, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.SaveButton, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.DisplayCheckBox, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.SaveCheckBox, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.CompareButton, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.ProgressBar1, 0, 4)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 28)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(848, 157)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'FileToButton
        '
        Me.FileToButton.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FileToButton.Location = New System.Drawing.Point(719, 33)
        Me.FileToButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.FileToButton.Name = "FileToButton"
        Me.FileToButton.Size = New System.Drawing.Size(126, 27)
        Me.FileToButton.TabIndex = 4
        Me.FileToButton.Text = "Choose"
        Me.FileToButton.UseVisualStyleBackColor = True
        '
        'FileTwoLabel
        '
        Me.FileTwoLabel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.FileTwoLabel.AutoSize = True
        Me.FileTwoLabel.Location = New System.Drawing.Point(15, 38)
        Me.FileTwoLabel.Name = "FileTwoLabel"
        Me.FileTwoLabel.Size = New System.Drawing.Size(79, 17)
        Me.FileTwoLabel.TabIndex = 0
        Me.FileTwoLabel.Text = "Excel File 2"
        '
        'FileOneComboBox
        '
        Me.FileOneComboBox.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.FileOneComboBox.FormattingEnabled = True
        Me.FileOneComboBox.Location = New System.Drawing.Point(113, 3)
        Me.FileOneComboBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.FileOneComboBox.Name = "FileOneComboBox"
        Me.FileOneComboBox.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FileOneComboBox.Size = New System.Drawing.Size(599, 24)
        Me.FileOneComboBox.TabIndex = 1
        '
        'FileOneLabel
        '
        Me.FileOneLabel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.FileOneLabel.AutoSize = True
        Me.FileOneLabel.Location = New System.Drawing.Point(15, 7)
        Me.FileOneLabel.Name = "FileOneLabel"
        Me.FileOneLabel.Size = New System.Drawing.Size(79, 17)
        Me.FileOneLabel.TabIndex = 0
        Me.FileOneLabel.Text = "Excel File 1"
        '
        'FileOneButton
        '
        Me.FileOneButton.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FileOneButton.Location = New System.Drawing.Point(719, 2)
        Me.FileOneButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.FileOneButton.Name = "FileOneButton"
        Me.FileOneButton.Size = New System.Drawing.Size(126, 27)
        Me.FileOneButton.TabIndex = 2
        Me.FileOneButton.Text = "Choose"
        Me.FileOneButton.UseVisualStyleBackColor = True
        '
        'FileTwoComboBox
        '
        Me.FileTwoComboBox.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.FileTwoComboBox.FormattingEnabled = True
        Me.FileTwoComboBox.Location = New System.Drawing.Point(113, 34)
        Me.FileTwoComboBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.FileTwoComboBox.Name = "FileTwoComboBox"
        Me.FileTwoComboBox.Size = New System.Drawing.Size(599, 24)
        Me.FileTwoComboBox.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 69)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Save to"
        '
        'SaveComboBox
        '
        Me.SaveComboBox.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.SaveComboBox.FormattingEnabled = True
        Me.SaveComboBox.Location = New System.Drawing.Point(113, 65)
        Me.SaveComboBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SaveComboBox.Name = "SaveComboBox"
        Me.SaveComboBox.Size = New System.Drawing.Size(599, 24)
        Me.SaveComboBox.TabIndex = 5
        '
        'SaveButton
        '
        Me.SaveButton.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SaveButton.Location = New System.Drawing.Point(719, 64)
        Me.SaveButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.Size = New System.Drawing.Size(126, 27)
        Me.SaveButton.TabIndex = 6
        Me.SaveButton.Text = "Choose"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'DisplayCheckBox
        '
        Me.DisplayCheckBox.AutoSize = True
        Me.DisplayCheckBox.Checked = True
        Me.DisplayCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DisplayCheckBox.Location = New System.Drawing.Point(113, 126)
        Me.DisplayCheckBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DisplayCheckBox.Name = "DisplayCheckBox"
        Me.DisplayCheckBox.Size = New System.Drawing.Size(76, 21)
        Me.DisplayCheckBox.TabIndex = 8
        Me.DisplayCheckBox.Text = "Display"
        Me.DisplayCheckBox.UseVisualStyleBackColor = True
        '
        'SaveCheckBox
        '
        Me.SaveCheckBox.AutoSize = True
        Me.SaveCheckBox.Checked = True
        Me.SaveCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SaveCheckBox.Location = New System.Drawing.Point(113, 95)
        Me.SaveCheckBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SaveCheckBox.Name = "SaveCheckBox"
        Me.SaveCheckBox.Size = New System.Drawing.Size(62, 21)
        Me.SaveCheckBox.TabIndex = 7
        Me.SaveCheckBox.Text = "Save"
        Me.SaveCheckBox.UseVisualStyleBackColor = True
        '
        'CompareButton
        '
        Me.CompareButton.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CompareButton.Location = New System.Drawing.Point(719, 95)
        Me.CompareButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CompareButton.Name = "CompareButton"
        Me.TableLayoutPanel1.SetRowSpan(Me.CompareButton, 2)
        Me.CompareButton.Size = New System.Drawing.Size(126, 60)
        Me.CompareButton.TabIndex = 9
        Me.CompareButton.Text = "Compare"
        Me.CompareButton.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(3, 126)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(104, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(5, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(848, 28)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveToolStripMenuItem, Me.QuitToolStripMenuItem, Me.TestTextToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(139, 26)
        Me.SaveToolStripMenuItem.Text = "Save"
        '
        'QuitToolStripMenuItem
        '
        Me.QuitToolStripMenuItem.Name = "QuitToolStripMenuItem"
        Me.QuitToolStripMenuItem.Size = New System.Drawing.Size(139, 26)
        Me.QuitToolStripMenuItem.Text = "Quit"
        '
        'TestTextToolStripMenuItem
        '
        Me.TestTextToolStripMenuItem.Name = "TestTextToolStripMenuItem"
        Me.TestTextToolStripMenuItem.Size = New System.Drawing.Size(139, 26)
        Me.TestTextToolStripMenuItem.Text = "TestText"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "txt"
        Me.SaveFileDialog1.Filter = "Text File|*.txt|JPG|*.jpg|All Files|*.*"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(848, 185)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "Form1"
        Me.Text = "Compare Excel Files"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents QuitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FileOneComboBox As ComboBox
    Friend WithEvents FileOneLabel As Label
    Friend WithEvents FileTwoLabel As Label
    Friend WithEvents FileToButton As Button
    Friend WithEvents FileTwoComboBox As ComboBox
    Friend WithEvents FileOneButton As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents SaveComboBox As ComboBox
    Friend WithEvents SaveButton As Button
    Friend WithEvents DisplayCheckBox As CheckBox
    Friend WithEvents SaveCheckBox As CheckBox
    Friend WithEvents CompareButton As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents TestTextToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
