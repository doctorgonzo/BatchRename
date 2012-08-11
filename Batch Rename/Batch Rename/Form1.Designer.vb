<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.lstFiles = New System.Windows.Forms.ListView()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.grpOptions = New System.Windows.Forms.GroupBox()
        Me.txtExtension = New System.Windows.Forms.TextBox()
        Me.txtSuffix = New System.Windows.Forms.TextBox()
        Me.txtPrefix = New System.Windows.Forms.TextBox()
        Me.boxOptions = New System.Windows.Forms.CheckedListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNewName = New System.Windows.Forms.TextBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.lblInstructions = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkRemoveChars = New System.Windows.Forms.CheckBox()
        Me.rdoTrailing = New System.Windows.Forms.RadioButton()
        Me.rdoLeading = New System.Windows.Forms.RadioButton()
        Me.numLeading = New System.Windows.Forms.NumericUpDown()
        Me.numTrailing = New System.Windows.Forms.NumericUpDown()
        Me.lblLeading = New System.Windows.Forms.Label()
        Me.lblTrailing = New System.Windows.Forms.Label()
        Me.btnChop = New System.Windows.Forms.Button()
        Me.grpOptions.SuspendLayout()
        CType(Me.numLeading, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numTrailing, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstFiles
        '
        Me.lstFiles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstFiles.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.lstFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstFiles.HideSelection = False
        Me.lstFiles.Location = New System.Drawing.Point(341, 47)
        Me.lstFiles.Name = "lstFiles"
        Me.lstFiles.Size = New System.Drawing.Size(570, 273)
        Me.lstFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstFiles.TabIndex = 0
        Me.lstFiles.UseCompatibleStateImageBehavior = False
        Me.lstFiles.View = System.Windows.Forms.View.List
        '
        'btnBrowse
        '
        Me.btnBrowse.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBrowse.Location = New System.Drawing.Point(155, 12)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(99, 23)
        Me.btnBrowse.TabIndex = 1
        Me.btnBrowse.Text = "Browse For Files"
        Me.btnBrowse.UseVisualStyleBackColor = False
        '
        'grpOptions
        '
        Me.grpOptions.BackColor = System.Drawing.Color.Transparent
        Me.grpOptions.Controls.Add(Me.txtExtension)
        Me.grpOptions.Controls.Add(Me.txtSuffix)
        Me.grpOptions.Controls.Add(Me.txtPrefix)
        Me.grpOptions.Controls.Add(Me.boxOptions)
        Me.grpOptions.Controls.Add(Me.Label2)
        Me.grpOptions.Controls.Add(Me.txtNewName)
        Me.grpOptions.Controls.Add(Me.btnGo)
        Me.grpOptions.Controls.Add(Me.Label1)
        Me.grpOptions.Location = New System.Drawing.Point(2, 41)
        Me.grpOptions.Name = "grpOptions"
        Me.grpOptions.Size = New System.Drawing.Size(333, 204)
        Me.grpOptions.TabIndex = 3
        Me.grpOptions.TabStop = False
        Me.grpOptions.Text = "Options"
        '
        'txtExtension
        '
        Me.txtExtension.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.txtExtension.Location = New System.Drawing.Point(227, 184)
        Me.txtExtension.MaxLength = 4
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.Size = New System.Drawing.Size(100, 20)
        Me.txtExtension.TabIndex = 7
        '
        'txtSuffix
        '
        Me.txtSuffix.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.txtSuffix.Location = New System.Drawing.Point(227, 162)
        Me.txtSuffix.Name = "txtSuffix"
        Me.txtSuffix.Size = New System.Drawing.Size(100, 20)
        Me.txtSuffix.TabIndex = 6
        '
        'txtPrefix
        '
        Me.txtPrefix.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.txtPrefix.Location = New System.Drawing.Point(227, 139)
        Me.txtPrefix.Name = "txtPrefix"
        Me.txtPrefix.Size = New System.Drawing.Size(100, 20)
        Me.txtPrefix.TabIndex = 5
        '
        'boxOptions
        '
        Me.boxOptions.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.boxOptions.CheckOnClick = True
        Me.boxOptions.FormattingEnabled = True
        Me.boxOptions.Items.AddRange(New Object() {"Numbering", "Numbering with parentheses", "Numbering with hyphens", "Add a prefix", "Add a suffix", "Change the file extension"})
        Me.boxOptions.Location = New System.Drawing.Point(6, 89)
        Me.boxOptions.Name = "boxOptions"
        Me.boxOptions.Size = New System.Drawing.Size(215, 109)
        Me.boxOptions.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(218, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "What naming scheme would you like to use?"
        '
        'txtNewName
        '
        Me.txtNewName.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.txtNewName.Location = New System.Drawing.Point(6, 41)
        Me.txtNewName.Name = "txtNewName"
        Me.txtNewName.Size = New System.Drawing.Size(255, 20)
        Me.txtNewName.TabIndex = 0
        '
        'btnGo
        '
        Me.btnGo.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnGo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGo.Location = New System.Drawing.Point(268, 37)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(52, 23)
        Me.btnGo.TabIndex = 2
        Me.btnGo.Text = "Go!"
        Me.btnGo.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(179, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "What should we rename the files to?"
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnClear.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClear.Location = New System.Drawing.Point(260, 12)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 4
        Me.btnClear.Text = "Clear Form"
        Me.btnClear.UseVisualStyleBackColor = False
        '
        'lblInstructions
        '
        Me.lblInstructions.AutoSize = True
        Me.lblInstructions.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstructions.Location = New System.Drawing.Point(467, 15)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.Size = New System.Drawing.Size(455, 20)
        Me.lblInstructions.TabIndex = 5
        Me.lblInstructions.Text = "Select the files to rename from the list below before clicking Go!"
        '
        'Timer1
        '
        Me.Timer1.Interval = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(349, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(573, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "If you choose files with more than one extension, the renamed files will have the" & _
    " same extension as the last file in your list"
        '
        'chkRemoveChars
        '
        Me.chkRemoveChars.AutoSize = True
        Me.chkRemoveChars.Location = New System.Drawing.Point(8, 251)
        Me.chkRemoveChars.Name = "chkRemoveChars"
        Me.chkRemoveChars.Size = New System.Drawing.Size(184, 17)
        Me.chkRemoveChars.TabIndex = 7
        Me.chkRemoveChars.Text = "Remove characters from filename"
        Me.chkRemoveChars.UseVisualStyleBackColor = True
        '
        'rdoTrailing
        '
        Me.rdoTrailing.AutoSize = True
        Me.rdoTrailing.Location = New System.Drawing.Point(8, 297)
        Me.rdoTrailing.Name = "rdoTrailing"
        Me.rdoTrailing.Size = New System.Drawing.Size(112, 17)
        Me.rdoTrailing.TabIndex = 9
        Me.rdoTrailing.TabStop = True
        Me.rdoTrailing.Text = "Trailing characters"
        Me.rdoTrailing.UseVisualStyleBackColor = True
        Me.rdoTrailing.Visible = False
        '
        'rdoLeading
        '
        Me.rdoLeading.AutoSize = True
        Me.rdoLeading.Location = New System.Drawing.Point(8, 274)
        Me.rdoLeading.Name = "rdoLeading"
        Me.rdoLeading.Size = New System.Drawing.Size(116, 17)
        Me.rdoLeading.TabIndex = 10
        Me.rdoLeading.TabStop = True
        Me.rdoLeading.Text = "Leading characters"
        Me.rdoLeading.UseVisualStyleBackColor = True
        Me.rdoLeading.Visible = False
        '
        'numLeading
        '
        Me.numLeading.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.numLeading.Location = New System.Drawing.Point(130, 274)
        Me.numLeading.Name = "numLeading"
        Me.numLeading.Size = New System.Drawing.Size(52, 20)
        Me.numLeading.TabIndex = 11
        Me.numLeading.Visible = False
        '
        'numTrailing
        '
        Me.numTrailing.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.numTrailing.Location = New System.Drawing.Point(130, 297)
        Me.numTrailing.Name = "numTrailing"
        Me.numTrailing.Size = New System.Drawing.Size(52, 20)
        Me.numTrailing.TabIndex = 11
        Me.numTrailing.Visible = False
        '
        'lblLeading
        '
        Me.lblLeading.AutoSize = True
        Me.lblLeading.Location = New System.Drawing.Point(188, 278)
        Me.lblLeading.Name = "lblLeading"
        Me.lblLeading.Size = New System.Drawing.Size(129, 13)
        Me.lblLeading.TabIndex = 12
        Me.lblLeading.Text = "# of characters to remove"
        Me.lblLeading.Visible = False
        '
        'lblTrailing
        '
        Me.lblTrailing.AutoSize = True
        Me.lblTrailing.Location = New System.Drawing.Point(188, 299)
        Me.lblTrailing.Name = "lblTrailing"
        Me.lblTrailing.Size = New System.Drawing.Size(129, 13)
        Me.lblTrailing.TabIndex = 12
        Me.lblTrailing.Text = "# of characters to remove"
        Me.lblTrailing.Visible = False
        '
        'btnChop
        '
        Me.btnChop.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.btnChop.Location = New System.Drawing.Point(226, 251)
        Me.btnChop.Name = "btnChop"
        Me.btnChop.Size = New System.Drawing.Size(57, 23)
        Me.btnChop.TabIndex = 13
        Me.btnChop.Text = "Chop!"
        Me.btnChop.UseVisualStyleBackColor = False
        Me.btnChop.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(923, 332)
        Me.Controls.Add(Me.btnChop)
        Me.Controls.Add(Me.lblTrailing)
        Me.Controls.Add(Me.lblLeading)
        Me.Controls.Add(Me.numTrailing)
        Me.Controls.Add(Me.numLeading)
        Me.Controls.Add(Me.rdoLeading)
        Me.Controls.Add(Me.rdoTrailing)
        Me.Controls.Add(Me.chkRemoveChars)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblInstructions)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.grpOptions)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.lstFiles)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "The Amazing Batch Renamer"
        Me.grpOptions.ResumeLayout(False)
        Me.grpOptions.PerformLayout()
        CType(Me.numLeading, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numTrailing, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstFiles As System.Windows.Forms.ListView
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents grpOptions As System.Windows.Forms.GroupBox
    Friend WithEvents boxOptions As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNewName As System.Windows.Forms.TextBox
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSuffix As System.Windows.Forms.TextBox
    Friend WithEvents txtPrefix As System.Windows.Forms.TextBox
    Friend WithEvents txtExtension As System.Windows.Forms.TextBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents lblInstructions As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkRemoveChars As System.Windows.Forms.CheckBox
    Friend WithEvents rdoTrailing As System.Windows.Forms.RadioButton
    Friend WithEvents rdoLeading As System.Windows.Forms.RadioButton
    Friend WithEvents numLeading As System.Windows.Forms.NumericUpDown
    Friend WithEvents numTrailing As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblLeading As System.Windows.Forms.Label
    Friend WithEvents lblTrailing As System.Windows.Forms.Label
    Friend WithEvents btnChop As System.Windows.Forms.Button

End Class
