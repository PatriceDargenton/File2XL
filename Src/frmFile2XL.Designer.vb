<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFile2XL
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFile2XL))
        Me.cmdRemoveContextMenu = New System.Windows.Forms.Button()
        Me.lblContextMenu = New System.Windows.Forms.Label()
        Me.cmdAddContextMenu = New System.Windows.Forms.Button()
        Me.cmdCreateTestFiles = New System.Windows.Forms.Button()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdShow = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdRemoveContextMenu
        '
        Me.cmdRemoveContextMenu.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRemoveContextMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRemoveContextMenu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRemoveContextMenu.Location = New System.Drawing.Point(138, 13)
        Me.cmdRemoveContextMenu.Name = "cmdRemoveContextMenu"
        Me.cmdRemoveContextMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRemoveContextMenu.Size = New System.Drawing.Size(25, 25)
        Me.cmdRemoveContextMenu.TabIndex = 2
        Me.cmdRemoveContextMenu.Text = "-"
        Me.cmdRemoveContextMenu.UseVisualStyleBackColor = False
        '
        'lblContextMenu
        '
        Me.lblContextMenu.BackColor = System.Drawing.SystemColors.Control
        Me.lblContextMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblContextMenu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblContextMenu.Location = New System.Drawing.Point(23, 19)
        Me.lblContextMenu.Name = "lblContextMenu"
        Me.lblContextMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblContextMenu.Size = New System.Drawing.Size(78, 18)
        Me.lblContextMenu.TabIndex = 0
        Me.lblContextMenu.Text = "Context menu :"
        '
        'cmdAddContextMenu
        '
        Me.cmdAddContextMenu.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddContextMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddContextMenu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddContextMenu.Location = New System.Drawing.Point(107, 13)
        Me.cmdAddContextMenu.Name = "cmdAddContextMenu"
        Me.cmdAddContextMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddContextMenu.Size = New System.Drawing.Size(25, 25)
        Me.cmdAddContextMenu.TabIndex = 1
        Me.cmdAddContextMenu.Text = "+"
        Me.cmdAddContextMenu.UseVisualStyleBackColor = False
        '
        'cmdCreateTestFiles
        '
        Me.cmdCreateTestFiles.Location = New System.Drawing.Point(209, 15)
        Me.cmdCreateTestFiles.Name = "cmdCreateTestFiles"
        Me.cmdCreateTestFiles.Size = New System.Drawing.Size(105, 23)
        Me.cmdCreateTestFiles.TabIndex = 3
        Me.cmdCreateTestFiles.Text = "Create test files"
        Me.cmdCreateTestFiles.UseVisualStyleBackColor = True
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 160)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(395, 25)
        Me.ToolStrip1.TabIndex = 8
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.AutoSize = False
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(350, 22)
        Me.ToolStripLabel1.Text = "ToolStripLabel1"
        Me.ToolStripLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(46, 43)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(75, 38)
        Me.cmdStart.TabIndex = 4
        Me.cmdStart.Text = "Start"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(160, 43)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 38)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdShow
        '
        Me.cmdShow.Location = New System.Drawing.Point(275, 43)
        Me.cmdShow.Name = "cmdShow"
        Me.cmdShow.Size = New System.Drawing.Size(75, 38)
        Me.cmdShow.TabIndex = 6
        Me.cmdShow.Text = "Show"
        Me.cmdShow.UseVisualStyleBackColor = True
        '
        'lblInfo
        '
        Me.lblInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInfo.Location = New System.Drawing.Point(12, 101)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(371, 39)
        Me.lblInfo.TabIndex = 7
        Me.lblInfo.Text = "Messages"
        '
        'frmFile2XL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(395, 185)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.cmdShow)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.cmdCreateTestFiles)
        Me.Controls.Add(Me.cmdRemoveContextMenu)
        Me.Controls.Add(Me.lblContextMenu)
        Me.Controls.Add(Me.cmdAddContextMenu)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmFile2XL"
        Me.Text = "File2XL"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

End Sub
    Friend WithEvents cmdRemoveContextMenu As System.Windows.Forms.Button
    Friend WithEvents lblContextMenu As System.Windows.Forms.Label
    Friend WithEvents cmdAddContextMenu As System.Windows.Forms.Button
    Friend WithEvents cmdCreateTestFiles As System.Windows.Forms.Button
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdShow As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
