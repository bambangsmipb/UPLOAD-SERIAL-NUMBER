<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMenu))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.LpbToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LpbToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SERIALNUMBERRETURTOKOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SERIALNUMBERMUTASIGSBSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SNMUTASIDCOUTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SNMUTASIDCINToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CekSnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SNToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SNRETURSUPPLIERToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LpbToolStripMenuItem, Me.CekSnToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1094, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'LpbToolStripMenuItem
        '
        Me.LpbToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LpbToolStripMenuItem1, Me.ToToolStripMenuItem, Me.SERIALNUMBERRETURTOKOToolStripMenuItem, Me.SERIALNUMBERMUTASIGSBSToolStripMenuItem, Me.SNMUTASIDCOUTToolStripMenuItem, Me.SNMUTASIDCINToolStripMenuItem, Me.SNRETURSUPPLIERToolStripMenuItem})
        Me.LpbToolStripMenuItem.Name = "LpbToolStripMenuItem"
        Me.LpbToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.LpbToolStripMenuItem.Text = "UPLOAD"
        '
        'LpbToolStripMenuItem1
        '
        Me.LpbToolStripMenuItem1.Name = "LpbToolStripMenuItem1"
        Me.LpbToolStripMenuItem1.Size = New System.Drawing.Size(179, 22)
        Me.LpbToolStripMenuItem1.Text = "SN LPB"
        '
        'ToToolStripMenuItem
        '
        Me.ToToolStripMenuItem.Name = "ToToolStripMenuItem"
        Me.ToToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.ToToolStripMenuItem.Text = "SN TO TOKO"
        '
        'SERIALNUMBERRETURTOKOToolStripMenuItem
        '
        Me.SERIALNUMBERRETURTOKOToolStripMenuItem.Name = "SERIALNUMBERRETURTOKOToolStripMenuItem"
        Me.SERIALNUMBERRETURTOKOToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.SERIALNUMBERRETURTOKOToolStripMenuItem.Text = "SN  RETUR TOKO"
        '
        'SERIALNUMBERMUTASIGSBSToolStripMenuItem
        '
        Me.SERIALNUMBERMUTASIGSBSToolStripMenuItem.Name = "SERIALNUMBERMUTASIGSBSToolStripMenuItem"
        Me.SERIALNUMBERMUTASIGSBSToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.SERIALNUMBERMUTASIGSBSToolStripMenuItem.Text = "SN MUTASI GS/BS"
        '
        'SNMUTASIDCOUTToolStripMenuItem
        '
        Me.SNMUTASIDCOUTToolStripMenuItem.Name = "SNMUTASIDCOUTToolStripMenuItem"
        Me.SNMUTASIDCOUTToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.SNMUTASIDCOUTToolStripMenuItem.Text = "SN MUTASI DC OUT"
        '
        'SNMUTASIDCINToolStripMenuItem
        '
        Me.SNMUTASIDCINToolStripMenuItem.Name = "SNMUTASIDCINToolStripMenuItem"
        Me.SNMUTASIDCINToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.SNMUTASIDCINToolStripMenuItem.Text = "SN MUTASI DC IN"
        '
        'CekSnToolStripMenuItem
        '
        Me.CekSnToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SNToolStripMenuItem})
        Me.CekSnToolStripMenuItem.Name = "CekSnToolStripMenuItem"
        Me.CekSnToolStripMenuItem.Size = New System.Drawing.Size(63, 20)
        Me.CekSnToolStripMenuItem.Text = "CARI SN"
        '
        'SNToolStripMenuItem
        '
        Me.SNToolStripMenuItem.Name = "SNToolStripMenuItem"
        Me.SNToolStripMenuItem.Size = New System.Drawing.Size(89, 22)
        Me.SNToolStripMenuItem.Text = "SN"
        '
        'SNRETURSUPPLIERToolStripMenuItem
        '
        Me.SNRETURSUPPLIERToolStripMenuItem.Name = "SNRETURSUPPLIERToolStripMenuItem"
        Me.SNRETURSUPPLIERToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.SNRETURSUPPLIERToolStripMenuItem.Text = "SN RETUR SUPPLIER"
        '
        'FrmMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1094, 604)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MENU UTAMA"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents LpbToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LpbToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SERIALNUMBERRETURTOKOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CekSnToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SNToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SERIALNUMBERMUTASIGSBSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SNMUTASIDCOUTToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SNMUTASIDCINToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SNRETURSUPPLIERToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
