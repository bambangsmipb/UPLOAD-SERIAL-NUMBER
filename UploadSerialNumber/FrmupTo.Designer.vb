<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frmupto
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
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtnoto = New System.Windows.Forms.TextBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.BtnValidasi = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ListView2 = New System.Windows.Forms.ListView()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtlocal = New System.Windows.Forms.TextBox()
        Me.btnbrows = New System.Windows.Forms.Button()
        Me.btnProses = New System.Windows.Forms.Button()
        Me.dgminmax = New System.Windows.Forms.DataGridView()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.pbload = New System.Windows.Forms.ProgressBar()
        Me.opdg = New System.Windows.Forms.OpenFileDialog()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbtotrows = New System.Windows.Forms.Label()
        Me.lbhi = New System.Windows.Forms.Label()
        Me.lbdc = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lbnama = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Panel3.SuspendLayout()
        CType(Me.dgminmax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.Label15)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.txtnoto)
        Me.Panel3.Controls.Add(Me.Button5)
        Me.Panel3.Controls.Add(Me.Label12)
        Me.Panel3.Controls.Add(Me.BtnValidasi)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.ListView2)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.txtlocal)
        Me.Panel3.Controls.Add(Me.btnbrows)
        Me.Panel3.Controls.Add(Me.btnProses)
        Me.Panel3.Controls.Add(Me.dgminmax)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Location = New System.Drawing.Point(1, 1)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1073, 457)
        Me.Panel3.TabIndex = 13
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(654, 428)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(155, 29)
        Me.Button1.TabIndex = 36
        Me.Button1.Text = "Cancel  Validasi  TO"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(996, 248)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(13, 15)
        Me.Label15.TabIndex = 41
        Me.Label15.Text = "0"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(874, 248)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(108, 15)
        Me.Label2.TabIndex = 40
        Me.Label2.Text = "Total Rows Validasi:"
        '
        'txtnoto
        '
        Me.txtnoto.BackColor = System.Drawing.SystemColors.Window
        Me.txtnoto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtnoto.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnoto.ForeColor = System.Drawing.Color.Black
        Me.txtnoto.Location = New System.Drawing.Point(13, 81)
        Me.txtnoto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtnoto.Name = "txtnoto"
        Me.txtnoto.ReadOnly = True
        Me.txtnoto.Size = New System.Drawing.Size(172, 29)
        Me.txtnoto.TabIndex = 36
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(202, 78)
        Me.Button5.Margin = New System.Windows.Forms.Padding(2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(128, 29)
        Me.Button5.TabIndex = 39
        Me.Button5.Text = "Cari Nomor TO"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(894, 97)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 13)
        Me.Label12.TabIndex = 37
        Me.Label12.Text = "Total QTY TO :"
        '
        'BtnValidasi
        '
        Me.BtnValidasi.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.BtnValidasi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnValidasi.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnValidasi.ForeColor = System.Drawing.Color.White
        Me.BtnValidasi.Location = New System.Drawing.Point(557, 428)
        Me.BtnValidasi.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnValidasi.Name = "BtnValidasi"
        Me.BtnValidasi.Size = New System.Drawing.Size(90, 29)
        Me.BtnValidasi.TabIndex = 35
        Me.BtnValidasi.Text = "Validasi"
        Me.BtnValidasi.UseVisualStyleBackColor = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(982, 97)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(65, 13)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Jumlah QTY"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(646, 97)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(89, 13)
        Me.Label11.TabIndex = 36
        Me.Label11.Text = "Jumlah SKU TO :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(11, 69)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 13)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "NOMOR TO"
        '
        'ListView2
        '
        Me.ListView2.Location = New System.Drawing.Point(13, 113)
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(1048, 129)
        Me.ListView2.TabIndex = 30
        Me.ListView2.UseCompatibleStateImageBehavior = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(746, 97)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 13)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "Jumlah Item"
        '
        'txtlocal
        '
        Me.txtlocal.BackColor = System.Drawing.SystemColors.Window
        Me.txtlocal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtlocal.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtlocal.ForeColor = System.Drawing.Color.Black
        Me.txtlocal.Location = New System.Drawing.Point(102, 427)
        Me.txtlocal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtlocal.Name = "txtlocal"
        Me.txtlocal.ReadOnly = True
        Me.txtlocal.Size = New System.Drawing.Size(451, 29)
        Me.txtlocal.TabIndex = 27
        '
        'btnbrows
        '
        Me.btnbrows.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.btnbrows.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnbrows.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnbrows.ForeColor = System.Drawing.Color.White
        Me.btnbrows.Location = New System.Drawing.Point(8, 426)
        Me.btnbrows.Margin = New System.Windows.Forms.Padding(2)
        Me.btnbrows.Name = "btnbrows"
        Me.btnbrows.Size = New System.Drawing.Size(90, 28)
        Me.btnbrows.TabIndex = 26
        Me.btnbrows.Text = "BROWSE"
        Me.btnbrows.UseVisualStyleBackColor = False
        '
        'btnProses
        '
        Me.btnProses.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.btnProses.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnProses.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProses.ForeColor = System.Drawing.Color.White
        Me.btnProses.Location = New System.Drawing.Point(971, 429)
        Me.btnProses.Margin = New System.Windows.Forms.Padding(2)
        Me.btnProses.Name = "btnProses"
        Me.btnProses.Size = New System.Drawing.Size(90, 28)
        Me.btnProses.TabIndex = 25
        Me.btnProses.Text = "UPLOAD"
        Me.btnProses.UseVisualStyleBackColor = False
        '
        'dgminmax
        '
        Me.dgminmax.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgminmax.Location = New System.Drawing.Point(9, 265)
        Me.dgminmax.Margin = New System.Windows.Forms.Padding(2)
        Me.dgminmax.Name = "dgminmax"
        Me.dgminmax.RowHeadersWidth = 51
        Me.dgminmax.RowTemplate.Height = 24
        Me.dgminmax.Size = New System.Drawing.Size(1052, 158)
        Me.dgminmax.TabIndex = 21
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SteelBlue
        Me.Label5.ForeColor = System.Drawing.Color.SteelBlue
        Me.Label5.Location = New System.Drawing.Point(10, 57)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(1070, 406)
        Me.Label5.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SteelBlue
        Me.Label4.Location = New System.Drawing.Point(6, 62)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(1066, 391)
        Me.Label4.TabIndex = 19
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(0, 1)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1061, 54)
        Me.Panel1.TabIndex = 18
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Location = New System.Drawing.Point(1033, 5)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(26, 28)
        Me.Button2.TabIndex = 14
        Me.Button2.Text = "x"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(36, 25)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "---"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(110, 69)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(10, 13)
        Me.Label14.TabIndex = 38
        Me.Label14.Text = "."
        '
        'pbload
        '
        Me.pbload.Location = New System.Drawing.Point(-8, 630)
        Me.pbload.Margin = New System.Windows.Forms.Padding(2)
        Me.pbload.Name = "pbload"
        Me.pbload.Size = New System.Drawing.Size(1056, 10)
        Me.pbload.TabIndex = 28
        Me.pbload.Visible = False
        '
        'opdg
        '
        '
        'BackgroundWorker1
        '
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(781, 603)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 15)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Total Rows :"
        '
        'lbtotrows
        '
        Me.lbtotrows.AutoSize = True
        Me.lbtotrows.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbtotrows.ForeColor = System.Drawing.Color.Black
        Me.lbtotrows.Location = New System.Drawing.Point(875, 603)
        Me.lbtotrows.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbtotrows.Name = "lbtotrows"
        Me.lbtotrows.Size = New System.Drawing.Size(13, 15)
        Me.lbtotrows.TabIndex = 29
        Me.lbtotrows.Text = "0"
        Me.lbtotrows.Visible = False
        '
        'lbhi
        '
        Me.lbhi.AutoSize = True
        Me.lbhi.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbhi.ForeColor = System.Drawing.Color.Black
        Me.lbhi.Location = New System.Drawing.Point(6, 475)
        Me.lbhi.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbhi.Name = "lbhi"
        Me.lbhi.Size = New System.Drawing.Size(22, 15)
        Me.lbhi.TabIndex = 30
        Me.lbhi.Text = "Hi."
        '
        'lbdc
        '
        Me.lbdc.AutoSize = True
        Me.lbdc.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbdc.ForeColor = System.Drawing.Color.Black
        Me.lbdc.Location = New System.Drawing.Point(369, 475)
        Me.lbdc.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbdc.Name = "lbdc"
        Me.lbdc.Size = New System.Drawing.Size(51, 15)
        Me.lbdc.TabIndex = 33
        Me.lbdc.Text = "namaDc"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(338, 475)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(29, 15)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "DC :"
        '
        'lbnama
        '
        Me.lbnama.AutoSize = True
        Me.lbnama.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbnama.ForeColor = System.Drawing.Color.Black
        Me.lbnama.Location = New System.Drawing.Point(25, 475)
        Me.lbnama.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbnama.Name = "lbnama"
        Me.lbnama.Size = New System.Drawing.Size(37, 15)
        Me.lbnama.TabIndex = 31
        Me.lbnama.Text = "nama"
        Me.lbnama.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(608, 472)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(75, 15)
        Me.Label10.TabIndex = 34
        Me.Label10.Text = "Jumlah SKU :"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(688, 472)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(13, 15)
        Me.Label13.TabIndex = 35
        Me.Label13.Text = "0"
        '
        'Frmupto
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(1081, 545)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.lbdc)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lbnama)
        Me.Controls.Add(Me.lbhi)
        Me.Controls.Add(Me.lbtotrows)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pbload)
        Me.Controls.Add(Me.Panel3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frmupto"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.dgminmax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtlocal As System.Windows.Forms.TextBox
    Friend WithEvents btnbrows As System.Windows.Forms.Button
    Friend WithEvents btnProses As System.Windows.Forms.Button
    Friend WithEvents dgminmax As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents opdg As System.Windows.Forms.OpenFileDialog
    Friend WithEvents pbload As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbtotrows As System.Windows.Forms.Label
    Friend WithEvents lbhi As System.Windows.Forms.Label
    Friend WithEvents lbdc As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lbnama As System.Windows.Forms.Label
    Friend WithEvents ListView2 As System.Windows.Forms.ListView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents BtnValidasi As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents txtnoto As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
