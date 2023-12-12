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
        Me.bWorker = New System.ComponentModel.BackgroundWorker()
        Me.pbar = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblfilesize = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblpercent = New System.Windows.Forms.Label()
        Me.lbldownloadbytes = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'bWorker
        '
        Me.bWorker.WorkerReportsProgress = True
        Me.bWorker.WorkerSupportsCancellation = True
        '
        'pbar
        '
        Me.pbar.Location = New System.Drawing.Point(12, 11)
        Me.pbar.Name = "pbar"
        Me.pbar.Size = New System.Drawing.Size(446, 31)
        Me.pbar.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "File Size :"
        '
        'lblfilesize
        '
        Me.lblfilesize.AutoSize = True
        Me.lblfilesize.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfilesize.ForeColor = System.Drawing.Color.Blue
        Me.lblfilesize.Location = New System.Drawing.Point(83, 50)
        Me.lblfilesize.Name = "lblfilesize"
        Me.lblfilesize.Size = New System.Drawing.Size(49, 13)
        Me.lblfilesize.TabIndex = 2
        Me.lblfilesize.Text = "0 Bytes"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(271, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 15)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Download bytes :"
        '
        'lblpercent
        '
        Me.lblpercent.AutoSize = True
        Me.lblpercent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblpercent.Location = New System.Drawing.Point(464, 20)
        Me.lblpercent.Name = "lblpercent"
        Me.lblpercent.Size = New System.Drawing.Size(29, 16)
        Me.lblpercent.TabIndex = 4
        Me.lblpercent.Text = "0%"
        '
        'lbldownloadbytes
        '
        Me.lbldownloadbytes.AutoSize = True
        Me.lbldownloadbytes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbldownloadbytes.ForeColor = System.Drawing.Color.Blue
        Me.lbldownloadbytes.Location = New System.Drawing.Point(393, 52)
        Me.lbldownloadbytes.Name = "lbldownloadbytes"
        Me.lbldownloadbytes.Size = New System.Drawing.Size(49, 13)
        Me.lbldownloadbytes.TabIndex = 5
        Me.lbldownloadbytes.Text = "0 Bytes"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(507, 75)
        Me.Controls.Add(Me.lbldownloadbytes)
        Me.Controls.Add(Me.lblpercent)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblfilesize)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pbar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form1"
        Me.Opacity = 0.7R
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.TransparencyKey = System.Drawing.Color.Lime
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents bWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents pbar As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblfilesize As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblpercent As System.Windows.Forms.Label
    Friend WithEvents lbldownloadbytes As System.Windows.Forms.Label

End Class
