Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms
Imports System.Data.OleDb
' tes buat jalan belakang layar
Imports System.ComponentModel
'Public Class FrmMenuUtama

Public Class FrmMenu
    Dim rscon, RsServer As New ADODB.Recordset
    Dim conn, ConnServer As New ADODB.Connection
    Dim sql, jamdigital As String
    Dim jam, menit, detik As Int64

    Dim frm As New Form
    Dim mulai As Boolean
    Dim munculJam As DateTime
    Private Sub LpbToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LpbToolStripMenuItem1.Click
        FrmUploaderlpb.ShowDialog()
    End Sub

    Private Sub ToToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToToolStripMenuItem.Click
        Frmupto.ShowDialog()
    End Sub

    Private Sub LbToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FrmMenu_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call getPathMdb()
    End Sub

    Private Sub FrmMenu_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'If conn.State = 0 Then
        '    GetStringKoneksi()
        '    conn.Open(StrKoneksi)
        'End If
        'sql = "exec spMstUser1022 'bukaupload','" & StrNamaUser & "','x','x',1,'2017-01-01','2017-01-01',1,1"
        'conn.Execute(sql)
        MsgBox("Terima Kasih", vbOKOnly + vbInformation, "PT. SMI")
        End
    End Sub

    Private Sub FrmMenu_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

    End Sub

    Private Sub FrmMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SERIALNUMBERRETURTOKOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SERIALNUMBERRETURTOKOToolStripMenuItem.Click
        Frmreturtoko.ShowDialog()
    End Sub

    Private Sub SNToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SNToolStripMenuItem.Click
        Form2.ShowDialog()
    End Sub

    Private Sub TesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub SERIALNUMBERMUTASIGSBSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SERIALNUMBERMUTASIGSBSToolStripMenuItem.Click
        FrmUpmutasi.ShowDialog()
    End Sub

    Private Sub SNMUTASIDCOUTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SNMUTASIDCOUTToolStripMenuItem.Click
        Frmupmutasidcout.ShowDialog()
    End Sub

    Private Sub SNMUTASIDCINToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SNMUTASIDCINToolStripMenuItem.Click
        Frmupmutasidcin.ShowDialog()
    End Sub

    Private Sub SNRETURSUPPLIERToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SNRETURSUPPLIERToolStripMenuItem.Click
        form4.ShowDialog()
    End Sub
End Class