Imports Microsoft.Win32
Imports System.IO
Public Class FrmLogin
    Private MouseIsDown As Boolean = False
    Private MouseIsDownLoc As Point = Nothing
    Dim Conn, ConnMDB As New ADODB.Connection
    Dim RsConn, RsMdb As New ADODB.Recordset
    Dim sql, passx, pcname, initial, cari As String
    Dim flaguser, flagpass As Boolean
    Dim sumNotif, nilaisum As Integer


    Private Sub FrmLogin_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'Call getPathMdb()
    End Sub

    Private Sub FrmLogin_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub
    Private Sub FrmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        'Call getPathMdb()
        Call gambar()
        Me.PictureBox1.Image = System.Drawing.Image.FromFile(logo)

        Call NamaPerusahaan()
        Call alamatftp()
        Label7.Text = "UPLOADER SERIAL NUMBER"

        Me.Text = namadc
        Me.Label3.Text = NamaPT
        Me.txtuser.Clear()
        Me.txtpass.Clear()
        Me.txtuser.Focus()
        Control.CheckForIllegalCrossThreadCalls = False
        BackgroundWorker2.RunWorkerAsync()
        Call namadcAktif()

        Call ceksatudc()

        Me.Text = namadc
        Me.Label3.Text = NamaPT

        Control.CheckForIllegalCrossThreadCalls = False
        'BackgroundWorker2.RunWorkerAsync()

        Label8.Text = "   Server : " & StrDB
        Label9.Text = "    Versi : " & System.Windows.Forms.Application.ProductVersion

    End Sub
    Sub rr()

        'SetWatermark(txtuser, "User")
        'SetWatermark(txtpass, "Password")

        'GetJamClosing()
        'GetJamSvr()
        'j = Format(jamSvr, "HH")
        'm = Format(jamSvr, "mm")
        'flagloginClose = True
        'If j = JamClosing AndAlso m >= DurasiClosing Then
        '    flagloginClose = False
        '    Dim frm As FrmPesanClosing
        '    frm = New FrmPesanClosing()
        '    frm.ShowDialog()
        '    Exit Sub
        'End If

        'Call gambar()
        'Me.PictureBox1.Image = System.Drawing.Image.FromFile(logo)

        'Call NamaPerusahaan()
        'Call alamatftp()


        'Me.Text = namadc
        'Me.Label3.Text = NamaPT

        'Control.CheckForIllegalCrossThreadCalls = False
        'BackgroundWorker2.RunWorkerAsync()

        'Label11.Text = "Versi : " & System.Windows.Forms.Application.ProductVersion
    End Sub
    Sub ceksatudc()

        Dim jmldc As Integer
        strsql = "SELECT count(iddc) as jmlaktif from MstDC WHERE statusData=1"
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            jmldc = RsConfig("jmlaktif").Value
        End If
        If jmldc = 1 Then
            'MsgBox("DC Ready...")

        Else
            MsgBox("Ada Kesalahan diSatatus DC, Hubungi IT...")
            End
        End If

    End Sub
    Public Sub namadcAktif()

        GetStringKoneksi()
        If ConnMDB.State = 0 Then
            ConnMDB.Open(StrKoneksi)
        End If



        strsql = "select a.*,b.namaKabKota  from mstdc a " & _
                 "inner join MstKabKota b on a.idPropinsi =b.idPropinsi and a.idKabKota =b.idKabKota  where a.statusdata=1 "
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            kodedc = RsConfig("kodedc").Value
            namadc = RsConfig("namadc").Value()
            IdDC = RsConfig("IdDC").Value
            alamatdc = RsConfig("Alamatdc").Value
            telpdc = RsConfig("telepon").Value
            kotadc = RsConfig("namaKabKota").Value
        End If
        Label4.Text = namadc

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call cek()

    End Sub

    Private Sub txtuser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtuser.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtpass.Focus()
            txtpass.SelectAll()
            e.SuppressKeyPress = True

        End If
    End Sub

    Private Sub txtpass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtpass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub cek()
        GetServerDate()
        If Conn.State = 0 Then
            GetStringKoneksi()
            Conn.Open(StrKoneksi)
        End If


        If txtuser.Text = "System" And txtpass.Text = "admin" Then
            'FrmMasterUser.ShowDialog()
        Else

            'strsql = "Select namauser,iduser,passuser,namagroupuser,statusUser from MstUser a inner join MstUserGroup b on a.idGroupUser =b.idGroupUser  where IdUser='" & txtuser.Text & "'"
            'strsql = "Select namauser,iduser,passuser,namagroupuser,statusUser from smiMstUserupload a inner join MstUserGroup b on a.idGroupUser =b.idGroupUser  where IdUser='" & txtuser.Text & "'"
            strsql = "exec spSnUploadlpb 'cariuser',0,0,0,'',0,'" & txtuser.Text & "'"
            RsConn = Conn.Execute(strsql)
            If Not RsConn.EOF Then
                If RsConn("statususer").Value = 0 Or RsConn("statususer").Value = 3 Then
                    MsgBox("Anda tidak berhak menggunakan aplikasi ini !" & vbCrLf _
                            & " Silahkan Hubungi Administrator !!!", vbOKOnly + vbCritical, "Informasi")
                    Exit Sub
                Else

                    passx = (Decrypt(RsConn("passuser").Value))
                    If Trim(passx) = txtpass.Text Then
                        StrNamaUser = RsConn("idUser").Value
                        StrUserid = RsConn("namauser").Value
                        VarBagian = RsConn("namagroupuser").Value


                        'sql = "exec spMstUser1022 'kunciupload','" & StrNamaUser & "','x','x',1,'2017-01-01','2017-01-01',1,1"
                        'Conn.Execute(sql)

                        Me.Hide()
                        FrmMenu.Show()

                    Else
                        MsgBox("Password yang anda masukan salah !!!", vbOKOnly + vbCritical, "Info")
                        txtpass.Focus()
                        txtpass.SelectAll()
                        Exit Sub
                    End If
                End If
            Else
                MsgBox("Username yang anda masukan salah / tidak terdaftar !!!", vbOKOnly + vbCritical, "Info")
                txtuser.Focus()
                txtuser.SelectAll()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtuser_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtuser.KeyPress
        If (e.KeyChar Like "[',]") Then e.Handled() = True
    End Sub

    Private Sub txtuser_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtuser.TextChanged
        If flaguser = True Then
            ''  txtuser.Clear()
            txtuser.Focus()
            txtuser.ForeColor = Color.Black
            flaguser = False
            flagpass = True
        Else
            Exit Sub
        End If
    End Sub

    Private Sub txtpass_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtpass.KeyPress
        If (e.KeyChar Like "[',]") Then e.Handled() = True
    End Sub




    Private Sub txtuser_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtuser.MouseClick
        flaguser = True
        If txtuser.Text = "User Id" Then
            txtuser.Clear()
            txtuser.Focus()
            txtuser.ForeColor = Color.Black
            flaguser = False
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        End
    End Sub


    Private Sub txtpass_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtpass.MouseClick
        flagpass = True
        If txtpass.Text = "Entry Your Password" Then
            txtpass.Clear()
            txtpass.Focus()
            txtpass.ForeColor = Color.Black
            flagpass = False
        Else
            Exit Sub
        End If
    End Sub

    Private Sub txtpass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpass.TextChanged
        If flagpass = True Then
            '' txtpass.Clear()
            txtpass.Focus()
            txtpass.ForeColor = Color.Black
            txtpass.PasswordChar = "*"
            flagpass = False
        Else
            'txtpass.ForeColor = Color.Black
            'txtpass.PasswordChar = "*"
            Exit Sub
        End If
    End Sub

    Private Sub Panel2_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel2.MouseMove
        If e.Button = MouseButtons.Left Then
            If MouseIsDown = False Then
                MouseIsDown = True
                MouseIsDownLoc = New Point(e.X, e.Y)
            End If

            Me.Location = New Point(Me.Location.X + e.X - MouseIsDownLoc.X, Me.Location.Y + e.Y - MouseIsDownLoc.Y)
        End If
    End Sub

    Private Sub Panel2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel2.MouseDown
        MouseIsDown = False
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Label1_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.MouseLeave
        Label1.BackColor = Color.Transparent
    End Sub

    Private Sub Label1_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.MouseHover
        Label1.BackColor = Color.Silver
    End Sub


    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 80
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        If Conn.State = 0 Then
            GetStringKoneksi()
            Conn.Open(StrKoneksi)
            ProgressBar1.Value = 20
        End If



    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        ProgressBar1.Maximum = 80
        'nilaisum = Registry.GetValue("HKEY_CURRENT_USER\StatusReadWADS", "Status", Nothing)
        'If nilaisum = 0 Then
        '    Label10.Text = ""
        'Else
        '    Label10.Text = sumNotif
        'End If
        ProgressBar1.Visible = False

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\StatusReadWADS",
        '  "Status", "0")
        'Label10.Text = ""

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
End Class