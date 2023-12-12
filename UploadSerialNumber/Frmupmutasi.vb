Imports System.IO
Imports System.Data.SqlClient
Imports System.Threading
Imports System.ComponentModel
Public Class FrmUpmutasi
    Dim Conn, ConnMDB As New ADODB.Connection
    Dim RsConn, RsMdb As New ADODB.Recordset
    Dim jmlitem, jmlqty, jml_item, jml_qty As Integer
    Dim sql, passx, nomorpb, nomorsn, kodeproduk, namapanjang, namauser As String
    Private MouseIsDown As Boolean = False
    Private MouseIsDownLoc As Point = Nothing
    Dim oReader As StreamReader
    Dim minor, maxst, kdtk, nilaigrid, sts, npersent, jumka As Integer
    Dim kdprd As String
    Dim strsql, strnomutasi, strmmutasi, strkdmutasi, strkdproduk, strnamaproduk As String
    Dim flagproses As Boolean

    Private Sub Panel1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If e.Button = MouseButtons.Left Then
            If MouseIsDown = False Then
                MouseIsDown = True
                MouseIsDownLoc = New Point(e.X, e.Y)
            End If

            Me.Location = New Point(Me.Location.X + e.X - MouseIsDownLoc.X, Me.Location.Y + e.Y - MouseIsDownLoc.Y)
        End If
    End Sub

    Private Sub Panel1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        MouseIsDown = False
    End Sub
    Private Sub cektemp()
        sql = "SELECT count(nomormutasi) as jmlrow from SMITblSnMutasiTmp where nomormutasi='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
        'sql = "exec spSnUploadto 'cekjmlrowtmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"


        RsConn = Conn.Execute(sql)
        Dim sntemp As Integer
        If Not RsConn.EOF Then
            sntemp = RsConn("jmlrow").Value
            If sntemp > 0 Then
                sql = "delete from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
                'sql = "delete from SMITblSNTOkeTokoTmp"
                'sql = "exec spSnUploadto 'hapustmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
                RsConn = Conn.Execute(sql)
                Exit Sub
            End If
        End If
        RsConn.Close()
    End Sub
    Private Sub btnbrows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbrows.Click
        BtnValidasi.Enabled = False
        btnProses.Enabled = False
        dgminmax.DataSource = Nothing
        dgminmax.Rows.Clear()
        lbtotrows.Text = 0
        Label13.Text = 0
        Label15.Text = 0
        Try
            opdg.FileName = ""
            opdg.Filter = "WPS SpreadSheets (*.xls)|*.xls|All Files (*.*)|*.*"
            If opdg.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtlocal.Text = System.IO.Path.GetFullPath(opdg.FileName)
            End If

            Call loaddata()
        Catch ex As Exception
            Exit Sub
        End Try


    End Sub

    Private Sub Button2_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.MouseHover
        Button2.BackColor = Color.Red
    End Sub

    Private Sub Button2_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.MouseLeave
        Button2.BackColor = Color.DimGray
    End Sub

    Private Sub loaddata()
        dgminmax.DataSource = Nothing
        dgminmax.Rows.Clear()
        dgminmax.Columns.Clear()


        Dim aCol() As Integer = {13, 13, 28, 13, 13, 13, 13}
        Dim icol As Integer
        ''  Dim sqlstring As String

        With dgminmax.ColumnHeadersDefaultCellStyle
            .BackColor = Color.DeepPink  'navy
            .ForeColor = Color.Navy
            .Font = New Font("Arial", 9, FontStyle.Bold)
        End With


        With dgminmax
            '.EditMode = DataGridViewEditMode.EditOnKeystroke
            .AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            .ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Raised
            .CellBorderStyle =
                DataGridViewCellBorderStyle.Single
            .GridColor = SystemColors.ActiveBorder
            .RowHeadersVisible = False
            .SelectionMode =
                DataGridViewSelectionMode.CellSelect
            .MultiSelect = False
            .BackgroundColor = Color.Honeydew
            .AllowUserToResizeColumns = False
        End With


        Dim MyConnection As OleDb.OleDbConnection
        'Dim cmd As OleDb.OleDbCommand
        Dim Ds As System.Data.DataSet
        Dim MyAdapter As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & txtlocal.Text & "';Extended Properties=Excel 8.0;")
        MyAdapter = New System.Data.OleDb.OleDbDataAdapter("Select RTRIM(LTRIM(nomormutasi)),RTRIM(LTRIM(Kodeproduk)),RTRIM(LTRIM(Namapanjang)),RTRIM(LTRIM(nomorsn)),'' as CekStatus,'' as ceksn from [SheetMutasi$]", MyConnection)
        Ds = New System.Data.DataSet
        MyAdapter.Fill(Ds)
        Me.dgminmax.DataSource = Ds.Tables(0)

        For icol = 0 To 5
            dgminmax.Columns(icol).Width = (dgminmax.Width / 100) * aCol(icol)
            Select Case icol


                Case 0
                    dgminmax.Columns(icol).HeaderText = "Nomor Mutasi"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 1
                    dgminmax.Columns(icol).HeaderText = "Kode Produk"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 2
                    dgminmax.Columns(icol).HeaderText = "Nama Panjang"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 3
                    dgminmax.Columns(icol).HeaderText = "Nomor SN"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 4
                    dgminmax.Columns(icol).HeaderText = "Cek SKU"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 5
                    dgminmax.Columns(icol).HeaderText = "Cek SN"
                    dgminmax.Columns(icol).ReadOnly = True
            End Select

        Next
        nilaigrid = dgminmax.RowCount
        lbtotrows.Visible = True
        lbtotrows.Text = 0
        'lbtotrows.Text = nilaigrid
        If dgminmax.RowCount > 0 Then
            BtnValidasi.Enabled = True
        Else
            BtnValidasi.Enabled = False
        End If
        btnProses.Enabled = False
    End Sub

    Private Sub btnProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProses.Click

        Dim sql1 As String
        Dim sql22 As String
        sql1 = "select * from SMITblSnMutasitmp where nomormutasi='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql1)
        If Not RsConn.EOF() Then
            'sql1 = "insert into SMITblSNTOkeToko select kodetoko,nomorto,idproduk,nomorsn,iduser,tglimport,null  from SMITblSNTOkeTokoTmp where nomorto='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
            'sql1 = "exec spSnUploadto 'insertMutasi','" & txtnoto.Text & "','" & strkdproduk & "','" & nomorsn & "',0,'" & StrNamaUser & "','" & kodetoko & "'"
            sql1 = "exec spSnUploadMutasi 'insertMutasi','" & txtnoto.Text & "','','',0,'" & StrNamaUser & "',''"
            RsConn = Conn.Execute(sql1)

            If TextBox1.Text = "MUTASI GS TO BS" Then
                sql22 = "exec spSnUploadMutasi 'UpdateStatusSNgsbs','" & txtnoto.Text & "','','',0,'" & StrNamaUser & "',''"
                RsConn = Conn.Execute(sql22)
            End If

            If TextBox1.Text = "MUTASI BS TO GS" Then
                sql22 = "exec spSnUploadMutasi 'UpdateStatusSNbsgs','" & txtnoto.Text & "','','',0,'" & StrNamaUser & "',''"
                RsConn = Conn.Execute(sql22)
            End If

            If TextBox1.Text = "BOOKING IN" Then
                sql22 = "exec spSnUploadMutasi 'UpdateStatusSNBookingIN','" & txtnoto.Text & "','','',0,'" & StrNamaUser & "',''"
                RsConn = Conn.Execute(sql22)
            End If

            If TextBox1.Text = "BOOKING OUT" Then
                sql22 = "exec spSnUploadMutasi 'UpdateStatusSNBookingOut','" & txtnoto.Text & "','','',0,'" & StrNamaUser & "',''"
                RsConn = Conn.Execute(sql22)
            End If

            'RsConn = Conn.Execute(sql22)

            MsgBox("Upload Data Sukses..")
            sql = "delete from SMITblSnMutasitmp where nomormutasi='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            Call buka_new()
        Else
            MsgBox("Upload Gagal, Silahkan Upload Ulang..")

            Call buka_new()
        End If

    End Sub

    Private Sub kunci()
        btnbrows.Enabled = False
        btnProses.Enabled = False
        BtnValidasi.Enabled = False

    End Sub


    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'Call getPathMdb()
    End Sub

    Private Sub FrmUploaderTo_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Conn.State = 0 Then
            GetStringKoneksi()
            Conn.Open(StrKoneksi)
        End If
        Control.CheckForIllegalCrossThreadCalls = False


        Call namadcAktif()
        lbdc.Text = namadc & ""
        Call kunci()
        lbnama.Visible = True
        lbnama.Text = StrNamaUser
        Label3.Text = "UPLOAD SERIAL NUMBER MUTASI BS/GS"
        Call buka_new()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Call cekdata()
    End Sub

    Private Sub cekdata()
        Dim x As Integer = dgminmax.CurrentCellAddress.X
        Dim y As Integer = dgminmax.CurrentCellAddress.Y
        pbload.Minimum = 0
        pbload.Maximum = nilaigrid
        pbload.Value = 0
        'Call cektemp()
        'Call insertemp()
        'Label15.Text = 0
        For Each row As DataGridViewRow In Me.dgminmax.Rows

            'Try

            If IsDBNull(row.Cells.Item(0).Value) Then
                strnomutasi = 0
            Else
                strnomutasi = row.Cells.Item(0).Value
            End If

            If IsDBNull(row.Cells.Item(1).Value) Then
                strkdproduk = 0
            Else
                strkdproduk = row.Cells.Item(1).Value
            End If

            If IsDBNull(row.Cells.Item(2).Value) Then
                strnamaproduk = 0
            Else
                strnamaproduk = row.Cells.Item(2).Value
            End If

            If IsDBNull(row.Cells.Item(3).Value) Then
                nomorsn = ""
            Else
                nomorsn = row.Cells.Item(3).Value
            End If



            sql = "SELECT * from MutasiSaldoDetail WHERE nomorMutasi='" & txtnoto.Text & "' and idproduk in(select idproduk from mstproduk where kodeproduk='" & strkdproduk & "' and idsnqr=1)"
            RsConn = Conn.Execute(sql)

            If strnomutasi = 0 Or strkdproduk = "" Or nomorsn = "" Then
            Else

                If Not RsConn.EOF() Then
                    row.Cells.Item(4).Value = "SKU OK"
                    Dim colorDown As Color = Color.LightGreen
                    row.Cells.Item(4).Style.BackColor = colorDown
                    row.Cells.Item(4).Selected = True
                Else

                    row.Cells.Item(4).Value = "SKU TIDAK ADA"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(4).Style.BackColor = colorDown
                    row.Cells.Item(4).Selected = True
                End If

                Label15.Text = (Label15.Text + 1)
                'Label2.Text = Label15.Text
                'lbtotrows.Text = (lbtotrows.Text - 1)
            End If
            'RsConn.Close()


            sql = "SELECT * from SMITblBankSN WHERE nomorsn='" & nomorsn & "'"
            'sql = "SELECT * from SMITblBankSN WHERE nomorsn='" & nomorsn & "' and statusdata in(1,5)"
            'sql = "exec spSnUploadto 'CekSn','" & txtnoto.Text & "',0,'',0,''"
            RsConn = Conn.Execute(sql)
            If strnomutasi = 0 Or strkdproduk = "" Or nomorsn = "" Then
            Else

                If Not RsConn.EOF() Then
                    sql = "SELECT nomorsn from SMITblSnMutasiTmp where nomorsn='" & nomorsn & "'"
                    'sql = "exec spSnUploadto 'CekSnduplikat',0,0,'" & nomorsn & "',0,''"
                    RsConn = Conn.Execute(sql)
                    If Not RsConn.EOF Then
                        row.Cells.Item(5).Value = "SN DUPLIKAT"
                        Dim colorDown As Color = Color.LightSalmon
                        row.Cells.Item(5).Style.BackColor = colorDown
                        row.Cells.Item(5).Selected = True
                    Else
                        If TextBox1.Text = "MUTASI GS TO BS" Then
                            sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData in(1)"
                            RsConn = Conn.Execute(sql)
                            If RsConn.EOF Then
                                row.Cells.Item(5).Value = "CEK STATUS SN"
                                Dim colorDown As Color = Color.LightSalmon
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            Else
                                row.Cells.Item(5).Value = "SN OK"
                                Dim colorDown As Color = Color.LightGreen
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            End If
                        End If
                        If TextBox1.Text = "MUTASI BS TO GS" Then
                            sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData in(7)"
                            RsConn = Conn.Execute(sql)
                            If RsConn.EOF Then
                                row.Cells.Item(5).Value = "CEK STATUS SN"
                                Dim colorDown As Color = Color.LightSalmon
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            Else
                                row.Cells.Item(5).Value = "SN OK"
                                Dim colorDown As Color = Color.LightGreen
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            End If
                        End If

                        If TextBox1.Text = "BOOKING IN" Then
                            sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData in(1,5,7,8)"
                            RsConn = Conn.Execute(sql)
                            If RsConn.EOF Then
                                row.Cells.Item(5).Value = "CEK STATUS SN"
                                Dim colorDown As Color = Color.LightSalmon
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            Else
                                row.Cells.Item(5).Value = "SN OK"
                                Dim colorDown As Color = Color.LightGreen
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            End If
                        End If

                        If TextBox1.Text = "BOOKING OUT" Then
                            sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData in(1,5,7,8)"
                            RsConn = Conn.Execute(sql)
                            If RsConn.EOF Then
                                row.Cells.Item(5).Value = "CEK STATUS SN"
                                Dim colorDown As Color = Color.LightSalmon
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            Else
                                row.Cells.Item(5).Value = "SN OK"
                                Dim colorDown As Color = Color.LightGreen
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            End If
                        End If
                    End If




                    'BOOKING OUT

                Else
                    row.Cells.Item(5).Value = "SN Tidak Terdaftar"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(5).Style.BackColor = colorDown
                    row.Cells.Item(5).Selected = True





                End If
            End If
            'RsConn.Close()

            If strnomutasi = 0 Or strkdproduk = "" Or nomorsn = "" Then
            Else

                If row.Cells.Item(5).Value = "SN OK" And row.Cells.Item(4).Value = "SKU OK" Then
                    Dim sql1 As String
                    sql1 = "insert into SMITblSnMutasiTmp values('" & strnomutasi & "','" & TextBox1.Text & "',(select idproduk from mstproduk where kodeproduk ='" & strkdproduk & "'),'" & nomorsn & "','" & StrNamaUser & "',getdate(),null)"
                    'sql1 = "exec spSnUploadto 'insertTotmp','" & txtnoto.Text & "','" & kdprd & "','" & nomorsn & "',0,'" & StrNamaUser & "','" & kodetoko & "'"
                    RsConn = Conn.Execute(sql1)
                Else

                End If

            End If

        Next
        Call cekqty()
        BtnValidasi.Enabled = False

    End Sub
    Private Sub opdg_FileOk(ByVal sender As Object, ByVal e As CancelEventArgs) Handles opdg.FileOk

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'MsgBox("Proses Upload Selesai,Silahkan Cek di Menu WADS")
        btnbrows.Enabled = True
        flagproses = False
        'btnProses.Enabled = True
        'btnminimin.Visible = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If flagproses = True Then
            MsgBox("Proses Upload sedang berlangsung, Harap tidak mematikan aplikasi terlebih dulu")
            Exit Sub
        Else
            Dim result2 As DialogResult = MessageBox.Show("Keluar ?",
        "Question ?",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)
            If result2 = DialogResult.Yes Then
                Me.Close()
                'FrmMenu.ShowDialog()
                'Application.Exit()
            Else
                Exit Sub
            End If
        End If
    End Sub




    Sub buka_new()
        Call kunci()
        dgminmax.DataSource = Nothing
        dgminmax.Rows.Clear()
        dgminmax.Columns.Clear()
        ListView2.Clear()
        Label8.Text = 0
        Label9.Text = 0
        lbtotrows.Text = 0
        Label13.Text = 0
        Label15.Text = 0
        txtlocal.Text = ""
        txtnoto.Text = ""
        TextBox1.Text = ""
    End Sub

    Private Sub BtnValidasi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnValidasi.Click
        sql = "SELECT * from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
        'sql = "exec spSnUploadto 'Cektmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        'Dim sntemp As Integer
        If Not RsConn.EOF Then
            sql = "delete from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
            'sql = "exec spSnUploadto 'hapustmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            flagproses = True
            pbload.Visible = True
            btnbrows.Enabled = False
            btnProses.Enabled = False
            BtnValidasi.Enabled = False
            Button1.Enabled = False
            BackgroundWorker1.RunWorkerAsync()
        Else
            sql = "SELECT * from SMITblSnMutasiTmp where nomormutasi='" & txtnoto.Text & "'"
            'sql = "exec spSnUploadto 'Ceknoto','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            'Dim sntemp As Integer
            If Not RsConn.EOF Then
                sql = "SELECT * from SMITblSnMutasiTmp where nomormutasi='" & txtnoto.Text & "' and iduser<>'" & StrNamaUser & "'"
                RsConn = Conn.Execute(sql)
                If Not RsConn.EOF Then
                    namauser = RsConn("iduser").Value
                    'lbtotrows.Text = RsConn("jmlsn").Value
                    MsgBox("Nomor Mutasi '" & txtnoto.Text & "' Sedang di proses user '" & namauser & "'")
                    BtnValidasi.Enabled = False
                    Call buka_new()
                End If

            Else
                flagproses = True
                pbload.Visible = True
                btnbrows.Enabled = False
                btnProses.Enabled = False
                BtnValidasi.Enabled = False
                Button1.Enabled = False
                BackgroundWorker1.RunWorkerAsync()

            End If
        End If

    End Sub

    Private Sub cekqty()
        sql = "SELECT count(nomormutasi) as jmlsn from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
        'sql = "exec spSnUploadto 'cekjmlsntmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jml_qty = RsConn("jmlsn").Value
            lbtotrows.Text = RsConn("jmlsn").Value
        End If
        RsConn.Close()

        sql = "SELECT count(distinct(idproduk)) as jmlsku from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
        'sql = "exec spSnUploadto 'cekjmlskutmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jml_item = RsConn("jmlsku").Value
            Label13.Text = RsConn("jmlsku").Value
        End If
        RsConn.Close()

        If jml_item <> jmlitem Or jml_qty <> jmlqty Then
            MsgBox("Jumlah SKU atau Total qty lpb dgn file excel tdk sama, Mohon Perbaiki file excel dan Browse Ulang.....")

            Call cektemp()
            btnProses.Enabled = False
            Button1.Enabled = True
            Exit Sub
        Else
            MsgBox("Validasi Data Sukses, Silahkan Klik Upload....")
            btnProses.Enabled = True
            'BtnValidasi.Enabled = True
            Button1.Enabled = True
        End If


    End Sub


    Sub tampilkansku()

        Dim intqty As Integer

        'Dim strTgl As Date
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        ListView2.Columns.Add("Nomor Mutasi", 100)
        ListView2.Columns.Add("Kode Movment", 100)
        ListView2.Columns.Add("Nama Movment", 130)
        ListView2.Columns.Add("Kode Produk", 100)
        ListView2.Columns.Add("Nama Panjang", 340)
        ListView2.Columns.Add("QTY", 100)
        strsql = "SELECT a.nomorMutasi,c.kodeMovment,d.namamovment,b.kodeproduk,b.namapanjang,a.qty from MutasiSaldoDetail a JOIN MstProduk b on a.idproduk=b.idproduk JOIN MutasiSaldoHeader c on  a.nomormutasi=c.nomormutasi and a.iddc=c.iddc JOIN  MstMovmentProduk d on c.kodemovment=d.kodemovment WHERE b.idsnqr=1 and a.nomormutasi='" & txtnoto.Text & "'"
        RsConn = Conn.Execute(strsql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnomutasi = RsConn("NomorMutasi").Value
                strkdmutasi = RsConn("Kodemovment").Value
                strmmutasi = RsConn("namamovment").Value
                strkdproduk = RsConn("kodeproduk").Value
                strnamaproduk = RsConn("namapanjang").Value
                intqty = RsConn("qty").Value

                Dim arr(5) As String
                Dim itm As ListViewItem

                arr(0) = strnomutasi
                arr(1) = strkdmutasi
                arr(2) = strmmutasi
                arr(3) = strkdproduk
                arr(4) = strnamaproduk
                arr(5) = intqty


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()

            Loop
            TextBox1.Text = strmmutasi
        End If
        RsConn.Close()

    End Sub
    Sub hitungskuTO()
        'cari total item terhadap to tersebut

        sql = "SELECT count(a.nomormutasi) as jumlahitem from MutasiSaldoDetail a JOIN MstProduk b on a.idproduk=b.idproduk WHERE b.idsnqr=1 and a.nomormutasi='" & txtnoto.Text & "'"
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            jmlitem = RsConn("jumlahitem").Value
            Label8.Text = RsConn("jumlahitem").Value

        End If
        RsConn.Close()

        sql = "SELECT isnull(sum(a.qty),0) as Jumlahqty from MutasiSaldoDetail a JOIN MstProduk b on a.idproduk=b.idproduk WHERE b.idsnqr=1 and a.nomormutasi='" & txtnoto.Text & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then

            jmlqty = RsConn("jumlahqty").Value
            Label9.Text = RsConn("jumlahqty").Value
            If jmlqty > 1 Then
            End If
        End If
        RsConn.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call buka_new()
        txtnoto.Text = ""

        txtnoto.Text = FrmFind.cari("Mutasi")
        If txtnoto.Text = "" Or txtnoto.Text = 0 Then
        Else
            jumka = 0
            jumka = txtnoto.Text.Count

            Call tampilkansku()
            Call hitungskuTO()
            btnbrows.Enabled = True
            Call cektodipakai()
        End If
    End Sub
    Sub cektodipakai()
        sql = "SELECT * from SMITblSnMutasiTmp where nomormutasi='" & txtnoto.Text & "' and iduser<>'" & lbnama.Text & "'"
        'strsql = "exec spSnUploadto 'Ceknoto','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            namauser = RsConn("iduser").Value
            'lbtotrows.Text = RsConn("jmlsn").Value
            MsgBox("NomorMutasi '" & txtnoto.Text & "' Sedang di proses user '" & namauser & "'")
            BtnValidasi.Enabled = False
            Call buka_new()
        End If
    End Sub
    ' pass 191141163
    Private Sub txtnoto_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtnoto.TextChanged
        If txtnoto.Text <> "" Then
            Button1.Enabled = True
            'Call cektodipakai()
            If txtnoto.Text = "0" Then
                Button1.Enabled = False
            End If
        Else

            Button1.Enabled = False

        End If
    End Sub
    Sub cektotemp()
        sql = "SELECT * from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
        'strsql = "exec spSnUploadto 'Cektmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            sql = "delete from SMITblSnMutasiTmp where iduser='" & StrNamaUser & "' and nomormutasi='" & txtnoto.Text & "'"
            'strsql = "exec spSnUploadto 'hapustmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            MsgBox("Cancel Validasi Mutasi Berhasil...")
            Call buka_new()
        Else
            MsgBox("Nomor Mutasi Validasi Tidak Ditemukan...")
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call cektotemp()
    End Sub
End Class
