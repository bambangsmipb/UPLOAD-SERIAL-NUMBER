Imports System.IO
Imports System.Data.SqlClient
Imports System.Threading
Imports System.ComponentModel
Public Class Frmreturtoko
    Dim Conn, ConnMDB As New ADODB.Connection
    Dim RsConn, RsMdb As New ADODB.Recordset
    Dim jmlitem, jmlqty, jml_item, jml_qty As Integer
    Dim sql, passx, nomorpb, nomorsn, kodeproduk, namapanjang, namauser, kdretur, nmretur, sqlid As String
    Private MouseIsDown As Boolean = False
    Private MouseIsDownLoc As Point = Nothing
    Dim oReader As StreamReader
    Dim minor, maxst, kdtk, nilaigrid, sts, npersent, jumka, idstok As Integer
    Dim kdprd As String
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
        sql = "SELECT count(kodetoko) as jmlrow from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and iduser='" & StrNamaUser & "' and kodetoko='" & TextBox1.Text & "'"
        'sql = "exec spSnUploadto 'cekjmlrowtmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"


        RsConn = Conn.Execute(sql)
        Dim sntemp As Integer
        If Not RsConn.EOF Then
            sntemp = RsConn("jmlrow").Value
            If sntemp > 0 Then
                sql = "delete from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
                'sql = "delete from SMITblSNReceiptReturTmp"
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


        Dim aCol() As Integer = {7, 13, 10, 25, 9, 18, 13}
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
        MyAdapter = New System.Data.OleDb.OleDbDataAdapter("Select RTRIM(LTRIM(kodetoko)),RTRIM(LTRIM(namatoko)),RTRIM(LTRIM(Kodeproduk)),RTRIM(LTRIM(Namapanjang)),RTRIM(LTRIM(nomorsn)),'' as CekKdToko,'' as ceksn from [SheetReturtoko$]", MyConnection)
        Ds = New System.Data.DataSet
        MyAdapter.Fill(Ds)
        Me.dgminmax.DataSource = Ds.Tables(0)

        For icol = 0 To 6
            dgminmax.Columns(icol).Width = (dgminmax.Width / 100) * aCol(icol)
            Select Case icol


                Case 0
                    dgminmax.Columns(icol).HeaderText = "Kode Toko"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 1
                    dgminmax.Columns(icol).HeaderText = "Nama Toko"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 2
                    dgminmax.Columns(icol).HeaderText = "Kode Produk"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 3
                    dgminmax.Columns(icol).HeaderText = "Nama Panjang"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 4
                    dgminmax.Columns(icol).HeaderText = "Nomor SN"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 5
                    dgminmax.Columns(icol).HeaderText = "Cek KdToko dan SKU"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 6
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
        Dim sql23 As String
        sql1 = "select * from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and iduser='" & StrNamaUser & "' and kodetoko='" & TextBox1.Text & "'"
        'sql1 = "exec spSnUploadto 'cekjmlrowtmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql1)
        If Not RsConn.EOF() Then
            'sql1 = "insert into SMITblSNReceiptRetur select kodetoko,nomorretur,idproduk,nomorsn,iduser,tglimport,koderetur,namaretur,idjenisstok,null  from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and iduser='" & StrNamaUser & "' and kodetoko='" & TextBox1.Text & "'"
            sql1 = "exec spSnUploadReceiptReturtoko 'insertRetur','" & TextBox1.Text & "','" & txtnoretur.Text & "','" & kdprd & "','" & nomorsn & "',0,'" & StrNamaUser & "','" & kdretur & "','" & nmretur & "','" & idstok & "'"
            RsConn = Conn.Execute(sql1)


            sql22 = "update SMITblBankSN set statusdata=1 where nomorsn in(select nomorsn from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "' and iduser='" & StrNamaUser & "') and statusdata=5"
            RsConn = Conn.Execute(sql22)

            sql23 = "update SMITblBankSN set statusdata=7 where nomorsn in(select nomorsn from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "' and iduser='" & StrNamaUser & "') and statusdata=4"
            RsConn = Conn.Execute(sql23)

            MsgBox("Upload Data Sukses..")
            'RsConn.Close()
            sql = "delete from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
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
        Label3.Text = "UPLOAD SERIAL NUMBER RETUR DARI TOKO"
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
                kodetoko = 0
            Else
                kodetoko = row.Cells.Item(0).Value
            End If
            If IsDBNull(row.Cells.Item(2).Value) Then
                kdprd = ""
            Else
                kdprd = row.Cells.Item(2).Value
            End If
            If IsDBNull(row.Cells.Item(4).Value) Then
                nomorsn = ""
            Else
                nomorsn = row.Cells.Item(4).Value
            End If


            'cek PO di excel harus sama dengan po yang di buka LPB



            sql = "SELECT * from SMITblReturSNkeDC WHERE nomorretur='" & txtnoretur.Text & "' and kodetoko='" & kodetoko & "' and kodeproduk in(select kodeproduk from mstproduk where kodeproduk='" & kdprd & "' and idsnqr=1)"
            RsConn = Conn.Execute(sql)
            'If kodetoko = 0 Or kdprd = 0 Or nomorsn = "" Then
            If kodetoko = 0 Or kdprd = "" Or nomorsn = "" Then
            Else

                If Not RsConn.EOF() Then


                    row.Cells.Item(5).Value = "KDTOKO/SKU OK"
                    Dim colorDown As Color = Color.LightGreen
                    row.Cells.Item(5).Style.BackColor = colorDown
                    row.Cells.Item(5).Selected = True
                Else

                    row.Cells.Item(5).Value = "KDTOKO/SKU TIDAK ADA"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(5).Style.BackColor = colorDown
                    row.Cells.Item(5).Selected = True
                End If

                Label15.Text = (Label15.Text + 1)
                'Label2.Text = Label15.Text
                'lbtotrows.Text = (lbtotrows.Text - 1)
            End If
            RsConn.Close()


            sql = "SELECT * from SMITblBankSN WHERE nomorsn='" & nomorsn & "' and statusdata in(4,5) and idproduk in(SELECT idproduk from MstProduk  WHERE idsnqr=1)"
            'sql = "exec spSnUploadto 'CekSn','" & txtnoto.Text & "',0,'',0,''"
            RsConn = Conn.Execute(sql)
            If kodetoko = 0 Or kdprd = "" Or nomorsn = "" Then
            Else

                If Not RsConn.EOF() Then
                    sql = "SELECT nomorsn from SMITblSNReceiptReturTmp where nomorsn='" & nomorsn & "'"
                    'sql = "exec spSnUploadto 'CekSnduplikat',0,0,'" & nomorsn & "',0,''"
                    RsConn = Conn.Execute(sql)
                    If Not RsConn.EOF Then
                        row.Cells.Item(6).Value = "SN DUPLIKAT"
                        Dim colorDown As Color = Color.LightSalmon
                        row.Cells.Item(6).Style.BackColor = colorDown
                        row.Cells.Item(6).Selected = True
                    Else

                        sql = "SELECT nomorsn from SMITblBankSN where nomorsn='" & nomorsn & "' and statusdata not in(4,5)"
                        RsConn = Conn.Execute(sql)
                        If Not RsConn.EOF Then
                            row.Cells.Item(6).Value = "STATUS SN TDK BISA DI RETUR"
                            Dim colorDown As Color = Color.LightSalmon
                            row.Cells.Item(6).Style.BackColor = colorDown
                            row.Cells.Item(6).Selected = True
                        Else
                            sqlid = "SELECT a.nomorRetur,a.kodeToko,b.namatoko,a.kodeProduk,a.namaPanjang,a.serialNumber,a.koderetur,e.namaretur,e.idjenisstok from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk JOIN MstRetur e on a.koderetur=e.koderetur WHERE c.idsnqr=1 and a.nomorRetur='" & txtnoretur.Text & "' and a.kodetoko='" & TextBox1.Text & "' and a.kodeProduk='" & kdprd & "' and a.serialNumber='" & nomorsn & "'"
                            RsConn = Conn.Execute(sqlid)
                            If Not RsConn.EOF Then
                                'MsgBox("sss")
                                kdretur = RsConn("koderetur").Value
                                nmretur = RsConn("namaretur").Value
                                idstok = RsConn("idjenisstok").Value
                                row.Cells.Item(6).Value = "SN OK"
                                Dim colorDown As Color = Color.LightGreen
                                row.Cells.Item(6).Style.BackColor = colorDown
                                row.Cells.Item(6).Selected = True
                            Else
                                row.Cells.Item(6).Value = "SN Tidak Bisa RETUR"
                                Dim colorDown As Color = Color.LightSalmon
                                row.Cells.Item(6).Style.BackColor = colorDown
                                row.Cells.Item(6).Selected = True



                            End If
                        End If


                    End If


                Else
                    row.Cells.Item(6).Value = "SN Tidak Bisa RETUR"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(6).Style.BackColor = colorDown
                    row.Cells.Item(6).Selected = True


                End If
            End If
            RsConn.Close()

            If kodetoko = 0 Or kdprd = "" Or nomorsn = "" Then
            Else

                If row.Cells.Item(6).Value = "SN OK" And row.Cells.Item(5).Value = "KDTOKO/SKU OK" Then
                    Dim sql1 As String
                    sql1 = "insert into SMITblSNReceiptReturTmp values('" & kodetoko & "','" & txtnoretur.Text & "',(select idproduk from mstproduk where kodeproduk ='" & kdprd & "'),'" & nomorsn & "','" & StrNamaUser & "',getdate(),'" & kdretur & "','" & nmretur & "','" & idstok & "',null)"
                    'sql1 = "exec spSnUploadReceiptReturtoko 'insertReturTmp','" & kodetoko & "','" & txtnoretur.Text & "','" & kdprd & "','" & nomorsn & "',0,'" & StrNamaUser & "','" & kdretur & "','" & nmretur & "','" & idstok & "'"
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
        txtnoretur.Text = ""
        TextBox1.Text = ""
    End Sub





    Private Sub BtnValidasi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnValidasi.Click
        sql = "SELECT * from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
        'sql = "exec spSnUploadto 'Cektmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        'Dim sntemp As Integer
        If Not RsConn.EOF Then
            sql = "delete from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
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
            sql = "SELECT * from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
            'sql = "exec spSnUploadto 'Ceknoto','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            'Dim sntemp As Integer
            If Not RsConn.EOF Then
                sql = "SELECT * from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "' and iduser<>'" & StrNamaUser & "'"
                RsConn = Conn.Execute(sql)
                If Not RsConn.EOF Then
                    namauser = RsConn("iduser").Value
                    'lbtotrows.Text = RsConn("jmlsn").Value
                    MsgBox("NomorRetur '" & txtnoretur.Text & "' Sedang di proses user '" & namauser & "'")
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
        sql = "SELECT count(nomorsn) as jmlsn from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
        'sql = "exec spSnUploadto 'cekjmlsntmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jml_qty = RsConn("jmlsn").Value
            lbtotrows.Text = RsConn("jmlsn").Value
        End If
        RsConn.Close()

        sql = "SELECT count(distinct(idproduk)) as jmlsku from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
        'sql = "exec spSnUploadto 'cekjmlskutmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jml_item = RsConn("jmlsku").Value
            Label13.Text = RsConn("jmlsku").Value
        End If
        RsConn.Close()

        If jml_item <> jmlitem Or jml_qty <> jmlqty Then
            MsgBox("Jumlah SKU atau Total qty dgn file excel tdk sama, Mohon Perbaiki file excel dan Browse Ulang.....")

            Call cektemp()
            btnProses.Enabled = False
            Button1.Enabled = True
            Exit Sub
        Else
            MsgBox("Validasi Data Sukses, Silahkan Klik Upload....")
            btnProses.Enabled = True
            Button1.Enabled = True
        End If


    End Sub


    Sub hitungskuRetur()
        'cari total item terhadap to tersebut

        sql = "SELECT count(DISTINCT(idProduk)) as jumlahitem from SMITblReturSNkeDC WHERE idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and nomorRetur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jmlitem = RsConn("jumlahitem").Value
            Label8.Text = RsConn("jumlahitem").Value

        End If
        RsConn.Close()

        sql = "SELECT count(nomorRetur) as Jumlahqty from SMITblReturSNkeDC WHERE idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and nomorRetur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"

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
        txtnoretur.Text = ""

        txtnoretur.Text = FrmFind.cari("Retur_toko")
        If txtnoretur.Text = "" Or txtnoretur.Text = 0 Then
        Else
            jumka = 0
            jumka = txtnoretur.Text.Count

            Call tampilkansku()
            Call hitungskuRetur()
            btnbrows.Enabled = True
            Call cektodipakai()
        End If
    End Sub

    Sub tampilkansku()
        Dim strsql, strnmtoko, strkdsku, strnamasku, serial, strkdretur, strnmretur As String
        Dim inoretur, ikdtoko, itjnsstok As Integer

        'Dim strTgl As Date
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        ListView2.Columns.Add("Nomor Retur", 80)
        ListView2.Columns.Add("Kode Toko", 80)
        ListView2.Columns.Add("Nama Toko", 250)
        ListView2.Columns.Add("Kode produk", 80)
        ListView2.Columns.Add("Nama Panjang", 320)
        ListView2.Columns.Add("Serial Number", 80)
        ListView2.Columns.Add("Kd Retur", 60)
        ListView2.Columns.Add("Nama Retur", 180)
        ListView2.Columns.Add("Id JNS", 60)


        'strsql = "SELECT a.nomorRetur,a.kodeToko,b.namatoko,a.kodeProduk,a.namaPanjang,a.serialNumber from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk WHERE c.idsnqr=1 and a.nomorRetur='" & txtnoretur.Text & "'"
        'strsql = "SELECT a.nomorRetur,a.kodeToko,b.namatoko,a.kodeProduk,a.namaPanjang,a.serialNumber,d.koderetur,e.namaretur,e.idjenisstok from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk JOIN ReturTokoKeDcDetail d on a.nomorretur=d.nomorretur  and a.kodetoko=d.kodetoko and a.idproduk=d.idproduk JOIN MstRetur e on d.koderetur=e.koderetur WHERE c.idsnqr=1 and a.nomorRetur='" & txtnoretur.Text & "' and a.kodetoko='" & TextBox1.Text & "'"
        strsql = "SELECT a.nomorRetur,a.kodeToko,b.namatoko,a.kodeProduk,a.namaPanjang,a.serialNumber,e.koderetur,e.namaretur,e.idjenisstok from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk JOIN MstRetur e on a.koderetur=e.koderetur WHERE c.idsnqr=1 and a.nomorRetur='" & txtnoretur.Text & "' and a.kodetoko='" & TextBox1.Text & "'"
        RsConn = Conn.Execute(strsql)
        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                inoretur = RsConn("NomorRetur").Value
                ikdtoko = RsConn("Kodetoko").Value
                strnmtoko = RsConn("namatoko").Value
                strkdsku = RsConn("Kodeproduk").Value
                strnamasku = RsConn("NamaPanjang").Value
                serial = RsConn("serialnumber").Value
                strkdretur = RsConn("koderetur").Value
                strnmretur = RsConn("namaretur").Value
                itjnsstok = RsConn("idjenisstok").Value

                Dim arr(8) As String
                Dim itm As ListViewItem

                arr(0) = inoretur
                arr(1) = ikdtoko
                arr(2) = strnmtoko
                arr(3) = strkdsku
                arr(4) = strnamasku
                arr(5) = serial
                arr(6) = strkdretur
                arr(7) = strnmretur
                arr(8) = itjnsstok


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()

            Loop

        End If
        RsConn.Close()

    End Sub
    Sub cektodipakai()
        sql = "SELECT * from SMITblSNReceiptReturTmp where nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "' and iduser<>'" & lbnama.Text & "'"
        'strsql = "exec spSnUploadto 'Ceknoto','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            namauser = RsConn("iduser").Value
            'lbtotrows.Text = RsConn("jmlsn").Value
            MsgBox("NomorRetur '" & txtnoretur.Text & "' Sedang di proses user '" & namauser & "'")
            BtnValidasi.Enabled = False
            Call buka_new()
        End If
    End Sub
    ' pass 191141163
    Private Sub txtnoto_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtnoretur.TextChanged
        If txtnoretur.Text <> "" Then
            Button1.Enabled = True
            'Call cektodipakai()
            If txtnoretur.Text = "0" Then
                Button1.Enabled = False
            End If
        Else

            Button1.Enabled = False

        End If
    End Sub
    Sub cektotemp()
        sql = "SELECT * from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
        'strsql = "exec spSnUploadto 'Cektmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            sql = "delete from SMITblSNReceiptReturTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoretur.Text & "' and kodetoko='" & TextBox1.Text & "'"
            'strsql = "exec spSnUploadto 'hapustmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            MsgBox("Cancel Validasi Retur Berhasil...")
            Call buka_new()
            'SMITblSNReceiptReturTmp
        Else
            MsgBox("Nomor Retur Validasi Tidak Ditemukan...")
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call cektotemp()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
