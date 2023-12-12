Imports System.IO
Imports System.Data.SqlClient
Imports System.Threading
Imports System.ComponentModel
Public Class form4
    Dim Conn, ConnMDB As New ADODB.Connection
    Dim RsConn, RsMdb As New ADODB.Recordset
    Dim jmlitem, jmlqty, jml_item, jml_qty As Integer
    Dim sql, passx, nomorpb, nomorsn, kodeproduk, namapanjang, namauser, intidsup, intidstok, intiddc As String
    Private MouseIsDown As Boolean = False
    Private MouseIsDownLoc As Point = Nothing
    Dim oReader As StreamReader
    Dim minor, maxst, kdtk, nilaigrid, sts, npersent, jumka As Integer
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
        sql = "SELECT count(nomorretur) as jmlrow from SMITblSNReturSupplierTmp where nomorretur='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
        'sql = "exec spSnUploadto 'cekjmlrowtmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"


        RsConn = Conn.Execute(sql)
        Dim sntemp As Integer
        If Not RsConn.EOF Then
            sntemp = RsConn("jmlrow").Value
            If sntemp > 0 Then
                sql = "delete from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"
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


        Dim aCol() As Integer = {7, 13, 25, 9, 18, 20}
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
        MyAdapter = New System.Data.OleDb.OleDbDataAdapter("Select RTRIM(LTRIM(NomorRetur)),RTRIM(LTRIM(Kodeproduk)),RTRIM(LTRIM(Namapanjang)),RTRIM(LTRIM(nomorsn)),'' as CekNomorRetur,'' as Ceksn from [SheetReturSup$]", MyConnection)
        Ds = New System.Data.DataSet
        MyAdapter.Fill(Ds)
        Me.dgminmax.DataSource = Ds.Tables(0)

        For icol = 0 To 5
            dgminmax.Columns(icol).Width = (dgminmax.Width / 100) * aCol(icol)
            Select Case icol


                Case 0
                    dgminmax.Columns(icol).HeaderText = "Nomor Retur"
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
                    dgminmax.Columns(icol).HeaderText = "Cek NoRetur"
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
        sql1 = "select * from SMITblSNReturSupplierTmp where nomorretur='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
        'sql1 = "exec spSnUploadto 'cekjmlrowtmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql1)
        If Not RsConn.EOF() Then
            'sql1 = "insert into SMITblSNTOkeToko select kodetoko,nomorto,idproduk,nomorsn,iduser,tglimport,null  from SMITblSNTOkeTokoTmp where nomorto='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
            sql1 = "exec spSnUploadretursupplier 'insertretursupplier','" & txtnoto.Text & "','" & txtnoto.Text & "','" & nomorsn & "',0,'" & StrNamaUser & "','" & kodetoko & "'"
            RsConn = Conn.Execute(sql1)


            'sql22 = "update SMITblBankSN set statusdata=2 where nomorsn in(select nomorsn from SMITblSNTOkeToko where nomorto='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "')"
            sql22 = "exec spSnUploadretursupplier 'UpdateStatusSN','" & txtnoto.Text & "','" & kdprd & "','" & nomorsn & "',0,'" & StrNamaUser & "','" & kodetoko & "'"
            RsConn = Conn.Execute(sql22)

            MsgBox("Upload Data Sukses..")
            'RsConn.Close()
            sql = "delete from SMITblSNReturSupplierTmp where nomorretur='" & txtnoto.Text & "' and iduser='" & StrNamaUser & "'"
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
    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
        Label3.Text = "UPLOAD SERIAL NUMBER RETUR KE SUPPLIER"
        Call buka_new()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Call cekdata()

        'MsgBox("NomorTO ")
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
            If IsDBNull(row.Cells.Item(1).Value) Then
                kdprd = ""
            Else
                kdprd = row.Cells.Item(1).Value
            End If
            If IsDBNull(row.Cells.Item(3).Value) Then
                nomorsn = ""
            Else
                nomorsn = row.Cells.Item(3).Value
            End If


            sql = "SELECT * from ReturDcDetailKeSupplier WHERE nomorretur='" & kodetoko & "' and idproduk in(select idproduk from mstproduk where kodeproduk='" & kdprd & "' and idsnqr=1)"

            RsConn = Conn.Execute(sql)
            If kodetoko = 0 Or kdprd = "" Or nomorsn = "" Then
            Else

                If Not RsConn.EOF() Then

                    sql = "SELECT top 1 idProduk from ReturDcDetailKeSupplier WHERE (SELECT count(nomorSN) from SMITblSNReturSuppliertmp WHERE nomorretur='" & txtnoto.Text & "' and idproduk in(select idproduk from mstproduk where kodeproduk='" & kdprd & "' and idsnqr=1)) >= (SELECT qty from SMITblSNReturSupplier WHERE nomorretur='" & txtnoto.Text & "' and idproduk in(select idproduk from mstproduk where kodeproduk='" & kdprd & "' and idsnqr=1))"
                    RsConn = Conn.Execute(sql)
                    If Not RsConn.EOF Then
                        row.Cells.Item(4).Value = "LEBIH SKU"
                        Dim colorDown As Color = Color.LightSalmon
                        row.Cells.Item(4).Style.BackColor = colorDown
                        row.Cells.Item(4).Selected = True
                    Else

                        row.Cells.Item(4).Value = "NORETUR/SKU OK"
                        Dim colorDown As Color = Color.LightGreen
                        row.Cells.Item(4).Style.BackColor = colorDown
                        row.Cells.Item(4).Selected = True
                    End If
                Else

                    row.Cells.Item(4).Value = "NORETUR/SKU TIDAK ADA"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(4).Style.BackColor = colorDown
                    row.Cells.Item(4).Selected = True
                End If

                Label15.Text = (Label15.Text + 1)
            End If
            RsConn.Close()


            sql = "SELECT * from SMITblBankSN WHERE nomorsn='" & nomorsn & "'"
            'sql = "SELECT * from SMITblBankSN WHERE nomorsn='" & nomorsn & "' and statusdata in(1,5)"
            'sql = "exec spSnUploadto 'CekSn','" & txtnoto.Text & "',0,'',0,''"
            RsConn = Conn.Execute(sql)
            If kodetoko = 0 Or kdprd = "" Or nomorsn = "" Then
            Else

                If Not RsConn.EOF() Then
                    sql = "SELECT nomorsn from SMITblSNReturSuppliertmp where nomorsn='" & nomorsn & "'"
                    'sql = "exec spSnUploadto 'CekSnduplikat',0,0,'" & nomorsn & "',0,''"
                    RsConn = Conn.Execute(sql)
                    If Not RsConn.EOF Then
                        row.Cells.Item(5).Value = "SN DUPLIKAT"
                        Dim colorDown As Color = Color.LightSalmon
                        row.Cells.Item(5).Style.BackColor = colorDown
                        row.Cells.Item(5).Selected = True
                    Else

                        sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData=9"
                        'sql = "SELECT nomorsn from SMITblSNTOkeToko where nomorsn='" & nomorsn & "'"
                        'sql = "exec spSnUploadto 'CekSudahupload',0,0,'" & nomorsn & "',0,''"
                        RsConn = Conn.Execute(sql)
                        If Not RsConn.EOF Then
                            row.Cells.Item(5).Value = "SN SUDAH DIUPLOAD RETUR"
                            Dim colorDown As Color = Color.LightSalmon
                            row.Cells.Item(5).Style.BackColor = colorDown
                            row.Cells.Item(5).Selected = True
                        Else
                            sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData in(1,7) and idProduk in(select idproduk from mstproduk where kodeproduk='" & kdprd & "')"
                            'sql = "SELECT * from SMITblBankSN where  nomorsn='" & nomorsn & "' and statusData in(1,5)"
                            RsConn = Conn.Execute(sql)
                            If Not RsConn.EOF Then
                                row.Cells.Item(5).Value = "SN OK"
                                Dim colorDown As Color = Color.LightGreen
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            Else
                                row.Cells.Item(5).Value = "SN & SKU TIDAK SESUAI/CEK STATUS SN"
                                Dim colorDown As Color = Color.LightSalmon
                                row.Cells.Item(5).Style.BackColor = colorDown
                                row.Cells.Item(5).Selected = True
                            End If

                        End If

                    End If


                Else
                    row.Cells.Item(5).Value = "SN Tidak Terdaftar"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(5).Style.BackColor = colorDown
                    row.Cells.Item(5).Selected = True





                End If
            End If
            RsConn.Close()

            If kodetoko = 0 Or kdprd = "" Or nomorsn = "" Then
            Else

                If row.Cells.Item(5).Value = "SN OK" And row.Cells.Item(4).Value = "NORETUR/SKU OK" Then
                    Dim sql1 As String
                    sql1 = "insert into SMITblSNReturSupplierTMP values('" & intidsup & "','" & txtnoto.Text & "',(select idproduk from mstproduk where kodeproduk ='" & kdprd & "'),'" & nomorsn & "','" & StrNamaUser & "',getdate(),'" & intidstok & "',null,'" & intiddc & "')"
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

    End Sub

    Private Sub BtnValidasi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnValidasi.Click
        sql = "SELECT * from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"
        'sql = "exec spSnUploadto 'Cektmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        'Dim sntemp As Integer
        If Not RsConn.EOF Then
            sql = "delete from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"
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
            sql = "SELECT * from SMITblSNReturSupplierTmp where nomorretur='" & txtnoto.Text & "'"
            'sql = "exec spSnUploadto 'Ceknoto','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            'Dim sntemp As Integer
            If Not RsConn.EOF Then
                sql = "SELECT * from SMITblSNReturSupplierTmp where nomorretur='" & txtnoto.Text & "' and iduser<>'" & StrNamaUser & "'"
                RsConn = Conn.Execute(sql)
                If Not RsConn.EOF Then
                    namauser = RsConn("iduser").Value
                    'lbtotrows.Text = RsConn("jmlsn").Value
                    MsgBox("NomorRetur '" & txtnoto.Text & "' Sedang di proses user '" & namauser & "'")
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
        sql = "SELECT count(nomorsn) as jmlsn from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jml_qty = RsConn("jmlsn").Value
            lbtotrows.Text = RsConn("jmlsn").Value
        End If
        RsConn.Close()


        sql = "SELECT count(distinct(idproduk)) as jmlsku from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"
        'sql = "exec spSnUploadto 'cekjmlskutmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jml_item = RsConn("jmlsku").Value
            Label13.Text = RsConn("jmlsku").Value
        End If
        RsConn.Close()

        If jml_item <> jmlitem Or jml_qty <> jmlqty Or Label15.Text <> jmlqty Then
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
        Dim strsql, strnama, strkd, strnopo, strnolpb As String
        Dim intqty As Integer

        'Dim strTgl As Date
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        ListView2.Columns.Add("Nomor Retur", 100)
        ListView2.Columns.Add("Nama Supplier", 100)
        ListView2.Columns.Add("Kode produk", 100)
        ListView2.Columns.Add("Nama Panjang", 440)
        ListView2.Columns.Add("QTY", 50)
        ListView2.Columns.Add("Jenis Stok", 90)
        ListView2.Columns.Add("idsupplier", 0)
        ListView2.Columns.Add("iddsc", 0)
        strsql = "SELECT a.nomorRetur,c.namasupplier,d.kodeProduk,d.namaPanjang,b.qty,b.idJenisStok,a.idSupplier,a.iddc from ReturDcHeaderKeSupplier a JOIN ReturDcDetailKeSupplier b on a.idSupplier=b.idSupplier and a.nomorRetur=b.nomorRetur JOIN MstSupplier c on a.idSupplier=c.idSupplier JOIN MstProduk d on b.idProduk=d.idProduk WHERE d.idSnQr=1 and a.nomorRetur='" & txtnoto.Text & "'"


        RsConn = Conn.Execute(strsql)
        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnopo = RsConn("nomorRetur").Value
                strnolpb = RsConn("namasupplier").Value
                strkd = RsConn("Kodeproduk").Value
                strnama = RsConn("NamaPanjang").Value
                intqty = RsConn("QTY").Value
                intidstok = RsConn("idjenisstok").Value
                intidsup = RsConn("idsupplier").Value
                intiddc = RsConn("iddc").Value
                Dim arr(7) As String
                Dim itm As ListViewItem

                arr(0) = strnopo
                arr(1) = strnolpb
                arr(2) = strkd
                arr(3) = strnama
                arr(4) = intqty
                arr(5) = intidstok
                arr(6) = intidsup
                arr(7) = intiddc
                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()

            Loop

        End If
        RsConn.Close()

    End Sub
    Sub hitungskuTO()
        'cari total item terhadap to tersebut

        sql = "SELECT count(a.nomorRetur) as jumlahitem from ReturDcDetailKeSupplier a JOIN MstProduk b on a.idproduk=b.idproduk WHERE b.idsnqr=1 and a.nomorRetur='" & txtnoto.Text & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jmlitem = RsConn("jumlahitem").Value
            Label8.Text = RsConn("jumlahitem").Value

        End If
        RsConn.Close()

        sql = "SELECT isnull(sum(a.qty),0) as Jumlahqty from ReturDcDetailKeSupplier a JOIN MstProduk b on a.idproduk=b.idproduk WHERE b.idsnqr=1 and a.nomorRetur='" & txtnoto.Text & "'"

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

        txtnoto.Text = FrmFind.cari("ReturSupplier")
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
        sql = "SELECT * from SMITblSNReturSupplierTmp where nomorRetur='" & txtnoto.Text & "' and iduser<>'" & lbnama.Text & "'"
        'strsql = "exec spSnUploadto 'Ceknoto','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            namauser = RsConn("iduser").Value
            'lbtotrows.Text = RsConn("jmlsn").Value
            MsgBox("NomorRetur '" & txtnoto.Text & "' Sedang di proses user '" & namauser & "'")
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
        sql = "SELECT * from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"
        'strsql = "exec spSnUploadto 'Cektmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            sql = "delete from SMITblSNReturSupplierTmp where iduser='" & StrNamaUser & "' and nomorretur='" & txtnoto.Text & "'"
            'strsql = "exec spSnUploadto 'hapustmp','" & txtnoto.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            MsgBox("Cancel Validasi No Retur Berhasil...")
            Call buka_new()
        Else
            MsgBox("Nomor No Retur Validasi Tidak Ditemukan...")
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call cektotemp()
    End Sub
End Class
