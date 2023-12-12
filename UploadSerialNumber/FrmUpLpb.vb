Imports System.IO
Imports System.Data.SqlClient
Imports System.Threading
Imports System.ComponentModel
Public Class FrmUploaderlpb
    Dim Conn, ConnMDB As New ADODB.Connection
    Dim RsConn, RsMdb As New ADODB.Recordset
    Dim jmlitem, jmlqty, jml_item, jml_qty As Integer
    Dim sql, passx, nomorpb, nomorsn, kodeproduk, namapanjang As String
    Private MouseIsDown As Boolean = False
    Private MouseIsDownLoc As Point = Nothing
    Dim oReader As StreamReader
    Dim minor, maxst, kdtk, nilaigrid, sts, npersent, kodeproduk1 As Integer
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


        Dim aCol() As Integer = {15, 13, 8, 30, 15, 9}
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
        MyAdapter = New System.Data.OleDb.OleDbDataAdapter("Select RTRIM(LTRIM(Nomorpo)),RTRIM(LTRIM(NomorSN)),RTRIM(LTRIM(Kodeproduk)),RTRIM(LTRIM(Namapanjang)),'' as cekpo,'' as ceksn from [Sheetlpb$]", MyConnection)
        Ds = New System.Data.DataSet
        MyAdapter.Fill(Ds)
        Me.dgminmax.DataSource = Ds.Tables(0)

        For icol = 0 To 5
            dgminmax.Columns(icol).Width = (dgminmax.Width / 100) * aCol(icol)
            Select Case icol


                Case 0
                    dgminmax.Columns(icol).HeaderText = "Nomor Po"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 1
                    dgminmax.Columns(icol).HeaderText = "Nomor SN"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 2
                    dgminmax.Columns(icol).HeaderText = "Kode Produk"
                    dgminmax.Columns(icol).ReadOnly = True
                Case 3
                    dgminmax.Columns(icol).HeaderText = "Nama Panjang"
                    dgminmax.Columns(icol).ReadOnly = True

                Case 4
                    dgminmax.Columns(icol).HeaderText = "Cek PO/SKU"
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
    End Sub
  
    Private Sub btnProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProses.Click
        'flagproses = True
        'pbload.Visible = True
        'btnbrows.Enabled = False
        'btnProses.Enabled = False
        'BackgroundWorker1.RunWorkerAsync()
        Dim sql1 As String

        'sql1 = "insert into SMITblBankSN select nomorpo,nomorlpb,idproduk,nomorsn,statusdata,iduser,tglimport,tglupdate  from SMITblBankSNtmp where nomorlpb='" & TXTNOLPB.Text & "' and nomorpo='" & Vnopo & "' and iduser='" & StrNamaUser & "'"
        sql1 = "exec spSnUploadlpb 'insertBankSN','0','" & TXTNOLPB.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql1)
        MsgBox("Upload Data Sukses..")
        'sql = "delete from SMITblBankSNtmp where iduser='" & StrNamaUser & "' and nomorlpb='" & TXTNOLPB.Text & "'"
        sql = "exec spSnUploadlpb 'hapustmp','0','" & TXTNOLPB.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        Call hitung()
        Call buka_new()
    End Sub

    Private Sub kunci()
        btnbrows.Enabled = False
        btnProses.Enabled = False
        BtnValidasi.Enabled = False
    End Sub

    Private Sub buka()
        btnbrows.Enabled = True
        'btnProses.Enabled = True


    End Sub
    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'Call getPathMdb()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Label16.Text = Format(Now, "dd-MM-yyyy")
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
        Label3.Text = "UPLOAD SERIAL NUMBER LPB"
        Call buka_new()
        'Label16.Text = Format(Now, "dd-MM-yyyy")
    End Sub


    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Call cek()
    End Sub
    Sub cek()
        Dim x As Integer = dgminmax.CurrentCellAddress.X
        Dim y As Integer = dgminmax.CurrentCellAddress.Y
        pbload.Minimum = 0
        pbload.Maximum = nilaigrid
        pbload.Value = 0
        'Call cektemp()
        'Call insertemp()
        For Each row As DataGridViewRow In Me.dgminmax.Rows

            'Try
            If IsDBNull(row.Cells.Item(0).Value) Then
                nomorpo = 0
            Else
                nomorpo = row.Cells.Item(0).Value
            End If
            If IsDBNull(row.Cells.Item(1).Value) Then
                nomorsn = ""
            Else
                nomorsn = row.Cells.Item(1).Value
            End If
            If IsDBNull(row.Cells.Item(2).Value) Then
                kdprd = ""
                'kodeproduk1 = ""
            Else
                kdprd = row.Cells.Item(2).Value
                'kodeproduk1 = row.Cells.Item(2).Value
            End If
            If IsDBNull(row.Cells.Item(3).Value) Then
                namapanjang = ""
            Else
                namapanjang = row.Cells.Item(3).Value
            End If




            sql = "SELECT * from SMITblBankSN WHERE nomorsn='" & nomorsn & "'"
            'sql = "exec spSnUploadlpb 'ceksnsudahada',0,0,0,'" & nomorsn & "',0,''"
            RsConn = Conn.Execute(sql)
            If nomorsn = "" Or nomorpo = 0 Or kdprd = "" Or namapanjang = "" Then
            Else

                If Not RsConn.EOF() Then
                    row.Cells.Item(5).Value = "SN SUDAH ADA"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(5).Style.BackColor = colorDown
                    row.Cells.Item(5).Selected = True
                Else
                    sql = "SELECT nomorsn from SMITblBankSNtmp where nomorsn='" & nomorsn & "'"
                    'sql = "exec spSnUploadlpb 'ceksnduplikat',0,0,0,'" & nomorsn & "',0,''"
                    RsConn = Conn.Execute(sql)

                    If Not RsConn.EOF Then


                        row.Cells.Item(5).Value = "SN DUPLIKAT"
                        Dim colorDown As Color = Color.LightSalmon
                        row.Cells.Item(5).Style.BackColor = colorDown
                        row.Cells.Item(5).Selected = True
                    Else


                        Dim t, s As String
                        Dim spasi As Integer
                        t = nomorsn
                        'spasi = 0
                        s = Len(t) - Len(Replace(t, " ", ""))
                        spasi = s
                        If spasi > 0 Then
                            row.Cells.Item(5).Value = "SN ADA SPASI"
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
            End If
            RsConn.Close()
            Dim sqlsku As String

            sqlsku = "SELECT * from LpbSupplierHeader a join LpbSupplierDetail b on a.nomorlpb=b.nomorlpb and a.nomorpo=b.nomorpo WHERE a.nomorlpb='" & TXTNOLPB.Text & "' and a.nomorpo='" & nomorpo & "' and b.idProduk in(select idproduk from mstproduk where kodeproduk='" & kdprd & "' and idsnqr=1)"
            'sqlsku = "exec spSnUploadlpb 'cekposku','" & nomorpo & "','" & TXTNOLPB.Text & "','" & kdprd & "','',0,'','2017-01-01','2017-01-01'"
            RsConn = Conn.Execute(sqlsku)
            If nomorsn = "" Or nomorpo = 0 Or kdprd = "" Or namapanjang = "" Then
            Else
                '
                If Not RsConn.EOF() Then
                    'jika ada
                    '191141163
                    row.Cells.Item(4).Value = "PO OK dan SKU OK"
                    Dim colorDown As Color = Color.LightGreen
                    row.Cells.Item(4).Style.BackColor = colorDown
                    row.Cells.Item(4).Selected = True
                Else
                    'jika tidak ada
                    row.Cells.Item(4).Value = "PO/SKU TIDAK ADA"
                    Dim colorDown As Color = Color.LightSalmon
                    row.Cells.Item(4).Style.BackColor = colorDown
                    row.Cells.Item(4).Selected = True

                End If
                Label15.Text = (Label15.Text + 1)
            End If
            RsConn.Close()
            If nomorsn = "" Or nomorpo = 0 Or kdprd = "" Or namapanjang = "" Then
            Else

                If row.Cells.Item(4).Value = "PO OK dan SKU OK" And row.Cells.Item(5).Value = "SN OK" Then



                    Dim sql1 As String
                    sql1 = "insert into SMITblBankSNtmp values('" & nomorpo & "','" & TXTNOLPB.Text & "',(select idproduk from mstproduk where kodeproduk ='" & kdprd & "'),'" & nomorsn & "',1,'" & StrNamaUser & "',getdate(),getdate())"
                    'sql1 = "exec spSnUploadlpb 'inserttmp'," & nomorpo & ",'" & TXTNOLPB.Text & "'," & kdprd & ",'" & nomorsn & "',1,'" & StrNamaUser & "'"
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

        btnbrows.Enabled = True
        flagproses = False


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
                'Application.Exit()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Label4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Label4.Click

    End Sub



    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub btnCetak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

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
        TXTNOLPB.Text = ""
        txtlocal.Text = ""

        Call hitung()
    End Sub


    Private Sub tampilkanskusn()
        Dim strsql, strnama, strkd, strnopo, strnolpb As String
        Dim intqty As Integer

        'Dim strTgl As Date
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        ListView2.Columns.Add("Nomor PO", 100)
        ListView2.Columns.Add("Nomor LPB", 100)
        ListView2.Columns.Add("Kode produk", 100)
        ListView2.Columns.Add("Nama Panjang", 440)
        ListView2.Columns.Add("QTY", 50)

        'strsql = "SELECT a.nomorpo,a.nomorlpb,b.kodeproduk,b.namapanjang,a.qty from LpbSupplierDetail a JOIN MstProduk b on a.idproduk=b.idproduk WHERE a.nomorlpb='" & TXTNOLPB.Text & "' and a.nomorpo='" & Vnopo & "' and b.idsnqr=1"
        strsql = "exec spSnUploadlpb 'TampilskuSN','" & Vnopo & "','" & TXTNOLPB.Text & "',0,'',0,''"
        RsConn = Conn.Execute(strsql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnopo = RsConn("NomorPo").Value
                strnolpb = RsConn("NomorLpb").Value
                strkd = RsConn("Kodeproduk").Value
                strnama = RsConn("NamaPanjang").Value
                intqty = RsConn("qty").Value

                Dim arr(4) As String
                Dim itm As ListViewItem

                arr(0) = strnopo
                arr(1) = strnolpb
                arr(2) = strkd
                arr(3) = strnama
                arr(4) = intqty


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()

            Loop

        End If
        'Label17.Text = StrNamaUser
        lbnama.Text = StrNamaUser
        RsConn.Close()
    End Sub





    Private Sub cekqty()
        sql = "SELECT count(nomorsn) as jmlsn from SMITblBankSNtmp where nomorlpb='" & TXTNOLPB.Text & "' and iduser='" & StrNamaUser & "'"
        'sql = "exec spSnUploadlpb 'cekJmlSNupload',0,'" & TXTNOLPB.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            'Label11.Text = RsConn("jmlsn").Value
            jml_qty = RsConn("jmlsn").Value
            lbtotrows.Text = RsConn("jmlsn").Value
        End If
        RsConn.Close()

        sql = "SELECT count(distinct(idproduk)) as jmlsku from SMITblBankSNtmp where nomorlpb='" & TXTNOLPB.Text & "' and iduser='" & StrNamaUser & "'"
        'sql = "exec spSnUploadlpb 'cekJmlsku',0,'" & TXTNOLPB.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            'Label12.Text = RsConn("jmlsku").Value
            jml_item = RsConn("jmlsku").Value
            Label13.Text = RsConn("jmlsku").Value
        End If
        RsConn.Close()

        If jml_item <> jmlitem Or jml_qty <> jmlqty Then
            MsgBox("Jumlah SKU/Total qty lpb dgn upload tdk sama, Mohon Cek lagi dan Upload ulang..")
            'Call cektemp()
            btnProses.Enabled = False
            Button1.Enabled = True
            Exit Sub
        Else

            MsgBox("Validasi Data Sukses, Silahkan Klik Upload....")
            btnProses.Enabled = True
            Button1.Enabled = True
        End If


    End Sub
    Private Sub cmbpo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

   

    Private Sub BtnValidasi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnValidasi.Click
        'sql = "SELECT * from SMITblBankSNtmp where iduser='" & StrNamaUser & "' and nomorlpb='" & TXTNOLPB.Text & "'"
        sql = "exec spSnUploadlpb 'BatalValidasi',0,'" & TXTNOLPB.Text & "',0,'',0,'" & StrNamaUser & "'"
        RsConn = Conn.Execute(sql)
        'Dim sntemp As Integer
        If Not RsConn.EOF Then
            'sql = "delete from SMITblBankSNtmp where iduser='" & StrNamaUser & "' and nomorlpb='" & TXTNOLPB.Text & "'"
            sql = "exec spSnUploadlpb 'hapustmp',0,'" & TXTNOLPB.Text & "',0,'',0,'" & StrNamaUser & "'"
            RsConn = Conn.Execute(sql)
            flagproses = True
            pbload.Visible = True
            btnbrows.Enabled = False
            btnProses.Enabled = False
            BtnValidasi.Enabled = False
            Button1.Enabled = False
            BackgroundWorker1.RunWorkerAsync()
        Else
            'sql = "SELECT * from SMITblBankSNtmp where nomorlpb='" & TXTNOLPB.Text & "'"
            sql = "exec spSnUploadlpb 'Cekuserproses',0,'" & TXTNOLPB.Text & "',0,'',0,''"
            RsConn = Conn.Execute(sql)
            'Dim sntemp As Integer
            If Not RsConn.EOF Then
                MsgBox("NomorLPB '" & TXTNOLPB.Text & "' Sedang di proses user '" & namauser & "'")
                BtnValidasi.Enabled = False
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Sub hitung()
        sql = "SELECT count(nomorlpb) as nomorlpb from LpbSupplierHeader WHERE nomorLpb in(SELECT DISTINCT nomorLpb from LpbSupplierDetail WHERE idproduk in(SELECT idProduk from MstProduk WHERE idsnQr=1)) and nomorPo not in(SELECT DISTINCT nomorPo from SMITblBankSN)"

        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            Label14.Text = "Jumlah LPB (" & RsConn("nomorlpb").Value & ")"
        End If
        RsConn.Close()
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TXTNOLPB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTNOLPB.TextChanged
        If TXTNOLPB.Text <> "" Then
            Button1.Enabled = True
            'Call cektodipakai()
            If TXTNOLPB.Text = "0" Then
                Button1.Enabled = False
            End If
        Else

            Button1.Enabled = False

        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call buka_new()
        TXTNOLPB.Text = ""
        TXTNOLPB.Text = FrmFind.cari("Nolpb")

        If TXTNOLPB.Text = "" Or TXTNOLPB.Text = 0 Then
        Else

            Call tampilkanskusn()
            Call hitungskulpb()
            Call ceklpbpakai()
        End If
    End Sub
    Sub ceklpbpakai()
        sql = "SELECT * from SMITblBankSNTmp where nomorlpb='" & TXTNOLPB.Text & "' and iduser<>'" & lbnama.Text & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            namauser = RsConn("iduser").Value
            'lbtotrows.Text = RsConn("jmlsn").Value
            MsgBox("NomorLPB '" & TXTNOLPB.Text & "' Sedang di proses user '" & namauser & "'")
            BtnValidasi.Enabled = False
            Call buka_new()
        End If
    End Sub
    Sub hitungskulpb()
        sql = "SELECT count(a.nomorpo) as jumlahitem from LpbSupplierDetail a JOIN MstProduk b on a.idproduk=b.idproduk WHERE a.nomorlpb='" & TXTNOLPB.Text & "' and a.nomorpo='" & Vnopo & "' and b.idsnqr=1"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            jmlitem = RsConn("jumlahitem").Value
            Label8.Text = RsConn("jumlahitem").Value
            Call buka()
        End If
        RsConn.Close()

        sql = "SELECT sum(a.qty) as Jumlahqty from LpbSupplierDetail a JOIN MstProduk b on a.idproduk=b.idproduk WHERE a.nomorlpb='" & TXTNOLPB.Text & "'  and a.nomorpo='" & Vnopo & "' and b.idsnqr=1"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then

            jmlqty = RsConn("jumlahqty").Value
            Label9.Text = RsConn("jumlahqty").Value

            Call buka()
        End If
        RsConn.Close()
    End Sub
    Sub ceklpbtemp()
        sql = "SELECT * from SMITblBankSNTmp where iduser='" & StrNamaUser & "' and nomorlpb='" & TXTNOLPB.Text & "'"
        RsConn = Conn.Execute(sql)
        If Not RsConn.EOF Then
            sql = "delete from SMITblBankSNTmp where iduser='" & StrNamaUser & "' and nomorlpb='" & TXTNOLPB.Text & "'"
            RsConn = Conn.Execute(sql)
            MsgBox("Cancel Validasi LPB Berhasil...")
            Call buka_new()
        Else
            MsgBox("Nomor LPB Validasi Tidak Ditemukan...")
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call ceklpbtemp()
    End Sub
End Class
