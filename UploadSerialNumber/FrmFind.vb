Imports System.Data.SqlClient
Public Class FrmFind
    Dim Conn, ConnMDB As New ADODB.Connection
    Dim RsConn, RsMdb As New ADODB.Recordset
    Dim sql As String
    Dim StrReturnValue, StrReturnValue1, StrFrmPemanggil As String
    Dim strkode, strnama, strfind, strNama2, strtglLpb, strpo, strnomutasi, strkdmutasi, strnmmutasi, strket As String
    Dim strtgl2, strtgl As Date
    Dim dr As SqlDataReader
    Dim cmd As SqlCommand
    Public Function cari(ByVal FrmPemanggil As String) As String
        StrReturnValue = 0
        Me.TopMost = True
        StrFrmPemanggil = FrmPemanggil
        Me.ShowDialog()
        cari = StrReturnValue
    End Function
   


    
    Private Sub LoadNomorTOtes()
       
    End Sub
 
    Private Sub ListView2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView2.SelectedIndexChanged
        Dim z As Integer
        z = ListView2.SelectedItems.Count

        If z = 0 Then
            Exit Sub
        Else



            'cari
            If StrFrmPemanggil = "NoTOManual" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                Me.Close()
            End If
            If StrFrmPemanggil = "Mutasi" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                Me.Close()
            End If

            If StrFrmPemanggil = "Nolpb" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                Vnopo = ListView2.SelectedItems.Item(0).SubItems(1).Text
                Me.Close()
            End If
            If StrFrmPemanggil = "Retur_toko" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                'StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(2).Text
                Frmreturtoko.TextBox1.Text = ListView2.SelectedItems.Item(0).SubItems(2).Text
                Me.Close()
            End If

            If StrFrmPemanggil = "Mutasidcout" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                'StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(2).Text
                Frmupmutasidcout.TextBox1.Text = ListView2.SelectedItems.Item(0).SubItems(2).Text
                Me.Close()
            End If


            If StrFrmPemanggil = "Mutasidcin" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                iddcpgn = ListView2.SelectedItems.Item(0).SubItems(4).Text
                'Frmupmutasidcin.TextBox1.Text = ListView2.SelectedItems.Item(0).SubItems(1).Text
                Me.Close()
            End If

            If StrFrmPemanggil = "ReturSupplier" Then
                StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(0).Text
                'StrReturnValue = ListView2.SelectedItems.Item(0).SubItems(2).Text
                'Frmupmutasidcin.TextBox1.Text = ListView2.SelectedItems.Item(0).SubItems(1).Text
                Me.Close()
            End If
        End If
    End Sub

    Private Sub TxtFind_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtFind.KeyPress
        If (e.KeyChar Like "[',]") Then e.Handled() = True

    End Sub

    Private Sub TxtFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtFind.TextChanged

        'Call cek()

    End Sub

    Private Sub cek()


        If StrFrmPemanggil = "NoTOManual" Then
            Button1.Text = "Cari Nama Toko"

            Call LoadNomorTO()
        End If

        If StrFrmPemanggil = "Nolpb" Then
            Button1.Text = "Cari Nama Supplier"

            Call LoadNomorLPB()
        End If

        If StrFrmPemanggil = "Retur_toko" Then
            Button1.Text = "Cari Nama Toko"
            Call LoadNoReturtoko()
        End If

        If StrFrmPemanggil = "Mutasi" Then
            Button1.Text = "Cari Nama Toko"
            Call LoadNoMutasi_gsbs()
        End If

        If StrFrmPemanggil = "Mutasidcout" Then
            Button1.Text = "Cari Nama Toko"
            Call LoadNoMutasidcout()
        End If


        If StrFrmPemanggil = "Mutasidcin" Then
            Button1.Text = "Cari Nama Toko"
            Call LoadNoMutasidcin()
        End If
        If StrFrmPemanggil = "ReturSupplier" Then
            Button1.Text = "Cari Nama Supplier"
            Call LoadNomorRetursupplier()
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If StrFrmPemanggil = "NoTOManual" Then
            Call LoadNomorTO()
        End If

        If StrFrmPemanggil = "Nolpb" Then
            Call LoadNomorLpb()
        End If

        If StrFrmPemanggil = "Retur_toko" Then
            Call LoadNoReturtoko()
        End If

        If StrFrmPemanggil = "ReturSupplier" Then
            Call LoadNomorRetursupplier()
        End If
    End Sub
    Private Sub LoadNomorTO()
        Label1.Text = "Nama Toko"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("No.TO", 70)
        ListView2.Columns.Add("Tgl TO", 80)
        ListView2.Columns.Add("Kode Toko", 80)
        ListView2.Columns.Add("Nama Toko", 210)

        If StrFrmPemanggil = "NoTOManual" Then
            ListView2.Columns.Add("JenisPB", 90)
        End If

        If StrFrmPemanggil = "NoTOManual" Then
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'Draft' as keterangan from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' ORDER BY a.tglto desc"
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Manual' as JenisPB from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' union SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Otomatis' as JenisPB from ToKeTokoHeader a JOIN ToKeTokoDetail b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' order by a.tglto desc"
            sql = "exec spFindSerialumber 'CariTO','',0,'%" & TxtFind.Text & "%'"
        End If
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strkode = RsConn("NomorTO").Value
                If StrFrmPemanggil = "NoTOManTMP" Or StrFrmPemanggil = "NoTObyPBTMP" Then
                    strtgl = Now.Date
                Else
                    strtgl = RsConn("tglTO").Value
                End If
                strnama = RsConn("kodetoko").Value
                strNama2 = RsConn("namatoko").Value


                Dim arr(5) As String
                Dim itm As ListViewItem

                arr(0) = strkode
                arr(1) = strtgl.Date
                arr(2) = strnama
                arr(3) = strNama2

                If StrFrmPemanggil = "NoTOManual" Then
                    arr(4) = RsConn("JenisPB").Value
                End If

                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub

    Private Sub LoadNoReturtoko()
        Label1.Text = "Nama Toko"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("No.Retur", 70)
        ListView2.Columns.Add("Tgl Retur", 80)
        ListView2.Columns.Add("Kode Toko", 80)
        ListView2.Columns.Add("Nama Toko", 210)

        'If StrFrmPemanggil = "NoTOManual" Then
        '    ListView2.Columns.Add("JenisPB", 90)
        'End If

        If StrFrmPemanggil = "Retur_toko" Then
            'sql = "SELECT DISTINCT(a.nomorRetur),d.tglretur,a.kodeToko,b.namatoko from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk JOIN ReturTokoKeDcHeader d on a.nomorRetur=d.nomorRetur and a.kodeToko=d.kodeToko WHERE c.idsnqr=1 and b.namaToko like '%" & TxtFind.Text & "%' order by d.tglretur desc"
            'sql = "SELECT DISTINCT a.nomorRetur,a.kodeToko,d.namatoko,a.tglretur from ReturTokoKeDcHeader a JOIN SMITblReturSNkeDC b on a.nomorRetur=b.nomorretur and a.kodeToko=b.kodeToko JOIN ReturTokoKeDcDetail c on a.nomorRetur=c.nomorretur and a.kodeToko=c.kodetoko JOIN MstToko d on a.kodeToko=d.kodeToko WHERE a.statusData=1 and a.nomorRetur not in(SELECT nomorRetur from SMITblSNReceiptRetur WHERE nomorRetur=a.nomorRetur and kodeToko=a.kodeToko)"
            sql = "exec spFindSerialumber 'CariReturToko','',0,'%" & TxtFind.Text & "%'"
        End If
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strkode = RsConn("NomorRetur").Value

                strtgl = RsConn("tglRetur").Value

                strkdtokocek = RsConn("kodetoko").Value
                strNama2 = RsConn("namatoko").Value


                Dim arr(5) As String
                Dim itm As ListViewItem

                arr(0) = strkode
                arr(1) = strtgl.Date
                arr(2) = strkdtokocek
                arr(3) = strNama2


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub
    Private Sub LoadNomorLpb()
        Label1.Text = "Nomor Lpb"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("Nomor Lpb", 100)
        ListView2.Columns.Add("Nomor PO", 100)
        ListView2.Columns.Add("Tanggal Lpb", 130)
        ListView2.Columns.Add("Nama Supplier", 270)


        'sql = "exec spLpbDariTokoHeaderDetail 'GetNoLpb',0,0,0,0,'%" & strfind & "%'"
        'strsql = "SELECT * from LpbSupplierHeader WHERE nomorLpb in(SELECT DISTINCT nomorLpb from LpbSupplierDetail WHERE idproduk in(SELECT idProduk from MstProduk WHERE idsnQr=1)) and nomorPo not in(SELECT DISTINCT nomorPo from SMITblBankSN) ORDER BY tglcreate desc"
        'sql = "SELECT a.nomorLpb,a.nomorpo,a.tglcreate as tgllpb,c.namaSupplier from LpbSupplierHeader a JOIN PoDcHeader b on a.nomorpo=b.nomorPo JOIN MstSupplier c on b.idSupplier=c.idSupplier WHERE a.nomorLpb in(SELECT DISTINCT nomorLpb from LpbSupplierDetail WHERE idproduk in(SELECT idProduk from MstProduk WHERE idsnQr=1)) and a.nomorPo not in(SELECT DISTINCT nomorPo from SMITblBankSN) and c.namasupplier like '%" & TxtFind.Text & "%' ORDER BY tgllpb desc"
        sql = "exec spFindSerialumber 'CariLPB','',0,'%" & TxtFind.Text & "%'"
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnama = RsConn("nomorLpb").Value
                strpo = RsConn("nomorpo").Value
                strtgl = RsConn("tglLpb").Value
                strNama2 = RsConn("namasupplier").Value

                Dim arr(4) As String
                Dim itm As ListViewItem

                arr(0) = strnama
                arr(1) = strpo
                arr(2) = strtgl
                arr(3) = strNama2

                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub
    Private Sub LoadNomorRetursupplier()
        Label1.Text = "Nomor Retur"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("Nomor Retur", 100)
        ListView2.Columns.Add("Tanggal Retur", 100)
        ListView2.Columns.Add("Kode Supplier", 130)
        ListView2.Columns.Add("Nama Supplier", 270)


        'sql = "exec spLpbDariTokoHeaderDetail 'GetNoLpb',0,0,0,0,'%" & strfind & "%'"
        'strsql = "SELECT * from LpbSupplierHeader WHERE nomorLpb in(SELECT DISTINCT nomorLpb from LpbSupplierDetail WHERE idproduk in(SELECT idProduk from MstProduk WHERE idsnQr=1)) and nomorPo not in(SELECT DISTINCT nomorPo from SMITblBankSN) ORDER BY tglcreate desc"
        'sql = "SELECT a.nomorLpb,a.nomorpo,a.tglcreate as tgllpb,c.namaSupplier from LpbSupplierHeader a JOIN PoDcHeader b on a.nomorpo=b.nomorPo JOIN MstSupplier c on b.idSupplier=c.idSupplier WHERE a.nomorLpb in(SELECT DISTINCT nomorLpb from LpbSupplierDetail WHERE idproduk in(SELECT idProduk from MstProduk WHERE idsnQr=1)) and a.nomorPo not in(SELECT DISTINCT nomorPo from SMITblBankSN) and c.namasupplier like '%" & TxtFind.Text & "%' ORDER BY tgllpb desc"
        sql = "exec spFindSerialumber 'CariRetursupplier','',0,'%" & TxtFind.Text & "%'"
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnama = RsConn("nomorretur").Value
                strpo = RsConn("kodesupplier").Value
                strtgl = RsConn("tglretur").Value
                strNama2 = RsConn("namasupplier").Value

                Dim arr(4) As String
                Dim itm As ListViewItem

                arr(0) = strnama

                arr(1) = strtgl
                arr(2) = strpo
                arr(3) = strNama2

                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub

    'LoadNoMutasi_gsbs
    Private Sub LoadNoMutasi_gsbs()
        Label1.Text = "Nama Toko"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("No.Mutasi", 70)
        ListView2.Columns.Add("Tgl Mutasi", 80)
        ListView2.Columns.Add("Kode Movment", 80)
        ListView2.Columns.Add("Nama Movement", 210)
        ListView2.Columns.Add("Keterangan", 210)
        'If StrFrmPemanggil = "NoTOManual" Then
        '    ListView2.Columns.Add("JenisPB", 90)
        'End If

        If StrFrmPemanggil = "Mutasi" Then
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'Draft' as keterangan from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' ORDER BY a.tglto desc"
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Manual' as JenisPB from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' union SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Otomatis' as JenisPB from ToKeTokoHeader a JOIN ToKeTokoDetail b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' order by a.tglto desc"
            sql = "exec spFindSerialumber 'Carimutasi','',0,'%" & TxtFind.Text & "%'"
        End If
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnomutasi = RsConn("nomormutasi").Value
                'If StrFrmPemanggil = "NoTOManTMP" Or StrFrmPemanggil = "NoTObyPBTMP" Then
                '    strtgl = Now.Date
                'Else
                '    strtgl = RsConn("tglTO").Value
                'End If
                strtgl = RsConn("tglmutasi").Value
                strkdmutasi = RsConn("kodemovment").Value
                strnmmutasi = RsConn("namamovment").Value
                strket = RsConn("keterangan").Value

                Dim arr(5) As String
                Dim itm As ListViewItem

                arr(0) = strnomutasi
                arr(1) = strtgl.Date
                arr(2) = strkdmutasi
                arr(3) = strnmmutasi
                arr(4) = strket


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub
    'LoadNoMutasidcout
    Private Sub LoadNoMutasidcout()
        Label1.Text = "Nama Toko"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("No.Mutasi", 70)
        ListView2.Columns.Add("Dc Pengirim", 120)
        ListView2.Columns.Add("Dc Penerima", 120)
        ListView2.Columns.Add("Tgl Mutasi", 210)
        'ListView2.Columns.Add("Keterangan", 210)
        'If StrFrmPemanggil = "NoTOManual" Then
        '    ListView2.Columns.Add("JenisPB", 90)
        'End If

        If StrFrmPemanggil = "Mutasidcout" Then
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'Draft' as keterangan from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' ORDER BY a.tglto desc"
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Manual' as JenisPB from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' union SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Otomatis' as JenisPB from ToKeTokoHeader a JOIN ToKeTokoDetail b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' order by a.tglto desc"
            sql = "exec spFindSerialumber 'Carimutasidcout',''," & IdDC & ",'%" & TxtFind.Text & "%'"
        End If
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnomutasi = RsConn("nomutasi").Value
                'If StrFrmPemanggil = "NoTOManTMP" Or StrFrmPemanggil = "NoTObyPBTMP" Then
                '    strtgl = Now.Date
                'Else
                '    strtgl = RsConn("tglTO").Value
                'End If
                strtgl = RsConn("tglmutasi").Value
                strkdmutasi = RsConn("dcpengirim").Value
                strnmmutasi = RsConn("dcpenerima").Value
                'strket = RsConn("keterangan").Value

                Dim arr(5) As String
                Dim itm As ListViewItem

                arr(0) = strnomutasi
                arr(1) = strkdmutasi
                arr(2) = strnmmutasi
                arr(3) = strtgl.Date
                'arr(4) = strket


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub

    Private Sub LoadNoMutasidcin()
        Label1.Text = "Nama Toko"
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True

        If TxtFind.Text = "" Then
            strfind = "%"
        Else
            strfind = TxtFind.Text
        End If

        ListView2.Columns.Add("No.Mutasi", 70)
        ListView2.Columns.Add("Dc Pengirim", 120)
        ListView2.Columns.Add("Dc Penerima", 120)
        ListView2.Columns.Add("Tgl Mutasi", 210)
        ListView2.Columns.Add("Iddc", 5)
        'If StrFrmPemanggil = "NoTOManual" Then
        '    ListView2.Columns.Add("JenisPB", 90)
        'End If

        If StrFrmPemanggil = "Mutasidcin" Then
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'Draft' as keterangan from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' ORDER BY a.tglto desc"
            'sql = "SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Manual' as JenisPB from ToKeTokoHeaderManual a JOIN ToKeTokoDetailManual b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' union SELECT DISTINCT a.nomorto,a.tglto,a.kodetoko,c.namatoko,'PB Otomatis' as JenisPB from ToKeTokoHeader a JOIN ToKeTokoDetail b on a.nomorto=b.nomorto and a.kodetoko=b.kodetoko JOIN MstToko c on a.kodetoko=c.kodetoko WHERE a.statusData=0 and b.idProduk in(SELECT idProduk from MstProduk WHERE idsnqr=1) and a.nomorto not in(SELECT DISTINCT nomorto from SMITblSNTOkeToko) and c.namatoko like '%" & TxtFind.Text & "%' order by a.tglto desc"
            sql = "exec spFindSerialumber 'Carimutasidcin',''," & IdDC & ",'%" & TxtFind.Text & "%'"
        End If
        RsConn = Conn.Execute(sql)

        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                strnomutasi = RsConn("nomutasi").Value
                'If StrFrmPemanggil = "NoTOManTMP" Or StrFrmPemanggil = "NoTObyPBTMP" Then
                '    strtgl = Now.Date
                'Else
                '    strtgl = RsConn("tglTO").Value
                'End If
                strtgl = RsConn("tglmutasi").Value
                strkdmutasi = RsConn("dcpengirim").Value
                strnmmutasi = RsConn("dcpenerima").Value
                iddcpgn = RsConn("iddcpengirim").Value
                'strket = RsConn("keterangan").Value

                Dim arr(5) As String
                Dim itm As ListViewItem

                arr(0) = strnomutasi
                arr(1) = strkdmutasi
                arr(2) = strnmmutasi
                arr(3) = strtgl.Date
                arr(4) = iddcpgn
                'arr(4) = strket


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()
            Loop

        End If
        RsConn.Close()
    End Sub
    Private Sub FrmFind_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TxtFind.Clear()
        If Conn.State = 0 Then
            Call GetStringKoneksi()
            Conn.Open(StrKoneksi)
        End If

        Call cek()
    End Sub


   
  
End Class