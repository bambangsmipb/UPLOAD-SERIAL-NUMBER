Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Threading
Imports System.Globalization
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Module PublicMod
    Public StrKoneksi, MDBKoneksi, StrFilter, PathMDB, StrKoneksi2, strsql, strsqlftp, nmaplikasi, nmbarux, nmlamax As String
    Public StrSA, StrPwd, StrDSN, StrDB, strNamaLokasi, VarBagian As String
    Public StrIdUser, VarTypeConfig, VarLoc, VarRegional As Integer
    Public StrUserid, StrNamaUser, termOP, Vnopo, namauser, strkdtokocek, strstatus As String
    Public servername, jenispromo As String
    Public ConnMDB, ConnConfig As New ADODB.Connection
    Public RsMDB, RsConfig As New ADODB.Recordset
    Public tglserver As Date
    Public iddiv, iddept, idsubdept, idctgr, iddcpgn As Integer
    Public iddcpengirim As Integer
    Public SQLConn As SqlConnection
    Public ConnSQLClient As String
    Public SQLConnStatus As Boolean
    Public bg, wp, atas, masuk, logo, statusToko, strnomutasi As String
    Public idsupjwk, iddcjwk, nomorUPO, nomorTO, nomorpo, Boleh, NomorPB As Int64
    Public idProduk, idjenispo As Integer
    Public NP, kodeproduk As Int64
    Public NamaPT, kodedc, namadc, alamatdc, telpdc, kotadc, nilaiterbilang, NamaSupplier, PKP, Ptag As String
    Public NamaProduk, namatoko, kodetipetoko, alamattoko, telptoko, kodejnslokasi, namapanjang, minorder, maxstok As String
    Public IdPT, IdDC, IdSupplier, IdPajak, kodetoko, LT As Integer
    Public harga, disk1, disk2, harganet, pajak, pajakpersen, hargakotor As Decimal
    Public hostname, passvalidasi As String
    Public vStatusPathMdb As Boolean
    Public vStatusKoneksiDc As Boolean
    Public POok As Boolean
    Public jumlahmenit, hasil, DurasiClosing, JamClosing, MenitClosing As Double
    Public icoadd, icoedit, icodelete, icoapprove, icosave, icoprint, icocancel, icoclear, icorefresh, icoproses, strnmmutasi As String
    Public FsavePsn, flagbefor, flagmain, flagstatusForm, flagloginClose As Boolean

    Public conn As OleDbConnection
    Public cmd As OleDbCommand
    Public RD As OleDbDataReader
    Public DA As OleDbDataAdapter
    Public DS As DataSet
    Public str As String

    Public Class MyCustomException
        Inherits System.ApplicationException
    End Class
    Public Sub jam()
        Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("id-ID")
    End Sub


    Public Sub symbol()
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")
    End Sub

    Public Sub GetMDBKoneksi()

        PathMDB = New System.IO.FileInfo(Application.ExecutablePath).DirectoryName
        MDBKoneksi = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & PathMDB & "\posCfg.mdb;Uid=Admin;Pwd=mypos234"

        ConnMDB.Open(MDBKoneksi)
        RsMDB = ConnMDB.Execute("select * from config where loc = 1")
        StrSA = RsMDB("UID").Value
        StrPwd = RsMDB("PID").Value
        StrDSN = RsMDB("DSN").Value
        servername = RsMDB("DSN").Value
        StrDB = RsMDB("DB").Value

        If ConnMDB.State = 1 Then
            ConnMDB.Close()
        End If
        If RsMDB.State = 1 Then
            RsMDB.Close()
        End If

        'Call getConnServer("DC", StrSA, StrPwd, StrDSN, StrDB)
        'If vStatusKoneksiDc = False Then
        '    frmConfig.ShowDialog()
        '    FrmLogin.Hide()
        '    Exit Sub
        'End If

    End Sub
    Public Sub GetJamClosing()
        If ConnConfig.State = 0 Then
            GetStringKoneksi()
            ConnConfig.Open(StrKoneksi)
        End If
        strsql = "select * from TblClosingDc"
        RsConfig = ConnConfig.Execute(strsql)
        JamClosing = RsConfig("JamClosing").Value
        MenitClosing = RsConfig("MenitClosing").Value
        DurasiClosing = RsConfig("DurasiClosing").Value
    End Sub
    Public Sub getConnServer(ByVal xLokasi As String, ByVal xUserID As String, ByVal xPassID As String, ByVal xServer As String, ByVal xDB As String)
        Dim connStr As String
        Dim rs As New ADODB.Recordset
        'Dim y As String

        On Error GoTo errGetconn

        Application.DoEvents()
        connStr = "Provider=SQLOLEDB.1;Password='" & xPassID & "';Persist Security Info=True;User ID='" & xUserID & "';Initial Catalog='" & xDB & "';Data Source='" & xServer & "' "
        If xLokasi = "DC" Then
            If ConnConfig.State = 1 Then ConnConfig.Close()
            ConnConfig.Open(connStr)
            rs.Open("exec spTest", ConnConfig, adOpenDynamic, adLockOptimistic)
            MsgBox(Err.Number)
        End If



errGetconn:

        If Err.Number = 0 Then

            If xLokasi = "DC" Then
                vStatusKoneksiDc = True
            End If

        Else
     
            If xLokasi = "DC" Then
                vStatusKoneksiDc = False
           
            End If

        End If
    End Sub




    Public Sub GetStringKoneksi()
        Try
            GetMDBKoneksi()
            StrKoneksi = "Provider=SQLOLEDB.1;Password=" & StrPwd & ";Persist Security Info=True;User ID=" & StrSA & ";Initial Catalog=" & StrDB & ";Data Source=" & StrDSN
            ConnSQLClient = "Data Source=" & StrDSN & ";Initial Catalog=" & StrDB & ";User ID=" & StrSA & ";Password=" & StrPwd & ""


        Catch ex As Exception
        End Try
    End Sub

    Public Sub gambar()

        PathMDB = New System.IO.FileInfo(Application.ExecutablePath).DirectoryName
        bg = System.IO.Path.Combine(PathMDB, "image\background2.jpg")
        wp = System.IO.Path.Combine(PathMDB, "image\wallpaper.jpg")
        atas = System.IO.Path.Combine(PathMDB, "image\header.jpg")
        masuk = System.IO.Path.Combine(PathMDB, "image\photo.jpg")
        logo = System.IO.Path.Combine(PathMDB, "image\smi.jpg")

        icoadd = System.IO.Path.Combine(PathMDB, "image\add.png")
        icoapprove = System.IO.Path.Combine(PathMDB, "image\approve.png")
        icodelete = System.IO.Path.Combine(PathMDB, "image\delete.png")
        icoedit = System.IO.Path.Combine(PathMDB, "image\edit.png")
        icosave = System.IO.Path.Combine(PathMDB, "image\save.png")
        icocancel = System.IO.Path.Combine(PathMDB, "image\cancel.png")
        icoprint = System.IO.Path.Combine(PathMDB, "image\print.png")
        icoclear = System.IO.Path.Combine(PathMDB, "image\clear.png")
        icorefresh = System.IO.Path.Combine(PathMDB, "image\refresh.png")
        icoproses = System.IO.Path.Combine(PathMDB, "image\proses.png")

    End Sub
    Public Sub NamaPerusahaan()
        GetStringKoneksi()
        If ConnMDB.State = 0 Then
            ConnMDB.Open(StrKoneksi)
        End If

        strsql = "Select NamaPerusahaan from mstperusahaan "
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            NamaPT = RsConfig("NamaPerusahaan").Value
        Else
            MsgBox("Perusahaan belum di daftarkan..", vbOKOnly + vbCritical, "Info")
            Exit Sub
        End If

    End Sub
    Public Sub alamatftp111()
        GetStringKoneksi()
        If ConnMDB.State = 0 Then
            ConnMDB.Open(StrKoneksi)
        End If

        strsqlftp = "Select * from mstftp where aplikasi='uploadserialnumber'"
        RsConfig = ConnMDB.Execute(strsqlftp)
        If Not RsConfig.EOF Then
            nmaplikasi = RsConfig("alamatftp").Value
            nmbarux = RsConfig("nmlama").Value
            nmlamax = RsConfig("nmbaru").Value
        Else

            End
            'MsgBox("Perusahaan belum di daftarkan..", vbOKOnly + vbCritical, "Info")
            'Exit Sub
        End If
        'ss
    End Sub
    Public Sub alamatftp()
        GetStringKoneksi()
        If ConnMDB.State = 0 Then
            ConnMDB.Open(StrKoneksi)
        End If
        Try
            strsqlftp = "Select * from mstftp where aplikasi='uploadserialnumber' and st_cekftp=1 and versi='" & System.Windows.Forms.Application.ProductVersion & "'"
            'strsqlftp = "Select * from mstftp where aplikasi='uploadserialnumber' and st_cekftp=1"
            RsConfig = ConnMDB.Execute(strsqlftp)
            If Not RsConfig.EOF Then
                'tt
                nmaplikasi = RsConfig("alamatftp").Value
                nmbarux = RsConfig("nmbaru").Value
                nmlamax = RsConfig("nmlama").Value
                strstatus = RsConfig("versi").Value
                If strstatus <> System.Windows.Forms.Application.ProductVersion Then

                    Dim exeServ As DateTime = GetExeFTP()
                    'Toleransi perbedaan jam dengan server = 2 jam
                    If exeServ > exeDate.AddHours(1) Then
                        MsgBox("Ada versi baru diserver, tanggal " & exeServ.ToString & " dibanding " & exeDate.ToString, MsgBoxStyle.Information)
                        Form1.ShowDialog()
                    End If
                Else


                End If
            Else
                MsgBox("Aplikasi Tidak di Izinkan Hub IT.!", vbOKOnly + vbInformation, "Info")
                End
                'MsgBox("Perusahaan belum di daftarkan..", vbOKOnly + vbCritical, "Info")
                'Exit Sub
            End If
        Catch ex As Exception
            Application.Exit()
        End Try
        'tttt
        'ee
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


    End Sub

    Public Sub GetServerDate()
        If ConnConfig.State = 0 Then
            GetStringKoneksi()
            ConnConfig.Open(StrKoneksi)
        End If
        strsql = "select convert(date,GETDATE()) as Tgl"
        RsConfig = ConnConfig.Execute(strsql)
        tglserver = RsConfig("Tgl").Value

    End Sub

    Public Sub pesan(ByVal kode As Single)
        If kode = 9 Then
            MsgBox("Data sudah pernah ada dengan status: Delete " & vbCrLf _
                 & "Silahkan hub Administrator !!", vbOKOnly + vbExclamation, "Info")
        ElseIf kode = 1 Then
            MsgBox("Data sudah ada ..." & vbCrLf _
                & "Silahkan cek kembali!!", vbOKOnly + vbExclamation, "Info")
        ElseIf kode = 3 Then
            MsgBox("Data berhasil disimpan ", vbOKOnly + vbInformation, "Info")
        ElseIf kode = 0 Then
            MsgBox("Data yang anda masukan belum lengkap !!" & vbCrLf _
                & "Silahkan lengkapi data..", vbOKOnly + vbExclamation, "Info")
        ElseIf (kode = 99) Then
            MsgBox("Data tidak terdaftar..", vbOKOnly + vbCritical, "Info")
        ElseIf kode = 100 Then
            MsgBox("Data Masih Kosong", vbOKOnly + vbCritical, "Info")
        End If

    End Sub

    Public Sub GetKoneksiSQLClient()
        Try
            SQLConn = New SqlConnection(ConnSQLClient)
            SQLConn.Open()
            SQLConnStatus = False
        Catch ex As Exception
            SQLConnStatus = True
        End Try
    End Sub

    Public Sub getSupplier(ByVal kode As String)

        strsql = "select a.*,b.namajenisperusahaan from mstsupplier a " & _
                 "left join MstJenisPerusahaan b on a.idjenisperusahaan=b.idjenisperusahaan " & _
                 " where kodesupplier='" & kode & "'"
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            PKP = RsConfig("namajenisperusahaan").Value
            IdSupplier = RsConfig("idsupplier").Value
            NamaSupplier = RsConfig("namasupplier").Value
            LT = RsConfig("leadtime").Value
            termOP = RsConfig("termOffpayment").Value
        End If
    End Sub

    Public Sub getNamaProduk(ByVal kode As String, ByVal idsup As Int16)
        If idsup = 0 Then

            strsql = "Select a.*,0 as harga,0 as disk1,0 as disk2,0 as harganet,0 as pajak,0 as hargakotor,0 as nilaipajakprosen from mstproduk a  where kodeproduk =" & kode
        Else
            strsql = "Select a.*,b.*,c.nilaiPajakProsen  from mstproduk a " & _
                    "inner join mstproduksupplier b on a.idproduk=b.idproduk  " & _
                    "inner join MstPajak c on a.idPajak =c.idPajak " & _
                   " where kodeproduk ='" & kode & "' and idsupplier=" & IdSupplier
        End If
        '" & kode & "'"
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            idProduk = RsConfig("idproduk").Value
            NamaProduk = RsConfig("namapanjang").Value
            harga = RsConfig("harga").Value
            disk1 = RsConfig("disk1").Value
            disk2 = RsConfig("disk2").Value
            harganet = RsConfig("harganet").Value
            pajak = RsConfig("pajak").Value
            hargakotor = RsConfig("hargakotor").Value
            Ptag = RsConfig("kodetag").Value
            pajakpersen = RsConfig("nilaiPajakProsen").Value
        Else
            idProduk = 0
        End If
    End Sub

    Public Sub combox(ByRef comboname As ComboBox)
        With comboname

            .AutoCompleteSource = Windows.Forms.AutoCompleteSource.ListItems
            .AutoCompleteMode = Windows.Forms.AutoCompleteMode.Suggest

        End With

        Dim index As Integer = comboname.FindString(comboname.Text)
        If index > -1 Then
            comboname.SelectedIndex = index
        Else
            Call pesan(99)
            comboname.Text = ""
            Beep()
            comboname.Focus()
            Exit Sub
        End If
    End Sub

    Public Sub GetIdSupplier(ByVal KodeSupplier As String)
        strsql = "Select idSupplier,namaSupplier from MstSupplier where KodeSupplier='" & KodeSupplier & "'"
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            IDSupplier = RsConfig("idSupplier").Value
            NamaSupplier = RsConfig("namaSupplier").Value
        End If
    End Sub

    Public Sub getjenispo()
        POok = False
        strsql = "Select count(*) as ttl from mstjenisposupplier"
        RsConfig = ConnMDB.Execute(strsql)
        If RsConfig("ttl").Value > 0 Then
            POok = True
        Else
            POok = False
        End If
    End Sub

    Public Sub GetIdProduk(ByVal KodeProduk As String)
        strsql = "Select dbo.FcGetIdProduk('" & KodeProduk & "')IdProduk"
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            idProduk = RsConfig("IdProduk").Value
        End If
    End Sub

    Public Sub gettoko(ByVal kodetoko As Int64)
        strsql = "Select namatoko,kodetipetoko,alamattoko,telepon,kodejenislokasi,kodestatustoko from msttoko where kodetoko=" & kodetoko
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            namatoko = RsConfig("namatoko").Value
            kodetipetoko = RsConfig("kodetipetoko").Value
            alamattoko = RsConfig("AlamatToko").Value
            telptoko = RsConfig("telepon").Value
            kodejnslokasi = RsConfig("kodejenislokasi").Value
            statusToko = RsConfig("kodestatustoko").Value
        End If
    End Sub

    Public Sub scanx(ByVal barcode As String)
        strsql = " select a.idproduk,kodeTipeToko,a.statusData ,b.kodeProduk ,b.namaPanjang,b.idpajak  from MstProdukTipeToko a " & _
                 " inner join MstProduk b on a.idProduk =b.idProduk and b.idJenisProduk=1 " & _
                 " inner join MstTagProduk c on b.kodeTag =c.kodeTag and c.flagDcTo = 1 " & _
                 " where a.kodeTipeToko ='" & kodetipetoko & "' and a.statusData =1 and ( b.barcode ='" & barcode & "' or b.kodeproduk='" & barcode & "')"
        RsConfig = ConnMDB.Execute(strsql)
        If Not RsConfig.EOF Then
            kodeproduk = RsConfig("kodeproduk").Value
            idProduk = RsConfig("IdProduk").Value
            NamaProduk = RsConfig("namapanjang").Value
            IdPajak = RsConfig("idpajak").Value
        Else
            idProduk = 0
        End If
    End Sub
    Public Sub scan(ByVal barcode As String)
        Try
            GetStringKoneksi()
            Using Con As New SqlConnection(ConnSQLClient)
                Con.Open()
                strsql = " select a.idproduk,kodeTipeToko,a.statusData ,b.kodeProduk ,b.namaPanjang,b.idpajak  from MstProdukTipeToko a " & _
               " inner join MstProduk b on a.idProduk =b.idProduk and b.idJenisProduk=1 " & _
               " inner join MstTagProduk c on b.kodeTag =c.kodeTag and c.flagDcTo = 1 " & _
               " where a.kodeTipeToko ='" & kodetipetoko & "' and a.statusData =1 and ( b.barcode ='" & barcode & "' or b.kodeproduk='" & barcode & "')"
                Using Com As New SqlCommand(strsql, Con)
                    Using RDR = Com.ExecuteReader()
                        RDR.Read()
                        If RDR.HasRows Then
                            kodeproduk = RDR.Item("kodeProduk")
                            idProduk = RDR.Item("idproduk")
                            NamaProduk = RDR.Item("namaPanjang")
                            IdPajak = RDR.Item("idpajak")

                        Else
                            idProduk = 0
                        End If
                    End Using
                End Using
                Con.Close()
            End Using

        Catch ex As SqlException
            MsgBox("Koneksi ke server terputus..!", vbOKOnly + vbInformation, "Info")
        End Try


    End Sub


    Public Sub ClearAllTextBoxes(ByVal ctl As Control)
        For Each MyControl As Control In ctl.Controls
            If TypeOf MyControl Is TextBox Then
                MyControl.Text = ""
            End If
            If MyControl.HasChildren Then ClearAllTextBoxes(MyControl)
        Next
    End Sub

    Public Sub namakomputer()
        hostname = My.Computer.Name
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> _
   <System.Runtime.CompilerServices.Extension()> _
    Public Function DataTable(ByVal sender As BindingSource) As DataTable
        Return DirectCast(sender.DataSource, DataTable)
    End Function
End Module
