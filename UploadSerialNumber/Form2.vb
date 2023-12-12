Imports System.IO
Imports System.Data.SqlClient
Imports System.Threading
Imports System.ComponentModel

Public Class Form2
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
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Conn.State = 0 Then
            GetStringKoneksi()
            Conn.Open(StrKoneksi)
        End If
        Control.CheckForIllegalCrossThreadCalls = False
        'dddd

        Call namadcAktif()
        lbdc.Text = namadc & ""

        lbnama.Visible = True
        lbnama.Text = StrNamaUser
        'Label3.Text = "UPLOAD SERIAL NUMBER RETUR DARI TOKO"
        'Call buka_new()
    End Sub


    Sub tampilkansku()
        Dim strsql, strnmtoko, strkdsku, strnamasku, serial, iduser, stglto, stgllpb, inolpb As String
        Dim inoto, ikdtoko, idproduk As Integer
        'Dim stgllpb As Date

        'Dim strTgl As Date
        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.View = Windows.Forms.View.Details
        ListView2.GridLines = True
        ListView2.FullRowSelect = True
        ListView2.Columns.Add("Nomor LPB", 80)
        ListView2.Columns.Add("Nomor TO", 80)
        ListView2.Columns.Add("Kode Toko", 80)
        ListView2.Columns.Add("Nama Toko", 200)
        ListView2.Columns.Add("ID produk", 80)
        ListView2.Columns.Add("Kode produk", 80)
        ListView2.Columns.Add("Nama Panjang", 320)
        ListView2.Columns.Add("Serial Number", 100)
        ListView2.Columns.Add("Status Data", 100)
        ListView2.Columns.Add("Tgl LPB", 80)
        ListView2.Columns.Add("Tgl TO", 80)


        'strsql = "SELECT a.nomorTO,a.kodetoko,b.namatoko,a.nomorSN,c.idproduk,c.kodeproduk,c.namapanjang,a.idUser,a.tglImport,isnull(a.tgldownload,'') as tgldownload from SMITblSNTOkeToko a JOIN MstToko b on a.kodeToko=b.kodeToko JOIN MstProduk c on a.idProduk=c.idproduk WHERE a.nomorsn like '%" & txtsn.Text & "%' ORDER BY a.kodeToko,a.nomorTO"
        'strsql = "SELECT a.nomorRetur,a.kodeToko,b.namatoko,a.kodeProduk,a.namaPanjang,a.serialNumber,d.koderetur,e.namaretur,e.idjenisstok from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk JOIN ReturTokoKeDcDetail d on a.nomorretur=d.nomorretur  and a.kodetoko=d.kodetoko and a.idproduk=d.idproduk JOIN MstRetur e on d.koderetur=e.koderetur WHERE c.idsnqr=1 and a.nomorRetur='" & txtnoretur.Text & "' and a.kodetoko='" & TextBox1.Text & "'"
        'strsql = "SELECT a.nomorRetur,a.kodeToko,b.namatoko,a.kodeProduk,a.namaPanjang,a.serialNumber,e.koderetur,e.namaretur,e.idjenisstok from SMITblReturSNkeDC a JOIN MstToko b on a.kodeToko=b.kodetoko JOIN MstProduk c on a.idProduk=c.idproduk JOIN MstRetur e on a.koderetur=e.koderetur WHERE c.idsnqr=1 and a.nomorRetur='" & txtnoretur.Text & "' and a.kodetoko='" & TextBox1.Text & "'"
        'strsql = "SELECT a.nomorTO,a.kodetoko,b.namatoko,a.nomorSN,c.idproduk,c.kodeproduk,c.namapanjang,a.idUser,a.tglImport from SMITblSNTOkeToko a JOIN MstToko b on a.kodeToko=b.kodeToko JOIN MstProduk c on a.idProduk=c.idproduk ORDER BY a.kodeToko,a.nomorTO"
        strsql = "exec spFindSerialumber 'carisn','',0,'%" & txtsn.Text & "%'"
        RsConn = Conn.Execute(strsql)



        If Not RsConn.EOF Then
            RsConn.MoveFirst()

            Do While Not RsConn.EOF
                inolpb = RsConn("nomorlpb").Value
                inoto = RsConn("nomorto").Value
                ikdtoko = RsConn("Kodetoko").Value
                strnmtoko = RsConn("namatoko").Value
                idproduk = RsConn("idproduk").Value
                strkdsku = RsConn("Kodeproduk").Value
                strnamasku = RsConn("NamaPanjang").Value
                serial = RsConn("nomorsn").Value
                iduser = RsConn("statusdata").Value
                stgllpb = RsConn("tgllpb").Value
                stglto = RsConn("tglto").Value
                'stgldownload = RsConn("m_Queue_rs.tgldownload").Value.ToString
                'If IsDBNull(row.Cells.Item(0).Value) Then
                '    kodetoko = 0
                'Else
                'itm.SubItems.Add(m_Queue_rs.Fields(6).Value)
                'Convert(varchar(30),'7/7/2011',102)
                Dim arr(10) As String
                Dim itm As ListViewItem
                arr(0) = inolpb
                arr(1) = inoto
                arr(2) = ikdtoko
                arr(3) = strnmtoko
                arr(4) = idproduk
                arr(5) = strkdsku
                arr(6) = strnamasku
                arr(7) = serial
                arr(8) = iduser
                arr(9) = stgllpb
                arr(10) = stglto


                itm = New ListViewItem(arr)
                ListView2.Items.Add(itm)

                RsConn.MoveNext()

            Loop
        Else
            MsgBox("Serial number tidak terdaftar.")
        End If
        RsConn.Close()
        Label6.Text = ListView2.Items.Count()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call tampilkansku()

    End Sub

    Private Sub txtsn_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsn.KeyPress
        If (e.KeyChar Like "[',]") Then e.Handled() = True
    End Sub

    Private Sub txtsn_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtsn.TextChanged

    End Sub
End Class