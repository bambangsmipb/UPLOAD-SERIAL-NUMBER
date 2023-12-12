Imports System.Data.SqlClient
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports UploadSN.Frmretursupplier

Module configMdl

    Public vKodeDc As Integer

    Public vStatusPathMdb As Boolean
    Public vStatusKoneksiPos As Boolean
    Public vStatusKoneksiDc As Boolean
    Public vStatusKoneksiHO As Boolean

    Public vPosUser, vPosPasswd, vPosDb, vPosDsn As String
    Public vDcUser, vDcPasswd, vDcDb, vDcDsn As String
    Public vHoUser, vHoPasswd, vHoDb, vHoDsn As String
    Public rsMdb As New ADODB.Recordset
    Public connMdb, connPos, connDc, connHO As New ADODB.Connection


    Public Sub getPathMdb()
        Dim pathMdb, getConnStringMdb As String
        On Error GoTo errConnMdb

        pathMdb = New System.IO.FileInfo(Application.ExecutablePath).DirectoryName
        getConnStringMdb = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & pathMdb & "\posCfg.mdb;Uid=Admin;Pwd=mypos234"

        If connMdb.State = 1 Then connMdb.Close()
        connMdb.Open(getConnStringMdb)

errConnMdb:
        If Err.Number <> 0 Then
            If Err.Number = -2147467259 Then
                MsgBox("Koneksi ke Database tidak ditemukan", vbOKOnly, "Koneksi gagal")
                vStatusPathMdb = False
                Exit Sub
            Else
                MsgBox("error conn " & Err.Number & " " & Err.Description)
                vStatusPathMdb = False
                Exit Sub
            End If
        Else
            vStatusPathMdb = True
            Call getUidServer()
        End If
    End Sub


    Public Sub getUidServer()

        rsMdb.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        If rsMdb.State = 1 Then rsMdb.Close()
        rsMdb.Open("select * from config", connMdb)
        With rsMdb
            If Not .EOF Then
                If .RecordCount > 1 Then
                    If rsMdb.State = 1 Then rsMdb.Close()
                    rsMdb.Open("delete * from config", connMdb)
                    MsgBox("Koneksi Tidak ditemukan")
                    Exit Sub
                End If

                Do
                    Select Case .Fields!loc.Value
                        Case 1
                            vHoUser = .Fields!uid.Value
                            vHoPasswd = .Fields!pid.Value
                            vHoDb = .Fields!db.Value
                            vHoDsn = .Fields!dsn.Value

                    End Select
                    .MoveNext()
                Loop Until .EOF = True

                Call getConnServer("DC", vHoUser, vHoPasswd, vHoDsn, vHoDb)
                If vStatusKoneksiDc = False Then
                    MsgBox("Koneksi Tidak ditemukan")
                    End
                Else
                    'FrmUploaderTo.Show()
                    FrmMenu.Show()
                    'End
                End If

            Else
                MsgBox("Silahkan seting koneksi ke Database terlebih dahulu", vbInformation, "Perhatian")
                'Exit Sub
            End If
        End With

    End Sub


    Public Sub getConnServer(ByVal xLokasi As String, ByVal xUserID As String, ByVal xPassID As String, ByVal xServer As String, ByVal xDB As String)
        Dim connStr As String
        Dim rs As New ADODB.Recordset
        'Dim y As String

        On Error GoTo errGetconn

        Application.DoEvents()
        connStr = "Provider=SQLOLEDB.1;Password='" & xPassID & "';Persist Security Info=True;User ID='" & xUserID & "';Initial Catalog='" & xDB & "';Data Source='" & xServer & "' "
        If xLokasi = "POS" Then
            If connPos.State = 1 Then connPos.Close()
            connPos.Open(connStr)
            rs.Open("exec spTest", connPos, adOpenDynamic, adLockOptimistic)
        Else
            If xLokasi = "DC" Then
                If connDc.State = 1 Then connDc.Close()
                connDc.Open(connStr)
                rs.Open("exec spTest", connDc, adOpenDynamic, adLockOptimistic)
            Else
                If connHO.State = 1 Then connHO.Close()
                connHO.Open(connStr)
                rs.Open("exec spTest", connHO, adOpenDynamic, adLockOptimistic)
            End If
        End If



errGetconn:
        If Err.Number = 0 Then
            If xLokasi = "POS" Then
                vStatusKoneksiPos = True
            Else
                If xLokasi = "DC" Then
                    vStatusKoneksiDc = True


                Else
                    If xLokasi = "HO" Then
                        vStatusKoneksiHO = True
                    End If
                End If
            End If
        Else
            If xLokasi = "POS" Then
                vStatusKoneksiPos = False
            Else
                If xLokasi = "DC" Then
                    vStatusKoneksiDc = False
                Else
                    If xLokasi = "HO" Then
                        vStatusKoneksiHO = False
                    End If
                End If
            End If
        End If
    End Sub


End Module
