Imports System.Net.Http
Imports System.Net
Imports System.Text
Imports System.IO
Module MdlUpdate
    Public exeServerUrl As String
    Public exeDate As DateTime = File.GetLastWriteTime(Application.ExecutablePath())
    'Public exeTemp As String = AppDomain.CurrentDomain.BaseDirectory & "UploadSerialNumber.new"
    'Public exeFTPUrl As String = "ftp://ftprms.planetban.co.id/RMS_TIGAS/SN/UploadSerialNumber.exe"
    Public userFTP As String = "fikri"
    Public passFTP As String = "fikri"
    Public Function GetExeServer() As DateTime
        Dim lastMod As DateTimeOffset
        Try
            Dim exeServerUrl = nmaplikasi
            Dim client = New HttpClient()
            Dim msg = New HttpRequestMessage(HttpMethod.Head, exeServerUrl)
            Dim resp = client.SendAsync(msg).Result
            lastMod = resp.Content.Headers.LastModified

        Catch ex As Exception
            lastMod = New DateTime(1900, 1, 1, 12, 0, 0)
            Application.Exit()
        End Try

        Dim exeServer As DateTime = lastMod.UtcDateTime

        Return exeServer
    End Function

    Public Sub DownloadExeServer()
        Dim exetemp As String = AppDomain.CurrentDomain.BaseDirectory & nmbarux
        If File.Exists(exetemp) Then File.Delete(exetemp)
        Try
            Dim webClient As New WebClient
            Dim exeServerUrl = nmaplikasi
            webClient.DownloadFile(exeServerUrl, exetemp)
        Catch ex As WebException
            MsgBox(ex.Message)

        End Try

        If Not File.Exists(exetemp) Then
            MsgBox("File server gagal didownload!", MsgBoxStyle.Critical)

        Else
            Try
                Dim exeOld As String = AppDomain.CurrentDomain.BaseDirectory & nmlamax
                If File.Exists(exeOld) Then File.Delete(exeOld)

                Rename(Application.ExecutablePath(), exeOld)
                Rename(exetemp, Application.ExecutablePath())

                MsgBox("Aplikasi akan dimatikan!", MsgBoxStyle.Information)
                Application.Exit()

            Catch ex As IO.IOException
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
    Public Function GetExeFTP() As DateTime

        Dim exeFTPUrl As String = nmaplikasi
        Dim lastMod As DateTime
        Try
            'dd
            Dim ftp As FtpWebRequest = FtpWebRequest.Create(exeFTPUrl)
            ftp.UsePassive = False
            ftp.Credentials = New NetworkCredential(userFTP, passFTP)
            ftp.Method = WebRequestMethods.Ftp.GetDateTimestamp
            Using response = CType(ftp.GetResponse(), FtpWebResponse)
                lastMod = response.LastModified
            End Using
        Catch ex As WebException
            lastMod = New DateTime(1900, 1, 1, 12, 0, 0)
            'MsgBox(ex.Message)
            MsgBox(" Aplikasi Anda Belum Update." & Environment.NewLine & " Gagal Update Karna Koneksi FTP Tidak ditemukan." & Environment.NewLine & " Hubungi IT !", MsgBoxStyle.Information)
            Application.Exit()
        End Try
        Return lastMod

    End Function

    Public Sub DownloadExeFTP()
        Dim exeTemp As String = AppDomain.CurrentDomain.BaseDirectory & nmbarux
        Dim exeFTPUrl As String = nmaplikasi
        If File.Exists(exeTemp) Then File.Delete(exeTemp)
        Try
            Dim request As FtpWebRequest = WebRequest.Create(exeFTPUrl)
            request.UsePassive = True
            request.Credentials = New NetworkCredential(userFTP, passFTP)
            request.Method = WebRequestMethods.Ftp.DownloadFile

            Using ftpStream As Stream = request.GetResponse().GetResponseStream(), fileStream As Stream = File.Create(exeTemp)
                ftpStream.CopyTo(fileStream)
            End Using

            Call ExeTempToAktif()
        Catch ex As WebException
            MessageBox.Show(ex.Message)
            Application.Exit()
        End Try
    End Sub
    Private Sub ExeTempToAktif()
        Dim exeTemp As String = AppDomain.CurrentDomain.BaseDirectory & nmbarux
        If Not File.Exists(exeTemp) Then
            MsgBox("File server gagal didownload!", MsgBoxStyle.Critical)

        Else
            Try
                Dim exeOld As String = AppDomain.CurrentDomain.BaseDirectory & nmlamax
                If File.Exists(exeOld) Then File.Delete(exeOld)

                Rename(Application.ExecutablePath(), exeOld)
                Rename(exeTemp, Application.ExecutablePath())

                MsgBox("Aplikasi akan dimatikan , Silahkan panggil ulang aplikasi!", MsgBoxStyle.Information)
                Application.Exit()

            Catch ex As IOException
                MsgBox(ex.Message)
            End Try

        End If
    End Sub
End Module
