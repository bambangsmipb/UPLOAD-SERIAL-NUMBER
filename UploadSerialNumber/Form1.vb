Imports System.Net
Public Class Form1
    Dim exeTemp As String = nmbarux
    Dim path As String = AppDomain.CurrentDomain.BaseDirectory

    Dim exeFTPUrl As String = nmaplikasi
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        If System.IO.File.Exists(exeTemp) Then System.IO.File.Delete(exeTemp)
        bWorker.RunWorkerAsync()
    End Sub

    Private Sub bWorker_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bWorker.DoWork
        Dim buffer(1023) As Byte
        Dim bytesin As Integer
        Dim totalbytesin As Integer
        Dim output As IO.Stream
        Dim fllength As Integer
        'ppp
        Try

            Dim ftprequest As FtpWebRequest = DirectCast(WebRequest.Create(exeFTPUrl), FtpWebRequest)
            ftprequest.Credentials = New NetworkCredential("fikri", "fikri")
            ftprequest.Method = Net.WebRequestMethods.Ftp.GetFileSize
            fllength = CInt(ftprequest.GetResponse.ContentLength)
            lblfilesize.Text = fllength & " bytes"
        Catch ex As Exception

        End Try

        Try
            Dim ftprequest As FtpWebRequest = DirectCast(WebRequest.Create(exeFTPUrl), FtpWebRequest)
            ftprequest.Credentials = New NetworkCredential("fikri", "fikri")
            ftprequest.Method = Net.WebRequestMethods.Ftp.DownloadFile
            Dim stream As System.IO.Stream = ftprequest.GetResponse.GetResponseStream
            'Dim outputfilepath As String = path & "\" & IO.Path.GetFileName(exeFTPUrl)
            Dim outputfilepath As String = path & "\" & exeTemp
            output = System.IO.File.Create(outputfilepath)
            bytesin = 1
            Do Until bytesin < 1
                bytesin = stream.Read(buffer, 0, 1024)
                If bytesin > 0 Then
                    output.Write(buffer, 0, bytesin)
                    totalbytesin += bytesin
                    lbldownloadbytes.Text = totalbytesin.ToString & " bytes"
                    If fllength > 0 Then
                        Dim perc As Integer = (totalbytesin / fllength) * 100
                        bWorker.ReportProgress(perc)

                    End If
                End If
            Loop
            output.Close()
            stream.Close()

        Catch ex As Exception
            MsgBox("Koneksi terputus", vbCritical, "Gagal download")
            'Application.Exit()
        End Try
    End Sub

    Private Sub bWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bWorker.ProgressChanged
        pbar.Value = e.ProgressPercentage
        lblpercent.Text = e.ProgressPercentage.ToString & "%"

    End Sub

    Private Sub bWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bWorker.RunWorkerCompleted
        Try
            Dim exeOld As String = AppDomain.CurrentDomain.BaseDirectory & nmlamax
            If System.IO.File.Exists(exeOld) Then System.IO.File.Delete(exeOld)

            Rename(Application.ExecutablePath(), exeOld)
            ' Rename("DistributionApplicationSystem.exe", "DistributionApplicationSystem.old")
            Rename(exeTemp, Application.ExecutablePath())
            ' Rename(exeTemp, "DistributionApplicationSystem.exe")

            MsgBox("Aplikasi akan dimatikan , Silahkan panggil ulang aplikasi!", MsgBoxStyle.Information)
            Application.Exit()

        Catch ex As Exception
            MsgBox(ex.Message)
            Application.Exit()
        End Try
        Me.Close()
    End Sub
End Class
