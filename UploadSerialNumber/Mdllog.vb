Imports System.IO
Imports System.IO.File
Module Mdllog
    Public waktulog = Format(Now, "yyyy-MM-dd" + "HH:mm:ss")
    Public logsql = " --    SQL  -- : "
    Public logError = " -- ERROR -- : "
    Public logmod = " --   MODUL -- : "
    Public Log As String = Application.StartupPath & "\LogFile\MyARMactivity.log"
    Public sw As StreamWriter = AppendText(Log)
    Public Sub Log_error_koneksi()
        Try
            sw.WriteLine(Now() & logError & "Number " & Err.Number & " " & Err.Description)
            sw.Close()
        Catch ex As Exception
            MsgBox("Priksa Koneksi Intenet Anda", vbOKOnly + vbCritical, "Info")
        End Try
    End Sub
    Public Sub Log_error()
        Try
            sw.WriteLine(Now() & logError & "Number " & Err.Number & " " & Err.Description)
            sw.Close()
        Catch ex As Exception
            IsNothing("!")
        End Try
    End Sub
  
End Module
