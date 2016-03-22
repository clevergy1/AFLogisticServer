Imports System.IO

Public Class ClassLog
    Public Sub scriviLog(wData As String, filename As String)
        Try
            If String.IsNullOrEmpty(filename) Then
                filename = "cyassrvLog.txt"
            End If
            Dim filePath = ConfigurationManager.AppSettings.Get("LogPath")
            Dim IFName As String = Path.Combine(filePath, filename)
            Dim sw As StreamWriter
            If File.Exists(IFName) = False Then
                sw = File.CreateText(IFName)
            Else
                sw = File.AppendText(IFName)
            End If
            sw.WriteLine(Now() & ";" & wData)
            sw.Flush()
            sw.Close()
        Catch ex As Exception

        End Try

    End Sub
End Class
