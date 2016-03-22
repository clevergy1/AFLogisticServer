Imports System.Web
Imports System.Web.Services
Imports System.Threading
Imports System.IO
Public Class hsBackDoor
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.ContentType = "text/plain"
        Dim retVal As Boolean = False
        Try
            retVal = hsBackDoor(context.Request.Url.Query.Replace("?", String.Empty))

        Catch ex As Exception

            retVal = False
        End Try
        context.Response.Write(retVal)

    End Sub

#Region "hslog backdoor"
    Public Function hsBackDoor(req As String) As Boolean
        If String.IsNullOrEmpty(req) Then
            Return False
        End If

        Dim markStatus As NameValueCollection = HttpUtility.ParseQueryString(req)
        Dim logFile As String = String.Empty

        Dim hsId As Integer = 0
        If Not markStatus.Get("ID") Is Nothing Then
            hsId = Convert.ToInt32(markStatus.Get("ID"))
            logFile = "log_" & hsId.ToString & ".txt"
        End If

        'Dim debug As String = String.Empty
        'For Each key In markStatus.Keys
        '    'If Trim(key).Contains("LUXM") Then
        '    '    Dim Callreq As String = markStatus.Get(key)
        '    '    debug += "key=" & key
        '    'End If

        '    If Mid(Trim(key).ToString, 1, 3) = "LUX" Then
        '        Dim Callreq As String = markStatus.Get(key)
        '        debug += "key=" & key
        '    End If

        '    'If Trim(key).Contains("LUX") Then
        '    '    Dim Callreq As String = markStatus.Get(key)
        '    '    debug += "key=" & key
        '    'End If
        'Next
        'Dim a As String = debug


        Try
            hs_BackDoor_scriviLog(req, logFile)
            Dim s As New HsRequest
            s.Elabor(req)
            'Dim thOldLog As New Thread(AddressOf s.Elabor)
            'thOldLog.IsBackground = True
            'thOldLog.Start(req)
            s = Nothing
        Catch ex As Exception
            hs_BackDoor_scriviLog("hsBackDoor " & ex.Message, String.Empty)
        End Try


        Return True
    End Function

    Private Sub hs_BackDoor_scriviLog(wData As String, logFile As String)
        Dim _txt As String = String.Empty
        If String.IsNullOrEmpty(logFile) Then
            _txt = "s2log.txt"
        Else
            _txt = logFile
        End If
        Try
            'LogPath
            Dim LogPath As String = ConfigurationManager.AppSettings("LogPath")
            Dim IFName As String = Path.Combine(LogPath, _txt)
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
#End Region

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class