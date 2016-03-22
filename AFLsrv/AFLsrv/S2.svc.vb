Imports System.ServiceModel
Imports System.ServiceModel.Activation
Imports System.ServiceModel.Web
Imports System.IO
Imports System.Threading


<ServiceContract(Namespace:="")>
<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
Public Class S2

    ' To use HTTP GET, add <WebGet()> attribute. (Default ResponseFormat is WebMessageFormat.Json)
    ' To create an operation that returns XML,
    '     add <WebGet(ResponseFormat:=WebMessageFormat.Xml)>,
    '     and include the following line in the operation body:
    '         WebOperationContext.Current.OutgoingResponse.ContentType = "text/xml"
    <OperationContract()>
    Public Function DoWork(req As String) As Boolean
        If String.IsNullOrEmpty(req) Then
            Return False
        End If

        Dim markStatus As NameValueCollection = HttpUtility.ParseQueryString(req)
        Dim logFile As String = String.Empty

        Dim hsId As Integer = 0
        If Not markStatus.Get("ID") Is Nothing Then
            hsId = Convert.ToInt32(markStatus.Get("ID"))
            logFile = "log_" & hsId.ToString & "_" & Now.Year.ToString & "_" & Now.Month.ToString.PadLeft(2, "0") & "_" & Now.Day.ToString.PadLeft(2, "0") & ".txt"
        End If


        Try
            scriviLog(req, logFile)
            Dim s As New HsRequest
            's.Elabor(req)
            Dim thOldLog As New Thread(AddressOf s.Elabor)
            thOldLog.IsBackground = True
            thOldLog.Start(req)
            s = Nothing

        Catch ex As Exception
            scriviLog("S2.svc " & ex.Message, String.Empty)
        End Try


        Return True
    End Function


#Region "private methods"
    Private Sub scriviLog(wData As String, logFile As String)
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

End Class
