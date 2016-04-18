'
' Copyright (c) Clevergy 
'
' The SOFTWARE, as well as the related copyrights and intellectual property rights, are the exclusive property of Clevergy srl. 
' Licensee acquires no title, right or interest in the SOFTWARE other than the license rights granted herein.
'
' conceived and developed by Marco Fagnano (D.R.T.C.)
' il software è ricorsivo, nel tempo rigenera se stesso.
'

Imports System
Imports System.Collections.Generic
Imports SCP.DAL
Imports Microsoft.AspNet.SignalR

Namespace SCP.BLL
    Public Class hs_Cron

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, False, 0, String.Empty, String.Empty, 0, String.Empty, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0, 0, 0, 0)
        End Sub
        Public Sub New(m_CronId As Integer,
                       m_hsId As Integer,
                       m_SetPoint As Decimal,
                       m_isOn As Boolean,
                       m_TLocal As Decimal,
                       m_CronCod As String,
                       m_CronDescr As String,
                       m_stato As Integer,
                       m_NoteInterne As String,
                       m_marcamodello As String,
                       m_installationDate As Date,
                       m_Latitude As Decimal,
                       m_Longitude As Decimal,
                       m_CronType As Integer,
                       m_RemoteConnId As Integer)
            CronId = m_CronId
            hsId = m_hsId
            SetPoint = m_SetPoint
            isOn = m_isOn
            TLocal = m_TLocal
            CronCod = m_CronCod
            CronDescr = m_CronDescr
            stato = m_stato
            NoteInterne = m_NoteInterne
            marcamodello = m_marcamodello
            installationDate = m_installationDate
            Latitude = m_Latitude
            Longitude = m_Longitude
            CronType = m_CronType
            RemoteConnId = m_RemoteConnId
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer,
                                   CronCod As String,
                                   CronDescr As String,
                                   NoteInterne As String,
                                   UserName As String,
                                   marcamodello As String,
                                   installationDate As Date,
                                   CronType As Integer) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return False
            End If
            If String.IsNullOrEmpty(CronDescr) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Cron_Add(hsId, CronCod, CronDescr, NoteInterne, UserName, marcamodello, installationDate, CronType)
        End Function

        Public Shared Function Del(CronId As Integer) As Boolean
            If CronId <= 0 Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.hs_Cron_Del(CronId)
            If retVal = True Then
                DataAccessHelper.GetDataAccess.hs_Cron_Profile_Clear(CronId)
            End If
            Return retVal
        End Function

        Public Shared Function List(hsId As Integer) As List(Of hs_Cron)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Cron_List(hsId)
        End Function

        Public Shared Function ListOnOff(hsid As Integer) As List(Of hs_Cron)
            If hsid <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Cron_ListOnOff(hsid)
        End Function

        Public Shared Function Read(CronId As Integer) As hs_Cron
            If CronId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Cron_Read(CronId)
        End Function

        Public Shared Function ReadByCronCod(hsId As Integer, CronCod As String) As hs_Cron
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Cron_ReadByCronCod(hsId, CronCod)
        End Function

        Public Shared Function Update(CronId As Integer, _
                                      CronCod As String, _
                                      CronDescr As String, _
                                      NoteInterne As String, _
                                      UserName As String, _
                                      marcamodello As String, _
                                      installationDate As Date, _
                                      CronType As Integer) As Boolean
            If CronId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return False
            End If
            If String.IsNullOrEmpty(CronDescr) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Cron_Update(CronId, CronCod, CronDescr, NoteInterne, UserName, marcamodello, installationDate, CronType)
        End Function

        Public Shared Function UpdateMarcamodello(Id As Integer, marcamodello As String, installationDate As Date, UserName As String) As Boolean
            If Id <= 0 Then
                Return False
            End If
            If installationDate.Year <= 1900 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_Cron_UpdateMarcamodello(Id, marcamodello, installationDate, UserName)
        End Function

        Public Shared Function setStatus(hsId As Integer, CronCod As String, SetPoint As Decimal, stato As Integer) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.hs_Cron_setStatus(hsId, CronCod, SetPoint, stato)
            Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
            If Not _s Is Nothing Then
                Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                m_clientsHub.Clients.Group(_s.IdImpianto).received_hs_Cron_setStatus(hsId, CronCod, SetPoint, stato)
                m_clientsHub = Nothing
                _s = Nothing
            End If
            Return retVal
        End Function

        Public Shared Function setGeoLocation(CronId As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            If CronId <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_Cron_setGeoLocation(CronId, Latitude, Longitude)
        End Function

        Public Shared Function Restart(CronId As Integer) As Boolean
            Dim retVal As Boolean = False

            If CronId <= 0 Then Return False

            Dim CronCod As String = String.Empty
            Dim hsId As Integer = 0
            Dim RemoteConnId As Integer = 0
            Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.Read(CronId)
            If Not m_hs_Cron Is Nothing Then
                CronCod = m_hs_Cron.CronCod
                hsId = m_hs_Cron.hsId
                RemoteConnId = m_hs_Cron.RemoteConnId
                m_hs_Cron = Nothing
            End If

            Dim connectionType As Integer = 0
            Dim remoteAddress As String = String.Empty
            Dim remotePort As Integer = 0
            If RemoteConnId <= 0 Then
                Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsId)
                If Not hs Is Nothing Then
                    If hs.VPNConnectionId > 0 Then
                        Dim remoteConnection As SCP.BLL.Impianti_RemoteConnections = SCP.BLL.Impianti_RemoteConnections.Read(hs.IdImpianto, hs.VPNConnectionId)
                        If Not remoteConnection Is Nothing Then
                            If remoteConnection.connectionType = 2 Or remoteConnection.connectionType = 3 Then
                                connectionType = remoteConnection.connectionType
                                remoteAddress = remoteConnection.remoteAddress
                                remotePort = remoteConnection.remotePort
                            End If
                            remoteConnection = Nothing
                        End If
                    End If
                    hs = Nothing
                End If

            Else
                Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsId)
                If Not hs Is Nothing Then
                    Dim remoteConnection As SCP.BLL.Impianti_RemoteConnections = SCP.BLL.Impianti_RemoteConnections.Read(hs.IdImpianto, RemoteConnId)
                    If Not remoteConnection Is Nothing Then
                        If remoteConnection.connectionType = 2 Or remoteConnection.connectionType = 3 Then
                            connectionType = remoteConnection.connectionType
                            remoteAddress = remoteConnection.remoteAddress
                            remotePort = remoteConnection.remotePort
                        End If
                        remoteConnection = Nothing
                    End If
                End If
            End If
            If remotePort <= 0 Then Return False

            Dim sendObj As String = "0C00"
            Dim inx As Integer = CInt(Replace(CronCod, "CRON", String.Empty))
            Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00" ' Replace(CronCod, "CRON", String.Empty) & "00"
            Dim sendCmd As String = "0900"
            Dim sendLen As String = "0000"

            Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen

            Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
            Dim _wr As New hsdsm
            m_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
            _wr = Nothing

            If m_gmDsm.returnValue = True Then
                If m_gmDsm.returnData.Length >= 16 Then
                    Dim responseObj As String = Mid(m_gmDsm.returnData, 1, 4)
                    Dim responseIdx As String = Mid(m_gmDsm.returnData, 5, 4)
                    Dim responseCod As String = Mid(m_gmDsm.returnData, 9, 4)
                    If responseCod = "0000" Then
                        retVal = True
                    End If
                End If
            End If
            Return retVal
        End Function
#End Region

#Region "public properties"
        Public Property CronId As Integer
        Public Property hsId As Integer
        Public Property SetPoint As Decimal
        Public Property isOn As Boolean
        Public Property TLocal As Decimal
        Public Property CronCod As String
        Public Property CronDescr As String
        Public Property stato As Integer
        Public Property NoteInterne As String
        Public Property marcamodello As String
        Public Property installationDate As Date
        Public Property Latitude As Decimal
        Public Property Longitude As Decimal
        Public Property CronType As Integer
        Public Property RemoteConnId As Integer
#End Region
    End Class
End Namespace
