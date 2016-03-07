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
    Public Class HeatingSystem
        Private _hsId As Integer
        Private _IdImpianto As String
        Private _Descr As String
        Private _Indirizzo As String
        Private _Latitude As Decimal
        Private _Longitude As Decimal
        Private _AltSLM As Integer
        Private _MaintenanceMode As Boolean
        Private _VPNConnectionId As Integer
        Private _MapId As Integer
        Private _TempExt As Decimal
        Private _stato As Integer
        Private _Note As String
        Private _lastRec As Date
        Private _isOnline As Boolean
        Private _isEnabled As Boolean
        Private _IwMonitoringId As Integer
        Private _IwMonitoringDes As String

        Private _DesImpianto As String

#Region "constructor"
        Public Sub New()
            Me.New(0, String.Empty, String.Empty, String.Empty, 0, 0, 0, False, 0, 0, 0, 0, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), False, False, 0, String.Empty, String.Empty)
        End Sub
        Public Sub New(hsId As Integer, _
                       IdImpianto As String, _
                       Descr As String, _
                       Indirizzo As String, _
                       Latitude As Decimal, _
                       Longitude As Decimal, _
                       AltSLM As Integer, _
                       MaintenanceMode As Boolean, _
                       VPNConnectionId As Integer, _
                       MapId As Integer, _
                       TempExt As Decimal, _
                       stato As Integer, _
                       Note As String, _
                       lastRec As Date, _
                       isOnline As Boolean, _
                       isEnabled As Boolean, _
                       IwMonitoringId As Integer, _
                       IwMonitoringDes As String, _
                       DesImpianto As String)
            Me._hsId = hsId
            Me._IdImpianto = IdImpianto
            Me._Descr = Descr
            Me._Indirizzo = Indirizzo
            Me._Latitude = Latitude
            Me._Longitude = Longitude
            Me._AltSLM = AltSLM
            Me._MaintenanceMode = MaintenanceMode
            Me._VPNConnectionId = VPNConnectionId
            _MapId = MapId
            _TempExt = TempExt
            _stato = stato
            _Note = Note
            Me._lastRec = lastRec
            _isOnline = isOnline
            _isEnabled = isEnabled
            _IwMonitoringId = IwMonitoringId
            _IwMonitoringDes = IwMonitoringDes

            _DesImpianto = DesImpianto
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(IdImpianto As String, Descr As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Integer, UserName As String) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            If String.IsNullOrEmpty(Indirizzo) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.HeatingSystem_Add(IdImpianto, Descr, Indirizzo, Latitude, Longitude, AltSLM, UserName)
            If retVal = True Then
                Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                m_clientsHub.Clients.Group(IdImpianto).received_HeatingSystem_Add(IdImpianto)
                m_clientsHub = Nothing
            End If
            Return retVal
        End Function

        Public Shared Function Del(hsId As Integer) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)

            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.HeatingSystem_Del(hsId)
            If retVal = True Then
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_HeatingSystem_Del(hsId)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function getTotMap(IdImpianto As String) As Integer
            If String.IsNullOrEmpty(IdImpianto) Then Return 0
            Return DataAccessHelper.GetDataAccess.HeatingSystem_getTotMap(IdImpianto)
        End Function

        Public Shared Function List(IdImpianto As String) As List(Of HeatingSystem)
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_List(IdImpianto)
        End Function

        Public Shared Function ListEnabled(IdImpianto As String) As List(Of HeatingSystem)
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_ListEnabled(IdImpianto)
        End Function

        Public Shared Function ListAll() As List(Of HeatingSystem)
            Return DataAccessHelper.GetDataAccess.HeatingSystem_ListAll()
        End Function

        Public Shared Function Read(hsId As Integer) As HeatingSystem
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
        End Function

        Public Shared Function setMaintenanceMode(hsId As Integer, MaintenanceMode As Boolean) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.HeatingSystem_setMaintenanceMode(hsId, MaintenanceMode)
            If retVal = True Then
                Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_HeatingSystem_setMaintenanceMode(hsId, MaintenanceMode)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function setIwMonitoring(hsId As Integer, IwMonitoringId As Integer, IwMonitoringDes As String) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_setIwMonitoring(hsId, IwMonitoringId, Trim(IwMonitoringDes))
        End Function

        Public Shared Function setisOnline(hsId As Integer, isOnline As Boolean) As Boolean
            If hsId <= 0 Then
                Return False
            End If

            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.HeatingSystem_setisOnline(hsId, isOnline)
            If retVal = True Then
                Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_HeatingSystem_setisOnline(hsId, isOnline)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function setlastRec(hsId As Integer, lastRec As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_setlastRec(hsId, lastRec)
        End Function

        Public Shared Function setNote(hsId As Integer, Note As String, UserName As String) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_setNote(hsId, Note, UserName)
        End Function

        Public Shared Function setStatus(hsId As Integer, stato As Integer) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.HeatingSystem_setStatus(hsId, stato)
            If retVal = True Then
                Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_HeatingSystem_setStatus(hsId, stato)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function setIsEnabled(hsId As Integer, isEnabled As Boolean) As Boolean
            If hsId <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.HeatingSystem_setIsEnabled(hsId, isEnabled)
        End Function

        Public Shared Function Update(hsId As Integer, _
                                      Descr As String, _
                                      Indirizzo As String, _
                                      Latitude As Decimal, _
                                      Longitude As Decimal, _
                                      AltSLM As Integer, _
                                      VPNConnectionId As Integer, _
                                      MapId As Integer, _
                                      UserName As String) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Indirizzo) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.HeatingSystem_Update(hsId, Descr, Indirizzo, Latitude, Longitude, AltSLM, VPNConnectionId, MapId, UserName)
        End Function

        Public Shared Function resetSystemStatus(hsId As Integer) As Boolean
            Dim retVal As Boolean = False
            If hsId <= 0 Then Return False

            Dim connectionType As Integer = 0
            Dim remoteAddress As String = String.Empty
            Dim remotePort As Integer = 0
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
            If remotePort <= 0 Then Return False

            Dim sendObj As String = "0000"
            Dim sendIdx As String = "0000"
            Dim sendCmd As String = "0200"
            Dim sendData As String = "0000"
            Dim sendString As String = sendObj & sendIdx & sendCmd & sendData

            Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
            Dim _wr As New hsdsm
            m_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
            _wr = Nothing
            If m_gmDsm.returnValue = True Then
                Dim responseObj As String = Mid(m_gmDsm.returnData, 1, 4)
                Dim responseIdx As String = Mid(m_gmDsm.returnData, 5, 4)
                Dim responseCod As String = Mid(m_gmDsm.returnData, 9, 4)
                If responseCod = "0000" Then
                    retVal = True
                    SCP.BLL.HeatingSystem.setStatus(hsId, 0)
                End If
            End If
            Return retVal
        End Function

        Public Shared Function requestLog(hsId As Integer) As Boolean
            Dim retVal As Boolean = False
            If hsId <= 0 Then Return False

            Dim connectionType As Integer = 0
            Dim remoteAddress As String = String.Empty
            Dim remotePort As Integer = 0
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
            If remotePort <= 0 Then Return False

            Dim sendObj As String = "0000"
            Dim sendIdx As String = "0000"
            Dim sendCmd As String = "0100"
            Dim sendData As String = "0000"
            Dim sendString As String = sendObj & sendIdx & sendCmd & sendData

            Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
            Dim _wr As New hsdsm
            m_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
            _wr = Nothing
            If m_gmDsm.returnValue = True Then
                Dim responseObj As String = Mid(m_gmDsm.returnData, 1, 4)
                Dim responseIdx As String = Mid(m_gmDsm.returnData, 5, 4)
                Dim responseCod As String = Mid(m_gmDsm.returnData, 9, 4)
                If responseCod = "0000" Then
                    retVal = True
                End If
            End If
            Return retVal
        End Function
#End Region

#Region "public properties"
        Public Property hsId As Integer
            Get
                Return Me._hsId
            End Get
            Set(value As Integer)
                Me._hsId = value
            End Set
        End Property
        Public Property IdImpianto() As String
            Get
                Return Me._IdImpianto
            End Get
            Set(value As String)
                Me._IdImpianto = value
            End Set
        End Property
        Public Property Descr As String
            Get
                Return Me._Descr
            End Get
            Set(value As String)
                Me._Descr = value
            End Set
        End Property
        Public Property Indirizzo As String
            Get
                Return Me._Indirizzo
            End Get
            Set(value As String)
                Me._Indirizzo = value
            End Set
        End Property
        Public Property Latitude As Decimal
            Get
                Return Me._Latitude
            End Get
            Set(value As Decimal)
                Me._Latitude = value
            End Set
        End Property
        Public Property Longitude As Decimal
            Get
                Return Me._Longitude
            End Get
            Set(value As Decimal)
                Me._Longitude = value
            End Set
        End Property
        Public Property AltSLM As Integer
            Get
                Return Me._AltSLM
            End Get
            Set(value As Integer)
                Me._AltSLM = value
            End Set
        End Property
        Public Property MaintenanceMode As Boolean
            Get
                Return Me._MaintenanceMode
            End Get
            Set(value As Boolean)
                Me._MaintenanceMode = value
            End Set
        End Property
        Public Property VPNConnectionId As Integer
            Get
                Return Me._VPNConnectionId
            End Get
            Set(value As Integer)
                Me._VPNConnectionId = value
            End Set
        End Property
        Public Property MapId As Integer
            Get
                Return Me._MapId
            End Get
            Set(value As Integer)
                Me._MapId = value
            End Set
        End Property
        Public Property TempExt As Decimal
            Get
                Return Me._TempExt
            End Get
            Set(value As Decimal)
                Me._TempExt = value
            End Set
        End Property
        Public Property stato As Integer
            Get
                Return Me._stato
            End Get
            Set(value As Integer)
                Me._stato = value
            End Set
        End Property
        Public Property Note As String
            Get
                Return _Note
            End Get
            Set(value As String)
                _Note = value
            End Set
        End Property
        Public Property lastRec As Date
            Get
                Return Me._lastRec
            End Get
            Set(value As Date)
                Me._lastRec = value
            End Set
        End Property
        Public Property isOnline As Boolean
            Get
                Return Me._isOnline
            End Get
            Set(value As Boolean)
                Me._isOnline = value
            End Set
        End Property
        Public Property isEnabled As Boolean
            Get
                Return Me._isEnabled
            End Get
            Set(value As Boolean)
                Me._isEnabled = value
            End Set
        End Property
        Public Property IwMonitoringId As Integer
            Get
                Return _IwMonitoringId
            End Get
            Set(value As Integer)
                _IwMonitoringId = value
            End Set
        End Property
        Public Property IwMonitoringDes As String
            Get
                Return _IwMonitoringDes
            End Get
            Set(value As String)
                _IwMonitoringDes = value
            End Set
        End Property

        Public Property DesImpianto As String
            Get
                Return Me._DesImpianto
            End Get
            Set(value As String)
                Me._DesImpianto = value
            End Set
        End Property
#End Region
    End Class
End Namespace
