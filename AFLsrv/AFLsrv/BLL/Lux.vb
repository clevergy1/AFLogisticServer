'
' Copyright (c) Clevergy 
'
' The SOFTWARE, as well as the related copyrights and intellectual property rights, are the exclusive property of Clevergy srl. 
' Licensee acquires no title, right or interest in the SOFTWARE other than the license rights granted herein.
'
' conceived and developed by Marco Fagnano (D.R.T.C.)
' il software � ricorsivo, nel tempo rigenera se stesso.
'

Imports System
Imports System.Collections.Generic
Imports SCP.DAL
Imports Microsoft.AspNet.SignalR

Namespace SCP.BLL
    Public Class Lux

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, String.Empty, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0, 0, 0, 0, 0, False, 0, False, False, False, 0, 0)
        End Sub
        Public Sub New(m_Id As Integer,
                       m_hsId As Integer,
                       m_Cod As String,
                       m_Descr As String,
                       m_lastReceived As Date,
                       m_marcamodello As String,
                       m_installationDate As Date,
                       m_Latitude As Decimal,
                       m_Longitude As Decimal,
                       m_stato As Integer,
                       m_WorkingTimeCounter As Decimal,
                       m_PowerOnCycleCounter As Decimal,
                       m_LightON As Boolean,
                       m_CurrentMode As Integer,
                       m_forcedOn As Boolean,
                       m_forcedOff As Boolean,
                       m_isManual As Boolean,
                       m_IdAmbiente As Integer,
                       m_RemoteConnId As Integer)


            Id = m_Id
            hsId = m_hsId
            Cod = m_Cod
            Descr = m_Descr

            lastReceived = m_lastReceived

            marcamodello = m_marcamodello
            installationDate = m_installationDate
            Latitude = m_Latitude
            Longitude = m_Longitude

            stato = m_stato
            WorkingTimeCounter = m_WorkingTimeCounter
            PowerOnCycleCounter = m_PowerOnCycleCounter
            LightON = m_LightON

            CurrentMode = m_CurrentMode

            forcedOn = m_forcedOn
            forcedOff = m_forcedOff
            isManual = m_isManual
            IdAmbiente = m_IdAmbiente
            RemoteConnId = m_RemoteConnId
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            If String.IsNullOrEmpty(Descr) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.Lux_Add(hsId, Cod, Descr, UserName, marcamodello, installationDate)
            If retVal = True Then
                Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_Lux_add(hsId, Cod)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Lux_Del(Id)
        End Function

        Public Shared Function List(hsId As Integer) As List(Of Lux)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Lux_List(hsId)
        End Function

        Public Shared Function Read(Id As Integer) As Lux
            If Id <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Lux_Read(Id)
        End Function

        Public Shared Function ReadByCod(hsId As Integer, Cod As String) As Lux
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Lux_ReadByCod(hsId, Cod)
        End Function

        Public Shared Function ReadByAmb(hsId As Integer, IdAmbiente As Integer) As List(Of Lux)
            If hsId <= 0 Then
                Return Nothing
            End If

            Return DataAccessHelper.GetDataAccess.Lux_ReadByAmb(hsId, IdAmbiente)
        End Function


        Public Shared Function Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            If Id <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            If String.IsNullOrEmpty(Descr) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Lux_Update(Id, Cod, Descr, UserName, marcamodello, installationDate)
        End Function

        Public Shared Function setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.Lux_setStatus(hsId, Cod, stato)
            Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
            If Not _s Is Nothing Then
                Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                m_clientsHub.Clients.Group(_s.IdImpianto).received_Lux_setStatus(hsId, Cod, stato)
                m_clientsHub = Nothing
                _s = Nothing
            End If
            Return retVal
        End Function

        Public Shared Function setValue(hsId As Integer,
                                        Cod As String,
                                        LightON As Boolean,
                                        WorkingTimeCounter As Decimal,
                                        PowerOnCycleCounter As Decimal,
                                        CurrentMode As Integer,
                                        forcedOn As Boolean,
                                        forcedOff As Boolean) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.Lux_setValue(hsId, Cod, LightON, WorkingTimeCounter, PowerOnCycleCounter, CurrentMode, forcedOn, forcedOff)
            If retVal = True Then
                Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_Lux_changed(hsId, Cod)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.Lux_setGeoLocation(Id, Latitude, Longitude)
        End Function

        Public Shared Function setLightByCod(hsId As Integer, Cod As String, LightON As Boolean) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.Lux_setLightByCod(hsId, Cod, LightON)
            If retVal = True Then
                Dim _s As SCP.BLL.HeatingSystem = DataAccessHelper.GetDataAccess.HeatingSystem_Read(hsId)
                If Not _s Is Nothing Then
                    Dim m_clientsHub = GlobalHost.ConnectionManager.GetHubContext(Of SCP.clientsHub)()
                    m_clientsHub.Clients.Group(_s.IdImpianto).received_Lux_changed(hsId, Cod)
                    m_clientsHub = Nothing
                    _s = Nothing
                End If
            End If
            Return retVal
        End Function

        Public Shared Function setCurrentMode(Id As Integer, currentMode As Integer) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.Lux_setCurrentMode(Id, currentMode)
        End Function

        Public Shared Function setisManual(Id As Integer, isManual As Boolean) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.Lux_setisManual(Id, isManual)
        End Function

        Public Shared Function cmd_LightOn(Id As Integer) As Boolean
            Dim retVal As Boolean = False

            If Id <= 0 Then Return False

            Dim RemoteConnId As Integer = 0
            Dim Cod As String = String.Empty
            Dim hsId As Integer = 0
            Dim m_Lux As SCP.BLL.Lux = SCP.BLL.Lux.Read(Id)
            If Not m_Lux Is Nothing Then
                Cod = m_Lux.Cod
                hsId = m_Lux.hsId
                RemoteConnId = m_Lux.RemoteConnId
                m_Lux = Nothing
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

            Dim i As Int16 = 0
            Dim s As String = String.Empty

            Dim sendObj As String = "1E00"
            Dim inx As Integer = CInt(Replace(Cod, "LUX", String.Empty))
            Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00"
            Dim sendCmd As String = "0200"
            Dim sendLen As String = "0000"
            Dim sendData As String = String.Empty

            Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

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
            If retVal = True Then
                setLightByCod(hsId, Cod, True)
                setCurrentMode(Id, 0)
            End If
            Return retVal
        End Function

        Public Shared Function cmd_LightOff(Id As Integer) As Boolean
            Dim retVal As Boolean = False

            If Id <= 0 Then Return False
            Dim RemoteConnId As Integer = 0
            Dim Cod As String = String.Empty
            Dim hsId As Integer = 0
            Dim m_Lux As SCP.BLL.Lux = SCP.BLL.Lux.Read(Id)
            If Not m_Lux Is Nothing Then
                Cod = m_Lux.Cod
                hsId = m_Lux.hsId
                RemoteConnId = m_Lux.RemoteConnId
                m_Lux = Nothing
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

            Dim i As Int16 = 0
            Dim s As String = String.Empty

            Dim sendObj As String = "1E00"
            Dim inx As Integer = CInt(Replace(Cod, "LUX", String.Empty))
            Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00"
            Dim sendCmd As String = "0300"
            Dim sendLen As String = "0000"
            Dim sendData As String = String.Empty

            Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

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
            If retVal = True Then
                setLightByCod(hsId, Cod, False)
                setCurrentMode(Id, 0)
            End If
            Return retVal
        End Function

        Public Shared Function cmd_RestoreWorkingMode(Id As Integer) As Boolean
            Dim retVal As Boolean = False

            If Id <= 0 Then Return False

            Dim RemoteConnId As Integer = 0
            Dim Cod As String = String.Empty
            Dim hsId As Integer = 0
            Dim m_Lux As SCP.BLL.Lux = SCP.BLL.Lux.Read(Id)
            If Not m_Lux Is Nothing Then
                Cod = m_Lux.Cod
                hsId = m_Lux.hsId
                RemoteConnId = m_Lux.RemoteConnId
                m_Lux = Nothing
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

            Dim i As Int16 = 0
            Dim s As String = String.Empty

            Dim sendObj As String = "1E00"
            Dim inx As Integer = CInt(Replace(Cod, "LUX", String.Empty))
            Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00"
            Dim sendCmd As String = "0000"
            Dim sendLen As String = "0000"
            Dim sendData As String = String.Empty

            Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

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
        Public Property Id As Integer
        Public Property hsId As Integer
        Public Property Cod As String
        Public Property Descr As String

        Public Property lastReceived As Date

        Public Property marcamodello As String
        Public Property installationDate As Date
        Public Property Latitude As Decimal
        Public Property Longitude As Decimal

        Public Property stato As Integer
        Public Property WorkingTimeCounter As Decimal
        Public Property PowerOnCycleCounter As Decimal
        Public Property LightON As Boolean
        Public Property CurrentMode As Integer

        Public Property forcedOn As Boolean
        Public Property forcedOff As Boolean
        Public Property isManual As Boolean
        Public Property IdAmbiente As Integer
        Public Property RemoteConnId As Integer
#End Region
    End Class
End Namespace
