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
    Public Class Lux

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, String.Empty, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0, 0, 0, 0, 0, False)
        End Sub
        Public Sub New(m_Id As Integer, _
                       m_hsId As Integer, _
                       m_Cod As String, _
                       m_Descr As String, _
                       m_lastReceived As Date, _
                       m_marcamodello As String, _
                       m_installationDate As Date, _
                       m_Latitude As Decimal, _
                       m_Longitude As Decimal, _
                       m_stato As Integer, _
                       m_WorkingTimeCounter As Decimal, _
                       m_PowerOnCycleCounter As Decimal, _
                       m_LightON As Boolean)
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

        Public Shared Function setValue(hsId As Integer, _
                                        Cod As String, _
                                        WorkingTimeCounter As Decimal, _
                                        PowerOnCycleCounter As Decimal) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Dim retVal As Boolean = False
            retVal = DataAccessHelper.GetDataAccess.Lux_setValue(hsId, Cod, WorkingTimeCounter, PowerOnCycleCounter)
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
#End Region
    End Class
End Namespace
