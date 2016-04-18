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
    Public Class hs_Cron_Profile
#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, String.Empty, New Decimal(95) {})
        End Sub
        Public Sub New(_CronId As Integer, _ProfileY As Integer, _ProfileNr As Integer, _descr As String, _ProfileData As Decimal())
            CronId = _CronId
            ProfileY = _ProfileY
            ProfileNr = _ProfileNr
            descr = _descr
            ProfileData = _ProfileData
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData As Decimal()) As Boolean
            If CronId <= 0 Then Return False
            If ProfileY <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            If ProfileData.Length <= 0 Then Return False
            Dim _d As hs_Cron_Profile_Descr = hs_Cron_Profile_Descr.Read(CronId, ProfileNr)
            If _d Is Nothing Then
                hs_Cron_Profile_Descr.Add(CronId, ProfileNr, descr)
            Else
                _d = Nothing
            End If

            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Add(CronId, ProfileY, ProfileNr, descr, ProfileData)
        End Function

        Public Shared Function Clear(CronId As Integer) As Boolean
            If CronId <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Clear(CronId)
        End Function

        Public Shared Function getProfile(hsid As Integer, CronCod As String, ProfileYear As Integer, ProfileNr As Integer) As hs_Cron_Profile
            Dim retVal As hs_Cron_Profile = Nothing

            If hsid <= 0 Then Return Nothing
            If String.IsNullOrEmpty(CronCod) Then Return Nothing

            Dim CronId As Integer = 0
            Dim RemoteConnId As Integer = 0
            Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.ReadByCronCod(hsid, CronCod)
            If Not m_hs_Cron Is Nothing Then
                CronId = m_hs_Cron.CronId
                RemoteConnId = m_hs_Cron.RemoteConnId
                m_hs_Cron = Nothing
            End If
            If CronId <= 0 Then Return Nothing

            Dim connectionType As Integer = 0
            Dim remoteAddress As String = String.Empty
            Dim remotePort As Integer = 0
            If RemoteConnId <= 0 Then
                Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsid)
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
                Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsid)
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

            If remotePort <= 0 Then Return Nothing

            Dim i As Int16 = 0
            Dim s As String = String.Empty

            Dim sendObj As String = "0C00"
            Dim inx As Integer = CInt(Replace(CronCod, "CRON", String.Empty))
            Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00" ' Replace(CronCod, "CRON", String.Empty) & "00"
            Dim sendCmd As String = "0100"
            Dim sendLen As String = "0300"
            Dim sendData As String = String.Empty

            i = Convert.ToInt16(ProfileYear)
            s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
            sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(ProfileNr, 16).PadLeft(2, "0")

            Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

            Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
            Dim _wr As New hsdsm
            m_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
            _wr = Nothing

            If m_gmDsm.returnValue = True Then



                Dim retCode As String = Mid(m_gmDsm.returnData, 9, 4)
                If retCode = "0000" Then

                    Dim lBytes As Byte() = Array.CreateInstance(GetType(Byte), 2)
                    lBytes(0) = CInt("&H" & Mid(m_gmDsm.returnData, 13, 2))
                    lBytes(1) = CInt("&H" & Mid(m_gmDsm.returnData, 15, 2))

                    'lBytes(1) = CInt("&H" & Mid(m_gmDsm.returnData, 13, 2))
                    'lBytes(0) = CInt("&H" & Mid(m_gmDsm.returnData, 15, 2))

                    Dim dataLen As Integer = BitConverter.ToInt16(lBytes, 0)

                    If dataLen > 0 Then
                        Dim ar() As String = New String(95) {}
                        Dim data2Elabor As String = Mid(m_gmDsm.returnData, 17)
                        Dim ii As Integer = 0
                        For x As Integer = 0 To (dataLen * 2) - 1 Step 4
                            ar(ii) = Mid(data2Elabor, x + 1, 4)
                            ii = ii + 1
                        Next
                        retVal = New hs_Cron_Profile
                        retVal.CronId = CronId
                        retVal.ProfileY = ProfileYear
                        retVal.ProfileNr = ProfileNr

                        Dim profileData() As Decimal = New Decimal(95) {}
                        For x As Integer = 0 To 95
                            Dim tset As Decimal = 0
                            lBytes(0) = CInt("&H" & Mid(ar(x), 1, 2))
                            lBytes(1) = CInt("&H" & Mid(ar(x), 3, 2))
                            'lBytes(1) = CInt("&H" & Mid(ar(x), 1, 2))
                            'lBytes(0) = CInt("&H" & Mid(ar(x), 3, 2))
                            tset = BitConverter.ToInt16(lBytes, 0) / 10
                            ' If tset > 100 Then tset = tset / 10
                            profileData(x) = tset
                        Next
                        retVal.ProfileData = profileData

                        Dim obj As SCP.BLL.hs_Cron_Profile = SCP.BLL.hs_Cron_Profile.Read(retVal.CronId, retVal.ProfileY, retVal.ProfileNr)
                        If Not obj Is Nothing Then
                            SCP.BLL.hs_Cron_Profile.Update(retVal.CronId, retVal.ProfileY, retVal.ProfileNr, obj.descr, retVal.ProfileData)
                            retVal.descr = obj.descr
                            obj = Nothing
                        Else
                            Dim profileDescr As String = String.Empty
                            If retVal.ProfileNr = 0 Then
                                profileDescr = "OFF"
                            Else
                                profileDescr = "profilo " & retVal.ProfileNr
                            End If
                            SCP.BLL.hs_Cron_Profile.Add(retVal.CronId, retVal.ProfileY, retVal.ProfileNr, profileDescr, retVal.ProfileData)
                            retVal.descr = profileDescr
                        End If


                    End If ' datalen >0

                End If ' recCode OK

            End If


            Return retVal
        End Function

        Public Shared Function getProfileById(CronId As Integer, ProfileYear As Integer, ProfileNr As Integer) As hs_Cron_Profile
            Dim retVal As hs_Cron_Profile = Nothing

            Dim hsId As Integer = 0
            Dim CronCod As String = String.Empty

            Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.Read(CronId)
            If Not m_hs_Cron Is Nothing Then
                CronCod = m_hs_Cron.CronCod
                hsId = m_hs_Cron.hsId
                m_hs_Cron = Nothing
            End If
            If CronId <= 0 Then Return Nothing

            Dim connectionType As Integer = 0
            Dim remoteAddress As String = String.Empty
            Dim remotePort As Integer = 0
            Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsid)
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
            If remotePort <= 0 Then Return Nothing

            Dim i As Int16 = 0
            Dim s As String = String.Empty

            Dim sendObj As String = "0C00"
            Dim sendIdx As String = Replace(CronCod, "CRON", String.Empty) & "00"
            Dim sendCmd As String = "0100"
            Dim sendLen As String = "0300"
            Dim sendData As String = String.Empty

            i = Convert.ToInt16(ProfileYear)
            s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
            sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(ProfileNr, 16).PadLeft(2, "0")

            Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

            Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
            Dim _wr As New hsdsm
            m_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
            _wr = Nothing

            If m_gmDsm.returnValue = True Then

                Dim retCode As String = Mid(m_gmDsm.returnData, 9, 4)
                If retCode = "0000" Then

                    Dim lBytes As Byte() = Array.CreateInstance(GetType(Byte), 2)
                    lBytes(0) = CInt("&H" & Mid(m_gmDsm.returnData, 13, 2))
                    lBytes(1) = CInt("&H" & Mid(m_gmDsm.returnData, 15, 2))
                    Dim dataLen As Integer = BitConverter.ToInt16(lBytes, 0)

                    If dataLen > 0 Then
                        Dim ar() As String = New String(95) {}
                        Dim data2Elabor As String = Mid(m_gmDsm.returnData, 17)
                        Dim ii As Integer = 0
                        For x As Integer = 0 To (dataLen * 2) - 1 Step 4
                            ar(ii) = Mid(data2Elabor, x + 1, 4)
                            ii = ii + 1
                        Next
                        retVal = New hs_Cron_Profile
                        retVal.CronId = CronId
                        retVal.ProfileY = ProfileYear
                        retVal.ProfileNr = ProfileNr

                        Dim profileData() As Decimal = New Decimal(95) {}
                        For x As Integer = 0 To 95
                            Dim tset As Decimal = 0
                            lBytes(0) = CInt("&H" & Mid(ar(x), 1, 2))
                            lBytes(1) = CInt("&H" & Mid(ar(x), 3, 2))
                            tset = BitConverter.ToInt16(lBytes, 0) / 10
                            profileData(x) = tset
                        Next
                        retVal.ProfileData = profileData

                        Dim obj As SCP.BLL.hs_Cron_Profile = SCP.BLL.hs_Cron_Profile.Read(retVal.CronId, retVal.ProfileY, retVal.ProfileNr)
                        If Not obj Is Nothing Then
                            SCP.BLL.hs_Cron_Profile.Update(retVal.CronId, retVal.ProfileY, retVal.ProfileNr, obj.descr, retVal.ProfileData)
                            retVal.descr = obj.descr
                            obj = Nothing
                        Else
                            Dim profileDescr As String = String.Empty
                            If retVal.ProfileNr = 0 Then
                                profileDescr = "OFF"
                            Else
                                profileDescr = "profilo " & retVal.ProfileNr
                            End If
                            SCP.BLL.hs_Cron_Profile.Add(retVal.CronId, retVal.ProfileY, retVal.ProfileNr, profileDescr, retVal.ProfileData)
                            retVal.descr = profileDescr
                        End If


                    End If ' datalen >0

                End If ' recCode OK

            End If


            Return retVal
        End Function

        Public Shared Function List(CronId As Integer, ProfileY As Integer) As List(Of hs_Cron_Profile)
            If CronId <= 0 Then Return Nothing
            If ProfileY <= 0 Then Return Nothing
            'Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_List(CronId, ProfileY)

            Dim retVal As New List(Of hs_Cron_Profile)
            For x As Integer = 0 To 10
                Dim _obj As hs_Cron_Profile = DataAccessHelper.GetDataAccess.hs_Cron_Profile_Read(CronId, ProfileY, x)
                If _obj Is Nothing Then
                    retVal.Add(getProfileById(CronId, ProfileY, x))
                Else
                    retVal.Add(_obj)
                    _obj = Nothing
                End If
            Next
            Return retVal
        End Function

        Public Shared Function Read(CronId As Integer, ProfileY As Integer, ProfileNr As Integer) As hs_Cron_Profile
            If CronId <= 0 Then Return Nothing
            If ProfileY <= 0 Then Return Nothing
            If ProfileNr < 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Read(CronId, ProfileY, ProfileNr)
        End Function

        Public Shared Function Update(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData As Decimal()) As Boolean
            If CronId <= 0 Then Return False
            If ProfileY <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            If ProfileData.Length <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Update(CronId, ProfileY, ProfileNr, descr, ProfileData)
        End Function

        Public Shared Function setProfileNow(CronId As Integer, ProfileNr As Integer) As Boolean
            Dim retVal As Boolean = False

            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False

            Dim CronCod As String = String.Empty
            Dim hsId As Integer = 0
            Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.Read(CronId)
            If Not m_hs_Cron Is Nothing Then
                CronCod = m_hs_Cron.CronCod
                hsId = m_hs_Cron.hsId
                m_hs_Cron = Nothing
            End If

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

            Dim Calyear As Integer = Now.Year
            Dim Calmonth As Integer = Now.Month
            Dim _calendar As hs_Cron_Calendar = DataAccessHelper.GetDataAccess.hs_Cron_Calendar_Read(CronId, Calyear, Calmonth)
            If Not _calendar Is Nothing Then

                Dim i As Int16 = 0
                Dim s As String = String.Empty

                Dim sendObj As String = "0C00"
                Dim inx As Integer = CInt(Replace(CronCod, "CRON", String.Empty))
                Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00" ' Replace(CronCod, "CRON", String.Empty) & "00"
                Dim sendCmd As String = "0600"
                Dim sendLen As String = "2300"
                Dim sendData As String = String.Empty

                i = Convert.ToInt16(Calyear)
                s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
                sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(Calmonth, 16).PadLeft(2, "0")  'anno e mese

                Dim daytoChange As Integer = Now.Day - 1

                For x As Integer = 0 To 31
                    If x = daytoChange Then
                        sendData += Convert.ToString(ProfileNr, 16).PadLeft(2, "0")
                    Else
                        sendData += Convert.ToString(_calendar.DesiredMonthData(x), 16).PadLeft(2, "0")
                    End If
                Next

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

                _calendar = Nothing
            End If

            If retVal = True Then
                SCP.BLL.hs_Cron.Restart(CronId)
                SCP.BLL.HeatingSystem.requestLog(hsId)
            End If

            Return retVal
        End Function


        Public Shared Function setProfile(CronId As Integer, ProfileY As Integer, ProfileNr As Integer) As Boolean
            Dim retVal As Boolean = False

            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False

            Dim CronCod As String = String.Empty
            Dim hsId As Integer = 0
            Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.Read(CronId)
            If Not m_hs_Cron Is Nothing Then
                CronCod = m_hs_Cron.CronCod
                hsId = m_hs_Cron.hsId
                m_hs_Cron = Nothing
            End If

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

            Dim profile As SCP.BLL.hs_Cron_Profile = SCP.BLL.hs_Cron_Profile.Read(CronId, ProfileY, ProfileNr)
            If Not profile Is Nothing Then

                Dim i As Int16 = 0
                Dim s As String = String.Empty

                Dim sendObj As String = "0C00"
                ' Dim sendIdx As String = Replace(CronCod, "CRON", String.Empty) & "00"

                Dim inx As Integer = CInt(Replace(CronCod, "CRON", String.Empty))
                Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00" ' Replace(CronCod, "CRON", String.Empty) & "00"


                Dim sendCmd As String = "0200"
                Dim sendLen As String = "C300"
                Dim sendData As String = String.Empty

                i = Convert.ToInt16(ProfileY)
                s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
                sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(ProfileNr, 16).PadLeft(2, "0")

                Dim lBytes As Byte() = Array.CreateInstance(GetType(Byte), 2)

                For x As Integer = 0 To 95
                    i = Convert.ToInt16(profile.ProfileData(x) * 10)
                    s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
                    Dim TsetHex As String = Mid(s, 3, 2) & Mid(s, 1, 2)
                    sendData += TsetHex
                Next


                profile = Nothing


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

                            retVal = SCP.BLL.hs_Cron.Restart(CronId)

                        End If
                    End If
                End If
            End If



            Return retVal

        End Function
#End Region

#Region "public properties"
        Public Property CronId As Integer
        Public Property ProfileY As Integer
        Public Property ProfileNr As Integer
        Public Property descr As String
        Public Property ProfileData As Decimal()
#End Region

    End Class
End Namespace
