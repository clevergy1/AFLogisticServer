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

Namespace SCP.BLL
    Public Class hs_amb_Calendar

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, New Integer(31) {}, New Integer(31) {}, New Integer(31) {}, FormatDateTime("01/01/1900", DateFormat.GeneralDate))
        End Sub
        Public Sub New(_CronId As Integer, _Calyear As Integer, _Calmonth As Integer, _RealMonthData As Integer(), _DesiredMonthData As Integer(), _TasksForDesired As Integer(), _LastSend As Date)
            CronId = _CronId
            Calyear = _Calyear
            Calmonth = _Calmonth
            RealMonthData = _RealMonthData
            DesiredMonthData = _DesiredMonthData
            TasksForDesired = _TasksForDesired
            LastSend = _LastSend
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(CronId As Integer, Calyear As Integer, Calmonth As Integer, monthData As Integer()) As Boolean
            If CronId <= 0 Then Return False
            If Calyear <= 0 Then Return False
            If Calmonth <= 0 Then Return False
            If Calmonth > 12 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Calendar_Add(CronId, Calyear, Calmonth, monthData)
        End Function

        Public Shared Function CalendarGet(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
            'Dim retVal As Boolean = False
            Dim retVal As hs_amb_Calendar = Nothing

            If CronId <= 0 Then Return Nothing
            Dim CronCod As String = String.Empty
            Dim hsId As Integer = 0
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
            If remotePort <= 0 Then Return Nothing


            Dim i As Int16 = 0
            Dim s As String = String.Empty

            Dim sendObj As String = "0C00"
            Dim inx As Integer = CInt(Replace(CronCod, "CRON", String.Empty))
            Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00" ' Replace(CronCod, "CRON", String.Empty) & "00"
            Dim sendCmd As String = "0500"
            Dim sendLen As String = "0300"
            Dim sendData As String = String.Empty

            i = Convert.ToInt16(Calyear)
            s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
            sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(Calmonth, 16).PadLeft(2, "0")

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
                        Dim lBytes As Byte() = Array.CreateInstance(GetType(Byte), 2)
                        lBytes(0) = CInt("&H" & Mid(m_gmDsm.returnData, 13, 2))
                        lBytes(1) = CInt("&H" & Mid(m_gmDsm.returnData, 15, 2))
                        Dim dataLen As Integer = BitConverter.ToInt16(lBytes, 0)
                        If dataLen > 0 Then
                            Dim ar() As String = New String(31) {}
                            Dim data2Elabor As String = Mid(m_gmDsm.returnData, 17)
                            Dim ii As Integer = 0
                            For x As Integer = 0 To (dataLen * 2) - 1 Step 2
                                ar(ii) = Mid(data2Elabor, x + 1, 2)
                                ii = ii + 1
                            Next
                            Dim monthData() As Integer = New Integer(31) {}
                            For x As Integer = 0 To 31
                                Dim tset As Integer = 0
                                lBytes(0) = CInt("&H" & Mid(ar(x), 1, 2))
                                tset = BitConverter.ToInt16(lBytes, 0)
                                monthData(x) = tset
                            Next

                            'Dim _calendar As SCP.BLL.hs_amb_Calendar = Read(CronId, Calyear, Calmonth)
                            'If _calendar Is Nothing Then
                            '    retVal = Add(CronId, Calyear, Calmonth, monthData)
                            'Else
                            '    retVal = Update(CronId, Calyear, Calmonth, monthData)
                            'End If

                            retVal = New hs_amb_Calendar(CronId, Calyear, Calmonth, monthData, Nothing, New Integer(31) {}, FormatDateTime("01/01/1900", DateFormat.GeneralDate))
                        End If
                    End If
                End If
            End If

            Return retVal
        End Function

        Public Shared Function CalendarSet(CronId As Integer, Calyear As Integer, Calmonth As Integer) As Boolean
            Dim retVal As Boolean = False

            If CronId <= 0 Then Return False
            If Calyear <= 0 Then Return False
            If Calmonth <= 0 Then Return False
            If Calmonth > 12 Then Return False

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

            Dim _calendar As hs_amb_Calendar = DataAccessHelper.GetDataAccess.hs_amb_Calendar_Read(CronId, Calyear, Calmonth)
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


                For x As Integer = 0 To 31
                    'If _calendar.DesiredMonthData(x) <= 0 Then
                    '    _calendar.DesiredMonthData(x) = _calendar.DesiredMonthData(x) * -1
                    'End If
                    sendData += Convert.ToString(_calendar.DesiredMonthData(x), 16).PadLeft(2, "0")
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
                Dim _obj As SCP.BLL.hs_amb_Calendar = SCP.BLL.hs_amb_Calendar.ReadFromDB(CronId, Calyear, Calmonth)
                If Not _obj Is Nothing Then
                    DataAccessHelper.GetDataAccess.hs_amb_Calendar_UpdateReal(CronId, Calyear, Calmonth, _obj.DesiredMonthData)
                    _obj = Nothing
                End If
                SCP.BLL.HeatingSystem.requestLog(hsId)
            End If

            Return retVal
        End Function

        Public Shared Function Clear(CronId) As Boolean
            If CronId <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Calendar_Clear(CronId)
        End Function

        Public Shared Function Read(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
            If CronId <= 0 Then Return Nothing
            If Calyear <= 0 Then Return Nothing
            If Calmonth <= 0 Then Return Nothing
            If Calmonth > 12 Then Return Nothing

            Dim retval As hs_amb_Calendar = DataAccessHelper.GetDataAccess.hs_amb_Calendar_Read(CronId, Calyear, Calmonth)
            If retval Is Nothing Then
                retval = CalendarGet(CronId, Calyear, Calmonth)
                If Not retval Is Nothing Then
                    retval.DesiredMonthData = retval.RealMonthData
                End If
                Add(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData)
            End If
            UpdateReal(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData)
            Dim DesiredMonthData() As Integer = New Integer(31) {}

            Dim TasksForDesired() As Integer = New Integer(31) {}
            Dim TaskArray As String = String.Empty
            Dim profile_task_list As List(Of hs_amb_Profile_Tasks) = hs_amb_Profile_Tasks.ListByMonth(CronId, Calyear, Calmonth)
            If Not profile_task_list Is Nothing Then
                Dim x As Integer = 0
                For Each _task As hs_amb_Profile_Tasks In profile_task_list
                    TasksForDesired(x) += _task.ProfileNr
                    TaskArray += _task.Subject & ";"
                    x = x + 1
                    If x > 31 Then Exit For
                Next
                Dim a As Integer = TaskArray.Length
                retval.TasksForDesired = TasksForDesired
            End If
            If Not retval.DesiredMonthData Is Nothing Then
                For x As Integer = 0 To 31
                    If TasksForDesired(x) <> retval.DesiredMonthData(x) Then
                        DesiredMonthData(x) = retval.DesiredMonthData(x)
                    Else
                        DesiredMonthData(x) = TasksForDesired(x)
                    End If
                Next
            Else
                For x As Integer = 0 To 31
                    DesiredMonthData(x) = TasksForDesired(x)
                Next
            End If

            retval.DesiredMonthData = DesiredMonthData
            Update(CronId, Calyear, Calmonth, TasksForDesired, DesiredMonthData)
            Return retval
        End Function

        Public Shared Function OLDRead(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
            If CronId <= 0 Then Return Nothing
            If Calyear <= 0 Then Return Nothing
            If Calmonth <= 0 Then Return Nothing
            If Calmonth > 12 Then Return Nothing

            'Dim rv As Boolean = CalendarGet(CronId, Calyear, Calmonth)

            'Dim retval As hs_amb_Calendar = DataAccessHelper.GetDataAccess.hs_amb_Calendar_Read(CronId, Calyear, Calmonth)
            'If retval Is Nothing Then
            '    retval = CalendarGet(CronId, Calyear, Calmonth)
            '    If Not retval Is Nothing Then
            '        If Add(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData) = False Then
            '            UpdateReal(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData)
            '        End If
            '    End If
            'End If
            Dim retval As hs_amb_Calendar = CalendarGet(CronId, Calyear, Calmonth)
            If Not retval Is Nothing Then
                Dim tempObj As SCP.BLL.hs_amb_Calendar = DataAccessHelper.GetDataAccess.hs_amb_Calendar_Read(CronId, Calyear, Calmonth)
                If tempObj Is Nothing Then
                    If Add(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData) = False Then
                        UpdateReal(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData)
                    End If
                Else
                    tempObj = Nothing
                    UpdateReal(retval.CronId, retval.Calyear, retval.Calmonth, retval.RealMonthData)
                End If
            End If

            If Not retval Is Nothing Then

                Dim Cron_Calendar As hs_amb_Calendar = DataAccessHelper.GetDataAccess.hs_amb_Calendar_Read(CronId, Calyear, Calmonth)
                If Not Cron_Calendar Is Nothing Then
                    retval.DesiredMonthData = Cron_Calendar.DesiredMonthData
                    Cron_Calendar = Nothing
                End If

                Dim DesiredMonthData() As Integer = New Integer(31) {}

                Dim TasksForDesired() As Integer = New Integer(31) {}
                Dim TaskArray As String = String.Empty
                Dim profile_task_list As List(Of hs_amb_Profile_Tasks) = hs_amb_Profile_Tasks.ListByMonth(CronId, Calyear, Calmonth)
                If Not profile_task_list Is Nothing Then
                    Dim x As Integer = 0
                    For Each _task As hs_amb_Profile_Tasks In profile_task_list
                        TasksForDesired(x) += _task.ProfileNr
                        TaskArray += _task.Subject & ";"
                        x = x + 1
                        If x > 31 Then Exit For
                    Next
                    Dim a As Integer = TaskArray.Length
                    '
                    profile_task_list = Nothing
                End If
                retval.TasksForDesired = TasksForDesired
                If Not retval.DesiredMonthData Is Nothing Then
                    For x As Integer = 0 To 31
                        'If retval.DesiredMonthData(x) < 0 Then
                        '    DesiredMonthData(x) = retval.DesiredMonthData(x) '* -1
                        'Else
                        '    DesiredMonthData(x) = TasksForDesired(x)
                        'End If

                        If TasksForDesired(x) <> retval.DesiredMonthData(x) Then
                            DesiredMonthData(x) = retval.DesiredMonthData(x)
                        Else
                            DesiredMonthData(x) = TasksForDesired(x)
                        End If
                    Next
                Else
                    For x As Integer = 0 To 31
                        DesiredMonthData(x) = TasksForDesired(x)
                    Next
                End If

                retval.DesiredMonthData = DesiredMonthData
                Update(CronId, Calyear, Calmonth, TasksForDesired, DesiredMonthData)
            End If
            Return retval
        End Function

        Public Shared Function ReadFromDB(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
            If CronId <= 0 Then Return Nothing
            If Calyear <= 0 Then Return Nothing
            If Calmonth <= 0 Then Return Nothing
            If Calmonth > 12 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_amb_Calendar_Read(CronId, Calyear, Calmonth)
        End Function

        Public Shared Function Update(CronId As Integer, Calyear As Integer, Calmonth As Integer, TasksForDesired As Integer(), DesiredMonthData As Integer()) As Boolean
            If CronId <= 0 Then Return False
            If Calyear <= 0 Then Return False
            If Calmonth <= 0 Then Return False
            If Calmonth > 12 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Calendar_Update(CronId, Calyear, Calmonth, TasksForDesired, DesiredMonthData)
        End Function

        Public Shared Function UpdateDesired(CronId As Integer, Calyear As Integer, Calmonth As Integer, DesiredMonthData As Integer()) As Boolean
            If CronId <= 0 Then Return False
            If Calyear <= 0 Then Return False
            If Calmonth <= 0 Then Return False
            If Calmonth > 12 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Calendar_UpdateDesired(CronId, Calyear, Calmonth, DesiredMonthData)
        End Function

        Public Shared Function UpdateReal(CronId As Integer, Calyear As Integer, Calmonth As Integer, RealMonthData As Integer()) As Boolean
            If CronId <= 0 Then Return False
            If Calyear <= 0 Then Return False
            If Calmonth <= 0 Then Return False
            If Calmonth > 12 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Calendar_UpdateReal(CronId, Calyear, Calmonth, RealMonthData)
        End Function
#End Region

#Region "public properties"
        Public Property CronId As Integer
        Public Property Calyear As Integer
        Public Property Calmonth As Integer
        Public Property RealMonthData As Integer()
        Public Property DesiredMonthData As Integer()
        Public Property TasksForDesired As Integer()
        Public Property LastSend As Date
#End Region
    End Class
End Namespace
