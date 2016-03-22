Imports System.Threading

Public Class HsRequest
    Public Sub Elabor(req As String)
        Try
            doWork(req)
        Catch ex As Exception
            Dim _l As New ClassLog
            _l.scriviLog("AfRequest " & ex.Message & Space(1) & req, "Afrequest.txt")
            _l = Nothing
        End Try
    End Sub

    Private Sub doWork(req As String)
        Dim markStatus As NameValueCollection = HttpUtility.ParseQueryString(req)

        Dim HeatingSystemId_Valid As Boolean = False

        Dim hsId As Integer = 0
        If Not markStatus.Get("ID") Is Nothing Then
            hsId = Convert.ToInt32(markStatus.Get("ID"))
        End If

        Dim PacketVersion As Integer = 0
        If Not markStatus.Get("V") Is Nothing Then
            PacketVersion = Convert.ToInt32(markStatus.Get("V"))
        End If

        'Dim PacketType As Integer = 0
        'If Not markStatus.Get("T") Is Nothing Then
        '    PacketType = Convert.ToInt32(markStatus.Get("T"))
        'End If

        Dim PacketDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If Not markStatus.Get("D") Is Nothing Then
            Dim reqPacketDate As String = markStatus.Get("D")
            '2015-01-14-14:21:51
            Dim strPacketDate As String = Mid(reqPacketDate, 9, 2) & "/" & Mid(reqPacketDate, 6, 2) & "/" & Mid(reqPacketDate, 1, 4) & Space(1) & Mid(reqPacketDate, 12)
            PacketDate = FormatDateTime(strPacketDate, DateFormat.GeneralDate)
        End If

        Dim SystemStatus As Integer = 0
        If Not markStatus.Get("S") Is Nothing Then
            SystemStatus = Convert.ToInt32(markStatus.Get("S"))
        End If

        Dim ErrorData As String = String.Empty
        If Not markStatus("ERR") Is Nothing Then
            ErrorData = markStatus("ERR")
        End If

        Dim HeatingSystem As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsId)
        If Not HeatingSystem Is Nothing Then
            HeatingSystemId_Valid = True
            If HeatingSystem.stato <> SystemStatus Then
                SCP.BLL.HeatingSystem.setStatus(hsId, SystemStatus)

                'solo se non sono in modalità di manutenzione mando allarme
                If HeatingSystem.MaintenanceMode = False Then

                    If SystemStatus >= 2 Then
                        'se ho ricevuto un evento di allarme. mando mail di allarme
                        If HeatingSystem.hsId > 0 Then
                            Dim s As New mailSender
                            Dim thOldLog As New Thread(AddressOf s.hsSendStatusAlarmON2clevergy)
                            thOldLog.IsBackground = True
                            thOldLog.Start(HeatingSystem)
                            s = Nothing
                        End If
                    Else
                        If HeatingSystem.stato >= 2 Then
                            'se il sistema è ancora in allarme ma ho ricevuto messaggio di allarme rientrato: mando mail di rientro
                            If SystemStatus < 2 Then
                                If HeatingSystem.hsId > 0 Then
                                    Dim s As New mailSender
                                    Dim thOldLog As New Thread(AddressOf s.hsSendStatusAlarmOFF2clevergy)
                                    thOldLog.IsBackground = True
                                    thOldLog.Start(HeatingSystem)
                                    s = Nothing
                                End If
                            End If
                        End If
                    End If


                End If 'MaintenanceMode false

            End If 'HeatingSystem.stato <> SystemStatus

            If HeatingSystem.MaintenanceMode = False Then
                If HeatingSystem.isOnline = False Then
                    Dim s As New mailSender
                    Dim thOldLog As New Thread(AddressOf s.hsSendConnectionRestored2Clevergy)
                    thOldLog.IsBackground = True
                    thOldLog.Start(HeatingSystem)
                    s = Nothing
                    SCP.BLL.hs_ErrorLog.Add(Now, HeatingSystem.hsId, "VPN01", "VPN", 0, "Connection restored")
                End If
            End If

            SCP.BLL.HeatingSystem.setlastRec(hsId, Now())

            HeatingSystem = Nothing
        End If

        If Not String.IsNullOrEmpty(ErrorData) Then
            elabor_error(hsId, ErrorData, PacketDate)
        Else
            If HeatingSystemId_Valid = True Then
                elabor_lux(hsId, markStatus, PacketDate)
                elabor_luxM(hsId, markStatus, PacketDate)
                elabor_psg(hsId, markStatus, PacketDate)
                elabor_Zrel(hsId, markStatus, PacketDate)
                elabor_cron(hsId, markStatus, PacketDate)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="hsId"></param>
    ''' <param name="ErrorData"></param>
    ''' <remarks></remarks>
    Private Sub elabor_error(hsId As Integer, ErrorData As String, PacketDate As Date)
        Dim Level As Integer = 0
        Dim Obj As Integer = 0
        Dim ObjIndex As Integer = 0
        Dim ErrorCode As Integer = 0
        Dim ErrorMsg As String = String.Empty

        Dim ar() As String = Split(ErrorData, ";")
        Level = Convert.ToInt32(ar(0))
        Obj = Convert.ToInt32(ar(1))
        ObjIndex = Convert.ToInt32(ar(2))
        ErrorCode = Convert.ToInt32(ar(3))
        ErrorMsg = ar(4)

        '
        'If Obj = "TP" Then Obj = "S"

        Dim elementCode As String = String.Empty
        Dim _hsElem As SCP.BLL.tbhsElem = SCP.BLL.tbhsElem.Read(Obj)
        If Not _hsElem Is Nothing Then
            elementCode = RTrim(_hsElem.name)
            _hsElem = Nothing
        End If
        Dim hselement As String = elementCode & ObjIndex.ToString.PadLeft(2, "0")

        SCP.BLL.hs_ErrorLog.Add(PacketDate, hsId, hselement, elementCode, ErrorCode, ErrorMsg)


        Select Case Obj
            Case Is = 0
                error_SYS(hsId, hselement, elementCode, ErrorCode, ErrorMsg, PacketDate)

            Case Is = 1
                'VPN
            Case Is = 2
                'DOOR

            Case Is = 10
                'CAL

            Case Is = 11
                'CIR

            Case Is = 12
                'CRON

            Case Is = 13
                'VRD

            Case Is = 14
                'CTL

            Case Is = 15
                'CTB

            Case Is = 16
                'TP

            Case Is = 17
                'PP
            Case Is = 18
                'PB
   
            Case Is = 19
                'TB
   
            Case Is = 20
                'PM
   
            Case Is = 21
                'GRU
   
            Case Is = 22
                'PDC

            Case Is = 25
                'GAS
   
            Case Is = 26
                'PORT
   
            Case Is = 27
                'FPV

            Case Is = 28
                'ANZ
   
            Case Is = 29
                'SCA

            Case Is = 30
                'LUX
                error_lux(hsId, hselement, elementCode, ErrorCode, ErrorMsg, PacketDate)

            Case Is = 31
                'LUXM
                error_luxM(hsId, hselement, elementCode, ErrorCode, ErrorMsg, PacketDate)

            Case Is = 32
                'psg
                error_psg(hsId, hselement, elementCode, ErrorCode, ErrorMsg, PacketDate)

            Case Is = 33
                'HVAC

            Case Is = 202
                'ZREL
                error_zrel(hsId, hselement, elementCode, ErrorCode, ErrorMsg, PacketDate)

            Case Is = 1003
                'Bus - MODBUS TCP
                error_Modbus_tcp(hsId, hselement, elementCode, ErrorCode, ErrorMsg, PacketDate)

        End Select
    End Sub

#Region "elaborazione dati ricevuti"

    ''' <summary>
    ''' base illumination point
    ''' </summary>
    ''' <param name="hsId"></param>
    ''' <param name="markStatus"></param>
    ''' <param name="PacketDate"></param>
    ''' <remarks></remarks>
    Private Sub elabor_lux(hsId As Integer, markStatus As NameValueCollection, PacketDate As Date)
        'Dim key As String = String.Empty
        'For Each key In markStatus.Keys
        '    If Trim(key).Contains("LUX") Then
        '        Dim Callreq As String = markStatus.Get(key)
        '        Dim _obj As SCP.BLL.Lux = SCP.BLL.Lux.ReadByCod(hsId, key)
        '        If _obj Is Nothing Then
        '            SCP.BLL.Lux.Add(hsId, key, key, "SYS", String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate))
        '            'SCP.BLL.Lux.setValue(hsId, key, WorkingTimeCounter, PowerOnCycleCounter)
        '            'SCP.BLL.Lux.setStatus(hsId, key, stato)
        '            ''SCP.BLL.log_Lux.Add(hsId, key, key, WorkingTimeCounter, PowerOnCycleCounter, stato, PacketDate)
        '        End If
        '    End If
        'Next key


        Dim _list As List(Of SCP.BLL.Lux) = SCP.BLL.Lux.List(hsId)
        If Not _list Is Nothing Then
            For Each _obj As SCP.BLL.Lux In _list
                Dim Callreq As String = markStatus.Get(_obj.Cod)
                If Not String.IsNullOrEmpty(Callreq) Then
                    '-------------------------------------------------------
                    Dim LightON As Boolean = False
                    Dim WorkingTimeCounter As Decimal = 0
                    Dim PowerOnCycleCounter As Decimal = 0
                    Dim CurrentMode As Integer = 0
                    Dim stato As Integer = 0

                    Dim ar() As String = Split(Callreq, ";")
                    If ar.Length >= 1 Then
                        If ar(0) = "FALSE" Then
                            LightON = False
                        Else
                            LightON = True
                        End If
                    End If
                    If ar.Length >= 2 Then
                        WorkingTimeCounter = CDec(ar(1))
                    End If
                    If ar.Length >= 3 Then
                        PowerOnCycleCounter = CDec(ar(2))
                    End If

                    If ar.Length >= 4 Then
                        CurrentMode = CInt(ar(3))
                    End If
                    If _obj.isManual = True Then
                        CurrentMode = 0 ' manual mode
                    End If

                    If ar.Length >= 5 Then
                        stato = CInt(ar(4))
                    End If

                    Dim forcedOn As Boolean = False
                    If ar.Length >= 6 Then
                        forcedOn = CBool(ar(5))
                    End If

                    Dim forcedOff As Boolean = False
                    If ar.Length >= 7 Then
                        forcedOff = CBool(ar(6))
                    End If
                    '-------------------------------------------------------

                    SCP.BLL.Lux.setValue(hsId, _obj.Cod, LightON, WorkingTimeCounter, PowerOnCycleCounter, CurrentMode, forcedOn, forcedOff)
                    SCP.BLL.Lux.setStatus(hsId, _obj.Cod, stato)

                    Dim ok2Write As Boolean = False
                    Dim _lastLog As SCP.BLL.Lux_last = SCP.BLL.Lux_last.Read(hsId, _obj.Cod)

                    If Not _lastLog Is Nothing Then

                        'se maggiore dell'ultimo log devono essere passati almeno 5 min
                        If PacketDate > _lastLog.lastdtLog Then
                            If DateDiff(DateInterval.Minute, _lastLog.lastdtLog, PacketDate) >= 5 Then ok2Write = True
                        End If
                        _lastLog = Nothing
                    Else
                        ok2Write = True
                    End If

                    If ok2Write = True Then

                        Dim lastidlog As Integer = SCP.BLL.log_Lux.Add(hsId, _obj.Cod, _obj.Descr, WorkingTimeCounter, PowerOnCycleCounter, stato, _obj.LightON, PacketDate)
                        SCP.BLL.Lux_last.Upsert(hsId, _obj.Cod, lastidlog, PacketDate)

                    End If




                    ''SCP.BLL.log_Lux.Add(hsId, _obj.Cod, _obj.Descr, WorkingTimeCounter, PowerOnCycleCounter, stato, _obj.LightON, PacketDate)

                End If
            Next
            _list = Nothing
        End If

    End Sub

    ''' <summary>
    ''' metered illumination point
    ''' </summary>
    ''' <param name="hsId"></param>
    ''' <param name="markStatus"></param>
    ''' <param name="PacketDate"></param>
    ''' <remarks></remarks>
    Private Sub elabor_luxM(hsId As Integer, markStatus As NameValueCollection, PacketDate As Date)

        Dim _list As List(Of SCP.BLL.LuxM) = SCP.BLL.LuxM.List(hsId)
        If Not _list Is Nothing Then
            For Each _obj As SCP.BLL.LuxM In _list
                Dim Callreq As String = markStatus.Get(_obj.Cod)
                If Not String.IsNullOrEmpty(Callreq) Then
                    Dim Voltage As Decimal = 0
                    Dim Curr As Decimal = 0
                    Dim EnergyCounter As Decimal
                    Dim WorkingTimeCounter As Decimal = 0
                    Dim PowerOnCycleCounter As Decimal = 0
                    Dim Temp As Decimal = 0
                    Dim stato As Integer = 0

                    Dim ar() As String = Split(Callreq, ";")
                    If ar.Length >= 1 Then
                        Voltage = CDec(ar(0)) / 10
                    End If
                    If ar.Length >= 2 Then
                        Curr = CDec(ar(1)) / 10
                    End If
                    If ar.Length >= 3 Then
                        EnergyCounter = CDec(ar(2)) / 10
                    End If
                    If ar.Length >= 4 Then
                        WorkingTimeCounter = CDec(ar(3))
                    End If
                    If ar.Length >= 5 Then
                        PowerOnCycleCounter = CDec(ar(4))
                    End If
                    If ar.Length >= 6 Then
                        Temp = CDec(ar(5)) / 10
                    End If
                    If ar.Length >= 7 Then
                        stato = CInt(ar(6))
                    End If

                    Dim forcedOn As Boolean = False
                    If ar.Length >= 8 Then
                        forcedOn = CBool(ar(7))
                    End If

                    Dim forcedOff As Boolean = False
                    If ar.Length >= 9 Then
                        forcedOff = CBool(ar(8))
                    End If

                    SCP.BLL.LuxM.setStatus(hsId, _obj.Cod, stato)
                    SCP.BLL.LuxM.setValue(hsId, _obj.Cod, Voltage, Curr, EnergyCounter, WorkingTimeCounter, PowerOnCycleCounter, Temp)


                    Dim ok2Write As Boolean = False
                    Dim _lastLog As SCP.BLL.LuxM_last = SCP.BLL.LuxM_last.Read(hsId, _obj.Cod)

                    If Not _lastLog Is Nothing Then

                        'se maggiore dell'ultimo log devono essere passati almeno 5 min
                        If PacketDate > _lastLog.lastdtLog Then
                            If DateDiff(DateInterval.Minute, _lastLog.lastdtLog, PacketDate) >= 5 Then ok2Write = True
                        End If
                        _lastLog = Nothing
                    Else
                        ok2Write = True
                    End If

                    If ok2Write = True Then

                        Dim lastidlog As Integer = SCP.BLL.log_LuxM.Add(hsId, _obj.Cod, _obj.Descr, Voltage, Curr, EnergyCounter, WorkingTimeCounter, PowerOnCycleCounter, Temp, stato, _obj.LightON, PacketDate)
                        SCP.BLL.LuxM_last.Upsert(hsId, _obj.Cod, lastidlog, PacketDate)

                    End If
                End If

            Next
            _list = Nothing
        End If
    End Sub

    ''' <summary>
    ''' contatori di passaggio
    ''' </summary>
    ''' <param name="hsId"></param>
    ''' <param name="markStatus"></param>
    ''' <param name="PacketDate"></param>
    ''' <remarks></remarks>
    Private Sub elabor_psg(hsId As Integer, markStatus As NameValueCollection, PacketDate As Date)
        Dim key As String = String.Empty
        For Each key In markStatus.Keys
            If Trim(key).Contains("PSG") Then
                Dim Callreq As String = markStatus.Get(key)
                Dim _obj As SCP.BLL.Psg = SCP.BLL.Psg.ReadByCod(hsId, key)
                If _obj Is Nothing Then
                    SCP.BLL.Psg.Add(hsId, key, key, "SYS", String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate))
                    'SCP.BLL.Lux.setValue(hsId, key, WorkingTimeCounter, PowerOnCycleCounter)
                    'SCP.BLL.Lux.setStatus(hsId, key, stato)
                    ''SCP.BLL.log_Lux.Add(hsId, key, key, WorkingTimeCounter, PowerOnCycleCounter, stato, PacketDate)
                End If
            End If
        Next key
        'Dim psg As String = markStatus.Get("PSG")
        'If Not String.IsNullOrEmpty(psg) Then
        '    Dim ar() As String = Split(psg, ";")
        '    For x As Integer = 0 To ar.Length - 1
        '        Dim Cod As String = "PSG" & (x + 1).ToString.PadLeft(2, "0")
        '        Dim Val As Integer = Convert.ToInt32(ar(x))

        '        Dim _obj As SCP.BLL.Psg = SCP.BLL.Psg.ReadByCod(hsId, Cod)
        '        If Not _obj Is Nothing Then
        '            SCP.BLL.Psg.setValue(_obj.hsId, _obj.Cod, Val)
        '            SCP.BLL.log_Psg.Add(_obj.hsId, _obj.Cod, _obj.Descr, Val, _obj.stato, PacketDate)
        '            _obj = Nothing
        '        Else
        '            If SCP.BLL.Psg.Add(hsId, Cod, Cod, "SYS", String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate)) = True Then
        '                SCP.BLL.Psg.setValue(hsId, Cod, Val)
        '                SCP.BLL.log_Psg.Add(hsId, Cod, Cod, Val, 0, PacketDate)
        '            End If
        '        End If
        '    Next
        'End If

    End Sub

    ''' <summary>
    ''' ZigBee(X-Monitor) Actuator Rele
    ''' </summary>
    ''' <param name="hsId"></param>
    ''' <param name="markStatus"></param>
    ''' <param name="PacketDate"></param>
    ''' <remarks></remarks>
    Private Sub elabor_Zrel(hsId As Integer, markStatus As NameValueCollection, PacketDate As Date)
        'Dim key As String = String.Empty
        'For Each key In markStatus.Keys
        '    If Trim(key).Contains("ZREL") Then
        '        Dim Callreq As String = markStatus.Get(key)

        '        Dim CurrentId As Integer = 0
        '        Dim LQI As Integer = 0
        '        Dim Temperature As Decimal = 0
        '        Dim MeshParentId As Integer = 0
        '        Dim stato As Integer = 0

        '        Dim ar() As String = Split(Callreq, ";")
        '        If ar.Length >= 1 Then
        '            Temperature = CDec(ar(0)) / 10
        '        End If
        '        If ar.Length >= 2 Then
        '            MeshParentId = CInt(ar(1))
        '        End If
        '        If ar.Length >= 3 Then
        '            LQI = CInt(ar(2))
        '        End If
        '        If ar.Length >= 4 Then
        '            stato = CInt(ar(3))
        '        End If


        '        Dim _obj As SCP.BLL.Zrel = SCP.BLL.Zrel.ReadByCod(hsId, key)
        '        If _obj Is Nothing Then
        '            SCP.BLL.Zrel.Add(hsId, key, key, "SYS", String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate))
        '            SCP.BLL.Zrel.setValue(hsId, key, LQI, Temperature, MeshParentId, CurrentId)
        '            SCP.BLL.Zrel.setStatus(hsId, key, stato)
        '            SCP.BLL.log_Zrel.Add(hsId, key, key, LQI, Temperature, MeshParentId, CurrentId, stato, PacketDate)
        '        Else
        '            SCP.BLL.Zrel.setValue(hsId, key, LQI, Temperature, MeshParentId, CurrentId)
        '            SCP.BLL.Zrel.setStatus(hsId, key, stato)
        '            SCP.BLL.log_Zrel.Add(hsId, key, key, LQI, Temperature, MeshParentId, CurrentId, stato, PacketDate)

        '            _obj = Nothing
        '        End If

        '    End If
        'Next key

    End Sub
    ''' <summary>
    ''' Cron
    ''' </summary>
    ''' <param name="hsId"></param>
    ''' <param name="markStatus"></param>
    ''' <remarks></remarks>
    Private Sub elabor_cron(hsId As Integer, markStatus As NameValueCollection, PacketDate As Date)

        'Dim key As String = String.Empty
        'For Each key In markStatus.Keys
        '    If Trim(key).Contains("CRON") Then
        '        Dim Callreq As String = markStatus.Get(key)
        '        Dim _obj As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.ReadByCronCod(hsId, key)
        '        If _obj Is Nothing Then
        '            SCP.BLL.hs_Cron.Add(hsId, key, key, "nota", "SYS", String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 1)
        '            'SCP.BLL.Lux.setValue(hsId, key, WorkingTimeCounter, PowerOnCycleCounter)
        '            'SCP.BLL.Lux.setStatus(hsId, key, stato)
        '            ''SCP.BLL.log_Lux.Add(hsId, key, key, WorkingTimeCounter, PowerOnCycleCounter, stato, PacketDate)
        '        End If
        '    End If
        'Next key


        'Try
        '    Dim _list As List(Of SCP.BLL.hs_Cron) = SCP.BLL.hs_Cron.List(hsId)
        '    If Not _list Is Nothing Then
        '        For Each _obj As SCP.BLL.hs_Cron In _list

        '            Dim m_done As Boolean = False

        '            Dim Callreq As String = markStatus.Get(_obj.CronCod)
        '            If Not String.IsNullOrEmpty(Callreq) Then
        '                Dim ar() As String = Split(Callreq, ";")
        '                Dim setPoint As Decimal = Convert.ToInt32(ar(0)) / 10
        '                Dim stato As Integer = Convert.ToInt32(ar(1))
        '                SCP.BLL.hs_Cron.setStatus(hsId, _obj.CronCod, setPoint, stato)

        '                Dim ok2Write As Boolean = False
        '                Dim _lastLog As SCP.BLL.log_hs_Cron = SCP.BLL.log_hs_Cron.ReadLast(hsId, _obj.CronCod)
        '                If Not _lastLog Is Nothing Then
        '                    'se minore dell'ultimo log sto caricando un dato storico
        '                    If PacketDate < _lastLog.dtLog Then ok2Write = True
        '                    'se maggiore dell'ultimo log devono essere passati almeno 5 min
        '                    If PacketDate > _lastLog.dtLog Then
        '                        If DateDiff(DateInterval.Minute, _lastLog.dtLog, PacketDate) >= 5 Then ok2Write = True
        '                    End If
        '                    _lastLog = Nothing
        '                Else
        '                    ok2Write = True
        '                End If
        '                If ok2Write = True Then
        '                    SCP.BLL.log_hs_Cron.Add(hsId, setPoint, _obj.CronCod, _obj.CronDescr, stato, PacketDate)
        '                End If

        '                m_done = True
        '            Else
        '                'se non trovo CRON, provo con THER
        '                Dim therCod As String = Replace(_obj.CronCod, "CRON", "THER")
        '                Callreq = markStatus.Get(therCod)
        '                If Not String.IsNullOrEmpty(Callreq) Then
        '                    Dim ar() As String = Split(Callreq, ";")
        '                    Dim setPoint As Decimal = Convert.ToInt32(ar(0)) / 10 'ATTENZIONE. pare non si debba dividere per 10
        '                    Dim stato As Integer = Convert.ToInt32(ar(1))
        '                    SCP.BLL.hs_Cron.setStatus(hsId, _obj.CronCod, setPoint, stato)

        '                    Dim ok2Write As Boolean = False
        '                    Dim _lastLog As SCP.BLL.log_hs_Cron = SCP.BLL.log_hs_Cron.ReadLast(hsId, _obj.CronCod)
        '                    If Not _lastLog Is Nothing Then
        '                        'se minore dell'ultimo log sto caricando un dato storico
        '                        If PacketDate < _lastLog.dtLog Then ok2Write = True
        '                        'se maggiore dell'ultimo log devono essere passati almeno 5 min
        '                        If PacketDate > _lastLog.dtLog Then
        '                            If DateDiff(DateInterval.Minute, _lastLog.dtLog, PacketDate) >= 5 Then ok2Write = True
        '                        End If
        '                        _lastLog = Nothing
        '                    Else
        '                        ok2Write = True
        '                    End If
        '                    If ok2Write = True Then
        '                        SCP.BLL.log_hs_Cron.Add(hsId, setPoint, _obj.CronCod, _obj.CronDescr, stato, PacketDate)
        '                    End If

        '                    m_done = True
        '                End If
        '            End If

        '            If m_done = False Then
        '                'ho sicuramente dei cronotermostati
        '                'ma non sono riuscito ad aggiornarli perchè il log non arriva o non è corretto
        '                'elabor_cron_withNoLog(_obj, PacketDate)
        '            End If
        '        Next
        '        _list = Nothing
        '    End If
        'Catch ex As Exception

        'End Try
    End Sub
#End Region

#Region "elaborazione degli errori"
    Private Sub error_SYS(hsId As Integer, hselement As String, elementCode As String, ErrorCode As String, ErrorMsg As String, PacketDate As Date)
        'SYS
        If ErrorCode = 2 Then
            'log Logger connection timeout
            'SCP.BLL.HeatingSystem.setStatus(hsId, 0)
        End If
        If ErrorCode = 1 Or ErrorCode = 3 Then
            'ErrorCode = 1 warning System boot and software version
            'ErrorCode = 3 warning Logger warning lot of pending messages
            'SCP.BLL.HeatingSystem.setStatus(hsId, 1)
            SCP.BLL.hs_ErrorLog.Add(Now, hsId, "SYS", "SYS", 0, "System restart")
        End If
        If ErrorCode = 4 Then
            'Alarm PLC Watchdog
            'SCP.BLL.HeatingSystem.setStatus(hsId, 2)
        End If
        'SCP.BLL.HeatingSystem.setStatus(hsId, 0)
        'SCP.BLL.HeatingSystem.resetSystemStatus(hsId)
    End Sub

    Private Sub error_lux(hsId As Integer, hselement As String, elementCode As String, ErrorCode As Integer, ErrorMsg As String, PacketDate As Date)
        Dim currentStatus As Integer = 0

        If ErrorCode = 1 Then
            'log light OK
            currentStatus = 0
        End If

        If ErrorCode = 2 Then
            'log light OFF
            currentStatus = 0
            SCP.BLL.Lux.setLightByCod(hsId, hselement, False)
        End If
        If ErrorCode = 3 Then
            'log light ON
            currentStatus = 0
            SCP.BLL.Lux.setLightByCod(hsId, hselement, True)
        End If


        If ErrorCode = 4 Then
            'warning load low
            currentStatus = 1
        End If
        If ErrorCode = 5 Then
            'warning load hight
            currentStatus = 1
        End If
        If ErrorCode = 6 Then
            'alarm disconnected
            currentStatus = 2
        End If
        If ErrorCode = 7 Then
            'alarm overload
            currentStatus = 2
        End If
        If ErrorCode = 8 Then
            'alarm control failure
            currentStatus = 2
        End If
        If ErrorCode = 9 Then
            'alarm feedback failure
            currentStatus = 2
        End If

        Dim _obj As SCP.BLL.Lux = SCP.BLL.Lux.ReadByCod(hsId, hselement)
        If Not _obj Is Nothing Then
            SCP.BLL.Lux.setStatus(hsId, _obj.Cod, currentStatus)
            SCP.BLL.log_Lux.Add(hsId, _obj.Cod, _obj.Descr, _obj.WorkingTimeCounter, _obj.PowerOnCycleCounter, currentStatus, _obj.LightON, PacketDate)

            If currentStatus = 2 Then
                'ticket generation
                '--------------------------------------------------------------------------------------------
                Dim ticketDescription As String = ErrorMsg
                Dim errorCodes As SCP.BLL.hs_ErrorCodes = SCP.BLL.hs_ErrorCodes.Read(elementCode, ErrorCode)
                If Not errorCodes Is Nothing Then
                    ticketDescription = errorCodes.DescIT
                End If
                creatTicket(hsId, hselement, _obj.Id, _obj.Descr, ticketDescription)
                '--------------------------------------------------------------------------------------------
            End If

            _obj = Nothing
        End If
    End Sub

    Private Sub error_luxM(hsId As Integer, hselement As String, elementCode As String, ErrorCode As Integer, ErrorMsg As String, PacketDate As Date)
        Dim currentStatus As Integer = 0

        If ErrorCode = 1 Then
            'log light OK
            currentStatus = 0
        End If

        If ErrorCode = 2 Then
            'log light OFF
            currentStatus = 0
            SCP.BLL.LuxM.setLightByCod(hsId, hselement, False)
        End If
        If ErrorCode = 3 Then
            'log light ON
            currentStatus = 0
            SCP.BLL.LuxM.setLightByCod(hsId, hselement, True)
        End If

        If ErrorCode = 4 Then
            'warning load low
            currentStatus = 1
        End If
        If ErrorCode = 5 Then
            'warning load hight
            currentStatus = 1
        End If
        If ErrorCode = 6 Then
            'alarm disconnected
            currentStatus = 2
        End If
        If ErrorCode = 7 Then
            'alarm overload
            currentStatus = 2
        End If
        If ErrorCode = 8 Then
            'alarm control failure
            currentStatus = 2
        End If
        If ErrorCode = 9 Then
            'alarm feedback failure
            currentStatus = 2
        End If

        Dim _obj As SCP.BLL.LuxM = SCP.BLL.LuxM.ReadByCod(hsId, hselement)
        If Not _obj Is Nothing Then
            SCP.BLL.LuxM.setStatus(hsId, _obj.Cod, currentStatus)
            SCP.BLL.log_LuxM.Add(hsId, _obj.Cod, _obj.Descr, _obj.Voltage, _obj.Curr, _obj.EnergyCounter, _obj.WorkingTimeCounter, _obj.PowerOnCycleCounter, _obj.Temp, currentStatus, _obj.LightON, PacketDate)

            If currentStatus = 2 Then
                'ticket generation
                '--------------------------------------------------------------------------------------------
                Dim ticketDescription As String = ErrorMsg
                Dim errorCodes As SCP.BLL.hs_ErrorCodes = SCP.BLL.hs_ErrorCodes.Read(elementCode, ErrorCode)
                If Not errorCodes Is Nothing Then
                    ticketDescription = errorCodes.DescIT
                End If
                creatTicket(hsId, hselement, _obj.Id, _obj.Descr, ticketDescription)
                '--------------------------------------------------------------------------------------------
            End If

            _obj = Nothing
        End If
    End Sub

    Private Sub error_psg(hsId As Integer, hselement As String, elementCode As String, ErrorCode As Integer, ErrorMsg As String, PacketDate As Date)
        Dim currentStatus As Integer = 0

        If ErrorCode = 2 Then
            'alarm Occupied barrier alarm timeout
            currentStatus = 2
        End If
        If ErrorCode = 3 Then
            'warning Occupied barrier warning timeout
            currentStatus = 1
        End If
        If ErrorCode = 4 Then
            'warning Inactive barrier warning timeout
            currentStatus = 1
        End If
        If ErrorCode = 5 Then
            'warning Passage counter overflow
            currentStatus = 1
        End If

        Dim _obj As SCP.BLL.Psg = SCP.BLL.Psg.ReadByCod(hsId, hselement)
        If Not _obj Is Nothing Then
            SCP.BLL.Psg.setStatus(hsId, _obj.Cod, currentStatus)
            SCP.BLL.log_Psg.Add(hsId, _obj.Cod, _obj.Descr, _obj.currentValue, currentStatus, PacketDate)

            If currentStatus = 2 Then
                'ticket generation
                '--------------------------------------------------------------------------------------------
                Dim ticketDescription As String = ErrorMsg
                Dim errorCodes As SCP.BLL.hs_ErrorCodes = SCP.BLL.hs_ErrorCodes.Read(elementCode, ErrorCode)
                If Not errorCodes Is Nothing Then
                    ticketDescription = errorCodes.DescIT
                End If
                creatTicket(hsId, hselement, _obj.Id, _obj.Descr, ticketDescription)
                '--------------------------------------------------------------------------------------------
            End If

            _obj = Nothing
        End If

    End Sub

    Private Sub error_zrel(hsId As Integer, hselement As String, elementCode As String, ErrorCode As Integer, ErrorMsg As String, PacketDate As Date)
        Dim currentStatus As Integer = 0

        If ErrorCode = 2 Then
            'warning Low quality connection (LQI<100)
            currentStatus = 1
        End If
        If ErrorCode = 3 Then
            'Alarm Node offline
            currentStatus = 2
        End If
        If ErrorCode = 4 Then
            'Alarm Modbus communication error
            currentStatus = 2
        End If
        If ErrorCode = 5 Then
            'warning Temperature warning (T<-5 or T>60)
            currentStatus = 1
        End If
        If ErrorCode = 6 Then
            'Alarm Hardware failure
            currentStatus = 2
        End If

        Dim _obj As SCP.BLL.Zrel = SCP.BLL.Zrel.ReadByCod(hsId, hselement)
        If Not _obj Is Nothing Then
            SCP.BLL.Zrel.setStatus(hsId, _obj.Cod, currentStatus)
            SCP.BLL.log_Zrel.Add(hsId, _obj.Cod, _obj.Descr, _obj.LQI, _obj.Temperature, _obj.MeshParentId, _obj.CurrentId, currentStatus, PacketDate)

            If currentStatus = 2 Then
                'ticket generation
                '--------------------------------------------------------------------------------------------
                Dim ticketDescription As String = ErrorMsg
                Dim errorCodes As SCP.BLL.hs_ErrorCodes = SCP.BLL.hs_ErrorCodes.Read(elementCode, ErrorCode)
                If Not errorCodes Is Nothing Then
                    ticketDescription = errorCodes.DescIT
                End If
                creatTicket(hsId, hselement, _obj.Id, _obj.Descr, ticketDescription)
                '--------------------------------------------------------------------------------------------
            End If

            _obj = Nothing
        End If

    End Sub

    Private Sub error_Modbus_tcp(hsId As Integer, hselement As String, elementCode As String, ErrorCode As Integer, ErrorMsg As String, PacketDate As Date)
        Dim currentStatus As Integer = 0

        If ErrorCode = 1 Then
            'Alarm Interface error
            currentStatus = 2
        End If
        If ErrorCode = 2 Then
            'log Interface OK
            currentStatus = 0
        End If
        If ErrorCode = 3 Then
            'warning Request response error
            currentStatus = 1
        End If
        If ErrorCode = 4 Then
            'warning Request extended error from device (Fault code in string)
            currentStatus = 1
        End If
        If ErrorCode = 5 Then
            'log Request response OK
            currentStatus = 0
        End If
        Dim heatingsystem As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsId)
        If Not heatingsystem Is Nothing Then
            If currentStatus > heatingsystem.stato Then
                SCP.BLL.HeatingSystem.setStatus(hsId, currentStatus)
            End If
            heatingsystem = Nothing
        End If

        SCP.BLL.hs_ErrorLog.Add(Now, hsId, "SYS", "MODBUS_TCP", ErrorCode, ErrorMsg)
    End Sub

    Private Sub creatTicket(hsid As Integer, elementName As String, elementId As Integer, title As String, description As String)
        Dim ok2go As Boolean = False
        Dim oldTickets As SCP.BLL.hs_Tickets = SCP.BLL.hs_Tickets.readByElement(hsid, elementName, elementId)
        If oldTickets Is Nothing Then
            ok2go = True
        Else
            oldTickets = Nothing
        End If

        If ok2go = True Then
            Dim requester As String = String.Empty
            Dim heatingsystem As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsid)
            If Not heatingsystem Is Nothing Then
                requester = heatingsystem.DesImpianto & Space(1) & heatingsystem.Descr
                heatingsystem = Nothing
            End If

            Dim ticketExecutor As SCP.BLL.hs_Tickets_Executors = Nothing
            Dim executorList As List(Of SCP.BLL.hs_Tickets_Executors) = SCP.BLL.hs_Tickets_Executors.List(hsid)
            If Not executorList Is Nothing Then
                ticketExecutor = executorList(0)
                executorList = Nothing
            End If

            If Not ticketExecutor Is Nothing Then
               SCP.BLL.hs_Tickets.AddBySystem(hsid, title, requester, String.Empty, description, ticketExecutor.Nome, ticketExecutor.emailaddress, "system", elementName, elementId) 
            End If
        End If

    End Sub
#End Region

End Class
