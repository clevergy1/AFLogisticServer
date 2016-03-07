﻿'
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

Imports Telerik.Windows.Controls.Scheduler
Imports Telerik.Windows.Controls.Scheduler.ICalendar


Namespace SCP.BLL
    Public Class hs_Cron_Profile_Tasks
#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, String.Empty, False, String.Empty, String.Empty)
        End Sub
        Public Sub New(_TaskId As Integer, _
                       _CronId As Integer, _
                       _ProfileNr As Integer, _
                       _Subject As String, _
                       _StartDate As Date, _
                       _EndDate As Date, _
                       _RecurrencePattern As String, _
                       _ExceptionAppointments As String, _
                       _yearsRepeatable As Boolean, _
                       _CronCod As String, _
                       _CronDescr As String)
            TaskId = _TaskId
            CronId = _CronId
            ProfileNr = _ProfileNr
            Subject = _Subject
            StartDate = _StartDate
            EndDate = _EndDate
            RecurrencePattern = _RecurrencePattern
            ExceptionAppointments = _ExceptionAppointments
            yearsRepeatable = _yearsRepeatable
            CronCod = _CronCod
            CronDescr = _CronDescr
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(CronId As Integer, _
                                   ProfileNr As Integer, _
                                   Subject As String, _
                                   StartDate As String, _
                                   EndDate As String, _
                                   chkMonday As Boolean, _
                                   chkTuesday As Boolean, _
                                   chkWednesday As Boolean, _
                                   chkThursday As Boolean, _
                                   chkFriday As Boolean, _
                                   chkSaturday As Boolean, _
                                   chkSunday As Boolean, _
                                   yearsRepeatable As Boolean) As Boolean
            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            If String.IsNullOrEmpty(Subject) Then Return False

            Dim m_RecurrencePattern As String = String.Empty
            Dim m_RecurrenceByDay As String = String.Empty
            Dim m_RecurrenceCount As String = String.Empty
            Dim m_RecurrenceUntil As String = String.Empty

            Dim m_taskDateTime As New DateTime
            Dim m_endTask As Date = FormatDateTime("01/01/1900", DateFormat.LongTime)
            Dim m_EndBy As New DateTime
            Dim m_Mon As Boolean = False
            Dim m_Tue As Boolean = False
            Dim m_Wed As Boolean = False
            Dim m_Thu As Boolean = False
            Dim m_Fri As Boolean = False
            Dim m_Sat As Boolean = False
            Dim m_Sun As Boolean = False
            Dim m_subject As String = String.Empty


            Dim txtStartDate As String = StartDate
            Dim txtEndDate As String = EndDate

            m_subject = Subject

            m_taskDateTime = DateTime.Parse(txtStartDate)
            m_endTask = DateAdd(DateInterval.Day, 1, m_taskDateTime)
            m_endTask = DateAdd(DateInterval.Minute, -1, m_endTask)

            m_Mon = chkMonday
            m_Tue = chkTuesday
            m_Wed = chkWednesday
            m_Thu = chkThursday
            m_Fri = chkFriday
            m_Sat = chkSaturday
            m_Sun = chkSunday

            If m_Mon = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "MO,"
            End If
            If m_Tue = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "TU,"
            End If
            If m_Wed = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "WE,"
            End If
            If m_Thu = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "TH,"
            End If
            If m_Fri = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "FR,"
            End If
            If m_Sat = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "SA,"
            End If
            If m_Sun = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "SU,"
            End If
            If Not String.IsNullOrEmpty(m_RecurrenceByDay) Then
                m_RecurrenceByDay = m_RecurrenceByDay.Remove(m_RecurrenceByDay.Length - 1)
            End If

            If IsDate(txtEndDate) Then
                'm_EndBy = DateAdd(DateInterval.Day, 1, m_EndBy)
                'm_EndBy = DateAdd(DateInterval.Minute, -1, m_EndBy)

                m_EndBy = DateTime.Parse(txtEndDate).Year & "/" & _
                    DateTime.Parse(txtEndDate).Month.ToString.PadLeft(2, "0") & "/" & _
                    DateTime.Parse(txtEndDate).Day.ToString.PadLeft(2, "0") & _
                    "T23:59"
                EndDate = m_EndBy
                m_RecurrenceUntil = "UNTIL=" & m_EndBy.Year.ToString & _
                               m_EndBy.Month.ToString.PadLeft(2, "0") & _
                               m_EndBy.Day.ToString.PadLeft(2, "0") & _
                              "T070000Z;"
                m_RecurrencePattern += m_RecurrenceUntil
            Else
                EndDate = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            End If
            If Not String.IsNullOrEmpty(m_RecurrenceByDay) Then
                m_RecurrencePattern += m_RecurrenceByDay
            End If

            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_Add(CronId, ProfileNr, Subject, StartDate, EndDate, m_RecurrencePattern, String.Empty, yearsRepeatable)
        End Function

        Public Shared Function Del(CronId As Integer, ProfileNr As Integer) As Boolean
            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_Del(CronId, ProfileNr)
        End Function

        Public Shared Function List(CronId As Integer) As List(Of hs_Cron_Profile_Tasks)
            If CronId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_List(CronId)
        End Function

        Public Shared Function ListAll() As List(Of hs_Cron_Profile_Tasks)
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_ListAll()
        End Function

        Public Shared Function ListCurrent(CronId As Integer, selDate As Date) As List(Of hs_Cron_Profile_Tasks)
            If CronId <= 0 Then Return Nothing

            Dim workList As New List(Of hs_Cron_Profile_Tasks)

            Dim profileTaskList As List(Of hs_Cron_Profile_Tasks) = DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_List(CronId)
            If Not profileTaskList Is Nothing Then
                For Each _obj As hs_Cron_Profile_Tasks In profileTaskList
                    Dim endDate As Date = DateAdd(DateInterval.Day, 1, selDate)
                    endDate = DateAdd(DateInterval.Second, -1, endDate)

                    Dim app As New Appointment
                    app.Start = _obj.StartDate
                    app.End = DateAdd(DateInterval.Minute, 1, _obj.StartDate)

                    If Not String.IsNullOrEmpty(_obj.RecurrencePattern) Then
                        Dim ar1() As String = _obj.RecurrencePattern.Split(";")
                        For x As Integer = 0 To ar1.Length - 1
                            Dim RecurrenceMode = Mid(ar1(x), 1, 5)
                            Select Case RecurrenceMode
                                Case "UNTIL"
                                    Dim UNTIL As String = ar1(x).Replace("UNTIL=", String.Empty)
                                    Dim UntilDate As Date
                                    If _obj.yearsRepeatable = False Then
                                        UntilDate = UNTIL.Substring(0, 4) & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                    Else
                                        app.Start = selDate.Year & "/" & _obj.StartDate.Month & "/" & _obj.StartDate.Day & "T" & _obj.StartDate.Hour.ToString.PadLeft(2, "0") & ":" & _obj.StartDate.Minute.ToString.PadLeft(2, "0")
                                        UntilDate = selDate.Year & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                    End If
                                    If IsDate(UntilDate) Then
                                        app.End = UntilDate
                                        endDate = UntilDate
                                    End If
                            End Select
                        Next

                        Dim pattern As New RecurrencePattern
                        RecurrencePatternHelper.TryParseRecurrencePattern(_obj.RecurrencePattern, pattern)
                        app.RecurrenceRule = New RecurrenceRule(pattern)
                    End If

                    If app.RecurrenceRule Is Nothing Then
                        Dim result1 As Integer = DateTime.Compare(_obj.StartDate, selDate)
                        If result1 >= 0 Then
                            Dim result2 As Integer = DateTime.Compare(_obj.StartDate, endDate)
                            If result2 < 0 Then
                                workList.Add(_obj)
                            End If
                        End If
                    Else
                        Dim rrrule As List(Of Occurrence) = app.GetOccurrences(selDate, endDate)
                        For Each occurrence As Occurrence In rrrule
                            If occurrence.Start.Date = selDate Then
                                Dim t1 As Date = occurrence.Start.Year & "/" & occurrence.Start.Month.ToString.PadLeft(2, "0") & "/" & occurrence.Start.Day.ToString.PadLeft(2, "0") & _
                                    "T" & _obj.StartDate.Hour.ToString.PadLeft(2, "0") & ":" & _obj.StartDate.Minute.ToString.PadLeft(2, "0")
                                Dim t2 As Date = occurrence.Start.Year & "/" & occurrence.Start.Month.ToString.PadLeft(2, "0") & "/" & occurrence.Start.Day.ToString.PadLeft(2, "0") & _
                                    "T" & _obj.StartDate.Hour.ToString.PadLeft(2, "0") & ":" & (_obj.StartDate.Minute + 1).ToString.PadLeft(2, "0")
                                _obj.StartDate = t1
                                If _obj.EndDate.Year = 1900 Then
                                    _obj.EndDate = t2
                                End If
                                workList.Add(_obj)
                            End If
                        Next
                    End If

                Next

                profileTaskList = Nothing
            End If

            If workList.Count > 0 Then
                Dim idxTask As Integer = workList.Count
                Dim retVal As New List(Of hs_Cron_Profile_Tasks)
                retVal.Add(workList(idxTask - 1))
                Return retVal
            Else
                Return Nothing
            End If
        End Function

        Public Shared Function ListByMonth(CronId As Integer, calYear As Integer, calMonth As Integer) As List(Of hs_Cron_Profile_Tasks)
            Dim startDate As Date = "01/" & calMonth.ToString.PadLeft(2, "0") & "/" & calYear.ToString
            Dim endDate As Date = DateAdd(DateInterval.Month, 1, startDate)
            endDate = DateAdd(DateInterval.Day, -1, endDate)
            endDate = DateAdd(DateInterval.Day, 1, endDate)
            'endDate = DateAdd(DateInterval.Minute, -1, endDate)

            Dim retval As New List(Of hs_Cron_Profile_Tasks)

            For Each Day As DateTime In Enumerable.Range(0, (endDate - startDate).Days).Select(Function(i) startDate.AddDays(i))
                Dim _list As List(Of hs_Cron_Profile_Tasks) = ListCurrent(CronId, Day)
                If Not _list Is Nothing Then
                    For Each _obj As hs_Cron_Profile_Tasks In _list
                        retval.Add(_obj)
                    Next
                End If
            Next Day
            If retval.Count > 0 Then
                Return retval
            Else
                Return Nothing
            End If
        End Function

        Public Shared Function EXListByMonth(CronId As Integer, calYear As Integer, calMonth As Integer) As List(Of hs_Cron_Profile_Tasks)
            If CronId <= 0 Then Return Nothing

            Dim retVal As New List(Of hs_Cron_Profile_Tasks)

            Dim profileTaskList As List(Of hs_Cron_Profile_Tasks) = DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_List(CronId)
            If Not profileTaskList Is Nothing Then
                For Each _obj As hs_Cron_Profile_Tasks In profileTaskList
                    Dim selDate As Date = "01/" & calMonth.ToString.PadLeft(2, "0") & "/" & calYear.ToString
                    Dim endDate As Date = DateAdd(DateInterval.Month, 1, selDate)
                    endDate = DateAdd(DateInterval.Day, -1, endDate)
                    endDate = DateAdd(DateInterval.Day, 1, endDate)
                    endDate = DateAdd(DateInterval.Minute, -1, endDate)
                    Dim app As New Appointment
                    'app.Start = _obj.StartDate
                    'app.End = DateAdd(DateInterval.Minute, 1, _obj.StartDate)
                    'app.Start = selDate

                    app.Start = _obj.StartDate
                    app.End = endDate

                    If Not String.IsNullOrEmpty(_obj.RecurrencePattern) Then
                        Dim ar1() As String = _obj.RecurrencePattern.Split(";")
                        For x As Integer = 0 To ar1.Length - 1
                            Dim RecurrenceMode = Mid(ar1(x), 1, 5)
                            Select Case RecurrenceMode
                                Case "UNTIL"
                                    Dim UNTIL As String = ar1(x).Replace("UNTIL=", String.Empty)
                                    Dim UntilDate As Date
                                    If _obj.yearsRepeatable = False Then
                                        UntilDate = UNTIL.Substring(0, 4) & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                    Else
                                        UntilDate = selDate.Year & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                    End If
                                    If IsDate(UntilDate) Then
                                        app.End = UntilDate
                                        'endDate = UntilDate
                                    End If
                            End Select
                        Next

                        Dim pattern As New RecurrencePattern
                        RecurrencePatternHelper.TryParseRecurrencePattern(_obj.RecurrencePattern, pattern)
                        app.RecurrenceRule = New RecurrenceRule(pattern)
                    End If

                    If app.End >= endDate Then

                        If app.RecurrenceRule Is Nothing Then
                            Dim result1 As Integer = DateTime.Compare(_obj.StartDate, selDate)
                            If result1 >= 0 Then
                                Dim result2 As Integer = DateTime.Compare(_obj.StartDate, endDate)
                                If result2 < 0 Then
                                    retVal.Add(_obj)
                                End If
                            End If
                        Else
                            Dim rrrule As List(Of Occurrence) = app.GetOccurrences(selDate, endDate)
                            For Each occurrence As Occurrence In rrrule
                                If occurrence.Start.Date >= selDate Then
                                    Dim taskObj As New SCP.BLL.hs_Cron_Profile_Tasks
                                    taskObj.CronId = _obj.CronId
                                    taskObj.ProfileNr = _obj.ProfileNr
                                    taskObj.Subject = _obj.Subject
                                    taskObj.StartDate = occurrence.Start.Date
                                    If taskObj.EndDate.Year = 1900 Then
                                        taskObj.EndDate = occurrence.End
                                    Else
                                        taskObj.EndDate = _obj.EndDate
                                    End If
                                    taskObj.RecurrencePattern = _obj.RecurrencePattern
                                    taskObj.ExceptionAppointments = _obj.ExceptionAppointments
                                    retVal.Add(taskObj)
                                End If
                            Next
                        End If

                    End If

                Next

                profileTaskList = Nothing
            End If
            Return retVal
        End Function

        Public Shared Function ListByDates(CronId As Integer, startDate As Date, endDate As Date) As List(Of hs_Cron_Profile_Tasks)
            If CronId <= 0 Then Return Nothing

            Dim selDate As Date = startDate
            Dim retVal As New List(Of hs_Cron_Profile_Tasks)

            Dim profileTaskList As List(Of hs_Cron_Profile_Tasks) = DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_List(CronId)
            If Not profileTaskList Is Nothing Then
                For Each _obj As hs_Cron_Profile_Tasks In profileTaskList
                    Dim ok2Elabor As Boolean = False

                    If _obj.StartDate <= startDate Then
                        Dim app As New Appointment
                        app.Start = _obj.StartDate
                        app.End = _obj.EndDate

                        If Not String.IsNullOrEmpty(_obj.RecurrencePattern) Then
                            Dim ar1() As String = _obj.RecurrencePattern.Split(";")
                            For x As Integer = 0 To ar1.Length - 1
                                Dim RecurrenceMode = Mid(ar1(x), 1, 5)
                                Select Case RecurrenceMode
                                    Case "UNTIL"
                                        Dim UNTIL As String = ar1(x).Replace("UNTIL=", String.Empty)
                                        Dim UntilDate As Date
                                        If _obj.yearsRepeatable = False Then
                                            UntilDate = UNTIL.Substring(0, 4) & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                        Else
                                            UntilDate = selDate.Year & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                        End If
                                        If IsDate(UntilDate) Then
                                            app.End = UntilDate
                                            'endDate = UntilDate
                                        End If
                                End Select
                            Next

                            Dim pattern As New RecurrencePattern
                            RecurrencePatternHelper.TryParseRecurrencePattern(_obj.RecurrencePattern, pattern)
                            app.RecurrenceRule = New RecurrenceRule(pattern)
                        End If

                        'If app.End >= endDate Then

                        'End If 'app.End <= endDate

                        If app.RecurrenceRule Is Nothing Then
                            Dim result1 As Integer = DateTime.Compare(_obj.StartDate, selDate)
                            If result1 >= 0 Then
                                Dim result2 As Integer = DateTime.Compare(_obj.StartDate, endDate)
                                If result2 < 0 Then
                                    retVal.Add(_obj)
                                End If
                            End If
                        Else
                            Dim rrrule As List(Of Occurrence) = app.GetOccurrences(selDate, app.End)
                            For Each occurrence As Occurrence In rrrule
                                If occurrence.Start.Date >= selDate Then
                                    If occurrence.Start.Date <= endDate Then
                                        Dim taskObj As New SCP.BLL.hs_Cron_Profile_Tasks
                                        taskObj.CronId = _obj.CronId
                                        taskObj.ProfileNr = _obj.ProfileNr
                                        taskObj.Subject = _obj.Subject

                                        Dim occurrenceStart As Date = occurrence.Start.Year & "/" & occurrence.Start.Month.ToString.PadLeft(2, "0") & "/" & occurrence.Start.Day.ToString.PadLeft(2, "0") & "T00:00"
                                        Dim occurrenceEnd As Date = occurrence.Start.Year & "/" & occurrence.Start.Month.ToString.PadLeft(2, "0") & "/" & occurrence.Start.Day.ToString.PadLeft(2, "0") & "T23:59"
                                        taskObj.StartDate = occurrenceStart
                                        taskObj.EndDate = occurrenceEnd
                                        taskObj.RecurrencePattern = _obj.RecurrencePattern
                                        taskObj.ExceptionAppointments = _obj.ExceptionAppointments
                                        taskObj.CronCod = _obj.CronId
                                        taskObj.CronDescr = _obj.CronDescr
                                        retVal.Add(taskObj)
                                    End If
                                End If
                            Next

                        End If



                    End If '_obj.StartDate >= startDate

                Next

                profileTaskList = Nothing
            End If

            Return retVal

        End Function

        Public Shared Function Read(TaskId As Integer) As hs_Cron_Profile_Tasks
            If TaskId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_Read(TaskId)
        End Function

        Public Shared Function Update(TaskId As Integer, _
                                      ProfileNr As Integer, _
                                      Subject As String, _
                                      StartDate As String, _
                                      EndDate As String, _
                                      chkMonday As Boolean, _
                                      chkTuesday As Boolean, _
                                      chkWednesday As Boolean, _
                                      chkThursday As Boolean, _
                                      chkFriday As Boolean, _
                                      chkSaturday As Boolean, _
                                      chkSunday As Boolean, _
                                      yearsRepeatable As Boolean) As Boolean
            If TaskId <= 0 Then Return Nothing
            If ProfileNr < 0 Then Return Nothing
            If String.IsNullOrEmpty(Subject) Then Return Nothing

            Dim m_RecurrencePattern As String = String.Empty
            Dim m_RecurrenceByDay As String = String.Empty
            Dim m_RecurrenceCount As String = String.Empty
            Dim m_RecurrenceUntil As String = String.Empty

            Dim m_taskDateTime As New DateTime
            Dim m_endTask As Date = FormatDateTime("01/01/1900", DateFormat.LongTime)
            Dim m_EndBy As New DateTime
            Dim m_Mon As Boolean = False
            Dim m_Tue As Boolean = False
            Dim m_Wed As Boolean = False
            Dim m_Thu As Boolean = False
            Dim m_Fri As Boolean = False
            Dim m_Sat As Boolean = False
            Dim m_Sun As Boolean = False
            Dim m_subject As String = String.Empty


            Dim txtStartDate As String = StartDate
            Dim txtEndDate As String = EndDate

            m_subject = Subject

            m_taskDateTime = DateTime.Parse(txtStartDate)
            m_endTask = DateAdd(DateInterval.Minute, 1, m_taskDateTime)

            m_Mon = chkMonday
            m_Tue = chkTuesday
            m_Wed = chkWednesday
            m_Thu = chkThursday
            m_Fri = chkFriday
            m_Sat = chkSaturday
            m_Sun = chkSunday

            If m_Mon = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "MO,"
            End If
            If m_Tue = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "TU,"
            End If
            If m_Wed = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "WE,"
            End If
            If m_Thu = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "TH,"
            End If
            If m_Fri = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "FR,"
            End If
            If m_Sat = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "SA,"
            End If
            If m_Sun = True Then
                If String.IsNullOrEmpty(m_RecurrenceByDay) Then
                    m_RecurrenceByDay = "FREQ=WEEKLY;BYDAY="
                End If
                m_RecurrenceByDay += "SU,"
            End If
            If Not String.IsNullOrEmpty(m_RecurrenceByDay) Then
                m_RecurrenceByDay = m_RecurrenceByDay.Remove(m_RecurrenceByDay.Length - 1)
            End If

            If IsDate(txtEndDate) Then
                'm_EndBy = DateTime.Parse(txtEndDate)

                m_EndBy = DateTime.Parse(txtEndDate).Year & "/" & _
                    DateTime.Parse(txtEndDate).Month.ToString.PadLeft(2, "0") & "/" & _
                    DateTime.Parse(txtEndDate).Day.ToString.PadLeft(2, "0") & _
                    "T23:59"
                EndDate = m_EndBy
                m_RecurrenceUntil = "UNTIL=" & m_EndBy.Year.ToString & _
                               m_EndBy.Month.ToString.PadLeft(2, "0") & _
                               m_EndBy.Day.ToString.PadLeft(2, "0") & _
                              "T070000Z;"
                m_RecurrencePattern += m_RecurrenceUntil
            Else
                EndDate = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            End If
            If Not String.IsNullOrEmpty(m_RecurrenceByDay) Then
                m_RecurrencePattern += m_RecurrenceByDay
            End If

            'Dim oldTask As hs_Cron_Profile_Tasks = SCP.BLL.hs_Cron_Profile_Tasks.Read(TaskId)
            'If Not oldTask Is Nothing Then

            '    oldTask = Nothing
            'End If

            Return DataAccessHelper.GetDataAccess.hs_Cron_Profile_Tasks_Update(TaskId, ProfileNr, Subject, StartDate, EndDate, m_RecurrencePattern, String.Empty, yearsRepeatable)
        End Function
#End Region

#Region "public properties"
        Public Property TaskId As Integer
        Public Property CronId As Integer
        Public Property ProfileNr As Integer
        Public Property Subject As String
        Public Property StartDate As Date
        Public Property EndDate As Date
        Public Property RecurrencePattern As String
        Public Property ExceptionAppointments As String
        Public Property yearsRepeatable As Boolean
        Public Property CronCod As String
        Public Property CronDescr As String
#End Region

    End Class
End Namespace
