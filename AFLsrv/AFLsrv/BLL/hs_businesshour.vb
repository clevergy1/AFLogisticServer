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

Imports Telerik.Windows.Controls.Scheduler
Imports Telerik.Windows.Controls.Scheduler.ICalendar

Namespace SCP.BLL
    Public Class hs_businesshour
#Region "constructor"
        Public Sub New()
            Me.New(0, 0, False, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, String.Empty)
        End Sub
        Public Sub New(mId As Integer, mhsId As Integer, misClosedTime As Boolean, mSubject As String, mStartDate As Date, mEndDate As Date, mRecurrencePattern As String, mExceptionAppointments As String)
            Id = mId
            hsId = mhsId
            isClosedTime = misClosedTime
            Subject = mSubject
            StartDate = mStartDate
            EndDate = mEndDate
            RecurrencePattern = mRecurrencePattern
            ExceptionAppointments = mExceptionAppointments
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, _
                                   isClosedTime As Boolean, _
                                   Subject As String, _
                                   StartDate As String, _
                                   EndDate As String, _
                                   chkMonday As Boolean, _
                                   chkTuesday As Boolean, _
                                   chkWednesday As Boolean, _
                                   chkThursday As Boolean, _
                                   chkFriday As Boolean, _
                                   chkSaturday As Boolean, _
                                   chkSunday As Boolean) As Boolean
            If hsId <= 0 Then Return False
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

            Return DataAccessHelper.GetDataAccess.hs_businesshour_Add(hsId, isClosedTime, Subject, StartDate, EndDate, m_RecurrencePattern, String.Empty)
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_businesshour_Del(Id)
        End Function

        Public Shared Function List(hsId As Integer) As List(Of hs_businesshour)
            If hsId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_businesshour_List(hsId)
        End Function

        Public Shared Function ListCurrent(hsId As Integer, selDate As Date) As List(Of hs_businesshour)
            If hsId <= 0 Then Return Nothing
            Dim workList As New List(Of hs_businesshour)

            Dim TaskList As List(Of hs_businesshour) = DataAccessHelper.GetDataAccess.hs_businesshour_List(hsId)
            If Not TaskList Is Nothing Then

                For Each _obj As hs_businesshour In TaskList
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
                                    UntilDate = UNTIL.Substring(0, 4) & "/" & UNTIL.Substring(4, 2) & "/" & UNTIL.Substring(6, 2) & "T" & UNTIL.Substring(9, 2).PadLeft(2, "0") & ":" & UNTIL.Substring(11, 2).PadLeft(2, "0")
                                    If IsDate(UntilDate) Then
                                        app.End = UntilDate
                                        endDate = UntilDate
                                    End If
                            End Select
                        Next

                        Dim pattern As New RecurrencePattern
                        RecurrencePatternHelper.TryParseRecurrencePattern(_obj.RecurrencePattern, pattern)
                        app.RecurrenceRule = New RecurrenceRule(pattern)
                    End If '_obj.RecurrencePattern

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
                    End If 'app.RecurrenceRule

                Next ' for each

                TaskList = Nothing
            End If 'Tasklist

            'Dim idxTask As Integer = workList.Count
            'Dim retVal As New List(Of hs_businesshour)
            'retVal.Add(workList(idxTask - 1))
            'Return retVal
            Return workList
        End Function

        Public Shared Function ListByMonth(hsId As Integer, calYear As Integer, calMonth As Integer) As List(Of hs_businesshour)
            Dim startDate As Date = "01/" & calMonth.ToString.PadLeft(2, "0") & "/" & calYear.ToString
            Dim endDate As Date = DateAdd(DateInterval.Month, 1, startDate)
            endDate = DateAdd(DateInterval.Day, -1, endDate)
            endDate = DateAdd(DateInterval.Day, 1, endDate)
            'endDate = DateAdd(DateInterval.Minute, -1, endDate)

            Dim retval As New List(Of hs_businesshour)

            For Each Day As DateTime In Enumerable.Range(0, (endDate - startDate).Days).Select(Function(i) startDate.AddDays(i))
                Dim _list As List(Of hs_businesshour) = ListCurrent(hsId, Day)
                If Not _list Is Nothing Then
                    For Each _obj As hs_businesshour In _list
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

        Public Shared Function Read(Id As Integer) As hs_businesshour
            If Id <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_businesshour_Read(Id)
        End Function

        Public Shared Function Update(Id As Integer, _
                                      isClosedTime As Boolean, _
                                      Subject As String, _
                                      StartDate As String, _
                                      EndDate As String, _
                                      chkMonday As Boolean, _
                                      chkTuesday As Boolean, _
                                      chkWednesday As Boolean, _
                                      chkThursday As Boolean, _
                                      chkFriday As Boolean, _
                                      chkSaturday As Boolean, _
                                      chkSunday As Boolean) As Boolean
            If Id <= 0 Then Return Nothing
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
            Return DataAccessHelper.GetDataAccess.hs_businesshour_Update(Id, isClosedTime, Subject, StartDate, EndDate, m_RecurrencePattern, String.Empty)
        End Function

#End Region

#Region "public properties"
        Public Property Id As Integer
        Public Property hsId As Integer
        Public Property isClosedTime As Boolean
        Public Property Subject As String
        Public Property StartDate As Date
        Public Property EndDate As Date
        Public Property RecurrencePattern As String
        Public Property ExceptionAppointments As String
#End Region

    End Class
End Namespace
