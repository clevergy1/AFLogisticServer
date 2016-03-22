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
    Public Class log_hs_Cron
        Private _LogId As Integer
        Private _dtLog As Date
        Private _hsId As Integer
        Private _SetPoint As Decimal
        Private _CronCod As String
        Private _CronDescr As String
        Private _stato As Integer

#Region "constructor"
        Public Sub New()
            Me.New(0, FormatDateTime("01/01/1900"), 0, 0, String.Empty, String.Empty, 0)
        End Sub
        Public Sub New(LogId As Integer, dtLog As Date, hsId As Integer, SetPoint As Decimal, CronCod As String, CronDescr As String, stato As Integer)
            _LogId = LogId
            _dtLog = dtLog
            _hsId = hsId
            _SetPoint = SetPoint
            _CronCod = CronCod
            _CronDescr = CronDescr
            _stato = stato
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, SetPoint As Decimal, CronCod As String, CronDescr As String, stato As Integer, dtLog As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_Add(hsId, SetPoint, CronCod, CronDescr, stato, dtLog)
        End Function

        Public Shared Function List(hsId As Integer, CronCod As String, fromDate As Date, toDate As Date) As List(Of log_hs_Cron)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_List(hsId, CronCod, fromDate, toDate)
        End Function

        Public Shared Function ListPaged(hsId As Integer, CronCod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_hs_Cron)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_ListPaged(hsId, CronCod, fromDate, toDate, rowNumber)
        End Function

        Public Shared Function ListAll(hsId As Integer, fromDate As Date, toDate As Date) As List(Of log_hs_Cron)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_ListAll(hsId, fromDate, toDate)
        End Function

        Public Shared Function logNotSent(hsId As Integer, CronCod As String) As log_hs_Cron
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_logNotSent(hsId, CronCod)
        End Function

        Public Shared Function setIsSent(Logid As Integer) As Boolean
            If Logid <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_setIsSent(Logid)
        End Function


        Public Shared Function ReadLast(hsId As Integer, CronCod As String) As log_hs_Cron
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(CronCod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_hs_Cron_ReadLast(hsId, CronCod)
        End Function
#End Region

#Region "public properties"
        Public Property LogId As Integer
            Get
                Return Me._LogId
            End Get
            Set(value As Integer)
                Me._LogId = value
            End Set
        End Property
        Public Property dtLog As Date
            Get
                Return Me._dtLog
            End Get
            Set(value As Date)
                Me._dtLog = value
            End Set
        End Property
        Public Property hsId As Integer
            Get
                Return Me._hsId
            End Get
            Set(value As Integer)
                Me._hsId = value
            End Set
        End Property
        Public Property SetPoint As Decimal
            Get
                Return Me._SetPoint
            End Get
            Set(value As Decimal)
                Me._SetPoint >>= value
            End Set
        End Property
        Public Property CronCod As String
            Get
                Return Me._CronCod
            End Get
            Set(value As String)
                Me._CronCod = value
            End Set
        End Property
        Public Property CronDescr As String
            Get
                Return Me._CronDescr
            End Get
            Set(value As String)
                Me._CronDescr = value
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
#End Region
    End Class
End Namespace
