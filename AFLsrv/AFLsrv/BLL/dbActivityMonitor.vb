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
    Public Class dbActivityMonitor
        Private _SessionId As Integer
        Private _App As String
        Private _WaitTime_ms As Integer
        Private _TotalCPU_ms As Integer
        Private _TotalPhyIO_mb As Integer
        Private _MemUsage_kb As Integer
        Private _OpenTrans As Integer
        Private _LoginTime As Date
        Private _LastReqStartTime As Date
        Private _HostName As String
        Private _NetworkAddr As String

#Region "constructor"
        Public Sub New()
            Me.New(0, String.Empty, 0, 0, 0, 0, 0, FormatDateTime("01/01/1900", DateFormat.GeneralDate), FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, String.Empty)
        End Sub
        Public Sub New(SessionId As Integer, _
                       App As String, _
                       WaitTime_ms As Integer, _
                       TotalCPU_ms As Integer, _
                       TotalPhyIO_mb As Integer, _
                       MemUsage_kb As Integer, _
                       OpenTrans As Integer, _
                       LoginTime As Date, _
                       LastReqStartTime As Date, _
                       HostName As String, _
                       NetworkAddr As String)
            Me._SessionId = SessionId
            Me._App = App
            Me._WaitTime_ms = WaitTime_ms
            Me._TotalCPU_ms = TotalCPU_ms
            Me._TotalPhyIO_mb = TotalPhyIO_mb
            Me._MemUsage_kb = MemUsage_kb
            Me._OpenTrans = OpenTrans
            Me._LoginTime = LoginTime
            Me._LastReqStartTime = LastReqStartTime
            Me._HostName = HostName
            Me._NetworkAddr = NetworkAddr
        End Sub
#End Region

#Region "methods"
        Public Shared Function List() As List(Of dbActivityMonitor)
            Return DataAccessHelper.GetDataAccess.dbActivityMonitor_List
        End Function
#End Region

#Region "properties"
        Public Property SessionId As Integer
            Get
                Return Me._SessionId
            End Get
            Set(value As Integer)
                Me._SessionId = value
            End Set
        End Property
        Public Property App As String
            Get
                Return Me._App
            End Get
            Set(value As String)
                Me._App = value
            End Set
        End Property
        Public Property WaitTime_ms As Integer
            Get
                Return Me._WaitTime_ms
            End Get
            Set(value As Integer)
                Me._WaitTime_ms = value
            End Set
        End Property
        Public Property TotalCPU_ms As Integer
            Get
                Return Me._TotalCPU_ms
            End Get
            Set(value As Integer)
                Me._TotalCPU_ms = value
            End Set
        End Property
        Public Property TotalPhyIO_mb As Integer
            Get
                Return Me._TotalPhyIO_mb
            End Get
            Set(value As Integer)
                Me._TotalPhyIO_mb = value
            End Set
        End Property
        Public Property MemUsage_kb As Integer
            Get
                Return Me._MemUsage_kb
            End Get
            Set(value As Integer)
                Me._MemUsage_kb = value
            End Set
        End Property
        Public Property OpenTrans As Integer
            Get
                Return Me._OpenTrans
            End Get
            Set(value As Integer)
                Me._OpenTrans = value
            End Set
        End Property
        Public Property LoginTime As Date
            Get
                Return Me._LoginTime
            End Get
            Set(value As Date)
                Me._LoginTime = value
            End Set
        End Property
        Public Property LastReqStartTime As Date
            Get
                Return Me._LastReqStartTime
            End Get
            Set(value As Date)
                Me._LastReqStartTime = value
            End Set
        End Property
        Public Property HostName As String
            Get
                Return Me._HostName
            End Get
            Set(value As String)
                Me._HostName = value
            End Set
        End Property
        Public Property NetworkAddr As String
            Get
                Return Me._NetworkAddr
            End Get
            Set(value As String)
                Me._NetworkAddr = value
            End Set
        End Property
#End Region
    End Class
End Namespace
