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
    Public Class dbstats
        Private _cpu_perc As Decimal
        Private _avg_io As Decimal
        Private _batch_request As Integer

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0)
        End Sub
        Public Sub New(cpu_perc As Decimal, avg_io As Decimal, batch_request As Integer)
            Me._cpu_perc = cpu_perc
            Me._avg_io = avg_io
            Me._batch_request = batch_request
        End Sub
#End Region

#Region "methods"
        Public Shared Function Read() As dbstats
            Return DataAccessHelper.GetDataAccess.dbstats_Read
        End Function
#End Region

#Region "properties"
        Public Property cpu_perc As Decimal
            Get
                Return Me._cpu_perc
            End Get
            Set(value As Decimal)
                Me._cpu_perc = value
            End Set
        End Property
        Public Property avg_io As Decimal
            Get
                Return Me._avg_io
            End Get
            Set(value As Decimal)
                Me._avg_io = value
            End Set
        End Property
        Public Property batch_request As Integer
            Get
                Return Me._batch_request
            End Get
            Set(value As Integer)
                Me._batch_request = value
            End Set
        End Property
#End Region
    End Class
End Namespace
