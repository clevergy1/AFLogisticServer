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
    Public Class aspnetusersstat
        Private _totUser As Integer
        Private _Approved As Integer
        Private _notApproved As Integer
        Private _Locked As Integer
        Private _notLocked As Integer

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, 0, 0)
        End Sub
        Public Sub New(totUser As Integer, Approved As Integer, notApproved As Integer, Locked As Integer, notLocked As Integer)
            _totUser = totUser
            _Approved = Approved
            _notApproved = notApproved
            _Locked = Locked
            _notLocked = notLocked
        End Sub
#End Region

#Region "methods"
        Public Shared Function Read() As aspnetusersstat
            Return DataAccessHelper.GetDataAccess.aspnetusersstat_Read
        End Function
#End Region

#Region "properties"
        Public Property totUser As Integer
            Get
                Return Me._totUser
            End Get
            Set(value As Integer)
                Me._totUser = value
            End Set
        End Property
        Public Property Approved As Integer
            Get
                Return Me._Approved
            End Get
            Set(value As Integer)
                Me._Approved = value
            End Set
        End Property
        Public Property notApproved As Integer
            Get
                Return Me._notApproved
            End Get
            Set(value As Integer)
                Me._notApproved = value
            End Set
        End Property
        Public Property Locked As Integer
            Get
                Return Me._Locked
            End Get
            Set(value As Integer)
                Me._Locked = value
            End Set
        End Property
        Public Property notLocked As Integer
            Get
                Return Me._notLocked
            End Get
            Set(value As Integer)
                Me._notLocked = value
            End Set
        End Property
#End Region
    End Class
End Namespace
