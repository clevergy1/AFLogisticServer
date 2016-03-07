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
    Public Class hs_ControllerDetail
        Private _id As Integer
        Private _ControllerId As Integer
        Private _Descr As String
        Private _NoteInterne As String
        Private _qta As Integer

#Region "constructor"
        Public Sub New(id As Integer, ControllerId As Integer, Descr As String, NoteInterne As String, qta As Integer)
            Me.id = id
            Me._ControllerId = ControllerId
            Me._Descr = Descr
            Me._NoteInterne = NoteInterne
            Me._qta = qta
        End Sub
        Public Sub New()
            Me.New(0, 0, String.Empty, String.Empty, 0)
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(ControllerId As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
            If ControllerId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ControllerDetail_Add(ControllerId, Descr, NoteInterne, qta, UserName)
        End Function

        Public Shared Function Del(id As Integer) As Boolean
            If id <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ControllerDetail_Del(id)
        End Function

        Public Shared Function List(ControllerId As Integer) As List(Of hs_ControllerDetail)
            If ControllerId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ControllerDetail_List(ControllerId)
        End Function

        Public Shared Function Read(id As Integer) As hs_ControllerDetail
            If id <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ControllerDetail_Read(id)
        End Function

        Public Shared Function Update(id As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
            If id <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ControllerDetail_Update(id, Descr, NoteInterne, qta, UserName)
        End Function
#End Region

#Region "public properties"
        Public Property id As Integer
            Get
                Return Me._id
            End Get
            Set(value As Integer)
                Me._id = value
            End Set
        End Property
        Public Property ControllerId As Integer
            Get
                Return Me._ControllerId
            End Get
            Set(value As Integer)
                Me._ControllerId = value
            End Set
        End Property
        Public Property Descr As String
            Get
                Return Me._Descr
            End Get
            Set(value As String)
                Me._Descr = value
            End Set
        End Property
        Public Property NoteInterne As String
            Get
                Return Me._NoteInterne
            End Get
            Set(value As String)
                Me._NoteInterne = value
            End Set
        End Property
        Public Property qta As Integer
            Get
                Return Me._qta
            End Get
            Set(value As Integer)
                Me._qta = value
            End Set
        End Property
#End Region
    End Class
End Namespace
