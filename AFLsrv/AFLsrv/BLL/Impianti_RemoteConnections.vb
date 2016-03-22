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
    Public Class Impianti_RemoteConnections
        Private _IdImpianto As String
        Private _IdAddress As Integer
        Private _Descr As String
        Private _remoteAddress As String
        Private _remotePort As Integer
        Private _connectionType As Integer
        Private _DconnectionType As String
        Private _NoteInterne As String

#Region "constructor"
        Public Sub New()
            Me.New(String.Empty, 0, String.Empty, String.Empty, 0, 0, String.Empty, String.Empty)
        End Sub
        Public Sub New(IdImpianto As String, IdAddress As Integer, Descr As String, remoteaddress As String, remotePort As Integer, connectionType As Integer, DconnectionType As String, NoteInterne As String)
            Me._IdImpianto = IdImpianto
            Me._IdAddress = IdAddress
            Me._Descr = Descr
            Me._remoteAddress = remoteaddress
            _remotePort = remotePort
            Me._connectionType = connectionType
            Me._DconnectionType = DconnectionType
            Me._NoteInterne = NoteInterne
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(IdImpianto As String, IdAddress As Integer, Descr As String, remoteaddress As String, connectionType As Integer, NoteInterne As String) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            If IdAddress <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Descr) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_RemoteConnections_Add(IdImpianto, IdAddress, Descr, remoteaddress, connectionType, NoteInterne)
        End Function

        Public Shared Function Del(IdImpianto As String, IdAddress As Integer) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            If IdAddress <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_RemoteConnections_Del(IdImpianto, IdAddress)
        End Function

        Public Shared Function getLastId() As Integer
            Return DataAccessHelper.GetDataAccess.Impianti_RemoteConnections_getLastId()
        End Function

        Public Shared Function List(IdImpianto As String) As List(Of Impianti_RemoteConnections)
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_RemoteConnections_List(IdImpianto)
        End Function

        Public Shared Function Read(IdImpianto As String, IdAddress As Integer) As Impianti_RemoteConnections
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            If IdAddress <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_RemoteConnections_Read(IdImpianto, IdAddress)
        End Function

        Public Shared Function Upd(IdImpianto As String, IdAddress As Integer, Descr As String, remoteaddress As String, connectionType As Integer, NoteInterne As String) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            If IdAddress <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_RemoteConnections_Upd(IdImpianto, IdAddress, Descr, remoteaddress, connectionType, NoteInterne)
        End Function

#End Region

#Region "public properties"
        Public Property IdImpianto As String
            Get
                Return Me._IdImpianto
            End Get
            Set(value As String)
                Me._IdImpianto = value
            End Set
        End Property
        Public Property IdAddress As Integer
            Get
                Return Me._IdAddress
            End Get
            Set(value As Integer)
                Me._IdAddress = value
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
        Public Property remoteAddress As String
            Get
                Return Me._remoteAddress
            End Get
            Set(value As String)
                Me._remoteAddress = value
            End Set
        End Property
        Public Property remotePort As Integer
            Get
                Return _remotePort
            End Get
            Set(value As Integer)
                _remotePort = value
            End Set
        End Property
        Public Property connectionType As Integer
            Get
                Return Me._connectionType
            End Get
            Set(value As Integer)
                Me._connectionType = value
            End Set
        End Property
        Public Property DconnectionType As String
            Get
                Return Me._DconnectionType
            End Get
            Set(value As String)
                Me._DconnectionType = value
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
#End Region
    End Class
End Namespace
