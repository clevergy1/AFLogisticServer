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
    Public Class hs_Controller
        Private _ControllerId As Integer
        Private _hsId As Integer
        Private _ControllerDescr As String
        Private _NoteInterne As String

#Region "constructor"
        Public Sub New(ControllerId As Integer, hsId As Integer, ControllerDescr As String, NoteInterne As String)
            Me._ControllerId = ControllerId
            Me._hsId = hsId
            Me._ControllerDescr = ControllerDescr
            Me._NoteInterne = NoteInterne
        End Sub
        Public Sub New()
            Me.New(0, 0, String.Empty, String.Empty)
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Controller_Add(hsId, ControllerDescr, NoteInterne, UserName)
        End Function

        Public Shared Function Del(ControllerId As Integer) As Boolean
            If ControllerId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Controller_Del(ControllerId)
        End Function

        Public Shared Function List(hsId As Integer) As List(Of hs_Controller)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Controller_List(hsId)
        End Function

        Public Shared Function Read(ControllerId As Integer) As hs_Controller
            If ControllerId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Controller_Read(ControllerId)
        End Function

        Public Shared Function Update(ControllerId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
            If ControllerId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Controller_Update(ControllerId, ControllerDescr, NoteInterne, UserName)
        End Function
#End Region

#Region "public properties"
        Public Property ControllerId As Integer
            Get
                Return Me._ControllerId
            End Get
            Set(value As Integer)
                Me._ControllerId = value
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
        Public Property ControllerDescr As String
            Get
                Return Me._ControllerDescr
            End Get
            Set(value As String)
                Me._ControllerDescr = value
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
