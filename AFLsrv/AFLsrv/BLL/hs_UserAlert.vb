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
    Public Class hs_UserAlert
        Private _Id As Integer
        Private _hsId As Integer
        Private _IdContatto As Integer

        Private _emailaddress As String

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, String.Empty)
        End Sub
        Public Sub New(Id As Integer, hsId As Integer, IdContatto As Integer, emailaddress As String)
            Me._Id = Id
            Me._hsId = hsId
            _IdContatto = IdContatto
            _emailaddress = emailaddress
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, IdContatto As Integer) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If IdContatto <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_UserAlert_Add(hsId, IdContatto)
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_UserAlert_Del(Id)
        End Function

        Public Shared Function List(hsId As Integer) As List(Of hs_UserAlert)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_UserAlert_List(hsId)
        End Function

        Public Shared Function Read(Id As Integer) As hs_UserAlert
            If Id <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_UserAlert_Read(Id)
        End Function
#End Region

#Region "public properties"
        Public Property Id As Integer
            Get
                Return Me._Id
            End Get
            Set(value As Integer)
                Me._Id = value
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
        Public Property IdContatto As Integer
            Get
                Return _IdContatto
            End Get
            Set(value As Integer)
                _IdContatto = value
            End Set
        End Property
        Public Property emailaddress As String
            Get
                Return _emailaddress
            End Get
            Set(value As String)
                _emailaddress = value
            End Set
        End Property
#End Region
    End Class
End Namespace
