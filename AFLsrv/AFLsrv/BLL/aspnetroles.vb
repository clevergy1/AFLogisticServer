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
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports SCP.DAL

Namespace SCP.BLL
    Public Class aspnetroles
        Private _ApplicationId As String
        Private _RoleId As String
        Private _RoleName As String
        Private _Description As String
        Private _countUser As Integer

#Region "constructor"
        Public Sub New()
            Me.New(String.Empty, String.Empty, String.Empty, String.Empty, 0)
        End Sub
        Public Sub New(ApplicationId As String, RoleId As String, RoleName As String, Description As String, countUser As Integer)
            Me._ApplicationId = ApplicationId
            Me._RoleId = RoleId
            Me._RoleName = RoleName
            Me._Description = Description
            Me._countUser = countUser
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(RoleName As String) As Boolean
            If String.IsNullOrEmpty(RoleName) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.aspnetroles_Add(RoleName)
        End Function

        Public Shared Function Del(RoleName As String) As Boolean
            If String.IsNullOrEmpty(RoleName) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.aspnetroles_Del(RoleName)
        End Function

        Public Shared Function List() As List(Of aspnetroles)
            Return DataAccessHelper.GetDataAccess.aspnetroles_List
        End Function

        Public Shared Function ListActive(IdImpianto As String) As List(Of aspnetroles)
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.aspnetroles_ListActive(IdImpianto)
        End Function

        Public Shared Function Update(RoleId As String, RoleName As String) As Boolean
            If String.IsNullOrEmpty(RoleId) Then
                Return False
            End If
            If String.IsNullOrEmpty(RoleName) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.aspnetroles_Update(RoleId, RoleName)
        End Function
#End Region

#Region "properties"
        Public Property ApplicationId As String
            Get
                Return Me._ApplicationId
            End Get
            Set(value As String)
                Me._ApplicationId = value
            End Set
        End Property
        Public Property RoleId As String
            Get
                Return Me._RoleId
            End Get
            Set(value As String)
                Me._RoleId = value
            End Set
        End Property
        Public Property RoleName As String
            Get
                Return Me._RoleName
            End Get
            Set(value As String)
                Me._RoleName = value
            End Set
        End Property
        Public Property Description As String
            Get
                Return Me._Description
            End Get
            Set(value As String)
                Me._Description = value
            End Set
        End Property
        Public Property countUser As Integer
            Get
                Return Me._countUser
            End Get
            Set(value As Integer)
                Me._countUser = value
            End Set
        End Property
#End Region
    End Class
End Namespace
