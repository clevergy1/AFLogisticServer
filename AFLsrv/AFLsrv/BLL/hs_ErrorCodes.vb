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
    Public Class hs_ErrorCodes
        Private _elementCode As String
        Private _errorCode As Integer
        Private _errorLevel As Integer
        Private _DescIT As String
        Private _DescEN As String

#Region "constructor"
        Public Sub New()
            Me.New(String.Empty, 0, 0, String.Empty, String.Empty)
        End Sub
        Public Sub New(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String)
            _elementCode = elementCode
            _errorCode = errorCode
            _errorLevel = errorLevel
            _DescIT = DescIT
            _DescEN = DescEN
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
            If String.IsNullOrEmpty(elementCode) Then
                Return False
            End If
            If errorCode <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(DescIT) Or String.IsNullOrEmpty(DescEN) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorCodes_Add(elementCode, errorCode, errorLevel, DescIT, DescEN)
        End Function

        Public Shared Function Del(elementCode As String, errorCode As Integer) As Boolean
            If String.IsNullOrEmpty(elementCode) Then
                Return False
            End If
            If errorCode <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorCodes_Del(elementCode, errorCode)
        End Function

        Public Shared Function getTot() As Integer
            Return DataAccessHelper.GetDataAccess.hs_ErrorCodes_getTot
        End Function

        Public Shared Function List(elementCode As String) As List(Of hs_ErrorCodes)
            If String.IsNullOrEmpty(elementCode) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorCodes_List(elementCode)
        End Function

        Public Shared Function Read(elementCode As String, errorCode As Integer) As hs_ErrorCodes
            If String.IsNullOrEmpty(elementCode) Then
                Return Nothing
            End If
            If errorCode <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorCodes_Read(elementCode, errorCode)
        End Function

        Public Shared Function Update(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
            If String.IsNullOrEmpty(elementCode) Then
                Return False
            End If
            If errorCode <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(DescIT) Or String.IsNullOrEmpty(DescEN) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorCodes_Update(elementCode, errorCode, errorLevel, DescIT, DescEN)
        End Function
#End Region

#Region "public properties"
        Public Property elementCode As String
            Get
                Return Me._elementCode
            End Get
            Set(value As String)
                Me._elementCode = value
            End Set
        End Property
        Public Property errorCode As Integer
            Get
                Return Me._errorCode
            End Get
            Set(value As Integer)
                Me._errorCode = value
            End Set
        End Property
        Public Property errorLevel As Integer
            Get
                Return _errorLevel
            End Get
            Set(value As Integer)
                _errorLevel = value
            End Set
        End Property
        Public Property DescIT As String
            Get
                Return Me._DescIT
            End Get
            Set(value As String)
                Me._DescIT = value
            End Set
        End Property
        Public Property DescEN As String
            Get
                Return Me._DescEN
            End Get
            Set(value As String)
                Me._DescEN = value
            End Set
        End Property
#End Region
    End Class
End Namespace
