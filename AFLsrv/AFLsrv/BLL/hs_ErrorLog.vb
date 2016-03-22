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
    Public Class hs_ErrorLog
        Private _Id As Integer
        Private _LogDate As Date
        Private _hsId As Integer
        Private _hselement As String
        Private _elementCode As String
        Private _errorCode As Integer
        Private _errorValue As String
        Private _errorLevel As Integer
        Private _DescIT As String
        Private _DescEN As String

#Region "constructor"
        Public Sub New()
            Me.New(0, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0, String.Empty, 0, 0, String.Empty, 0, String.Empty, String.Empty)
        End Sub
        Public Sub New(Id As Integer, _
                       LogDate As Date, _
                       hsId As Integer, _
                       hselement As String, _
                       elementCode As String, _
                       errorCode As Integer, _
                       errorValue As String, _
                       errorLevel As Integer, _
                       DescIT As String, _
                       DescEN As String)
            _Id = Id
            _LogDate = LogDate
            _hsId = hsId
            _hselement = hselement
            _elementCode = elementCode
            _errorCode = errorCode
            _errorValue = errorValue
            _errorLevel = errorLevel
            _DescIT = DescIT
            _DescEN = DescEN
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(LogDate As Date, hsId As Integer, hselement As String, elementCode As String, errorCode As Integer, errorValue As String) As Boolean
            If Not IsDate(LogDate) Then
                Return False
            End If
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(hselement) Then
                Return False
            End If
            If String.IsNullOrEmpty(elementCode) Then
                Return False
            End If
            If errorCode <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorLog_Add(LogDate, hsId, hselement, elementCode, errorCode, errorValue)
        End Function

        Public Shared Function List(hsId As Integer, fromDate As Date, toDate As Date) As List(Of hs_ErrorLog)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorLog_List(hsId, fromDate, toDate)
        End Function

        Public Shared Function ListAll(hsId As Integer, rowNumber As Integer) As List(Of hs_ErrorLog)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorLog_ListAll(hsId, rowNumber)
        End Function

        Public Shared Function ListByElement(hsId As Integer, hselement As String, rowNumber As Integer) As List(Of hs_ErrorLog)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(hselement) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_ErrorLog_ListByElement(hsId, hselement, rowNumber)
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
        Public Property LogDate As Date
            Get
                Return Me._LogDate
            End Get
            Set(value As Date)
                Me._LogDate = value
            End Set
        End Property
        Public Property hselement As String
            Get
                Return Me._hselement
            End Get
            Set(value As String)
                Me._hselement = value
            End Set
        End Property
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
        Public Property errorValue As String
            Get
                Return Me._errorValue
            End Get
            Set(value As String)
                Me._errorValue = value
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
