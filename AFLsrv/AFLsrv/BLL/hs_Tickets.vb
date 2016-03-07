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
    Public Class hs_Tickets
        Private _TicketId As Integer
        Private _hsId As Integer
        Private _TicketTitle As String
        Private _Requester As String
        Private _emailRequester As String
        Private _DateRequest As Date
        Private _Description As String
        Private _Executor As String
        Private _emailExecutor As String
        Private _DateExecution As Date
        Private _ExecutorComment As String
        Private _TicketStatus As Integer
        Private _TicketType As Integer

#Region "constructor"
        Public Sub New(TicketId As Integer, _
                       hsId As Integer, _
                       TicketTitle As String, _
                       Requester As String, _
                       emailRequester As String, _
                       DateRequest As Date, _
                       Description As String, _
                       Executor As String, _
                       emailExecutor As String, _
                       DateExecution As Date, _
                       ExecutorComment As String, _
                       TicketStatus As Integer, _
                       TicketType As Integer, _
                       m_elementName As String, _
                       m_elementId As Integer)

            _TicketId = TicketId
            _hsId = hsId
            _TicketTitle = TicketTitle
            _Requester = Requester
            _emailRequester = emailRequester
            _DateRequest = DateRequest
            _Description = Description
            _Executor = Executor
            _emailExecutor = emailExecutor
            _DateExecution = DateExecution
            _ExecutorComment = ExecutorComment
            _TicketStatus = TicketStatus
            _TicketType = TicketType
            m_elementName = elementName
            m_elementId = elementId
        End Sub
        Public Sub New()
            Me.New(0, _
                   0, _
                   String.Empty, _
                   String.Empty, _
                   String.Empty, _
                   FormatDateTime("01/01/1900", DateFormat.GeneralDate), _
                   String.Empty, _
                   String.Empty, _
                   String.Empty, _
                   FormatDateTime("01701/1900", DateFormat.GeneralDate), _
                   String.Empty, _
                   0, _
                   0, _
                   String.Empty, _
                   0)
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, _
                                   TicketTitle As String, _
                                   Requester As String, _
                                   emailRequester As String, _
                                   Description As String, _
                                   Executor As String, _
                                   emailExecutor As String, _
                                   UserName As String, _
                                   TicketType As Integer) As Boolean
            If String.IsNullOrEmpty(TicketTitle) Then
                Return False
            End If
            If String.IsNullOrEmpty(Requester) Then
                Return False
            End If
            If String.IsNullOrEmpty(emailRequester) Then
                Return False
            End If
            If String.IsNullOrEmpty(Description) Then
                Return False
            End If
            If String.IsNullOrEmpty(Executor) Then
                Return False
            End If
            If String.IsNullOrEmpty(emailExecutor) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Add(hsId, TicketTitle, Requester, emailRequester, Description, Executor, emailExecutor, UserName, TicketType)
        End Function

        Public Shared Function AddBySystem(hsId As Integer, _
                                           TicketTitle As String, _
                                           Requester As String, _
                                           emailRequester As String, _
                                           Description As String, _
                                           Executor As String, _
                                           emailExecutor As String, _
                                           UserName As String, _
                                           elementName As String, _
                                           elementId As Integer) As Boolean
            If String.IsNullOrEmpty(TicketTitle) Then
                Return False
            End If
            If String.IsNullOrEmpty(Requester) Then
                Return False
            End If
            If String.IsNullOrEmpty(emailRequester) Then
                Return False
            End If
            If String.IsNullOrEmpty(Description) Then
                Return False
            End If
            If String.IsNullOrEmpty(Executor) Then
                Return False
            End If
            If String.IsNullOrEmpty(emailExecutor) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_AddBySystem(hsId, TicketTitle, Requester, emailRequester, Description, Executor, emailExecutor, UserName, elementName, elementId)
        End Function

        Public Shared Function Del(TicketId As Integer) As Boolean
            If TicketId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Del(TicketId)
        End Function

        Public Shared Function getTotOpen(hsId As Integer) As Integer
            If hsId <= 0 Then
                Return 0
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_getTotOpen(hsId)
        End Function

        Public Shared Function List(hsId As Integer, TicketStatus As Integer, searchString As String) As List(Of hs_Tickets)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_List(hsId, TicketStatus, searchString)
        End Function

        Public Shared Function ListPaged(hsId As Integer, TicketStatus As Integer, searchString As String, RowNumber As Integer) As List(Of hs_Tickets)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_ListPaged(hsId, TicketStatus, searchString, RowNumber)
        End Function

        Public Shared Function Read(TicketId As Integer) As hs_Tickets
            If TicketId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Read(TicketId)
        End Function

        Public Shared Function readByElement(hsId As Integer, elementName As String, elementId As Integer) As hs_Tickets
            If hsId <= 0 Then Return Nothing
            If String.IsNullOrEmpty(elementName) Then Return Nothing
            If elementId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_Tickets_readByElement(hsId, elementName, elementId)
        End Function

        Public Shared Function Update(TicketId As Integer, _
                                      TicketTitle As String, _
                                      Requester As String, _
                                      emailRequester As String, _
                                      Description As String, _
                                      Executor As String, _
                                      emailExecutor As String, _
                                      UserName As String) As Boolean
            If TicketId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(TicketTitle) Then
                Return False
            End If
            If String.IsNullOrEmpty(Requester) Then
                Return False
            End If
            If String.IsNullOrEmpty(emailRequester) Then
                Return False
            End If
            If String.IsNullOrEmpty(Description) Then
                Return False
            End If
            If String.IsNullOrEmpty(Executor) Then
                Return False
            End If
            If String.IsNullOrEmpty(emailExecutor) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Update(TicketId, TicketTitle, Requester, emailRequester, Description, Executor, emailExecutor, UserName)
        End Function

        Public Shared Function ChangeStatus(TicketId As Integer, DateExecution As Date, ExecutorComment As String, TicketStatus As Integer) As Boolean
            If TicketId <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_ChangeStatus(TicketId, DateExecution, ExecutorComment, TicketStatus)
        End Function

        Public Shared Function setElement(TicketId As Integer, elementName As String, elementId As Integer) As Boolean
            If TicketId <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_Tickets_setElement(TicketId, elementName, elementId)
        End Function
#End Region

#Region "public properties"
        Public Property TicketId As Integer
            Get
                Return Me._TicketId
            End Get
            Set(value As Integer)
                Me._TicketId = value
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
        Public Property TicketTitle As String
            Get
                Return Me._TicketTitle
            End Get
            Set(value As String)
                Me._TicketTitle = value
            End Set
        End Property
        Public Property Requester As String
            Get
                Return Me._Requester
            End Get
            Set(value As String)
                Me._Requester = value
            End Set
        End Property
        Public Property emailRequester As String
            Get
                Return Me._emailRequester
            End Get
            Set(value As String)
                Me._emailRequester = value
            End Set
        End Property
        Public Property DateRequest As Date
            Get
                Return Me._DateRequest
            End Get
            Set(value As Date)
                Me._DateRequest = value
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
        Public Property Executor As String
            Get
                Return Me._Executor
            End Get
            Set(value As String)
                Me._Executor = value
            End Set
        End Property
        Public Property emailExecutor As String
            Get
                Return Me._emailExecutor
            End Get
            Set(value As String)
                Me._emailExecutor = value
            End Set
        End Property
        Public Property DateExecution As Date
            Get
                Return Me._DateExecution
            End Get
            Set(value As Date)
                Me._DateExecution = value
            End Set
        End Property
        Public Property ExecutorComment As String
            Get
                Return Me._ExecutorComment
            End Get
            Set(value As String)
                Me._ExecutorComment = value
            End Set
        End Property
        Public Property TicketStatus As Integer
            Get
                Return Me._TicketStatus
            End Get
            Set(value As Integer)
                Me._TicketStatus = value
            End Set
        End Property
        Public Property TicketType As Integer
            Get
                Return _TicketType
            End Get
            Set(value As Integer)
                _TicketType = value
            End Set
        End Property

        Public Property elementName As String
        Public Property elementId As Integer
#End Region
    End Class
End Namespace
