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
    Public Class hs_Tickets_Executors

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, String.Empty, String.Empty, String.Empty)
        End Sub
        Public Sub New(m_Id As Integer, m_hsId As Integer, m_IdContatto As Integer, m_emailaddress As String, m_Nome As String, m_Descrizione As String)
            Id = m_Id
            hsId = m_hsId
            IdContatto = m_IdContatto
            emailaddress = m_emailaddress
            Nome = m_Nome
            Descrizione = m_Descrizione
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
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Executors_Add(hsId, IdContatto)
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Executors_Del(Id)
        End Function

        Public Shared Function List(hsId As Integer) As List(Of hs_Tickets_Executors)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Executors_List(hsId)
        End Function

        Public Shared Function Read(Id As Integer) As hs_Tickets_Executors
            If Id <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Tickets_Executors_Read(Id)
        End Function
#End Region

#Region "public properties"
        Public Property Id As Integer
        Public Property hsId As Integer
        Public Property IdContatto As Integer
        Public Property emailaddress As String
        Public Property Nome As String
        Public Property Descrizione As String
#End Region
    End Class
End Namespace
