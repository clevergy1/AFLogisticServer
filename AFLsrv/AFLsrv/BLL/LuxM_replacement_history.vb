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
Imports System.Threading

Namespace SCP.BLL
    Public Class LuxM_replacement_history

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, String.Empty, FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, String.Empty)
        End Sub
        Public Sub New(m_Id As Integer, m_ParentId As Integer, m_marcamodello As String, m_installationDate As Date, m_note As String, m_userName As String)
            Id = m_Id
            ParentId = m_ParentId
            marcamodello = m_marcamodello
            installationDate = m_installationDate
            note = m_note
            m_userName = m_userName
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            If ParentId <= 0 Then Return False
            If installationDate.Year <= 1900 Then Return False
            Return DataAccessHelper.GetDataAccess.LuxM_replacement_history_Add(ParentId, marcamodello, installationDate, note, userName)
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.LuxM_replacement_history_Del(Id)
        End Function

        Public Shared Function List(ParentId As Integer) As List(Of LuxM_replacement_history)
            If ParentId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.LuxM_replacement_history_List(ParentId)
        End Function

        Public Shared Function Read(Id As Integer) As LuxM_replacement_history
            If Id <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.LuxM_replacement_history_Read(Id)
        End Function

        Public Shared Function Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            If Id <= 0 Then Return False
            If installationDate.Year <= 1900 Then Return False
            Return DataAccessHelper.GetDataAccess.LuxM_replacement_history_Update(Id, marcamodello, installationDate, note, userName)
        End Function
#End Region

#Region "public properties"
        Public Property Id As Integer
        Public Property ParentId As Integer
        Public Property marcamodello As String
        Public Property installationDate As Date
        Public Property note As String
        Public Property userName As String
#End Region
    End Class
End Namespace
