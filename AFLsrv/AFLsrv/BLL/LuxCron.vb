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
    Public Class LuxCron

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, String.Empty, String.Empty, String.Empty, String.Empty)
        End Sub
        Public Sub New(m_Id As Integer, m_LuxId As Integer, m_CronId As Integer, m_LuxCod As String, m_LuxDescr As String, m_CronCod As String, m_CronDescr As String)
            Id = m_Id
            LuxId = m_LuxId
            CronId = m_CronId
            LuxCod = m_LuxCod
            LuxDescr = m_LuxDescr
            CronCod = m_CronCod
            CronDescr = m_CronDescr
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(LuxId As Integer, CronId As Integer, UserName As String) As Boolean
            If LuxId <= 0 Then Return False
            If CronId <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.LuxCron_Add(LuxId, CronId, UserName)
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.LuxCron_Del(Id)
        End Function

        Public Shared Function List(LuxId As Integer) As List(Of LuxCron)
            If LuxId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.LuxCron_List(LuxId)
        End Function

        Public Shared Function Read(Id As Integer) As LuxCron
            If Id <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.LuxCron_Read(Id)
        End Function


#End Region

#Region "property"
        Public Property Id As Integer
        Public Property LuxId As Integer
        Public Property CronId As Integer
        Public Property LuxCod As String
        Public Property LuxDescr As String
        Public Property CronCod As String
        Public Property CronDescr As String
#End Region
    End Class
End Namespace
