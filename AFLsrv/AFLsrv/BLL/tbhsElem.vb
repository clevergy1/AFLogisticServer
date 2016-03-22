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
    Public Class tbhsElem

#Region "constructor"
        Public Sub New()
            Me.New(0, String.Empty, String.Empty)
        End Sub
        Public Sub New(m_id As Integer, m_name As String, m_descr As String)
            Id = m_id
            name = m_name
            descr = m_descr
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(Id As Integer, name As String, descr As String) As Boolean
            If Id <= 0 Then Return False
            If String.IsNullOrEmpty(name) Then Return False
            Return DataAccessHelper.GetDataAccess.tbhsElem_Add(Id, name, descr)
        End Function

        Public Shared Function Del(Id As Integer) As Boolean
            If Id <= 0 Then Return False
            Return DataAccessHelper.GetDataAccess.tbhsElem_Del(Id)
        End Function

        Public Shared Function List() As List(Of tbhsElem)
            Return DataAccessHelper.GetDataAccess.tbhsElem_List
        End Function

        Public Shared Function Read(Id As Integer) As tbhsElem
            If Id <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.tbhsElem_Read(Id)
        End Function

        Public Shared Function Update(Id As Integer, name As String, descr As String) As Boolean
            If Id <= 0 Then Return False
            If String.IsNullOrEmpty(name) Then Return False
            Return DataAccessHelper.GetDataAccess.tbhsElem_Update(Id, name, descr)
        End Function
#End Region

#Region "public properties"
        Public Property Id As Integer
        Public Property name As String
        Public Property descr As String
#End Region
    End Class
End Namespace
