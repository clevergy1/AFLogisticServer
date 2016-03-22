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
    Public Class hs_amb_Profile_Descr
#Region "constructor"
        Public Sub New()
            Me.New(0, 0, String.Empty)
        End Sub
        Public Sub New(_CronId As Integer, _ProfileNr As Integer, _descr As String)
            CronId = _CronId
            ProfileNr = _ProfileNr
            descr = _descr
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            If String.IsNullOrEmpty(descr) Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_Add(CronId, ProfileNr, descr)
        End Function

        Public Shared Function Del(CronId As Integer, ProfileNr As Integer) As Boolean
            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            Return DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_Del(CronId, ProfileNr)
        End Function

        Public Shared Function List(CronId As Integer) As List(Of hs_amb_Profile_Descr)
            If CronId <= 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_List(CronId)
        End Function

        Public Shared Function Read(CronId As Integer, ProfileNr As Integer) As hs_amb_Profile_Descr
            If CronId <= 0 Then Return Nothing
            If ProfileNr < 0 Then Return Nothing
            Return DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_Read(CronId, ProfileNr)
        End Function

        Public Shared Function Update(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
            If CronId <= 0 Then Return False
            If ProfileNr < 0 Then Return False
            If String.IsNullOrEmpty(descr) Then Return False
            If DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_Read(CronId, ProfileNr) Is Nothing Then
                Return DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_Add(CronId, ProfileNr, descr)
            Else
                Return DataAccessHelper.GetDataAccess.hs_amb_Profile_Descr_Update(CronId, ProfileNr, descr)
            End If
        End Function
#End Region

#Region "public properties"
        Public Property CronId As Integer
        Public Property ProfileNr As Integer
        Public Property descr As String
#End Region

    End Class
End Namespace
