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
Imports System.IO

Namespace SCP.BLL
    Public Class Impianti


#Region "constructor"
        Public Sub New()
            Me.New(String.Empty, String.Empty, String.Empty, 0, 0, 0, False)
        End Sub
        Public Sub New(m_IdImpianto As String, _
                       m_DesImpianto As String, _
                       m_Indirizzo As String, _
                       m_Latitude As Decimal, _
                       m_Longitude As Decimal, _
                       m_AltSLM As Decimal, _
                       m_IsActive As Boolean)
            IdImpianto = m_IdImpianto
            DesImpianto = m_DesImpianto
            Indirizzo = m_DesImpianto
            Latitude = m_Latitude
            Longitude = m_Longitude
            AltSLM = m_AltSLM
            IsActive = m_IsActive
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(DesImpianto As String, _
                                   Indirizzo As String, _
                                   Latitude As Decimal, _
                                   Longitude As Decimal, _
                                   AltSLM As Decimal, _
                                   IsActive As Boolean) As Boolean
            If String.IsNullOrEmpty(DesImpianto) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Add(DesImpianto, Indirizzo, Latitude, Longitude, AltSLM, IsActive)
        End Function

        Public Shared Function List(searchString As String) As List(Of Impianti)
            Return DataAccessHelper.GetDataAccess.Impianti_List(searchString)
        End Function

        Public Shared Function ListByUser(UserId As String, searchString As String) As List(Of Impianti)
            If String.IsNullOrEmpty(UserId) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_ListByUser(UserId, searchString)
        End Function

        Public Shared Function ListPaged(id As Integer, searchString As String) As List(Of Impianti)
            Return DataAccessHelper.GetDataAccess.Impianti_ListPaged(id, searchString)
        End Function

        Public Shared Function Read(IdImpianto As String) As Impianti
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Read(IdImpianto)
        End Function

        Public Shared Function setIsActive(IdImpianto As String, IsActive As Boolean) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_setIsActive(IdImpianto, IsActive)
        End Function

        Public Shared Function Upd(IdImpianto As String, DesImpianto As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Decimal, IsActive As Boolean) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            If String.IsNullOrEmpty(DesImpianto) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Upd(IdImpianto, DesImpianto, Indirizzo, Latitude, Longitude, AltSLM, IsActive)
        End Function
#End Region

#Region "public properties"
        Public Property IdImpianto As String
        Public Property DesImpianto As String
        Public Property Indirizzo As String
        Public Property Latitude As Decimal
        Public Property Longitude As Decimal
        Public Property AltSLM As Decimal
        Public Property IsActive As Boolean
#End Region
    End Class
End Namespace
