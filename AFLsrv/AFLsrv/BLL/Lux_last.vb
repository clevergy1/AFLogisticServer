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
    Public Class Lux_last
#Region "constructor"
        Public Sub New()
            Me.New(0, String.Empty, 0, FormatDateTime("01/01/1900", DateFormat.GeneralDate))
        End Sub
        Public Sub New(m_hsId As Integer,
                         m_Cod As String,
                       m_lastLog As Integer,
                       m_lastdtLog As Date)

            hsId = m_hsId
            Cod = m_Cod
            lastLog = m_lastLog
            lastdtLog = m_lastdtLog



        End Sub


#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer,
                                         Cod As String,
                                         lastLog As Integer,
                                         lastdtLog As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Lux_last_Add(hsId, Cod, lastLog, lastdtLog)

        End Function

        Public Shared Function Upd(hsId As Integer,
                                         Cod As String,
                                         lastLog As Integer,
                                         lastdtLog As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Lux_last_Upd(hsId, Cod, lastLog, lastdtLog)

        End Function

        Public Shared Function Read(hsId As Integer, Cod As String) As Lux_last
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Lux_last_Read(hsId, Cod)
        End Function

        Public Shared Function Upsert(hsId As Integer, Cod As String, lastLog As Integer, lastdtLog As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Dim _obj As Lux_last = Read(hsId, Cod)
            Dim retval As Boolean = False
            If _obj Is Nothing Then
                retval = Add(hsId, Cod, lastLog, lastdtLog)
            Else
                retval = Upd(hsId, Cod, lastLog, lastdtLog)
            End If
            Return retval
        End Function
#End Region

#Region "public properties"
        Public Property hsId As Integer
        Public Property Cod As String
        Public Property lastLog As Integer
        Public Property lastdtLog As Date
#End Region
    End Class
End Namespace
