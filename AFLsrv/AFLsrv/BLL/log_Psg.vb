﻿'
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
    Public Class log_Psg

#Region "constructor"
        Public Sub New()
            Me.New(0, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0, String.Empty, String.Empty, 0, 0)
        End Sub
        Public Sub New(m_LogId As Integer, _
                       m_dtLog As Date, _
                       m_hsId As Integer, _
                       m_Cod As String, _
                       m_Descr As String, _
                       m_stato As Integer, _
                       m_currentValue As integer)
            LogId = m_LogId
            dtLog = m_dtLog
            hsId = m_hsId
            Cod = m_Cod
            Descr = m_Descr

            stato = m_stato
            currentValue = m_currentValue
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, _
                                   Cod As String, _
                                   Descr As String, _
                                   currentValue As Integer, _
                                   stato As Integer, _
                                   dtLog As Date) As Boolean
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.log_Psg_Add(hsId, Cod, Descr, currentValue, stato, dtLog)
        End Function

        Public Shared Function List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Psg)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_Psg_List(hsId, Cod, fromDate, toDate)
        End Function

        Public Shared Function ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Psg)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_Psg_ListPaged(hsId, Cod, fromDate, toDate, rowNumber)
        End Function

        Public Shared Function logNotSent(hsId As Integer, Cod As String) As log_Psg
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_Psg_logNotSent(hsId, Cod)
        End Function

        Public Shared Function setIsSent(Logid As Integer) As Boolean
            If Logid <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.log_Psg_setIsSent(Logid)
        End Function

        Public Shared Function ReadLast(hsId As Integer, Cod As String) As log_Psg
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_Psg_ReadLast(hsId, Cod)
        End Function

#End Region

#Region "public properties"
        Public Property LogId As Integer
        Public Property dtLog As Date
        Public Property hsId As Integer
        Public Property Cod As String
        Public Property Descr As String

        Public Property stato As Integer
        Public Property currentValue As Integer


#End Region
    End Class
End Namespace
