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
    Public Class log_LuxM

#Region "constructor"
        Public Sub New()
            Me.New(0, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0, String.Empty, String.Empty, 0, 0, 0, 0, 0, 0, 0, False)
        End Sub
        Public Sub New(m_LogId As Integer, _
                       m_dtLog As Date, _
                       m_hsId As Integer, _
                       m_Cod As String, _
                       m_Descr As String, _
                       m_stato As Integer, _
                       m_Voltage As Decimal, _
                       m_Curr As Decimal, _
                       m_EnergyCounter As Decimal, _
                       m_WorkingTimeCounter As Decimal, _
                       m_PowerOnCycleCounter As Decimal, _
                       m_Temp As Decimal, _
                       m_LightON As Boolean)
            LogId = m_LogId
            dtLog = m_dtLog
            hsId = m_hsId
            Cod = m_Cod
            Descr = m_Descr

            stato = m_stato
            Voltage = m_Voltage
            Curr = m_Curr
            EnergyCounter = m_EnergyCounter
            WorkingTimeCounter = m_WorkingTimeCounter
            PowerOnCycleCounter = m_PowerOnCycleCounter
            Temp = m_Temp
            LightON = m_LightON
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer,
                                   Cod As String,
                                   Descr As String,
                                   Voltage As Decimal,
                                   Curr As Decimal,
                                   EnergyCounter As Decimal,
                                   WorkingTimeCounter As Decimal,
                                   PowerOnCycleCounter As Decimal,
                                   Temp As Decimal,
                                   stato As Integer,
                                   LightON As Boolean,
                                   dtLog As Date) As Integer
            If hsId <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.log_LuxM_Add(hsId, Cod, Descr, Voltage, Curr, EnergyCounter, WorkingTimeCounter, PowerOnCycleCounter, Temp, stato, LightON, dtLog)

        End Function

        Public Shared Function List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_LuxM)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_LuxM_List(hsId, Cod, fromDate, toDate)
        End Function

        Public Shared Function ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_LuxM)
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_LuxM_ListPaged(hsId, Cod, fromDate, toDate, rowNumber)
        End Function

        Public Shared Function logNotSent(hsId As Integer, Cod As String) As log_LuxM
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_LuxM_logNotSent(hsId, Cod)
        End Function

        Public Shared Function setIsSent(Logid As Integer) As Boolean
            If Logid <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.log_LuxM_setIsSent(Logid)
        End Function

        Public Shared Function ReadLast(hsId As Integer, Cod As String) As log_LuxM
            If hsId <= 0 Then
                Return Nothing
            End If
            If String.IsNullOrEmpty(Cod) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.log_LuxM_ReadLast(hsId, Cod)
        End Function
#End Region

#Region "public properties"
        Public Property LogId As Integer
        Public Property dtLog As Date
        Public Property hsId As Integer
        Public Property Cod As String
        Public Property Descr As String

        Public Property stato As Integer

        Public Property Voltage As Decimal
        Public Property Curr As Decimal
        Public Property EnergyCounter As Decimal
        Public Property WorkingTimeCounter As Decimal
        Public Property PowerOnCycleCounter As Decimal
        Public Property Temp As Decimal
        Public Property LightON As Boolean

#End Region
    End Class
End Namespace
