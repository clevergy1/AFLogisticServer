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
    Public Class hs_Elem

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        End Sub
        Public Sub New(m_totLux As Integer, m_statoLux As Integer,
                       m_totLuxM As Integer, m_statoLuxM As Integer,
                       m_totPsg As Integer, m_statoPsg As Integer,
                       m_totZrel As Integer, m_statoZrel As Integer,
                       m_totCronograph As Integer, m_statoCronograph As Integer)
            totLux = m_totLux
            statoLux = m_statoLux
            totLuxM = m_totLuxM
            statoLuxM = m_statoLuxM
            totPsg = m_totPsg
            statoPsg = m_statoPsg
            totZrel = m_totZrel
            statoZrel = m_statoZrel
            totCronograph = m_totCronograph
            statoCronograph = m_statoCronograph
        End Sub
#End Region

#Region "methods"
        Public Shared Function Read(hsId As Integer) As hs_Elem
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Elem_Read(hsId)
        End Function
#End Region

#Region "public properties"
        Public Property totLux As Integer
        Public Property statoLux As Integer
        Public Property totLuxM As Integer
        Public Property statoLuxM As Integer
        Public Property totPsg As Integer
        Public Property statoPsg As Integer
        Public Property totZrel As Integer
        Public Property statoZrel As Integer
        Public Property totCronograph As Integer
        Public Property statoCronograph As Integer
#End Region
    End Class
End Namespace
