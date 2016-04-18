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
    Public Class Ambienti
#Region "constructor"
        Public Sub New()
            Me.New(0, 0, String.Empty, String.Empty, String.Empty)
        End Sub
        Public Sub New(m_hsId As Integer,
                       m_IdAmbiente As Integer,
                       m_Cod As String,
                       m_DescrizioneAmbiente As String,
                       m_CronCod As String)

            hsId = m_hsId
            IdAmbiente = m_IdAmbiente
            Cod = m_Cod
            DescrizioneAmbiente = m_DescrizioneAmbiente
            CronCod = m_CronCod


        End Sub


#End Region

#Region "methods"
        Public Shared Function List(hsId As Integer) As List(Of Ambienti)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Ambienti_List(hsId)
        End Function

        Public Shared Function Read(hsId As Integer, IdAmbiente As Integer) As Ambienti
            If IdAmbiente <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Ambienti_Read(hsId, IdAmbiente)
        End Function

        Public Shared Function SetCron(hsId As Integer, IdAmbiente As Integer, CronCod As String) As Ambienti
            If IdAmbiente <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Ambienti_SetCron(hsId, IdAmbiente, CronCod)
        End Function



#End Region

#Region "public properties"
        Public Property hsId As Integer
        Public Property IdAmbiente As Integer
        Public Property Cod As String
        Public Property DescrizioneAmbiente As String
        Public Property CronCod As String
#End Region
    End Class
End Namespace
