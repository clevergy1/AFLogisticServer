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
    Public Class hs_Docs

#Region "constructor"
        Public Sub New()
            Me.New(0, 0, String.Empty, String.Empty, String.Empty, String.Empty)
        End Sub

        Public Sub New(m_IdDoc As Integer, _
                       m_hsId As Integer, _
                       m_DocName As String, _
                       m_ContentType As String, _
                       m_DocSize As String, _
                       m_fileExt As String)
            IdDoc = m_IdDoc
            hsId = m_hsId
            DocName = m_DocName
            ContentType = m_ContentType
            DocSize = m_DocSize
            fileExt = m_fileExt
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(hsId As Integer, _
                                   DocName As String, _
                                   Creator As String) As Integer
            If String.IsNullOrEmpty(DocName) Then
                Return -1
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_Add(hsId, DocName, Creator)
        End Function

        Public Shared Function Add(hsId As Integer, _
                                   DocName As String, _
                                   Creator As String, _
                                   ContentType As String, _
                                   DocSize As String, _
                                   BinaryData As String) As Integer
            If String.IsNullOrEmpty(DocName) Then
                Return -1
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_Add(hsId, DocName, Creator, ContentType, DocSize, BinaryData)
        End Function

        Public Shared Function Del(IdDoc As Integer) As Boolean
            If IdDoc <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_Del(IdDoc)
        End Function

        Public Shared Function List(hsId As Integer) As List(Of hs_Docs)
            If hsId <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_List(hsId)
        End Function

        Public Shared Function Read(IdDoc As Integer) As hs_Docs
            If IdDoc <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_Read(IdDoc)
        End Function

        Public Shared Function getBinaryData(IdDoc As Integer) As String
            If IdDoc <= 0 Then
                Return String.Empty
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_getBinaryData(IdDoc)
        End Function

        Public Shared Function getTot(hsId As Integer) As Integer
            If hsId <= 0 Then
                Return 0
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_getTot(hsId)
        End Function

        Public Shared Function setBinaryData(IdDoc As Integer, BinaryData As String, UserName As String) As Boolean
            If IdDoc <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(BinaryData) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_setBinaryData(IdDoc, BinaryData, UserName)
        End Function

        Public Shared Function setInfo(IdDoc As Integer, ContentType As String, DocSize As String, fileExt As String) As Boolean
            If IdDoc <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_setInfo(IdDoc, ContentType, DocSize, fileExt)
        End Function

        Public Shared Function Update(IdDoc As Integer, DocName As String, UserName As String) As Boolean
            If IdDoc <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(DocName) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.hs_Docs_Update(IdDoc, DocName, UserName)
        End Function
#End Region

#Region "public properties"
        Public Property IdDoc As Integer
        Public Property hsId As Integer
        Public Property DocName() As String
        Public Property ContentType() As String
        Public Property DocSize() As String
        Public Property fileExt As String
#End Region
    End Class
End Namespace
