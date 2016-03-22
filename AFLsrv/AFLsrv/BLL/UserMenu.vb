Imports System
Imports System.Collections.Generic
Imports SCP.DAL


Namespace SCP.BLL
    Public Class UserMenu


#Region "class implementation"
        Public Sub New()
            Me.New(0, String.Empty, 0, String.Empty, 0, String.Empty)
        End Sub
        Public Sub New(m_IdVoceMenu As Integer, _
                       m_DescrVoce As String, _
                       m_IdAzione As Integer, _
                       m_URLAzione As String, _
                       m_IdVocePadre As Integer, _
                       m_mainPage As String)
            IdVoceMenu = m_IdVoceMenu
            DescrVoce = m_DescrVoce
            IdAzione = m_IdAzione
            URLAzione = m_URLAzione
            IdVocePadre = m_IdVocePadre
            mainPage = m_mainPage
        End Sub
#End Region

#Region "methods"
        Public Shared Function List(UserName As String) As List(Of UserMenu)
            If String.IsNullOrEmpty(UserName) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.UserMenu_List(UserName)
        End Function

        Public Shared Function ListFather(UserName As String) As List(Of UserMenu)
            If String.IsNullOrEmpty(UserName) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.UserMenu_ListFather(UserName)
        End Function

        Public Shared Function ListChildren(UserName As String, IdVocePadre As Integer) As List(Of UserMenu)
            If String.IsNullOrEmpty(UserName) Then
                Return Nothing
            End If
            If IdVocePadre = 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.UserMenu_ListChildren(UserName, IdVocePadre)
        End Function
#End Region

#Region "properties"
        Public Property IdVoceMenu As Integer
        Public Property DescrVoce As String
        Public Property IdAzione As Integer
        Public Property URLAzione As String
        Public Property IdVocePadre As Integer
        Public Property mainPage As String
#End Region
    End Class
End Namespace
