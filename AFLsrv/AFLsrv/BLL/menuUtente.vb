Imports System
Imports System.Collections.Generic
Imports SCP.DAL


Namespace SCP.BLL
    Public Class menuUtente
        Private _Menu As String
        Private _VociMenu As List(Of UserMenu)

        Public Sub New()
            Me.New(String.Empty, New List(Of UserMenu))
        End Sub
        Public Sub New(Menu As String, VociMenu As List(Of UserMenu))
            Me._Menu = Menu
            Me._VociMenu = VociMenu
        End Sub

        Public Shared Function Read(username As String) As List(Of menuUtente)
            If String.IsNullOrEmpty(username) Then
                Return Nothing
            End If
            Dim retVal As New List(Of menuUtente)

            Dim lmenu As List(Of UserMenu) = DataAccessHelper.GetDataAccess.UserMenu_ListFather(username)
            If Not lmenu Is Nothing Then
                For Each m As UserMenu In lmenu
                    Dim _menuUtente As New menuUtente
                    _menuUtente.Menu = m.DescrVoce
                    Dim lvoci As List(Of UserMenu) = DataAccessHelper.GetDataAccess.UserMenu_ListChildren(username, m.IdVoceMenu)
                    If Not lvoci Is Nothing Then
                        _menuUtente.VociMenu = New List(Of UserMenu)
                        _menuUtente.VociMenu.AddRange(lvoci)
                        retVal.Add(_menuUtente)
                        _menuUtente = Nothing
                        lvoci = Nothing
                    End If
                Next
                lmenu = Nothing
            End If
            Return retVal
        End Function

        Public Property Menu As String
            Get
                Return Me._Menu
            End Get
            Set(value As String)
                Me._Menu = value
            End Set
        End Property
        Public Property VociMenu As List(Of UserMenu)
            Get
                Return Me._VociMenu
            End Get
            Set(value As List(Of UserMenu))
                Me._VociMenu = _VociMenu
            End Set
        End Property
    End Class
End Namespace
