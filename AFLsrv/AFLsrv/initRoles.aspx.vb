Public Class initRoles
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            'Administrators
            Dim a As Boolean = SCP.BLL.aspnetroles.Add("Administrators")
            If a = True Then
                Response.Write("role Administrators created")
                Dim u As Boolean = SCP.BLL.aspnetUsers.AddUser("sysadmin", "Kom2013", "Administrators", "super user", "sistemi@clevergy.it")
                If u = True Then
                    Response.Write("user: sysadmin created")
                Else
                    Response.Write("ERROR! cannot create user sysadmin")
                End If
            Else
                Response.Write("ERROR! cannot create role Administrators")
            End If

        End If
    End Sub

End Class