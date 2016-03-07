Imports System.Web.SessionState
Imports System
Imports System.Web
Imports System.Web.Routing

Imports Microsoft.Owin.Cors
Imports Owin
Imports Microsoft.AspNet.SignalR
Imports System.IO

Public Class Global_asax
    Inherits System.Web.HttpApplication

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application is started
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session is started
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        If Not HttpContext.Current.Request.Path.Contains("signalr/") Then
            HttpContext.Current.Response.AddHeader("Access-Control-Allow-Origin", "*")
            HttpContext.Current.Response.AddHeader("Access-Control-Allow-Credentials", "false")
            HttpContext.Current.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, Accept")
        End If

        If HttpContext.Current.Request.HttpMethod = "OPTIONS" Then
            HttpContext.Current.Response.AddHeader("Access-Control-Allow-Methods", "POST,GET,OPTIONS,PUT,DELETE")
            HttpContext.Current.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, Accept")
            HttpContext.Current.Response.End()
        End If
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub

End Class