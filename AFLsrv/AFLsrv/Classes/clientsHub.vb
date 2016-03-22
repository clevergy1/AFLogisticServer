Imports Microsoft.AspNet.SignalR
Namespace SCP

    Public Class clientsHub
        Inherits Hub

        Public Sub subscribe(IdImpianto As String)
            Groups.Add(Context.ConnectionId, IdImpianto)
        End Sub
        Public Sub unsubscribe(IdImpianto As String)
            If Not String.IsNullOrEmpty(IdImpianto) Then
                Groups.Remove(Context.ConnectionId, IdImpianto)
            End If
        End Sub

    End Class
End Namespace
