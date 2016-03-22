Imports Microsoft.Owin.Cors
Imports Owin
Imports Microsoft.AspNet.SignalR
Public Class Startup
    Public Sub Configuration(app As IAppBuilder)

        app.Map("/signalr", Sub(map)
                                map.UseCors(CorsOptions.AllowAll)
                                Dim hubConfiguration = New HubConfiguration
                                hubConfiguration.EnableJSONP = True
                                map.RunSignalR(hubConfiguration)
                            End Sub)

    End Sub
End Class
