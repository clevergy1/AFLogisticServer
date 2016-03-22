Imports System
Imports System.Configuration
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlTypes
Imports SCP.BLL

Namespace SCP.DAL

    Public MustInherit Class DataAccess

#Region "properties"
        Public Shared ReadOnly Property SCP() As String
            Get
                Dim connStr As String = ""
                If Not ConfigurationManager.ConnectionStrings("SCP") Is Nothing Then
                    connStr = ConfigurationManager.ConnectionStrings("SCP").ConnectionString
                End If

                If String.IsNullOrEmpty(connStr) Then
                    Throw New NullReferenceException("ConnectionString configuration is missing from you app.config. It should contain  <connectionStrings> <add key=""SCP"" value=""Server=(local);Integrated Security=True;Database=[database name]"" </connectionStrings>")
                Else
                    Return connStr
                End If
            End Get
        End Property
        Public Shared ReadOnly Property HS_LOG() As String
            Get
                Dim connStr As String = ""
                If Not ConfigurationManager.ConnectionStrings("HS_LOG") Is Nothing Then
                    connStr = ConfigurationManager.ConnectionStrings("HS_LOG").ConnectionString
                End If

                If String.IsNullOrEmpty(connStr) Then
                    Throw New NullReferenceException("ConnectionString configuration is missing from you app.config. It should contain  <connectionStrings> <add key=""HS_LOG"" value=""Server=(local);Integrated Security=True;Database=[database name]"" </connectionStrings>")
                Else
                    Return connStr
                End If
            End Get
        End Property
#End Region

#Region "aspnetroles"
        Public MustOverride Function aspnetroles_Add(RoleName As String) As Boolean
        Public MustOverride Function aspnetroles_Del(RoleName As String) As Boolean
        Public MustOverride Function aspnetroles_List() As List(Of aspnetroles)
        Public MustOverride Function aspnetroles_ListActive(IdImpianto As String) As List(Of aspnetroles)
        Public MustOverride Function aspnetroles_Update(RoleId As String, RoleName As String) As Boolean
#End Region

#Region "aspnetUsers"
        Public MustOverride Function aspnetUsers_approveUser(userId As String) As Boolean
        Public MustOverride Function aspnetUsers_disapproveUser(userId As String) As Boolean
        Public MustOverride Function aspnetUsers_lockUser(UserId As String) As Boolean
        Public MustOverride Function aspnetUsers_GetActiveUsersByRoleName(RoleName As String) As List(Of aspnetUsers)
        Public MustOverride Function aspnetUsers_GetTotActiveUsersByRoleName(RoleName As String) As Integer
        Public MustOverride Function aspnetUsers_GetUserByUserName(UserName As String) As aspnetUsers
        Public MustOverride Function aspnetUsers_GetUsersByRoleName(RoleName As String, searchString As String, IdImpianto As String) As List(Of aspnetUsers)
        Public MustOverride Function aspnetUsers_GetUsersAll() As List(Of aspnetUsers)
        Public MustOverride Function aspnetUsers_List(id As Integer, SearchString As String) As List(Of aspnetUsers)
        Public MustOverride Function aspnetUsers_Read(UserId As String) As aspnetUsers
        Public MustOverride Function aspnetUsers_whoIsOnline() As List(Of String)
#End Region

#Region "UserMenu"
        Public MustOverride Function UserMenu_List(UserName As String) As List(Of UserMenu)
        Public MustOverride Function UserMenu_ListFather(UserName As String) As List(Of UserMenu)
        Public MustOverride Function UserMenu_ListChildren(UserName As String, IdVocePadre As Integer) As List(Of UserMenu)
#End Region

#Region "aspnetusersstat"
        Public MustOverride Function aspnetusersstat_Read() As aspnetusersstat
#End Region

#Region "dbActivityMonitor"
        Public MustOverride Function dbActivityMonitor_List() As List(Of dbActivityMonitor)
#End Region

#Region "dbstats"
        Public MustOverride Function dbstats_Read() As dbstats
#End Region

#Region "Impianti"
        Public MustOverride Function Impianti_Add(DesImpianto As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Decimal, IsActive As Boolean) As Boolean
        Public MustOverride Function Impianti_List(searchString As String) As List(Of Impianti)
        Public MustOverride Function Impianti_ListByUser(UserId As String, searchString As String) As List(Of Impianti)
        Public MustOverride Function Impianti_ListPaged(id As Integer, searchString As String) As List(Of Impianti)
        Public MustOverride Function Impianti_Read(IdImpianto As String) As Impianti
        Public MustOverride Function Impianti_setIsActive(IdImpianto As String, IsActive As Boolean) As Boolean
        Public MustOverride Function Impianti_Upd(IdImpianto As String, DesImpianto As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Decimal, IsActive As Boolean) As Boolean
#End Region

#Region "Impianti_Contatti"
        Public MustOverride Function Impianti_Contatti_Add(IdImpianto As String, _
                                                           Descrizione As String, _
                                                           Indirizzo As String, _
                                                           Nome As String, _
                                                           TelFisso As String, _
                                                           TelMobile As String, _
                                                           emailaddress As String) As Boolean
        Public MustOverride Function Impianti_Contatti_Del(IdContatto As Integer) As Boolean
        Public MustOverride Function Impianti_Contatti_List(IdImpianto As String) As IList(Of Impianti_Contatti)
        Public MustOverride Function Impianti_Contatti_Read(IdContatto As Integer) As Impianti_Contatti
        Public MustOverride Function Impianti_Contatti_Update(IdContatto As Integer, _
                                                              Descrizione As String, _
                                                              Indirizzo As String, _
                                                              Nome As String, _
                                                              TelFisso As String, _
                                                              TelMobile As String, _
                                                              emailaddress As String) As Boolean
#End Region

#Region "Impianti_RemoteConnections"
        Public MustOverride Function Impianti_RemoteConnections_Add(IdImpianto As String, IdAddress As Integer, Descr As String, remoteaddress As String, connectionType As Integer, NoteInterne As String) As Boolean
        Public MustOverride Function Impianti_RemoteConnections_Del(IdImpianto As String, IdAddress As Integer) As Boolean
        Public MustOverride Function Impianti_RemoteConnections_getLastId() As Integer
        Public MustOverride Function Impianti_RemoteConnections_List(IdImpianto As String) As List(Of Impianti_RemoteConnections)
        Public MustOverride Function Impianti_RemoteConnections_Read(IdImpianto As String, IdAddress As Integer) As Impianti_RemoteConnections
        Public MustOverride Function Impianti_RemoteConnections_Upd(IdImpianto As String, IdAddress As Integer, Descr As String, remoteaddress As String, connectionType As Integer, NoteInterne As String) As Boolean
#End Region

#Region "tbhsElem"
        Public MustOverride Function tbhsElem_Add(Id As Integer, name As String, descr As String) As Boolean
        Public MustOverride Function tbhsElem_Del(Id As Integer) As Boolean
        Public MustOverride Function tbhsElem_List() As List(Of tbhsElem)
        Public MustOverride Function tbhsElem_Read(Id As Integer) As tbhsElem
        Public MustOverride Function tbhsElem_Update(Id As Integer, name As String, descr As String) As Boolean
#End Region

#Region "hs_Elem"
        Public MustOverride Function hs_Elem_Read(hsId As Integer) As hs_Elem
#End Region

#Region "hs_Docs"
        Public MustOverride Function hs_Docs_Add(hsId As Integer, _
                                                 DocName As String, _
                                                 Creator As String) As Integer
        Public MustOverride Function hs_Docs_Add(hsId As Integer, _
                                                 DocName As String, _
                                                 Creator As String, _
                                                 ContentType As String, _
                                                 DocSize As String, _
                                                 BinaryData As String) As Integer
        Public MustOverride Function hs_Docs_Del(IdDoc As Integer) As Boolean
        Public MustOverride Function hs_Docs_List(hsId As Integer) As List(Of hs_Docs)
        Public MustOverride Function hs_Docs_Read(IdDoc As Integer) As hs_Docs
        Public MustOverride Function hs_Docs_getBinaryData(IdDoc As Integer) As String
        Public MustOverride Function hs_Docs_getTot(hsId As Integer) As Integer
        Public MustOverride Function hs_Docs_setBinaryData(IdDoc As Integer, BinaryData As String, UserName As String) As Boolean
        Public MustOverride Function hs_Docs_setInfo(IdDoc As Integer, ContentType As String, DocSize As String, fileExt As String) As Boolean
        Public MustOverride Function hs_Docs_Update(IdDoc As Integer, DocName As String, UserName As String) As Boolean
#End Region

#Region "hs_ErrorCodes"
        Public MustOverride Function hs_ErrorCodes_Add(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
        Public MustOverride Function hs_ErrorCodes_Del(elementCode As String, errorCode As Integer) As Boolean
        Public MustOverride Function hs_ErrorCodes_getTot() As Integer
        Public MustOverride Function hs_ErrorCodes_List(elementCode As String) As List(Of hs_ErrorCodes)
        Public MustOverride Function hs_ErrorCodes_Read(elementCode As String, errorCode As Integer) As hs_ErrorCodes
        Public MustOverride Function hs_ErrorCodes_Update(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
#End Region

#Region "hs_ErrorLog"
        Public MustOverride Function hs_ErrorLog_Add(LogDate As Date, hsId As Integer, hselement As String, elementCode As String, errorCode As Integer, errorValue As String) As Boolean
        Public MustOverride Function hs_ErrorLog_List(hsId As Integer, fromDate As Date, toDate As Date) As List(Of hs_ErrorLog)
        Public MustOverride Function hs_ErrorLog_ListAll(hsId As Integer, rowNumber As Integer) As List(Of hs_ErrorLog)
        Public MustOverride Function hs_ErrorLog_ListByElement(hsId As Integer, hselement As String, rowNumber As Integer) As List(Of hs_ErrorLog)
#End Region

#Region "hs_Tickets"
        Public MustOverride Function hs_Tickets_Add(hsId As Integer, _
                                                    TicketTitle As String, _
                                                    Requester As String, _
                                                    emailRequester As String, _
                                                    Description As String, _
                                                    Executor As String, _
                                                    emailExecutor As String, _
                                                    UserName As String, _
                                                    TicketType As Integer) As Boolean
        Public MustOverride Function hs_Tickets_AddBySystem(hsId As Integer, _
                                                            TicketTitle As String, _
                                                            Requester As String, _
                                                            emailRequester As String, _
                                                            Description As String, _
                                                            Executor As String, _
                                                            emailExecutor As String, _
                                                            UserName As String, elementName As String, _
                                                            elementId As Integer) As Boolean
        Public MustOverride Function hs_Tickets_Del(TickedId As Integer) As Boolean
        Public MustOverride Function hs_Tickets_getTotOpen(hsid As Integer) As Integer
        Public MustOverride Function hs_Tickets_List(hsId As Integer, TicketStatus As Integer, searchString As String) As List(Of hs_Tickets)
        Public MustOverride Function hs_Tickets_ListPaged(hsId As Integer, TicketStatus As Integer, searchString As String, RowNumber As Integer) As List(Of hs_Tickets)
        Public MustOverride Function hs_Tickets_Read(TickedId As Integer) As hs_Tickets
        Public MustOverride Function hs_Tickets_readByElement(hsId As Integer, elementName As String, elementId As Integer) As hs_Tickets
        Public MustOverride Function hs_Tickets_Update(TickedId As Integer, _
                                                       TicketTitle As String, _
                                                       Requester As String, _
                                                       emailRequester As String, _
                                                       Description As String, _
                                                       Executor As String, _
                                                       emailExecutor As String, _
                                                       UserName As String) As Boolean
        Public MustOverride Function hs_Tickets_ChangeStatus(TickedId As Integer, _
                                                             DateExecution As Date, _
                                                             ExecutorComment As String, _
                                                             TicketStatus As Integer) As Boolean
        Public MustOverride Function hs_Tickets_setElement(TicketId As Integer, elementName As String, elementId As Integer) As Boolean
#End Region
#Region "hs_Tickets_Executors"
        Public MustOverride Function hs_Tickets_Executors_Add(hsId As Integer, IdContatto As Integer) As Boolean
        Public MustOverride Function hs_Tickets_Executors_Del(Id As Integer) As Boolean
        Public MustOverride Function hs_Tickets_Executors_List(hsId As Integer) As List(Of hs_Tickets_Executors)
        Public MustOverride Function hs_Tickets_Executors_Read(Id As Integer) As hs_Tickets_Executors
#End Region
#Region "hs_Tickets_Requesters"
        Public MustOverride Function hs_Tickets_Requesters_Add(hsId As Integer, IdContatto As Integer) As Boolean
        Public MustOverride Function hs_Tickets_Requesters_Del(Id As Integer) As Boolean
        Public MustOverride Function hs_Tickets_Requesters_List(hsId As Integer) As List(Of hs_Tickets_Requesters)
        Public MustOverride Function hs_Tickets_Requesters_Read(Id As Integer) As hs_Tickets_Requesters
#End Region

#Region "hs_UserAlert"
        Public MustOverride Function hs_UserAlert_Add(hsId As Integer, IdContatto As Integer) As Boolean
        Public MustOverride Function hs_UserAlert_Del(Id As Integer) As Boolean
        Public MustOverride Function hs_UserAlert_List(hsId As Integer) As List(Of hs_UserAlert)
        Public MustOverride Function hs_UserAlert_Read(Id As Integer) As hs_UserAlert
#End Region

#Region "HeatingSystem"
        Public MustOverride Function HeatingSystem_Add(IdImpianto As String, _
                                                       Descr As String, _
                                                       Indirizzo As String, _
                                                       Latitude As Decimal, _
                                                       Longitude As Decimal, _
                                                       AltSLM As Integer, _
                                                       UserName As String) As Boolean
        Public MustOverride Function HeatingSystem_Del(hsId As Integer) As Boolean
        Public MustOverride Function HeatingSystem_getTotMap(IdImpianto As String) As Integer
        Public MustOverride Function HeatingSystem_List(IdImpianto As String) As List(Of HeatingSystem)
        Public MustOverride Function HeatingSystem_ListEnabled(IdImpianto As String) As List(Of HeatingSystem)
        Public MustOverride Function HeatingSystem_ListAll() As List(Of HeatingSystem)
        Public MustOverride Function HeatingSystem_Read(hsId As Integer) As HeatingSystem
        Public MustOverride Function HeatingSystem_setMaintenanceMode(hsId As Integer, MaintenanceMode As Boolean) As Boolean
        Public MustOverride Function HeatingSystem_setisOnline(hsId As Integer, isOnline As Boolean) As Boolean
        Public MustOverride Function HeatingSystem_setIwMonitoring(hsId As Integer, IwMonitoringId As Integer, IwMonitoringDes As String) As Boolean
        Public MustOverride Function HeatingSystem_setlastRec(hsId As Integer, lastRec As Date) As Boolean
        Public MustOverride Function HeatingSystem_setNote(hsId As Integer, Note As String, UserName As String) As Boolean
        Public MustOverride Function HeatingSystem_setStatus(hsId As Integer, stato As Integer) As Boolean
        Public MustOverride Function HeatingSystem_setIsEnabled(hsId As Integer, isEnabled As Boolean) As Boolean
        Public MustOverride Function HeatingSystem_Update(hsId As Integer, Descr As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Integer, VPNConnectionId As Integer, MapId As Integer, UserName As String) As Boolean
#End Region

#Region "hs_Controller"
        Public MustOverride Function hs_Controller_Add(hsId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
        Public MustOverride Function hs_Controller_Del(ControllerId As Integer) As Boolean
        Public MustOverride Function hs_Controller_List(hsId As Integer) As List(Of hs_Controller)
        Public MustOverride Function hs_Controller_Read(ControllerId As Integer) As hs_Controller
        Public MustOverride Function hs_Controller_Update(ControllerId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
#End Region
#Region "hs_ControllerDetail"
        Public MustOverride Function hs_ControllerDetail_Add(ControllerId As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
        Public MustOverride Function hs_ControllerDetail_Del(id As Integer) As Boolean
        Public MustOverride Function hs_ControllerDetail_List(ControllerId As Integer) As List(Of hs_ControllerDetail)
        Public MustOverride Function hs_ControllerDetail_Read(id As Integer) As hs_ControllerDetail
        Public MustOverride Function hs_ControllerDetail_Update(id As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
#End Region

#Region "Lux"
        Public MustOverride Function Lux_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function Lux_Del(Id As Integer) As Boolean
        Public MustOverride Function Lux_List(hsId As Integer) As List(Of Lux)
        Public MustOverride Function Lux_Read(Id As Integer) As Lux
        Public MustOverride Function Lux_ReadByAmb(hsId As Integer, IdAmbiente As Integer) As List(Of Lux)
        Public MustOverride Function Lux_ReadByCod(hsId As Integer, Cod As String) As Lux
        Public MustOverride Function Lux_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function Lux_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        Public MustOverride Function Lux_setValue(hsId As Integer, Cod As String, LightON As Boolean, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, CurrentMode As Integer, forcedOn As Boolean, forcedOff As Boolean) As Boolean
        Public MustOverride Function Lux_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        Public MustOverride Function Lux_setLightByCod(hsId As Integer, Cod As String, LightON As Boolean) As Boolean
#End Region
#Region "Lux_replacement_history"
        Public MustOverride Function Lux_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
        Public MustOverride Function Lux_replacement_history_Del(Id As Integer) As Boolean
        Public MustOverride Function Lux_replacement_history_List(ParentId As Integer) As List(Of Lux_replacement_history)
        Public MustOverride Function Lux_replacement_history_Read(Id As Integer) As Lux_replacement_history
        Public MustOverride Function Lux_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
#End Region
#Region "log_Lux"
        Public MustOverride Function log_Lux_Add(hsId As Integer,
                                                 Cod As String,
                                                 Descr As String,
                                                 WorkingTimeCounter As Decimal,
                                                 PowerOnCycleCounter As Decimal,
                                                 stato As Integer,
                                                 LightON As Boolean,
                                                 dtLog As Date) As Integer
        Public MustOverride Function log_Lux_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Lux)
        Public MustOverride Function log_Lux_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Lux)
        Public MustOverride Function log_Lux_logNotSent(hsId As Integer, Cod As String) As log_Lux
        Public MustOverride Function log_Lux_setIsSent(Logid As Integer) As Boolean
#End Region

#Region "LuxM"
        Public MustOverride Function LuxM_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function LuxM_Del(Id As Integer) As Boolean
        Public MustOverride Function LuxM_List(hsId As Integer) As List(Of LuxM)
        Public MustOverride Function LuxM_Read(Id As Integer) As LuxM
        Public MustOverride Function LuxM_ReadByCod(hsId As Integer, Cod As String) As LuxM
        Public MustOverride Function LuxM_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function LuxM_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        Public MustOverride Function LuxM_setValue(hsId As Integer, Cod As String, Voltage As Decimal, Curr As Decimal, EnergyCounter As Decimal, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, Temp As Decimal) As Boolean
        Public MustOverride Function LuxM_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        Public MustOverride Function LuxM_setCurrentMode(Id As Integer, currentMode As Integer) As Boolean
        Public MustOverride Function LuxM_setLightByCod(hsId As Integer, Cod As String, LightON As Boolean) As Boolean
#End Region
#Region "LuxM_replacement_history"
        Public MustOverride Function LuxM_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
        Public MustOverride Function LuxM_replacement_history_Del(Id As Integer) As Boolean
        Public MustOverride Function LuxM_replacement_history_List(ParentId As Integer) As List(Of LuxM_replacement_history)
        Public MustOverride Function LuxM_replacement_history_Read(Id As Integer) As LuxM_replacement_history
        Public MustOverride Function LuxM_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
#End Region
#Region "log_LuxM"
        Public MustOverride Function log_LuxM_Add(hsId As Integer,
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
        Public MustOverride Function log_LuxM_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_LuxM)
        Public MustOverride Function log_LuxM_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_LuxM)
        Public MustOverride Function log_LuxM_logNotSent(hsId As Integer, Cod As String) As log_LuxM
        Public MustOverride Function log_LuxM_setIsSent(Logid As Integer) As Boolean
        Public MustOverride Function log_LuxM_ReadLast(hsId As Integer, Cod As String) As log_LuxM
#End Region

#Region "Psg"
        Public MustOverride Function Psg_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function Psg_Del(Id As Integer) As Boolean
        Public MustOverride Function Psg_List(hsId As Integer) As List(Of Psg)
        Public MustOverride Function Psg_Read(Id As Integer) As Psg
        Public MustOverride Function Psg_ReadByCod(hsId As Integer, Cod As String) As Psg
        Public MustOverride Function Psg_ReadByLux(hsId As Integer, luxCod As String) As List(Of Psg)
        Public MustOverride Function Psg_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function Psg_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        Public MustOverride Function Psg_setValue(hsId As Integer, Cod As String, currentValue As Integer) As Boolean
        Public MustOverride Function Psg_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
#End Region
#Region "Psg_replacement_history"
        Public MustOverride Function Psg_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
        Public MustOverride Function Psg_replacement_history_Del(Id As Integer) As Boolean
        Public MustOverride Function Psg_replacement_history_List(ParentId As Integer) As List(Of Psg_replacement_history)
        Public MustOverride Function Psg_replacement_history_Read(Id As Integer) As Psg_replacement_history
        Public MustOverride Function Psg_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
#End Region
#Region "log_Psg"
        Public MustOverride Function log_Psg_Add(hsId As Integer, _
                                                 Cod As String, _
                                                 Descr As String, _
                                                 currentValue As Integer, _
                                                 stato As Integer, _
                                                 dtLog As Date) As Boolean
        Public MustOverride Function log_Psg_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Psg)
        Public MustOverride Function log_Psg_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Psg)
        Public MustOverride Function log_Psg_logNotSent(hsId As Integer, Cod As String) As log_Psg
        Public MustOverride Function log_Psg_setIsSent(Logid As Integer) As Boolean
        Public MustOverride Function log_Psg_ReadLast(hsId As Integer, Cod As String) As log_Psg

#End Region

#Region "Zrel"
        Public MustOverride Function Zrel_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function Zrel_Del(Id As Integer) As Boolean
        Public MustOverride Function Zrel_List(hsId As Integer) As List(Of Zrel)
        Public MustOverride Function Zrel_Read(Id As Integer) As Zrel
        Public MustOverride Function Zrel_ReadByCod(hsId As Integer, Cod As String) As Zrel
        Public MustOverride Function Zrel_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        Public MustOverride Function Zrel_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        Public MustOverride Function Zrel_setValue(hsId As Integer, Cod As String, LQI As Integer, Temperature As Decimal, MeshParentId As Decimal, CurrentId As Integer) As Boolean
        Public MustOverride Function Zrel_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
#End Region
#Region "Zrel_replacement_history"
        Public MustOverride Function Zrel_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
        Public MustOverride Function Zrel_replacement_history_Del(Id As Integer) As Boolean
        Public MustOverride Function Zrel_replacement_history_List(ParentId As Integer) As List(Of Zrel_replacement_history)
        Public MustOverride Function Zrel_replacement_history_Read(Id As Integer) As Zrel_replacement_history
        Public MustOverride Function Zrel_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
#End Region
#Region "log_Zrel"
        Public MustOverride Function log_Zrel_Add(hsId As Integer, _
                                                 Cod As String, _
                                                 Descr As String, _
                                                 LQI As Integer, _
                                                 Temperature As Decimal, _
                                                 MeshParentId As Integer, _
                                                 CurrentId As Integer, _
                                                 stato As Integer, _
                                                 dtLog As Date) As Boolean
        Public MustOverride Function log_Zrel_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Zrel)
        Public MustOverride Function log_Zrel_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Zrel)
        Public MustOverride Function log_Zrel_logNotSent(hsId As Integer, Cod As String) As log_Zrel
        Public MustOverride Function log_Zrel_setIsSent(Logid As Integer) As Boolean
#End Region

#Region "hs_Cron"
        Public MustOverride Function hs_Cron_Add(hsId As Integer, CronCod As String, CronDescr As String, NoteInterne As String, UserName As String, marcamodello As String, installationDate As Date, CronType As Integer) As Boolean
        Public MustOverride Function hs_Cron_Del(CronId As Integer) As Boolean
        Public MustOverride Function hs_Cron_List(hsId As Integer) As List(Of hs_Cron)
        Public MustOverride Function hs_Cron_ListOnOff(hsid As Integer) As List(Of hs_Cron)
        Public MustOverride Function hs_Cron_Read(CronId As Integer) As hs_Cron
        Public MustOverride Function hs_Cron_ReadByCronCod(hsId As Integer, CronCod As String) As hs_Cron
        Public MustOverride Function hs_Cron_Update(CronId As Integer, CronCod As String, CronDescr As String, NoteInterne As String, UserName As String, marcamodello As String, installationDate As Date, CronType As Integer) As Boolean
        Public MustOverride Function hs_Cron_UpdateMarcamodello(Id As Integer, marcamodello As String, installationDate As Date, UserName As String) As Boolean
        Public MustOverride Function hs_Cron_setStatus(hsId As Integer, CronCod As String, SetPoint As Decimal, stato As Integer) As Boolean
        Public MustOverride Function hs_Cron_setGeoLocation(CronId As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
#End Region
#Region "hs_Cron_replacement_history"
        Public MustOverride Function hs_Cron_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
        Public MustOverride Function hs_Cron_replacement_history_Del(Id As Integer) As Boolean
        Public MustOverride Function hs_Cron_replacement_history_List(ParentId As Integer) As List(Of hs_Cron_replacement_history)
        Public MustOverride Function hs_Cron_replacement_history_Read(Id As Integer) As hs_Cron_replacement_history
        Public MustOverride Function hs_Cron_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
#End Region
#Region "hs_amb_Profile"
        Public MustOverride Function hs_amb_Profile_Add(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData As Decimal()) As Boolean
        Public MustOverride Function hs_amb_Profile_Clear(CronId As Integer) As Boolean
        Public MustOverride Function hs_amb_Profile_List(CronId As Integer, ProfileY As Integer) As List(Of hs_amb_Profile)
        Public MustOverride Function hs_amb_Profile_Read(CronId As Integer, ProfileY As Integer, ProfileNr As Integer) As hs_amb_Profile
        Public MustOverride Function hs_amb_Profile_Update(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData As Decimal()) As Boolean
#End Region
#Region "hs_amb_Profile_Descr"
        Public MustOverride Function hs_amb_Profile_Descr_Add(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
        Public MustOverride Function hs_amb_Profile_Descr_Del(CronId As Integer, ProfileNr As Integer) As Boolean
        Public MustOverride Function hs_amb_Profile_Descr_List(CronId As Integer) As List(Of hs_amb_Profile_Descr)
        Public MustOverride Function hs_amb_Profile_Descr_Read(CronId As Integer, ProfileNr As Integer) As hs_amb_Profile_Descr
        Public MustOverride Function hs_amb_Profile_Descr_Update(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
#End Region
#Region "hs_amb_Profile_Tasks"
        Public MustOverride Function hs_amb_Profile_Tasks_Add(CronId As Integer, ProfileNr As Integer, Subject As String, StartDate As Date, EndDate As Date, RecurrencePattern As String, ExceptionAppointments As String, yearsRepeatable As Boolean) As Boolean
        Public MustOverride Function hs_amb_Profile_Tasks_Del(CronId As Integer, ProfileNr As Integer) As Boolean
        Public MustOverride Function hs_amb_Profile_Tasks_List(CronId As Integer) As List(Of hs_amb_Profile_Tasks)
        Public MustOverride Function hs_amb_Profile_Tasks_ListAll() As List(Of hs_amb_Profile_Tasks)
        Public MustOverride Function hs_amb_Profile_Tasks_Read(TaskId As Integer) As hs_amb_Profile_Tasks
        Public MustOverride Function hs_amb_Profile_Tasks_Update(TaskId As Integer, ProfileNr As Integer, Subject As String, StartDate As Date, EndDate As Date, RecurrencePattern As String, ExceptionAppointments As String, yearsRepeatable As Boolean) As Boolean
#End Region
#Region "hs_amb_Calendar"
        Public MustOverride Function hs_amb_Calendar_Add(CronId As Integer, Calyear As Integer, Calmonth As Integer, monthData As Integer()) As Boolean
        Public MustOverride Function hs_amb_Calendar_Clear(CronId) As Boolean
        Public MustOverride Function hs_amb_Calendar_Read(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
        Public MustOverride Function hs_amb_Calendar_Update(CronId As Integer, Calyear As Integer, Calmonth As Integer, TasksForDesired As Integer(), DesiredMonthData As Integer()) As Boolean
        Public MustOverride Function hs_amb_Calendar_UpdateDesired(CronId As Integer, Calyear As Integer, Calmonth As Integer, DesiredMonthData As Integer()) As Boolean
        Public MustOverride Function hs_amb_Calendar_UpdateReal(CronId As Integer, Calyear As Integer, Calmonth As Integer, RealMonthData As Integer()) As Boolean
#End Region
#Region "log_hs_Cron"
        Public MustOverride Function log_hs_Cron_Add(hsId As Integer, SetPoint As Decimal, CronCod As String, CronDescr As String, stato As Integer, dtLog As Date) As Boolean
        Public MustOverride Function log_hs_Cron_List(hsId As Integer, CronCod As String, fromDate As Date, toDate As Date) As List(Of log_hs_Cron)
        Public MustOverride Function log_hs_Cron_ListPaged(hsId As Integer, CronCod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_hs_Cron)
        Public MustOverride Function log_hs_Cron_ListAll(hsId As Integer, fromDate As Date, toDate As Date) As List(Of log_hs_Cron)
        Public MustOverride Function log_hs_Cron_logNotSent(hsId As Integer, CronCod As String) As log_hs_Cron
        Public MustOverride Function log_hs_Cron_setIsSent(Logid As Integer) As Boolean
        Public MustOverride Function log_hs_Cron_ReadLast(hsId As Integer, CronCod As String) As log_hs_Cron
#End Region
#Region "LuxM_last"
        Public MustOverride Function LuxM_last_Add(hsId As Integer, Cod As String, lastLog As Integer, lastdtLog As Date) As Boolean
        Public MustOverride Function LuxM_last_Upd(hsId As Integer, Cod As String, lastLog As Integer, lastdtLog As Date) As Boolean
        Public MustOverride Function LuxM_last_Read(hsId As Integer, Cod As String) As LuxM_last
#End Region
#Region "Lux_last"
        Public MustOverride Function Lux_last_Add(hsId As Integer, Cod As String, lastLog As Integer, lastdtLog As Date) As Boolean
        Public MustOverride Function Lux_last_Upd(hsId As Integer, Cod As String, lastLog As Integer, lastdtLog As Date) As Boolean
        Public MustOverride Function Lux_last_Read(hsId As Integer, Cod As String) As Lux_last
#End Region
#Region "Ambienti"
        'Public MustOverride Function Ambienti_Add(IdAmbiente As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        ' Public MustOverride Function Ambienti_Del(Id As Integer) As Boolean
        Public MustOverride Function Ambienti_List(hsId As Integer) As List(Of Ambienti)
        Public MustOverride Function Ambienti_Read(IdAmbiente As Integer) As Ambienti
        'Public MustOverride Function Ambienti_ReadByCod(IdAmbiente As Integer, Cod As String) As Ambienti
        'Public MustOverride Function Ambienti_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
        'Public MustOverride Function Ambienti_setStatus(IdAmbiente As Integer, Cod As String, stato As Integer) As Boolean
        'Public MustOverride Function Ambienti_setValue(IdAmbiente As Integer, Cod As String, currentValue As Integer) As Boolean
        'Public MustOverride Function Ambienti_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
#End Region
    End Class
End Namespace
