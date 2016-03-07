Imports System
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Net.Sockets
Imports System.Linq

Imports SCP.BLL
Imports System.IO

Namespace SCP.DAL

    Public Class SQLDataAccess
        Inherits DataAccess

#Region "BASE CLASS IMPLEMENTATION"

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sqlCmd"></param>
        ''' <param name="cmdType"></param>
        ''' <param name="cmdText"></param>
        ''' <remarks></remarks>
        Private Sub SetCommandType(sqlCmd As SqlCommand, cmdType As CommandType, cmdText As String)
            sqlCmd.CommandType = cmdType
            sqlCmd.CommandText = cmdText
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sqlCmd"></param>
        ''' <param name="paramId"></param>
        ''' <param name="sqlType"></param>
        ''' <param name="paramSize"></param>
        ''' <param name="paramDirection"></param>
        ''' <param name="paramvalue"></param>
        ''' <remarks></remarks>
        Private Sub AddParamToSQLCmd(sqlCmd As SqlCommand, _
                                     paramId As String, _
                                     sqlType As SqlDbType, _
                                     paramSize As Integer, _
                                     paramDirection As ParameterDirection, _
                                     paramvalue As Object)

            If sqlCmd Is Nothing Then
                Throw New ArgumentNullException("sqlCmd")
            End If

            If paramId = String.Empty Then
                Throw New ArgumentOutOfRangeException("paramId")
            End If

            Dim newSqlParam As SqlParameter = New SqlParameter()
            newSqlParam.ParameterName = paramId
            newSqlParam.SqlDbType = sqlType
            newSqlParam.Direction = paramDirection

            If paramSize > 0 Then
                newSqlParam.Size = paramSize
            End If

            If Not paramvalue Is Nothing Then
                newSqlParam.Value = paramvalue
            End If

            sqlCmd.Parameters.Add(newSqlParam)

        End Sub
        Private Sub scriviLog(wData As String)
            Try
                'LogPath
                Dim LogPath As String = ConfigurationManager.AppSettings("LogPath")
                Dim IFName As String = Path.Combine(LogPath, "cyassrvLog.txt")
                Dim sw As StreamWriter
                If File.Exists(IFName) = False Then
                    sw = File.CreateText(IFName)
                Else
                    sw = File.AppendText(IFName)
                End If
                sw.WriteLine(Now() & ";" & wData)
                sw.Flush()
                sw.Close()
            Catch ex As Exception

            End Try

        End Sub
#End Region

#Region "aspnetroles"
        Public Overrides Function aspnetroles_Add(RoleName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlCmd As SqlCommand = New SqlCommand()
            SetCommandType(sqlCmd, CommandType.StoredProcedure, "aspnet_Roles_CreateRole")
            AddParamToSQLCmd(sqlCmd, "@ApplicationName", SqlDbType.NVarChar, 256, ParameterDirection.Input, "/")
            AddParamToSQLCmd(sqlCmd, "@RoleName", SqlDbType.NVarChar, 256, ParameterDirection.Input, RoleName)

            AddParamToSQLCmd(sqlCmd, "@ErrorCode", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Try
                sqlCmd.ExecuteNonQuery()
                If Convert.ToInt32(sqlCmd.Parameters("@ErrorCode").Value) = 0 Then
                    retVal = True
                End If
            Catch ex As Exception
                retVal = False
                scriviLog("aspnetroles_Add " & ex.Message)
            End Try
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function aspnetroles_Del(RoleName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlCmd As SqlCommand = New SqlCommand()
            SetCommandType(sqlCmd, CommandType.StoredProcedure, "aspnet_Roles_DeleteRole")
            AddParamToSQLCmd(sqlCmd, "@ApplicationName", SqlDbType.NVarChar, 256, ParameterDirection.Input, "/")
            AddParamToSQLCmd(sqlCmd, "@RoleName", SqlDbType.NVarChar, 256, ParameterDirection.Input, RoleName)
            AddParamToSQLCmd(sqlCmd, "@DeleteOnlyIfRoleIsEmpty", SqlDbType.Bit, 256, ParameterDirection.Input, 0)

            AddParamToSQLCmd(sqlCmd, "@ErrorCode", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Try
                sqlCmd.ExecuteNonQuery()
                If Convert.ToInt32(sqlCmd.Parameters("@ErrorCode").Value) = 0 Then
                    retVal = True
                End If
            Catch ex As Exception
                retVal = False
                scriviLog("aspnetroles_Del " & ex.Message)
            End Try
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function aspnetroles_List() As List(Of aspnetroles)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT  aspnet_Roles.[ApplicationId]"
            sqlString += "      ,aspnet_Roles.[RoleId]"
            sqlString += "      ,aspnet_Roles.[RoleName]"
            sqlString += "      ,aspnet_Roles.[Description]"
            sqlString += " ,(select count(*) "
            sqlString += "   from aspnet_UsersInRoles "
            sqlString += "   where aspnet_UsersInRoles.[RoleId]=[aspnet_Roles].[RoleId]) as countUser"
            sqlString += "  FROM [aspnet_Roles]"
            sqlString += "  inner join aspnet_Applications on aspnet_Applications.ApplicationId=aspnet_Roles.ApplicationId"
            sqlString += "  where aspnet_Applications.[ApplicationName]='/'"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _list As New List(Of aspnetroles)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    Dim _ApplicationId As String = String.Empty
                    Dim _RoleId As String = String.Empty
                    Dim _RoleName As String = String.Empty
                    Dim _Description As String = String.Empty
                    Dim _countUser As Integer = 0

                    If Not IsDBNull(reader("ApplicationId")) Then
                        _ApplicationId = reader("ApplicationId").ToString
                    End If
                    If Not IsDBNull(reader("RoleId")) Then
                        _RoleId = reader("RoleId").ToString
                    End If
                    If Not IsDBNull(reader("RoleName")) Then
                        _RoleName = CStr(reader("RoleName"))
                    End If
                    If Not IsDBNull(reader("Description")) Then
                        _Description = CStr(reader("Description"))
                    End If
                    If Not IsDBNull(reader("countUser")) Then
                        _countUser = CInt(reader("countUser"))
                    End If

                    Dim _obj As aspnetroles = New aspnetroles(_ApplicationId, _RoleId, _RoleName, _Description, _countUser)
                    _list.Add(_obj)
                Loop
            Catch ex As Exception
                scriviLog("aspnetroles_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function aspnetroles_ListActive(IdImpianto As String) As List(Of aspnetroles)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT  aspnet_Roles.[ApplicationId]"
            sqlString += " ,aspnet_Roles.[RoleId]"
            sqlString += " ,aspnet_Roles.[RoleName]"
            sqlString += " ,aspnet_Roles.[Description]"
            sqlString += " , ("
            sqlString += "             SELECT  count(*) as tot"
            sqlString += "             FROM Impianti "
            sqlString += "             INNER JOIN Impianti_Users ON Impianti.IdImpianto = Impianti_Users.IdImpianto "
            sqlString += "             RIGHT OUTER JOIN aspnet_Users "
            sqlString += "             INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += "             INNER JOIN aspnet_Roles as AAA ON aspnet_UsersInRoles.RoleId = AAA.RoleId "
            sqlString += "             INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId "
            sqlString += "             ON Impianti_Users.UserId = aspnet_Users.UserId"
            sqlString += "             where AAA.RoleName=aspnet_Roles.[RoleName]"
            sqlString += "             and Impianti.IdImpianto=@IdImpianto"
            sqlString += "             and (aspnet_Membership.IsApproved=1 and aspnet_Membership.IsLockedOut=0)"
            sqlString += " ) as tot"
            sqlString += " FROM [aspnet_Roles]"
            sqlString += " where aspnet_Roles.[RoleName] in('Supervisors','Mantainers','Endusers')"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Dim _list As New List(Of aspnetroles)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    Dim _ApplicationId As String = String.Empty
                    Dim _RoleId As String = String.Empty
                    Dim _RoleName As String = String.Empty
                    Dim _Description As String = String.Empty
                    Dim _countUser As Integer = 0

                    If Not IsDBNull(reader("ApplicationId")) Then
                        _ApplicationId = reader("ApplicationId").ToString
                    End If
                    If Not IsDBNull(reader("RoleId")) Then
                        _RoleId = reader("RoleId").ToString
                    End If
                    If Not IsDBNull(reader("RoleName")) Then
                        _RoleName = CStr(reader("RoleName"))
                    End If
                    If Not IsDBNull(reader("Description")) Then
                        _Description = CStr(reader("Description"))
                    End If
                    If Not IsDBNull(reader("tot")) Then
                        _countUser = CInt(reader("tot"))
                    End If

                    Dim _obj As aspnetroles = New aspnetroles(_ApplicationId, _RoleId, _RoleName, _Description, _countUser)
                    _list.Add(_obj)
                Loop
            Catch ex As Exception
                scriviLog("aspnetroles_ListActive " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function aspnetroles_Update(RoleId As String, RoleName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update aspnet_Roles"
            sqlString += " set RoleName=@RoleName"
            sqlString += " where [RoleId]=@RoleId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(RoleId)
            AddParamToSQLCmd(sqlCmd, "@RoleId", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)
            AddParamToSQLCmd(sqlCmd, "@RoleName", SqlDbType.NVarChar, 255, ParameterDirection.Input, RoleName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                Catch ex As Exception
                    scriviLog("aspnetroles_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "aspnetUsers"
        Private Function prepareRecord_aspnetusers(reader As SqlDataReader) As aspnetUsers
            Dim UserName As String = ""
            If Not IsDBNull(reader("UserName")) Then
                UserName = reader("UserName")
            End If

            Dim Email As String = ""
            If Not IsDBNull(reader("Email")) Then
                Email = reader("Email")
            End If

            Dim Password As String = ""
            If Not IsDBNull(reader("Password")) Then
                Password = reader("Password")
            End If

            Dim Comment As String = ""
            If Not IsDBNull(reader("Comment")) Then
                Comment = reader("Comment")
            End If

            Dim RoleName As String = String.Empty
            If Not IsDBNull(reader("RoleName")) Then
                RoleName = reader("RoleName")
            End If

            Dim IsApproved As Boolean = False
            If Not IsDBNull(reader("IsApproved")) Then
                IsApproved = CBool(reader("IsApproved"))
            End If

            Dim CreateDate As Date = CDate("01/01/1900")
            If Not IsDBNull(reader("CreateDate")) Then
                CreateDate = CDate(reader("CreateDate"))
            End If

            Dim LastLoginDate As Date = CDate("01/01/1900")
            If Not IsDBNull(reader("LastLoginDate")) Then
                LastLoginDate = CDate(reader("LastLoginDate"))
            End If

            Dim LastActivityDate As Date = CDate("01/01/1900")
            If Not IsDBNull(reader("LastActivityDate")) Then
                LastActivityDate = CDate(reader("LastActivityDate"))
            End If

            Dim LastPasswordChangedDate As Date = CDate("01/01/1900")
            If Not IsDBNull(reader("LastPasswordChangedDate")) Then
                LastPasswordChangedDate = CDate(reader("LastPasswordChangedDate"))
            End If

            Dim UserId As String = ""
            If Not IsDBNull(reader("UserId")) Then
                UserId = reader("UserId").ToString
            End If

            Dim IsLockedOut As Boolean = False
            If Not IsDBNull(reader("IsLockedOut")) Then
                IsLockedOut = CBool(reader("IsLockedOut"))
            End If

            Dim LastLockoutDate As Date = CDate("01/01/1900")
            If Not IsDBNull(reader("LastLockoutDate")) Then
                LastLockoutDate = CDate(reader("LastLockoutDate"))
            End If

            Dim _totImpianti As Integer = 0

            Dim _obj As aspnetUsers = New aspnetUsers(UserName, _
                                                      Email, _
                                                      Password, _
                                                      Comment, _
                                                      RoleName, _
                                                      IsApproved, _
                                                      CreateDate, _
                                                      LastLoginDate, _
                                                      LastActivityDate, _
                                                      LastPasswordChangedDate, _
                                                      UserId, _
                                                      IsLockedOut, _
                                                      LastLockoutDate, _
                                                      _totImpianti)

            Return _obj
        End Function
        Public Overrides Function aspnetUsers_approveUser(userId As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [aspnet_Membership]"
            sqlString += " set [IsApproved]=1"
            sqlString += " where UserId=@UserId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(userId)
            AddParamToSQLCmd(sqlCmd, "@UserId", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("aspnetUsers_approveUser " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function aspnetUsers_disapproveUser(userId As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [aspnet_Membership]"
            sqlString += " set [IsApproved]=0"
            sqlString += " where UserId=@UserId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(userId)
            AddParamToSQLCmd(sqlCmd, "@UserId", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("aspnetUsers_approveUser " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function aspnetUsers_lockUser(UserId As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [aspnet_Membership]"
            sqlString += " set [IsLockedOut]=1"
            sqlString += " where UserId=@UserId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(UserId)
            AddParamToSQLCmd(sqlCmd, "@UserId", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("aspnetUsers_lockUser " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function aspnetUsers_GetUserByUserName(UserName As String) As BLL.aspnetUsers
            Dim sqlString As String = String.Empty
            sqlString = "SELECT aspnet_Users.UserName"
            sqlString += " , aspnet_Membership.Email"
            sqlString += " , aspnet_Membership.Password"
            sqlString += " , aspnet_Membership.Comment"
            sqlString += " , aspnet_Roles.RoleName"
            sqlString += " , aspnet_Membership.IsApproved"
            sqlString += " , aspnet_Membership.CreateDate"
            sqlString += " , aspnet_Membership.LastLoginDate"
            sqlString += " , aspnet_Users.LastActivityDate"
            sqlString += " , aspnet_Membership.LastPasswordChangedDate"
            sqlString += " , aspnet_Users.UserId"
            sqlString += " , aspnet_Membership.IsLockedOut"
            sqlString += " , aspnet_Membership.LastLockoutDate"
            sqlString += " FROM aspnet_Users "
            sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"
            sqlString += " WHERE (aspnet_Users.UserName = @UserName)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Dim _list As New List(Of aspnetUsers)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_aspnetusers(reader))
                Loop
            Catch ex As Exception
                scriviLog("aspnetUsers_GetUserByUserName " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function aspnetUsers_GetActiveUsersByRoleName(RoleName As String) As List(Of aspnetUsers)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT aspnet_Users.UserName"
            sqlString += " , aspnet_Membership.Email"
            sqlString += " , aspnet_Membership.Password"
            sqlString += " , aspnet_Membership.Comment"
            sqlString += " , aspnet_Roles.RoleName"
            sqlString += " , aspnet_Membership.IsApproved"
            sqlString += " , aspnet_Membership.CreateDate"
            sqlString += " , aspnet_Membership.LastLoginDate"
            sqlString += " , aspnet_Users.LastActivityDate"
            sqlString += " , aspnet_Membership.LastPasswordChangedDate"
            sqlString += " , aspnet_Users.UserId"
            sqlString += " , aspnet_Membership.IsLockedOut"
            sqlString += " , aspnet_Membership.LastLockoutDate"
            sqlString += " FROM aspnet_Users "
            sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"
            sqlString += " WHERE (aspnet_Roles.RoleName = @RoleName)"
            sqlString += "  and (aspnet_Membership.IsApproved=1 and aspnet_Membership.IsLockedOut=0)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@RoleName", SqlDbType.NVarChar, 256, ParameterDirection.Input, RoleName)

            Dim UsersList As New List(Of aspnetUsers)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    UsersList.Add(prepareRecord_aspnetusers(reader))
                Loop
            Catch ex As Exception
                scriviLog("aspnetUsers_GetActiveUsersByRoleName " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If

            Return UsersList
        End Function

        Public Overrides Function aspnetUsers_GetTotActiveUsersByRoleName(RoleName As String) As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "SELECT  count(*) as tot"
            sqlString += " FROM aspnet_Users "
            sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"
            sqlString += " WHERE (aspnet_Roles.RoleName = @RoleName )"
            sqlString += "  and (aspnet_Membership.IsApproved=1 and aspnet_Membership.IsLockedOut=0)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@RoleName", SqlDbType.NVarChar, 256, ParameterDirection.Input, RoleName)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("tot")) Then
                        retVal = CBool(reader("tot"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("aspnetUsers_GetTotActiveUsersByRoleName " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If

            Return retVal
        End Function

        Public Overrides Function aspnetUsers_GetUsersByRoleName(RoleName As String, searchString As String, IdImpianto As String) As System.Collections.Generic.List(Of BLL.aspnetUsers)
            'Dim sqlString As String = String.Empty
            'sqlString = "SELECT aspnet_Users.UserName"
            'sqlString += " , aspnet_Membership.Email"
            'sqlString += " , aspnet_Membership.Password"
            'sqlString += " , aspnet_Membership.Comment"
            'sqlString += " , aspnet_Roles.RoleName"
            'sqlString += " , aspnet_Membership.IsApproved"
            'sqlString += " , aspnet_Membership.CreateDate"
            'sqlString += " , aspnet_Membership.LastLoginDate"
            'sqlString += " , aspnet_Users.LastActivityDate"
            'sqlString += " , aspnet_Membership.LastPasswordChangedDate"
            'sqlString += " , aspnet_Users.UserId"
            'sqlString += " , aspnet_Membership.IsLockedOut"
            'sqlString += " , aspnet_Membership.LastLockoutDate"
            'sqlString += " FROM aspnet_Users "
            'sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            'sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            'sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"
            'sqlString += " WHERE (aspnet_Roles.RoleName = @RoleName)"
            'If Not String.IsNullOrEmpty(searchString) Then
            '    sqlString += "	 and (aspnet_Users.UserName like @SearchString or"
            '    sqlString += "	     aspnet_Roles.RoleName like @SearchString)"
            'End If
            'Dim sqlCmd As SqlCommand = New SqlCommand()
            'sqlCmd.CommandText = sqlString

            Dim sqlCmd As SqlCommand = New SqlCommand()

            SetCommandType(sqlCmd, CommandType.StoredProcedure, "[App_GetUsers]")

            AddParamToSQLCmd(sqlCmd, "@RoleName", SqlDbType.NVarChar, 256, ParameterDirection.Input, RoleName)
            AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(searchString) & "%")
            If Not String.IsNullOrEmpty(IdImpianto) Then
                Dim ImpiantoGuid As Guid = New Guid(IdImpianto)
                AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, ImpiantoGuid)
            Else
                AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, DBNull.Value)
            End If


            Dim UsersList As New List(Of aspnetUsers)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    UsersList.Add(prepareRecord_aspnetusers(reader))
                Loop
            Catch ex As Exception
                scriviLog("aspnetUsers_GetUsersByRoleName " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If

            Return UsersList
        End Function

        Public Overrides Function aspnetUsers_GetUsersAll() As System.Collections.Generic.List(Of BLL.aspnetUsers)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT aspnet_Users.UserName"
            sqlString += " , aspnet_Membership.Email"
            sqlString += " , aspnet_Membership.Password"
            sqlString += " , aspnet_Membership.Comment"
            sqlString += " , aspnet_Roles.RoleName"
            sqlString += " , aspnet_Membership.IsApproved"
            sqlString += " , aspnet_Membership.CreateDate"
            sqlString += " , aspnet_Membership.LastLoginDate"
            sqlString += " , aspnet_Users.LastActivityDate"
            sqlString += " , aspnet_Membership.LastPasswordChangedDate"
            sqlString += " , aspnet_Users.UserId"
            sqlString += " , aspnet_Membership.IsLockedOut"
            sqlString += " , aspnet_Membership.LastLockoutDate"
            sqlString += " FROM aspnet_Users "
            sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim UsersList As New List(Of aspnetUsers)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    UsersList.Add(prepareRecord_aspnetusers(reader))
                Loop
            Catch ex As Exception
                scriviLog("aspnetUsers_GetUsersAll " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If

            Return UsersList
        End Function

        Public Overrides Function aspnetUsers_List(id As Integer, SearchString As String) As List(Of aspnetUsers)
            Dim sqlString As String = String.Empty
            sqlString = "select RowNumber"
            sqlString += " , UserName"
            sqlString += " , Email"
            sqlString += " , Password"
            sqlString += " , Comment"
            sqlString += " , RoleName"
            sqlString += " , IsApproved"
            sqlString += " , CreateDate"
            sqlString += " , LastLoginDate"
            sqlString += " , LastActivityDate"
            sqlString += " , LastPasswordChangedDate"
            sqlString += " , UserId"
            sqlString += " , IsLockedOut"
            sqlString += " , LastLockoutDate"
            sqlString += " from "
            sqlString += " ( select row_number() over(order by aspnet_Users.UserName) as RowNumber"
            sqlString += " , aspnet_Users.UserName"
            sqlString += " , aspnet_Membership.Email"
            sqlString += " , aspnet_Membership.Password"
            sqlString += " , aspnet_Membership.Comment"
            sqlString += " , aspnet_Roles.RoleName"
            sqlString += " , aspnet_Membership.IsApproved"
            sqlString += " , aspnet_Membership.CreateDate"
            sqlString += " , aspnet_Membership.LastLoginDate"
            sqlString += " , aspnet_Users.LastActivityDate"
            sqlString += " , aspnet_Membership.LastPasswordChangedDate"
            sqlString += " , aspnet_Users.UserId"
            sqlString += " , aspnet_Membership.IsLockedOut"
            sqlString += " , aspnet_Membership.LastLockoutDate"
            sqlString += " FROM aspnet_Users "
            sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"
            sqlString += "	 where aspnet_Users.UserName like @SearchString or"
            sqlString += "	       aspnet_Roles.RoleName like @SearchString) as Users"
            sqlString += " where RowNumber betWeen @first and @last"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(SearchString) & "%")
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, id + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, id + 10)

            Dim UsersList As New List(Of aspnetUsers)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    UsersList.Add(prepareRecord_aspnetusers(reader))
                Loop
            Catch ex As Exception
                scriviLog("aspnetUsers_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If

            Return UsersList
        End Function

        Public Overrides Function aspnetUsers_Read(UserId As String) As aspnetUsers
            Dim sqlString As String = String.Empty
            sqlString = "SELECT aspnet_Users.UserName"
            sqlString += " , aspnet_Membership.Email"
            sqlString += " , aspnet_Membership.Password"
            sqlString += " , aspnet_Membership.Comment"
            sqlString += " , aspnet_Roles.RoleName"
            sqlString += " , aspnet_Membership.IsApproved"
            sqlString += " , aspnet_Membership.CreateDate"
            sqlString += " , aspnet_Membership.LastLoginDate"
            sqlString += " , aspnet_Users.LastActivityDate"
            sqlString += " , aspnet_Membership.LastPasswordChangedDate"
            sqlString += " , aspnet_Users.UserId"
            sqlString += " , aspnet_Membership.IsLockedOut"
            sqlString += " , aspnet_Membership.LastLockoutDate"
            sqlString += " FROM aspnet_Users "
            sqlString += " INNER JOIN aspnet_UsersInRoles ON aspnet_Users.UserId = aspnet_UsersInRoles.UserId "
            sqlString += " INNER JOIN aspnet_Roles ON aspnet_UsersInRoles.RoleId = aspnet_Roles.RoleId "
            sqlString += " INNER JOIN aspnet_Membership ON aspnet_Users.UserId = aspnet_Membership.UserId"
            sqlString += " WHERE (aspnet_Users.UserId = @UserId)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim UserIdGuid As Guid = New Guid(UserId)
            AddParamToSQLCmd(sqlCmd, "@UserId", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, UserIdGuid)

            Dim _list As New List(Of aspnetUsers)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_aspnetusers(reader))
                Loop
            Catch ex As Exception
                scriviLog("aspnetUsers_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If

        End Function

        Public Overrides Function aspnetUsers_whoIsOnline() As List(Of String)
            Dim sqlCmd As SqlCommand = New SqlCommand()
            SetCommandType(sqlCmd, CommandType.StoredProcedure, "WhoIsOnline")
            Dim d As Date = DateAdd(DateInterval.Minute, -10, Now())
            Dim sd As String = d.Year.ToString & "-" & _
                d.Month.ToString.PadLeft(2, "0") & "-" & _
                d.Day.ToString.PadLeft(2, "0") & Space(1) & _
                d.Hour.ToString.PadLeft(2, "0") & ":" & _
                d.Minute.ToString.PadLeft(2, "0") & ":" & _
                d.Second.ToString.PadLeft(2, "0")
            AddParamToSQLCmd(sqlCmd, "@TimeWindow", SqlDbType.NVarChar, 50, ParameterDirection.Input, sd)

            Dim _List As New List(Of String)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim _UserName As String = String.Empty
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    If Not IsDBNull(reader("UserName")) Then
                        _UserName = reader("UserName")
                    End If
                Loop
                _List.Add(_UserName)
            Catch ex As Exception
                scriviLog("aspnetUsers_whoIsOnline " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "UserMenu"
        Private Function prepareRecord_UserMenu(reader As SqlDataReader) As UserMenu
            Dim _IdVoceMenu As Integer = 0
            Dim _DescrVoce As String = String.Empty
            Dim _IdAzione As Integer = 0
            Dim _URLAzione As String = String.Empty
            Dim _IdVocePadre As Integer = 0
            Dim mainPage As String = String.Empty

            If Not IsDBNull(reader("IdVoceMenu")) Then
                _IdVoceMenu = CInt(reader("IdVoceMenu"))
            End If
            If Not IsDBNull(reader("DescrVoce")) Then
                _DescrVoce = CStr(reader("DescrVoce"))
            End If
            If Not IsDBNull(reader("IdAzione")) Then
                _IdAzione = CInt(reader("IdAzione"))
            End If
            If Not IsDBNull(reader("IdVocePadre")) Then
                _IdVocePadre = CInt(reader("IdVocePadre"))
            End If
            If Not IsDBNull(reader("URLAzione")) Then
                _URLAzione = CStr(reader("URLAzione"))
            End If
            If Not IsDBNull(reader("mainPage")) Then
                mainPage = CStr(reader("mainPage"))
            End If
            Dim obj As UserMenu = New UserMenu(_IdVoceMenu, _DescrVoce, _IdAzione, _URLAzione, _IdVocePadre, mainPage)
            Return obj
        End Function

        Public Overrides Function UserMenu_List(UserName As String) As System.Collections.Generic.List(Of BLL.UserMenu)
            Dim sqlstring As String = ""
            sqlstring = "SELECT VociMenu.IdVoceMenu, VociMenu.DescrVoce,VociMenu.IdVocePadre, Azioni.IdAzione,Azioni.URLAzione"
            sqlstring += " ,MenuUtente.mainPage"
            sqlstring += " FROM MenuUtente"
            sqlstring += " INNER JOIN Menu ON MenuUtente.IdMenu = Menu.IdMenu"
            sqlstring += " INNER JOIN VociMenu ON Menu.IdMenu = VociMenu.IdMenu"
            sqlstring += " LEFT OUTER JOIN Azioni ON VociMenu.IdAzione = Azioni.IdAzione"
            sqlstring += " WHERE MenuUtente.UserName = @UserName"
            sqlstring += " ORDER BY VociMenu.IdMenu, VociMenu.Posizione"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlstring
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Dim UserMenuList As New List(Of UserMenu)()

            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    UserMenuList.Add(prepareRecord_UserMenu(reader))
                Loop
            Catch ex As Exception
                scriviLog("UserMenu_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            Return UserMenuList
        End Function

        Public Overrides Function UserMenu_ListFather(UserName As String) As System.Collections.Generic.List(Of BLL.UserMenu)
            Dim sqlstring As String = String.Empty
            sqlstring = "SELECT VociMenu.IdVoceMenu, VociMenu.DescrVoce,VociMenu.IdVocePadre, Azioni.IdAzione,Azioni.URLAzione"
            sqlstring += " ,MenuUtente.mainPage"
            sqlstring += " FROM MenuUtente"
            sqlstring += " INNER JOIN Menu ON MenuUtente.IdMenu = Menu.IdMenu"
            sqlstring += " INNER JOIN VociMenu ON Menu.IdMenu = VociMenu.IdMenu"
            sqlstring += " LEFT OUTER JOIN Azioni ON VociMenu.IdAzione = Azioni.IdAzione"
            sqlstring += " WHERE MenuUtente.UserName = @UserName"
            sqlstring += "   and VociMenu.IdVocePadre=0"
            sqlstring += " ORDER BY VociMenu.IdMenu, VociMenu.Posizione"
            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlstring
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Dim UserMenuList As New List(Of UserMenu)()

            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    UserMenuList.Add(prepareRecord_UserMenu(reader))
                Loop
            Catch ex As Exception
                scriviLog("UserMenu_ListFather " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            Return UserMenuList
        End Function

        Public Overrides Function UserMenu_ListChildren(UserName As String, IdVocePadre As Integer) As System.Collections.Generic.List(Of BLL.UserMenu)
            Dim sqlstring As String = ""

            sqlstring = "SELECT VociMenu.IdVoceMenu, VociMenu.DescrVoce,VociMenu.IdVocePadre, Azioni.IdAzione,Azioni.URLAzione"
            sqlstring += " ,MenuUtente.mainPage"
            sqlstring += " FROM MenuUtente"
            sqlstring += " INNER JOIN Menu ON MenuUtente.IdMenu = Menu.IdMenu"
            sqlstring += " INNER JOIN VociMenu ON Menu.IdMenu = VociMenu.IdMenu"
            sqlstring += " LEFT OUTER JOIN Azioni ON VociMenu.IdAzione = Azioni.IdAzione"
            sqlstring += " WHERE MenuUtente.UserName = @UserName"
            sqlstring += "   and VociMenu.IdVocePadre=@IdVocePadre"
            sqlstring += " ORDER BY VociMenu.IdMenu, VociMenu.Posizione"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlstring
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@IdVocePadre", SqlDbType.Int, 0, ParameterDirection.Input, IdVocePadre)

            Dim UserMenuList As New List(Of UserMenu)()

            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    UserMenuList.Add(prepareRecord_UserMenu(reader))
                Loop
            Catch ex As Exception
                scriviLog("UserMenu_ListChildren " & ex.Message)

            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            Return UserMenuList
        End Function
#End Region

#Region "aspnetusersstat"
        Public Overrides Function aspnetusersstat_Read() As aspnetusersstat
            Dim sqlString As String = String.Empty
            sqlString = "select"
            sqlString += " (select count(*) from aspnet_Membership) as totUser"
            sqlString += " ,(select count(*) from aspnet_Membership where IsApproved=1) as Approved"
            sqlString += " ,(select count(*) from aspnet_Membership where IsApproved=0) as notApproved"
            sqlString += " ,(select count(*) from aspnet_Membership where IsLockedOut=1) as Locked"
            sqlString += " ,(select count(*) from aspnet_Membership where IsLockedOut=0) as notLocked"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _list As New List(Of aspnetusersstat)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    Dim _totUser As Integer = 0
                    Dim _Approved As Integer = 0
                    Dim _notApproved As Integer = 0
                    Dim _Locked As Integer = 0
                    Dim _notLocked As Integer = 0

                    If Not IsDBNull(reader("totUser")) Then
                        _totUser = CInt(reader("totUser"))
                    End If
                    If Not IsDBNull(reader("Approved")) Then
                        _Approved = CInt(reader("Approved"))
                    End If
                    If Not IsDBNull(reader("notApproved")) Then
                        _notApproved = CInt(reader("notApproved"))
                    End If
                    If Not IsDBNull(reader("Locked")) Then
                        _Locked = CInt(reader("Locked"))
                    End If
                    If Not IsDBNull(reader("notLocked")) Then
                        _notLocked = CInt(reader("notLocked"))
                    End If
                    Dim _obj As aspnetusersstat = New aspnetusersstat(_totUser, _Approved, _notApproved, _Locked, _notLocked)
                    _list.Add(_obj)
                End If
            Catch ex As Exception
                scriviLog("aspnetusersstat_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "dbActivityMonitor"
        Public Overrides Function dbActivityMonitor_List() As List(Of dbActivityMonitor)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT "
            sqlString += "	   SessionId = s.session_id "
            sqlString += "	   ,App = ISNULL(s.program_name, N'')"
            sqlString += "	   ,WaitTime_ms = ISNULL(w.wait_duration_ms, 0)"
            sqlString += "	   ,TotalCPU_ms = s.cpu_time"
            sqlString += "	   ,TotalPhyIO_mb = (s.reads + s.writes) * 8 / 1024"
            sqlString += "	   ,MemUsage_kb = s.memory_usage * 8192 / 1024"
            sqlString += "	   ,OpenTrans = ISNULL(r.open_transaction_count,0)"
            sqlString += "	   ,LoginTime = s.login_time"
            sqlString += "	   ,LastReqStartTime = s.last_request_start_time"
            sqlString += "	   ,HostName = ISNULL(s.host_name, N'')"
            sqlString += "	   ,NetworkAddr = ISNULL(c.client_net_address, N'')"
            sqlString += "	FROM sys.dm_exec_sessions s LEFT OUTER JOIN sys.dm_exec_connections c ON (s.session_id = c.session_id)"
            sqlString += "	LEFT OUTER JOIN sys.dm_exec_requests r ON (s.session_id = r.session_id)"
            sqlString += "	LEFT OUTER JOIN sys.dm_os_tasks t ON (r.session_id = t.session_id AND r.request_id = t.request_id)"
            sqlString += "	LEFT OUTER JOIN "
            sqlString += "	("
            sqlString += "		SELECT *, ROW_NUMBER() OVER (PARTITION BY waiting_task_address ORDER BY wait_duration_ms DESC) AS row_num"
            sqlString += "		FROM sys.dm_os_waiting_tasks "
            sqlString += "	) w ON (t.task_address = w.waiting_task_address) AND w.row_num = 1"
            sqlString += "	LEFT OUTER JOIN sys.dm_exec_requests r2 ON (r.session_id = r2.blocking_session_id)"
            sqlString += "	OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) as st"
            sqlString += "	WHERE s.session_Id > 50  "
            sqlString += "	AND s.session_Id NOT IN (@@SPID) "
            sqlString += "	ORDER BY s.session_id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _list As New List(Of dbActivityMonitor)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    Dim _SessionId As Integer = 0
                    Dim _App As String = 0
                    Dim _WaitTime_ms As Integer = 0
                    Dim _TotalCPU_ms As Integer = 0
                    Dim _TotalPhyIO_mb As Integer = 0
                    Dim _MemUsage_kb As Integer = 0
                    Dim _OpenTrans As Integer = 0
                    Dim _LoginTime As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
                    Dim _LastReqStartTime As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
                    Dim _HostName As String = String.Empty
                    Dim _NetworkAddr As String = String.Empty

                    If Not IsDBNull(reader("SessionId")) Then
                        _SessionId = CInt(reader("SessionId"))
                    End If
                    If Not IsDBNull(reader("App")) Then
                        _App = CStr(reader("App"))
                    End If
                    If Not IsDBNull(reader("WaitTime_ms")) Then
                        _WaitTime_ms = CInt(reader("WaitTime_ms"))
                    End If
                    If Not IsDBNull(reader("TotalCPU_ms")) Then
                        _TotalCPU_ms = CInt(reader("TotalCPU_ms"))
                    End If
                    If Not IsDBNull(reader("TotalPhyIO_mb")) Then
                        _TotalPhyIO_mb = CInt(reader("TotalPhyIO_mb"))
                    End If
                    If Not IsDBNull(reader("MemUsage_kb")) Then
                        _MemUsage_kb = CInt(reader("MemUsage_kb"))
                    End If
                    If Not IsDBNull(reader("OpenTrans")) Then
                        _OpenTrans = CInt(reader("OpenTrans"))
                    End If
                    If Not IsDBNull(reader("LoginTime")) Then
                        If IsDate(reader("LoginTime")) Then
                            _LoginTime = reader("LoginTime")
                        End If
                    End If
                    If Not IsDBNull(reader("LastReqStartTime")) Then
                        If IsDate(reader("LastReqStartTime")) Then
                            _LastReqStartTime = reader("LastReqStartTime")
                        End If
                    End If
                    If Not IsDBNull(reader("HostName")) Then
                        _HostName = CStr(reader("HostName"))
                    End If
                    If Not IsDBNull(reader("NetworkAddr")) Then
                        _NetworkAddr = CStr(reader("NetworkAddr"))
                    End If

                    Dim _obj As dbActivityMonitor = New dbActivityMonitor(_SessionId, _App, _WaitTime_ms, _TotalCPU_ms, _TotalPhyIO_mb, _MemUsage_kb, _OpenTrans, _LoginTime, _LastReqStartTime, _HostName, _NetworkAddr)
                    _list.Add(_obj)
                Loop
            Catch ex As Exception
                scriviLog("dbActivityMonitorList " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "dbstats"
        Public Overrides Function dbstats_Read() As dbstats
            Dim sqlString As String = String.Empty
            sqlString = "SELECT (@@cpu_busy/(@@cpu_busy+@@idle*1.)*100) as cpu_perc"
            sqlString += " ,(SELECT "
            sqlString += "    avg( CAST((io_stall_read_ms + io_stall_write_ms)/(1.0 + num_of_reads + num_of_writes) AS NUMERIC(10,1))/1000) "
            sqlString += "  FROM sys.dm_io_virtual_file_stats(NULL,NULL) AS fs"
            sqlString += "  INNER JOIN sys.master_files AS mf"
            sqlString += "     ON fs.database_id = mf.database_id"
            sqlString += "     AND fs.[file_id] = mf.[file_id]) AS [avg_io]"
            sqlString += " ,(select cntr_value"
            sqlString += "  from sys.dm_os_performance_counters"
            sqlString += " where "
            sqlString += " 	object_name	= 'SQLServer:SQL Statistics'	"
            sqlString += " 	AND counter_name = 'Batch Requests/sec') as [batch_request]"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _list As New List(Of dbstats)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    Dim _cpu_perc As Decimal = 0
                    Dim _avg_io As Decimal = 0
                    Dim _batch_request As Integer = 0

                    If Not IsDBNull(reader("cpu_perc")) Then
                        _cpu_perc = CDec(reader("cpu_perc"))
                    End If
                    If Not IsDBNull(reader("avg_io")) Then
                        _avg_io = CDec(reader("avg_io"))
                    End If
                    If Not IsDBNull(reader("batch_request")) Then
                        _batch_request = CInt(reader("batch_request"))
                    End If

                    Dim _obj As dbstats = New dbstats(_cpu_perc, _avg_io, _batch_request)
                    _list.Add(_obj)
                End If
            Catch ex As Exception
                scriviLog("dbstats_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "Impianti"
        Private Function prepareRecord_Impianti(reader As SqlDataReader) As Impianti
            Dim IdImpianto As String = String.Empty
            Dim DesImpianto As String = String.Empty
            Dim Indirizzo As String = String.Empty
            Dim Latitude As Decimal = 0
            Dim Longitude As Decimal = 0
            Dim AltSLM As Decimal = 0
            Dim IsActive As Boolean = False

            If Not IsDBNull(reader("IdImpianto")) Then
                IdImpianto = reader("IdImpianto").ToString
            End If
            If Not IsDBNull(reader("DesImpianto")) Then
                DesImpianto = CStr(reader("DesImpianto"))
            End If
            If Not IsDBNull(reader("Indirizzo")) Then
                Indirizzo = CStr(reader("Indirizzo"))
            End If
            If Not IsDBNull(reader("Latitude")) Then
                Latitude = CDec(reader("Latitude"))
            End If
            If Not IsDBNull(reader("Longitude")) Then
                Longitude = CDec(reader("Longitude"))
            End If
            If Not IsDBNull(reader("AltSLM")) Then
                AltSLM = CDec(reader("AltSLM"))
            End If
            If Not IsDBNull(reader("IsActive")) Then
                IsActive = CBool(reader("IsActive"))
            End If
            Dim _obj As Impianti = New Impianti(IdImpianto, DesImpianto, Indirizzo, Latitude, Longitude, AltSLM, IsActive)
            Return _obj
        End Function

        Public Overrides Function Impianti_Add(DesImpianto As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Decimal, IsActive As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [Impianti]"
            sqlString += "           ([IdImpianto]"
            sqlString += "           ,[DesImpianto]"
            sqlString += "           ,[Indirizzo]"
            sqlString += "           ,[Latitude]"
            sqlString += "           ,[Longitude]"
            sqlString += "           ,[AltSLM]"
            sqlString += "           ,[IsActive])"
            sqlString += "     VALUES"
            sqlString += "           ((newid())"
            sqlString += "           ,@DesImpianto"
            sqlString += "           ,@Indirizzo"
            sqlString += "           ,@Latitude"
            sqlString += "           ,@Longitude"
            sqlString += "           ,@AltSLM"
            sqlString += "           ,@IsActive)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@DesImpianto", SqlDbType.NVarChar, 50, ParameterDirection.Input, DesImpianto)
            AddParamToSQLCmd(sqlCmd, "@Indirizzo", SqlDbType.NVarChar, 255, ParameterDirection.Input, Indirizzo)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Longitude)
            AddParamToSQLCmd(sqlCmd, "@AltSLM", SqlDbType.Decimal, 0, ParameterDirection.Input, AltSLM)
            AddParamToSQLCmd(sqlCmd, "@IsActive", SqlDbType.Bit, 0, ParameterDirection.Input, IsActive)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                Catch ex As Exception
                    scriviLog("Impianti_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Impianti_List(searchString As String) As List(Of Impianti)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " from Impianti"
            If Not String.IsNullOrEmpty(searchString) Then
                sqlString += " 	where Impianti.[DesImpianto] like @SearchString or"
                sqlString += " 		  Impianti.[Indirizzo] like @SearchString"
            End If

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            If Not String.IsNullOrEmpty(searchString) Then
                AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(searchString) & "%")
            End If

            Dim _list As New List(Of Impianti)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_Impianti(reader))
                Loop
            Catch ex As Exception
                scriviLog("Impianti_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_ListByUser(UserId As String, searchString As String) As List(Of Impianti)
            Dim sqlString As String = String.Empty
            sqlString = "select"
            sqlString += " 	Impianti.*"
            sqlString += " from Impianti"
            sqlString += " INNER JOIN Impianti_Users ON Impianti.IdImpianto = Impianti_Users.IdImpianto"
            sqlString += " WHERE     (Impianti_Users.UserId = @UserId)"
            If Not String.IsNullOrEmpty(searchString) Then
                sqlString += " 	and Impianti.[DesImpianto] like @SearchString or"
                sqlString += " 		Impianti.[Indirizzo] like @SearchString"
            End If

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim UserIdGuid As Guid = New Guid(UserId)
            AddParamToSQLCmd(sqlCmd, "@UserId", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, UserIdGuid)

            If Not String.IsNullOrEmpty(searchString) Then
                AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(searchString) & "%")
            End If

            Dim _list As New List(Of Impianti)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_Impianti(reader))
                Loop
            Catch ex As Exception
                scriviLog("Impianti_ListByUser " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_ListPaged(id As Integer, searchString As String) As List(Of Impianti)
            Dim sqlString As String = String.Empty
            sqlString = "select RowNumber"
            sqlString += " ,Impianti.*"
            sqlString += " from "
            sqlString += " 	(select row_number() over(order by Impianti.IdImpianto) as RowNumber"
            sqlString += " 	,Impianti.*"
            sqlString += " 	from Impianti"
            sqlString += " 	where Impianti.[DesImpianto] like @SearchString or"
            sqlString += " 		  Impianti.[Indirizzo] like @SearchString"
            sqlString += " 	) as Impianti"
            sqlString += " where RowNumber betWeen @first and @last"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(searchString) & "%")
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, id + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, id + 10)

            Dim _list As New List(Of Impianti)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_Impianti(reader))

                Loop
            Catch ex As Exception
                scriviLog("Impianti_ListPaged " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_Read(IdImpianto As String) As Impianti
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " from Impianti"
            sqlString += " where IdImpianto=@IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Dim _list As New List(Of Impianti)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(prepareRecord_Impianti(reader))
                End If
            Catch ex As Exception
                scriviLog("Impianti_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_setIsActive(IdImpianto As String, IsActive As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update Impianti"
            sqlString += " set IsActive=@IsActive"
            sqlString += " where [IdImpianto]=@IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)
            AddParamToSQLCmd(sqlCmd, "@IsActive", SqlDbType.Bit, 0, ParameterDirection.Input, IsActive)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                Catch ex As Exception
                    scriviLog("Impianti_setIsActive " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Impianti_Upd(IdImpianto As String, DesImpianto As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Decimal, IsActive As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update Impianti"
            sqlString += " set DesImpianto=@DesImpianto"
            sqlString += " ,Indirizzo=@Indirizzo"
            sqlString += " ,Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += " ,AltSLM=@AltSLM"
            sqlString += " ,IsActive=@IsActive"
            sqlString += " where [IdImpianto]=@IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)
            AddParamToSQLCmd(sqlCmd, "@DesImpianto", SqlDbType.NVarChar, 50, ParameterDirection.Input, DesImpianto)
            AddParamToSQLCmd(sqlCmd, "@Indirizzo", SqlDbType.NVarChar, 255, ParameterDirection.Input, Indirizzo)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Longitude)
            AddParamToSQLCmd(sqlCmd, "@AltSLM", SqlDbType.Decimal, 0, ParameterDirection.Input, AltSLM)
            AddParamToSQLCmd(sqlCmd, "@IsActive", SqlDbType.Bit, 0, ParameterDirection.Input, IsActive)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                Catch ex As Exception
                    scriviLog("Impianti_Upd " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "Impianti_Contatti"
        Private Function preparerecord_Impianti_contatti(reader As SqlDataReader) As Impianti_Contatti
            Dim _IdContatto As Integer = 0
            If Not IsDBNull(reader("IdContatto")) Then
                _IdContatto = CInt(reader("IdContatto"))
            End If

            Dim _IdImpianto As String = String.Empty
            If Not IsDBNull(reader("IdImpianto")) Then
                _IdImpianto = reader("IdImpianto").ToString
            End If

            Dim _Descrizione As String = String.Empty
            If Not IsDBNull(reader("Descrizione")) Then
                _Descrizione = reader("Descrizione").ToString
            End If

            Dim _Indirizzo As String = String.Empty
            If Not IsDBNull(reader("Indirizzo")) Then
                _Indirizzo = reader("Indirizzo").ToString
            End If

            Dim _Nome As String = String.Empty
            If Not IsDBNull(reader("Nome")) Then
                _Nome = reader("Nome").ToString
            End If

            Dim _TelFisso As String = String.Empty
            If Not IsDBNull(reader("telFisso")) Then
                _TelFisso = reader("telFisso").ToString
            End If

            Dim _TelMobile As String = String.Empty
            If Not IsDBNull(reader("telMobile")) Then
                _TelMobile = reader("telMobile").ToString
            End If

            Dim _emailaddress As String = String.Empty
            If Not IsDBNull(reader("emailaddress")) Then
                _emailaddress = reader("emailaddress").ToString
            End If

            Dim _obj As Impianti_Contatti = New Impianti_Contatti(_IdContatto, _IdImpianto, _Descrizione, _Indirizzo, _Nome, _TelFisso, _TelMobile, _emailaddress)
            Return _obj
        End Function

        Public Overrides Function Impianti_Contatti_Add(IdImpianto As String, _
                                                        Descrizione As String, _
                                                        Indirizzo As String, _
                                                        Nome As String, _
                                                        TelFisso As String, _
                                                        TelMobile As String, _
                                                        emailaddress As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [Impianti_Contatti]"
            sqlString += "           ([IdImpianto]"
            sqlString += "           ,[Descrizione]"
            sqlString += "           ,[Indirizzo]"
            sqlString += "           ,[Nome]"
            sqlString += "           ,[telFisso]"
            sqlString += "           ,[telMobile]"
            sqlString += "           ,[emailaddress])"
            sqlString += "     VALUES"
            sqlString += "           (@IdImpianto"
            sqlString += "           ,@Descrizione"
            sqlString += "           ,@Indirizzo"
            sqlString += "           ,@Nome"
            sqlString += "           ,@telFisso"
            sqlString += "           ,@telMobile"
            sqlString += "           ,@emailaddress)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)
            AddParamToSQLCmd(sqlCmd, "@Descrizione", SqlDbType.NVarChar, 100, ParameterDirection.Input, Descrizione)
            AddParamToSQLCmd(sqlCmd, "@Indirizzo", SqlDbType.NVarChar, 255, ParameterDirection.Input, Indirizzo)
            AddParamToSQLCmd(sqlCmd, "@Nome", SqlDbType.NVarChar, 100, ParameterDirection.Input, Nome)
            AddParamToSQLCmd(sqlCmd, "@telFisso", SqlDbType.NVarChar, 100, ParameterDirection.Input, TelFisso)
            AddParamToSQLCmd(sqlCmd, "@telMobile", SqlDbType.NVarChar, 100, ParameterDirection.Input, TelMobile)
            AddParamToSQLCmd(sqlCmd, "@emailaddress", SqlDbType.NVarChar, 100, ParameterDirection.Input, emailaddress)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    retVal = True
                Catch ex As Exception
                    scriviLog("Impianti_Contatti_Add " & ex.Message)
                    retVal = False
                End Try
            End Using
            Return retVal
        End Function

        Public Overrides Function Impianti_Contatti_Del(IdContatto As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Impianti_Contatti]"
            sqlString += " where IdContatto=@IdContatto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@IdContatto", SqlDbType.Int, 0, ParameterDirection.Input, IdContatto)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Impianti_Contatti_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Impianti_Contatti_List(IdImpianto As String) As IList(Of Impianti_Contatti)
            Dim sqlString As String = String.Empty
            sqlString = "select * from Impianti_Contatti"
            sqlString += " where IdImpianto=@IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Dim _List As New List(Of Impianti_Contatti)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_Impianti_contatti(reader))
                Loop
            Catch ex As Exception
                scriviLog("Impianti_Contatti_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_Contatti_Read(IdContatto As Integer) As Impianti_Contatti
            Dim sqlString As String = String.Empty
            sqlString = "select * from Impianti_Contatti"
            sqlString += " where IdContatto=@IdContatto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@IdContatto", SqlDbType.Int, 0, ParameterDirection.Input, IdContatto)

            Dim _List As New List(Of Impianti_Contatti)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_Impianti_contatti(reader))
                Loop
            Catch ex As Exception
                scriviLog("Impianti_Contatti_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_Contatti_Update(IdContatto As Integer, _
                                                           Descrizione As String, _
                                                           Indirizzo As String, _
                                                           Nome As String, _
                                                           TelFisso As String, _
                                                           TelMobile As String, _
                                                           emailaddress As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Impianti_Contatti]"
            sqlString += " set Descrizione=@Descrizione"
            sqlString += " , Indirizzo=@Indirizzo"
            sqlString += " , Nome=@Nome"
            sqlString += " , TelFisso=@TelFisso"
            sqlString += " , TelMobile=@TelMobile"
            sqlString += " , emailaddress=@emailaddress"
            sqlString += " where IdContatto=@IdContatto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@IdContatto", SqlDbType.Int, 0, ParameterDirection.Input, IdContatto)
            AddParamToSQLCmd(sqlCmd, "@Descrizione", SqlDbType.NVarChar, 100, ParameterDirection.Input, Descrizione)
            AddParamToSQLCmd(sqlCmd, "@Indirizzo", SqlDbType.NVarChar, 255, ParameterDirection.Input, Indirizzo)
            AddParamToSQLCmd(sqlCmd, "@Nome", SqlDbType.NVarChar, 100, ParameterDirection.Input, Nome)
            AddParamToSQLCmd(sqlCmd, "@telFisso", SqlDbType.NVarChar, 100, ParameterDirection.Input, TelFisso)
            AddParamToSQLCmd(sqlCmd, "@telMobile", SqlDbType.NVarChar, 100, ParameterDirection.Input, TelMobile)
            AddParamToSQLCmd(sqlCmd, "@emailaddress", SqlDbType.NVarChar, 100, ParameterDirection.Input, emailaddress)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Impianti_Contatti_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "Impianti_RemoteConnections"
        Private Function preparerecord_Impianti_RemoteConnection(reader As SqlDataReader) As Impianti_RemoteConnections
            Dim _IdImpianto As String = String.Empty
            If Not IsDBNull(reader("IdImpianto")) Then
                _IdImpianto = reader("IdImpianto").ToString
            End If

            Dim _IdAddress As Integer = 0
            If Not IsDBNull(reader("IdAddress")) Then
                _IdAddress = CInt(reader("IdAddress").ToString)
            End If

            Dim _Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                _Descr = reader("Descr")
            End If

            Dim _remoteAddress As String = String.Empty
            If Not IsDBNull(reader("remoteAddress")) Then
                _remoteAddress = reader("remoteAddress")
            End If

            Dim _remotePort As Integer = 0
            If Not IsDBNull(reader("remotePort")) Then
                _remotePort = CInt(reader("remotePort").ToString)
            End If

            Dim _connectionType As Integer = 0
            If Not IsDBNull(reader("connectionType")) Then
                _connectionType = CInt(reader("connectionType").ToString)
            End If
            Dim _DconnectionType As String = String.Empty
            If _connectionType = 0 Then
                _DconnectionType = "web"
            End If
            If _connectionType = 1 Then
                _DconnectionType = "rdp"
            End If
            If _connectionType = 2 Then
                _DconnectionType = "VPN"
            End If

            Dim _NoteInterne As String = String.Empty
            If Not IsDBNull(reader("NoteInterne")) Then
                _NoteInterne = reader("NoteInterne")
            End If

            Dim obj As Impianti_RemoteConnections = New Impianti_RemoteConnections(_IdImpianto, _IdAddress, _Descr, _remoteAddress, _remotePort, _connectionType, _DconnectionType, _NoteInterne)
            Return obj
        End Function

        Public Overrides Function Impianti_RemoteConnections_Add(IdImpianto As String, _
                                                                 IdAddress As Integer, _
                                                                 Descr As String, _
                                                                 remoteaddress As String, _
                                                                 connectionType As Integer, _
                                                                 NoteInterne As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Impianti_RemoteConnections]"
            sqlString += " ("
            sqlString += "IdImpianto,"
            sqlString += "[IdAddress],"
            sqlString += "[Descr],"
            sqlString += "[remoteAddress],"
            sqlString += "[connectionType],"
            sqlString += "[NoteInterne]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@IdImpianto,"
            sqlString += "@IdAddress,"
            sqlString += "@Descr,"
            sqlString += "@remoteAddress,"
            sqlString += "@connectionType,"
            sqlString += "@NoteInterne"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim ImpiantoGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, ImpiantoGuid)
            AddParamToSQLCmd(sqlCmd, "@IdAddress", SqlDbType.Int, 0, ParameterDirection.Input, IdAddress)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@remoteAddress", SqlDbType.NVarChar, 255, ParameterDirection.Input, remoteaddress)
            AddParamToSQLCmd(sqlCmd, "@connectionType", SqlDbType.Int, 0, ParameterDirection.Input, connectionType)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Impianti_RemoteConnections_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Impianti_RemoteConnections_Del(IdImpianto As String, IdAddress As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Impianti_RemoteConnections]"
            sqlString += " where IdImpianto=@IdImpianto"
            sqlString += "   and IdAddress=@IdAddress"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim ImpiantoGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, ImpiantoGuid)
            AddParamToSQLCmd(sqlCmd, "@IdAddress", SqlDbType.Int, 0, ParameterDirection.Input, IdAddress)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Impianti_RemoteConnections_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Impianti_RemoteConnections_getLastId() As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "SELECT max(IdAddress) as IdAddress"
            sqlString += " FROM Impianti_RemoteConnections"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString


            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("IdAddress")) Then
                        retVal = CInt(reader("IdAddress"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("Impianti_RemoteConnections_getLastId " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function Impianti_RemoteConnections_List(IdImpianto As String) As List(Of Impianti_RemoteConnections)
            Dim sqlString As String = String.Empty
            sqlString = "select * from Impianti_RemoteConnections"
            sqlString += " where IdImpianto=@IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim ImpiantoGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, ImpiantoGuid)


            Dim _List As New List(Of Impianti_RemoteConnections)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection



            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_Impianti_RemoteConnection(reader))
                Loop
            Catch ex As Exception
                scriviLog("Impianti_RemoteConnections_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_RemoteConnections_Read(IdImpianto As String, IdAddress As Integer) As Impianti_RemoteConnections
            Dim sqlString As String = String.Empty
            sqlString = "select * from Impianti_RemoteConnections"
            sqlString += " where IdImpianto=@IdImpianto"
            sqlString += "   and IdAddress=@IdAddress"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim ImpiantoGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, ImpiantoGuid)
            AddParamToSQLCmd(sqlCmd, "@IdAddress", SqlDbType.Int, 0, ParameterDirection.Input, IdAddress)

            Dim _List As New List(Of Impianti_RemoteConnections)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_Impianti_RemoteConnection(reader))
                End If
            Catch ex As Exception
                scriviLog("Impianti_RemoteConnections_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Impianti_RemoteConnections_Upd(IdImpianto As String, _
                                                                 IdAddress As Integer, _
                                                                 Descr As String, _
                                                                 remoteaddress As String, _
                                                                 connectionType As Integer, _
                                                                 NoteInterne As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update Impianti_RemoteConnections"
            sqlString += " set Descr=@Descr"
            sqlString += "   , remoteaddress=@remoteaddress"
            sqlString += "   , connectionType=@connectionType"
            sqlString += "   , NoteInterne=@NoteInterne"
            sqlString += " WHERE IdImpianto=@IdImpianto"
            sqlString += "   and IdAddress=@IdAddress"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim ImpiantoGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, ImpiantoGuid)
            AddParamToSQLCmd(sqlCmd, "@IdAddress", SqlDbType.Int, 0, ParameterDirection.Input, IdAddress)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@remoteAddress", SqlDbType.NVarChar, 255, ParameterDirection.Input, remoteaddress)
            AddParamToSQLCmd(sqlCmd, "@connectionType", SqlDbType.Int, 0, ParameterDirection.Input, connectionType)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Impianti_RemoteConnections_Upd " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "tbhsElem"
        Private Function prepareRecord_tbhsElem(reader As SqlDataReader) As tbhsElem
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim name As String = 0
            If Not IsDBNull(reader("name")) Then
                name = reader("name").ToString
            End If

            Dim descr As String = String.Empty
            If Not IsDBNull(reader("descr")) Then
                descr = CStr(reader("descr"))
            End If

            Dim _obj As tbhsElem = New tbhsElem(Id, name, descr)
            Return _obj
        End Function

        Public Overrides Function tbhsElem_Add(Id As Integer, name As String, descr As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [tbhsElem]"
            sqlString += "(Id,"
            sqlString += "[name],"
            sqlString += "[Descr],"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@Id,"
            sqlString += "@name,"
            sqlString += "@Descr,"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@name", SqlDbType.NVarChar, 50, ParameterDirection.Input, name)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, descr)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("tbhsElem_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function tbhsElem_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [tbhsElem]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("tbhsElem_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function tbhsElem_List() As List(Of tbhsElem)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM tbhsElem "

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _List As New List(Of tbhsElem)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_tbhsElem(reader))
                Loop
            Catch ex As Exception
                scriviLog("tbhsElem_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function tbhsElem_Read(Id As Integer) As tbhsElem
            Dim sqlString As String = String.Empty
            sqlString = "select tbhsElem.*"
            sqlString += " FROM tbhsElem "
            sqlString += " where tbhsElem.Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of tbhsElem)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_tbhsElem(reader))
                End If
            Catch ex As Exception
                scriviLog("tbhsElem_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function tbhsElem_Update(Id As Integer, name As String, descr As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [tbhsElem]"
            sqlString += " set name=@name"
            sqlString += " , descr=@descr"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@name", SqlDbType.NVarChar, 50, ParameterDirection.Input, name)
            AddParamToSQLCmd(sqlCmd, "@descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, descr)
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)


            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("tbhsElem_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "hs_Elem"
        Private Function preparerecord_hs_elem(reader As SqlDataReader) As hs_Elem
            Dim totLux As Integer = 0
            If Not IsDBNull(reader("totLux")) Then
                totLux = CInt(reader("totLux").ToString)
            End If

            Dim statoLux As Integer = 0
            If Not IsDBNull(reader("statoLux")) Then
                statoLux = CInt(reader("statoLux").ToString)
            End If

            Dim totLuxM As Integer = 0
            If Not IsDBNull(reader("totLuxM")) Then
                totLuxM = CInt(reader("totLuxM").ToString)
            End If

            Dim statoLuxM As Integer = 0
            If Not IsDBNull(reader("statoLuxM")) Then
                statoLuxM = CInt(reader("statoLuxM").ToString)
            End If

            Dim totPsg As Integer = 0
            If Not IsDBNull(reader("totPsg")) Then
                totPsg = CInt(reader("totPsg").ToString)
            End If

            Dim statoPsg As Integer = 0
            If Not IsDBNull(reader("statoPsg")) Then
                statoPsg = CInt(reader("statoPsg").ToString)
            End If

            Dim totZrel As Integer = 0
            If Not IsDBNull(reader("totZrel")) Then
                totZrel = CInt(reader("totZrel").ToString)
            End If

            Dim statoZrel As Integer = 0
            If Not IsDBNull(reader("statoZrel")) Then
                statoZrel = CInt(reader("statoZrel").ToString)
            End If

            Dim _totCronograph As Integer = 0
            If Not IsDBNull(reader("totCronograph")) Then
                _totCronograph = CInt(reader("totCronograph"))
            End If

            Dim _statoCronograph As Integer = 0
            If Not IsDBNull(reader("statoCronograph")) Then
                _statoCronograph = CInt(reader("statoCronograph"))
            End If



            Dim _obj As hs_Elem = New hs_Elem(totLux, statoLux, totLuxM, statoLuxM, totPsg, statoPsg, totZrel, statoZrel, _totCronograph, _statoCronograph)
            Return _obj
        End Function

        Public Overrides Function hs_Elem_Read(hsId As Integer) As hs_Elem
            Dim sqlString As String = String.Empty
            sqlString = "SELECT  hsId"
            sqlString += " ,(select count(*) from Lux  where Lux.hsId=HeatingSystem.hsId) as totLux"
            sqlString += " ,(select count(*) from LuxM where LuxM.hsId=HeatingSystem.hsId) as totLuxM"
            sqlString += " ,(select count(*) from Psg  where Psg.hsId=HeatingSystem.hsId) as totPsg"
            sqlString += " ,(select count(*) from Zrel where Zrel.hsId=HeatingSystem.hsId) as totZrel"
            sqlString += " ,(select count(*) from hs_cron where hs_cron.hsId=HeatingSystem.hsId and CronType=1) as totCronograph"

            sqlString += " ,(select max(stato) FROM [Lux]  where Lux.hsId=HeatingSystem.hsId ) statoLux"
            sqlString += " ,(select max(stato) FROM [LuxM] where LuxM.hsId=HeatingSystem.hsId ) statoLuxM"
            sqlString += " ,(select max(stato) from Psg    where Psg.hsId=HeatingSystem.hsId ) as statoPsg"
            sqlString += " ,(select max(stato) from Zrel   where Zrel.hsId=HeatingSystem.hsId ) as statoZrel"
            sqlString += " ,(select max(stato) from hs_Cron  where hs_Cron.hsId=HeatingSystem.hsId  and CronType=1) as statoCronograph"

            sqlString += " FROM HeatingSystem"
            sqlString += " where HeatingSystem.hsId = @hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _list As New List(Of hs_Elem)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(preparerecord_hs_elem(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Elem_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "hs_Docs"
        Private Function preparerecord_hs_Docs(reader As SqlDataReader) As hs_Docs
            Dim _IdDoc As Integer = 0
            If Not IsDBNull(reader("IdDoc")) Then
                _IdDoc = CInt(reader("IdDoc"))
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId"))
            End If

            Dim _DocName As String = String.Empty
            If Not IsDBNull(reader("DocName")) Then
                _DocName = CStr(reader("DocName"))
            End If

            Dim _ContentType As String = String.Empty
            If Not IsDBNull(reader("ContentType")) Then
                _ContentType = CStr(reader("ContentType"))
            End If

            Dim _DocSize As String = String.Empty
            If Not IsDBNull(reader("DocSize")) Then
                _DocSize = CStr(reader("DocSize"))
            End If

            Dim fileExt As String = String.Empty
            If Not IsDBNull(reader("fileExt")) Then
                fileExt = CStr(reader("fileExt"))
            End If

            Dim _obj As hs_Docs = New hs_Docs(_IdDoc, _hsId, _DocName, _ContentType, _DocSize, fileExt)
            Return _obj
        End Function

        Public Overrides Function hs_Docs_Add(hsId As Integer, DocName As String, Creator As String) As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Docs]"
            sqlString += " ("
            sqlString += " [hsId],"
            sqlString += " [DocName],"
            sqlString += " [Creator]"
            sqlString += " )"
            sqlString += " VALUES"
            sqlString += " ("
            sqlString += " @hsId,"
            sqlString += " @DocName,"
            sqlString += " @Creator"
            sqlString += " )"
            sqlString += "; SELECT CAST(scope_identity() AS int);"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@DocName", SqlDbType.NVarChar, 256, ParameterDirection.Input, DocName)
            AddParamToSQLCmd(sqlCmd, "@Creator", SqlDbType.NVarChar, 256, ParameterDirection.Input, Creator)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    retVal = Convert.ToInt32(sqlCmd.ExecuteScalar())
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("hs_Docs_Add " & ex.Message)
                    Return -1
                End Try
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Docs_Add(hsId As Integer, _
                                              DocName As String, _
                                              Creator As String, _
                                              ContentType As String, _
                                              DocSize As String, _
                                              BinaryData As String) As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Docs]"
            sqlString += " ("
            sqlString += " [hsId],"
            sqlString += " [DocName],"
            sqlString += " [ContentType],"
            sqlString += " [DocSize],"
            sqlString += " [BinaryData],"
            sqlString += " [Creator]"
            sqlString += " )"
            sqlString += " VALUES"
            sqlString += " ("
            sqlString += " @hsId,"
            sqlString += " @DocName,"
            sqlString += " @ContentType,"
            sqlString += " @DocSize,"
            sqlString += " @BinaryData,"
            sqlString += " @Creator"
            sqlString += " )"
            sqlString += "; SELECT CAST(scope_identity() AS int);"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@DocName", SqlDbType.NVarChar, 256, ParameterDirection.Input, DocName)
            AddParamToSQLCmd(sqlCmd, "@ContentType", SqlDbType.NVarChar, 50, ParameterDirection.Input, ContentType)
            AddParamToSQLCmd(sqlCmd, "@DocSize", SqlDbType.NVarChar, 50, ParameterDirection.Input, DocSize)
            AddParamToSQLCmd(sqlCmd, "@BinaryData", SqlDbType.NText, 0, ParameterDirection.Input, BinaryData)
            AddParamToSQLCmd(sqlCmd, "@Creator", SqlDbType.NVarChar, 256, ParameterDirection.Input, Creator)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    retVal = Convert.ToInt32(sqlCmd.ExecuteScalar())
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("hs_Docs_Add " & ex.Message)
                    retVal = -1
                End Try
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Docs_Del(IdDoc As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Docs]"
            sqlString += "where IdDoc=@IdDoc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@IdDoc", SqlDbType.Int, 0, ParameterDirection.Input, IdDoc)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Docs_Del " & ex.Message)
                    retVal = False
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Docs_List(hsId As Integer) As List(Of hs_Docs)
            Dim sqlString As String = String.Empty
            sqlString = "select [IdDoc],[hsId],[DocName],[ContentType],[DocSize],[fileExt] from hs_Docs"
            sqlString += " where hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of hs_Docs)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_Docs(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Docs_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Docs_Read(IdDoc As Integer) As hs_Docs
            Dim sqlString As String = String.Empty
            sqlString = "select [IdDoc],[hsId],[DocName],[ContentType],[DocSize],[fileExt] from hs_Docs"
            sqlString += " where IdDoc=@IdDoc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@IdDoc", SqlDbType.Int, 0, ParameterDirection.Input, IdDoc)

            Dim _List As New List(Of hs_Docs)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Docs(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Docs_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Docs_getBinaryData(IdDoc As Integer) As String
            Dim retVal As String = String.Empty
            Dim sqlString As String = String.Empty
            sqlString = "select BinaryData from hs_Docs"
            sqlString += " where IdDoc=@IdDoc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@IdDoc", SqlDbType.Int, 0, ParameterDirection.Input, IdDoc)

            Dim _List As New List(Of hs_Docs)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("BinaryData")) Then
                        retVal = CStr(reader("BinaryData"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("hs_Docs_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function hs_Docs_getTot(hsId As Integer) As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "select count([IdDoc])  as tot from hs_Docs"
            sqlString += " where hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("tot")) Then
                        retVal = CInt(reader("tot"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("hs_Docs_getTot " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function hs_Docs_setBinaryData(IdDoc As Integer, BinaryData As String, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update hs_Docs"
            sqlString += " set BinaryData=@BinaryData"
            sqlString += "   , Modifier=@UserName"
            sqlString += "   , LastUpdate=LastUpdate+1"
            sqlString += " where IdDoc=@IdDoc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@IdDoc", SqlDbType.Int, 0, ParameterDirection.Input, IdDoc)
            AddParamToSQLCmd(sqlCmd, "@BinaryData", SqlDbType.NText, 0, ParameterDirection.Input, BinaryData)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("hs_Docs_setBinaryData " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Docs_setInfo(IdDoc As Integer, ContentType As String, DocSize As String, fileExt As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update hs_Docs"
            sqlString += " set ContentType=@ContentType"
            sqlString += "   , DocSize=@DocSize"
            sqlString += "   , fileExt=@fileExt"
            sqlString += " where IdDoc=@IdDoc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@IdDoc", SqlDbType.Int, 0, ParameterDirection.Input, IdDoc)
            AddParamToSQLCmd(sqlCmd, "@ContentType", SqlDbType.NVarChar, 50, ParameterDirection.Input, ContentType)
            AddParamToSQLCmd(sqlCmd, "@DocSize", SqlDbType.NVarChar, 50, ParameterDirection.Input, DocSize)
            AddParamToSQLCmd(sqlCmd, "@fileExt", SqlDbType.NVarChar, 50, ParameterDirection.Input, fileExt)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("hs_Docs_setInfo " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Docs_Update(IdDoc As Integer, DocName As String, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update hs_Docs"
            sqlString += " set DocName=@DocName"
            sqlString += "   , Modifier=@UserName"
            sqlString += "   , LastUpdate=LastUpdate+1"
            sqlString += " where IdDoc=@IdDoc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@IdDoc", SqlDbType.Int, 0, ParameterDirection.Input, IdDoc)
            AddParamToSQLCmd(sqlCmd, "@DocName", SqlDbType.NVarChar, 256, ParameterDirection.Input, DocName)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("hs_Docs_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "hs_ErrorCodes"
        Private Function preparerecord_hs_ErrorCodes(reader As SqlDataReader) As hs_ErrorCodes
            Dim _elementCode As String = String.Empty
            If Not IsDBNull(reader("elementCode")) Then
                _elementCode = (reader("elementCode").ToString)
            End If

            Dim _errorCode As Integer = 0
            If Not IsDBNull(reader("errorCode")) Then
                _errorCode = CInt(reader("errorCode").ToString)
            End If

            Dim _errorLevel As Integer = 0
            If Not IsDBNull(reader("errorLevel")) Then
                _errorLevel = CInt(reader("errorLevel").ToString)
            End If

            Dim _DescIT As String = String.Empty
            If Not IsDBNull(reader("DescIT")) Then
                _DescIT = CStr(reader("DescIT"))
            End If

            Dim _DescEN As String = String.Empty
            If Not IsDBNull(reader("DescEN")) Then
                _DescEN = CStr(reader("DescEN"))
            End If

            Dim _obj As hs_ErrorCodes = New hs_ErrorCodes(_elementCode, _errorCode, _errorLevel, _DescIT, _DescEN)
            Return _obj
        End Function

        Public Overrides Function hs_ErrorCodes_Add(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_ErrorCodes]"
            sqlString += "(elementCode,"
            sqlString += "[errorCode],"
            sqlString += "[errorLevel],"
            sqlString += "[DescIT],"
            sqlString += "[DescEN]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@elementCode,"
            sqlString += "@errorCode,"
            sqlString += "@errorLevel,"
            sqlString += "@DescIT,"
            sqlString += "@DescEN"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@elementCode", SqlDbType.NVarChar, 10, ParameterDirection.Input, elementCode)
            AddParamToSQLCmd(sqlCmd, "@errorCode", SqlDbType.Int, 0, ParameterDirection.Input, errorCode)
            AddParamToSQLCmd(sqlCmd, "@errorLevel", SqlDbType.Int, 0, ParameterDirection.Input, errorLevel)
            AddParamToSQLCmd(sqlCmd, "@DescIT", SqlDbType.NVarChar, 50, ParameterDirection.Input, DescIT)
            AddParamToSQLCmd(sqlCmd, "@DescEN", SqlDbType.NVarChar, 50, ParameterDirection.Input, DescEN)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ErrorCodes_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_ErrorCodes_Del(elementCode As String, errorCode As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_ErrorCodes]"
            sqlString += " where elementCode=@elementCode"
            sqlString += "   and errorCode=@errorCode"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@elementCode", SqlDbType.NVarChar, 10, ParameterDirection.Input, elementCode)
            AddParamToSQLCmd(sqlCmd, "@errorCode", SqlDbType.Int, 0, ParameterDirection.Input, errorCode)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ErrorCodes_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_ErrorCodes_getTot() As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "select count(*)  as tot from hs_ErrorCodes"
            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString


            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("tot")) Then
                        retVal = CInt(reader("tot"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("hs_ErrorCodes_getTot " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function hs_ErrorCodes_List(elementCode As String) As List(Of hs_ErrorCodes)
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_ErrorCodes"
            sqlString += " where elementCode=@elementCode"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@elementCode", SqlDbType.NVarChar, 10, ParameterDirection.Input, elementCode)

            Dim _List As New List(Of hs_ErrorCodes)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_ErrorCodes(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_ErrorCodes_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_ErrorCodes_Read(elementCode As String, errorCode As Integer) As hs_ErrorCodes
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_ErrorCodes"
            sqlString += " where elementCode=@elementCode"
            sqlString += "   and errorCode=@errorCode"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@elementCode", SqlDbType.NVarChar, 10, ParameterDirection.Input, elementCode)
            AddParamToSQLCmd(sqlCmd, "@errorCode", SqlDbType.Int, 0, ParameterDirection.Input, errorCode)

            Dim _List As New List(Of hs_ErrorCodes)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_ErrorCodes(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_ErrorCodes_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_ErrorCodes_Update(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_ErrorCodes]"
            sqlString += " set DescIT=@DescIT"
            sqlString += " , DescEN=@DescEN"
            sqlString += " , errorLevel=@errorLevel"
            sqlString += "   where elementCode=@elementCode"
            sqlString += "     and errorCode=@errorCode"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@elementCode", SqlDbType.NVarChar, 10, ParameterDirection.Input, elementCode)
            AddParamToSQLCmd(sqlCmd, "@errorCode", SqlDbType.Int, 0, ParameterDirection.Input, errorCode)
            AddParamToSQLCmd(sqlCmd, "@DescIT", SqlDbType.NVarChar, 50, ParameterDirection.Input, DescIT)
            AddParamToSQLCmd(sqlCmd, "@DescEN", SqlDbType.NVarChar, 50, ParameterDirection.Input, DescEN)
            AddParamToSQLCmd(sqlCmd, "@errorLevel", SqlDbType.Int, 0, ParameterDirection.Input, errorLevel)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ErrorCodes_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "hs_ErrorLog"
        Private Function prepareRecord_hs_Errorlog(reader As SqlDataReader) As hs_ErrorLog
            Dim _Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                _Id = CInt(reader("Id").ToString)
            End If

            Dim _LogDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("LogDate")) Then
                If IsDate(reader("LogDate")) Then
                    _LogDate = reader("LogDate")
                End If
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _hselement As String = String.Empty
            If Not IsDBNull(reader("hselement")) Then
                _hselement = CStr(reader("hselement").ToString)
            End If

            Dim _elementCode As String = String.Empty
            If Not IsDBNull(reader("elementCode")) Then
                _elementCode = CStr(reader("elementCode").ToString)
            End If

            Dim _errorCode As Integer = 0
            If Not IsDBNull(reader("errorCode")) Then
                _errorCode = CInt(reader("errorCode").ToString)
            End If

            Dim _errorValue As String = String.Empty
            If Not IsDBNull(reader("errorValue")) Then
                _errorValue = CStr(reader("errorValue").ToString)
            End If

            Dim _errorLevel As Integer = 0
            If Not IsDBNull(reader("errorLevel")) Then
                _errorLevel = CInt(reader("errorLevel").ToString)
            End If

            Dim _DescIT As String = String.Empty
            If Not IsDBNull(reader("DescIT")) Then
                _DescIT = CStr(reader("DescIT").ToString)
            End If

            Dim _DescEN As String = String.Empty
            If Not IsDBNull(reader("DescEN")) Then
                _DescEN = CStr(reader("DescEN").ToString)
            End If

            Dim _obj As hs_ErrorLog = New hs_ErrorLog(_Id, _LogDate, _hsId, _hselement, _elementCode, _errorCode, _errorValue, _errorLevel, _DescIT, _DescEN)
            Return _obj
        End Function

        Public Overrides Function hs_ErrorLog_Add(LogDate As Date, hsId As Integer, hselement As String, elementCode As String, errorCode As Integer, errorValue As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_ErrorLog]"
            sqlString += "(LogDate,"
            sqlString += "[hsId],"
            sqlString += "[hselement],"
            sqlString += "[elementCode],"
            sqlString += "[errorCode],"
            sqlString += "[errorValue]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@LogDate,"
            sqlString += "@hsId,"
            sqlString += "@hselement,"
            sqlString += "@elementCode,"
            sqlString += "@errorCode,"
            sqlString += "@errorValue"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@LogDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, LogDate)
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@hselement", SqlDbType.NVarChar, 10, ParameterDirection.Input, hselement)
            AddParamToSQLCmd(sqlCmd, "@elementCode", SqlDbType.NVarChar, 10, ParameterDirection.Input, elementCode)
            AddParamToSQLCmd(sqlCmd, "@errorCode", SqlDbType.Int, 0, ParameterDirection.Input, errorCode)
            AddParamToSQLCmd(sqlCmd, "@errorValue", SqlDbType.NVarChar, 50, ParameterDirection.Input, errorValue)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ErrorLog_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_ErrorLog_List(hsId As Integer, fromDate As Date, toDate As Date) As List(Of hs_ErrorLog)
            Dim sqlString As String = String.Empty
            sqlString = "select hs_ErrorLog.*, hs_ErrorCodes.errorLevel,hs_ErrorCodes.errorlevel, hs_ErrorCodes.DescIT, hs_ErrorCodes.DescEN from hs_ErrorLog"
            sqlString += " LEFT OUTER JOIN  hs_ErrorCodes ON hs_ErrorLog.elementCode = hs_ErrorCodes.elementCode AND hs_ErrorLog.errorCode = hs_ErrorCodes.errorCode"
            sqlString += " where hs_ErrorLog.hsId=@hsId"
            sqlString += " and hs_ErrorLog.logDate between @fromDate and @toDate"
            sqlString += " order by hs_ErrorLog.[LogDate] desc"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of hs_ErrorLog)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_hs_Errorlog(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_ErrorLog_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_ErrorLog_ListAll(hsId As Integer, rowNumber As Integer) As List(Of hs_ErrorLog)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT RowNumber"
            sqlString += " ,0  as Id"
            sqlString += " , LogDate"
            sqlString += " , hsId"
            sqlString += " , hselement"
            sqlString += " , elementCode"
            sqlString += " , errorCode"
            sqlString += " , errorValue"
            sqlString += " , errorLevel"
            sqlString += " , DescIT"
            sqlString += " , DescEN"
            sqlString += " from"
            sqlString += " (select row_number() over(order by hs_ErrorLog.LogDate desc) as rowNumber"
            sqlString += " , 0 as Id"
            sqlString += " , hs_ErrorLog.LogDate"
            sqlString += " , hs_ErrorLog.hsId"
            sqlString += " , hs_ErrorLog.hselement"
            sqlString += " , hs_ErrorLog.elementCode"
            sqlString += " , hs_ErrorLog.errorCode"
            sqlString += " , hs_ErrorLog.errorValue"
            sqlString += " , hs_ErrorCodes.errorLevel"
            sqlString += " , hs_ErrorCodes.DescIT"
            sqlString += " , hs_ErrorCodes.DescEN"
            sqlString += " FROM hs_ErrorLog "
            sqlString += " LEFT OUTER JOIN hs_ErrorCodes ON hs_ErrorLog.elementCode = hs_ErrorCodes.elementCode "
            sqlString += "                                  AND hs_ErrorLog.errorCode = hs_ErrorCodes.errorCode"
            sqlString += " WHERE hs_ErrorLog.hsId = @hsId "
            sqlString += " group by   hs_ErrorLog.LogDate"
            sqlString += " , hs_ErrorLog.hsId"
            sqlString += " , hs_ErrorLog.hselement"
            sqlString += " , hs_ErrorLog.elementCode"
            sqlString += " , hs_ErrorLog.errorCode"
            sqlString += " , hs_ErrorLog.errorValue"
            sqlString += " , hs_ErrorCodes.errorLevel"
            sqlString += " , hs_ErrorCodes.DescIT"
            sqlString += " , hs_ErrorCodes.DescEN) as ErrorLog"
            sqlString += " where RowNumber betWeen @first and @last"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of hs_ErrorLog)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_hs_Errorlog(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_ErrorLog_ListAll " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_ErrorLog_ListByElement(hsId As Integer, hselement As String, rowNumber As Integer) As List(Of hs_ErrorLog)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT RowNumber"
            sqlString += " ,0  as Id"
            sqlString += " , LogDate"
            sqlString += " , hsId"
            sqlString += " , hselement"
            sqlString += " , elementCode"
            sqlString += " , errorCode"
            sqlString += " , errorValue"
            sqlString += " , errorLevel"
            sqlString += " , DescIT"
            sqlString += " , DescEN"
            sqlString += " from"
            sqlString += " (select row_number() over(order by hs_ErrorLog.LogDate desc) as rowNumber"
            sqlString += " , 0 as Id"
            sqlString += " , hs_ErrorLog.LogDate"
            sqlString += " , hs_ErrorLog.hsId"
            sqlString += " , hs_ErrorLog.hselement"
            sqlString += " , hs_ErrorLog.elementCode"
            sqlString += " , hs_ErrorLog.errorCode"
            sqlString += " , hs_ErrorLog.errorValue"
            sqlString += " , hs_ErrorCodes.errorLevel"
            sqlString += " , hs_ErrorCodes.DescIT"
            sqlString += " , hs_ErrorCodes.DescEN"
            sqlString += " FROM hs_ErrorLog "
            sqlString += " LEFT OUTER JOIN hs_ErrorCodes ON hs_ErrorLog.elementCode = hs_ErrorCodes.elementCode "
            sqlString += "                                  AND hs_ErrorLog.errorCode = hs_ErrorCodes.errorCode"
            sqlString += " WHERE hs_ErrorLog.hsId = @hsId "
            sqlString += " AND hs_ErrorLog.hselement = @hselement"
            sqlString += " group by   hs_ErrorLog.LogDate"
            sqlString += " , hs_ErrorLog.hsId"
            sqlString += " , hs_ErrorLog.hselement"
            sqlString += " , hs_ErrorLog.elementCode"
            sqlString += " , hs_ErrorLog.errorCode"
            sqlString += " , hs_ErrorLog.errorValue"
            sqlString += " , hs_ErrorCodes.errorLevel"
            sqlString += " , hs_ErrorCodes.DescIT"
            sqlString += " , hs_ErrorCodes.DescEN) as ErrorLog"
            sqlString += " where RowNumber betWeen @first and @last"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@hselement", SqlDbType.NVarChar, 10, ParameterDirection.Input, hselement)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of hs_ErrorLog)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_hs_Errorlog(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_ErrorLog_ListByElement " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "hs_Tickets"
        Private Function preparerecord_hs_Tickets(reader As SqlDataReader) As hs_Tickets
            Dim _TickedId As Integer = 0
            If Not IsDBNull(reader("TicketId")) Then
                _TickedId = CInt(reader("TicketId").ToString)
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _TicketTitle As String = String.Empty
            If Not IsDBNull(reader("TicketTitle")) Then
                _TicketTitle = CStr(reader("TicketTitle"))
            End If

            Dim _Requester As String = String.Empty
            If Not IsDBNull(reader("Requester")) Then
                _Requester = CStr(reader("Requester"))
            End If

            Dim _emailRequester As String = String.Empty
            If Not IsDBNull(reader("emailRequester")) Then
                _emailRequester = CStr(reader("emailRequester"))
            End If

            Dim _DateRequest As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("DateRequest")) Then
                If IsDate(reader("DateRequest")) Then
                    _DateRequest = reader("DateRequest")
                End If
            End If

            Dim _Description As String = String.Empty
            If Not IsDBNull(reader("Description")) Then
                _Description = CStr(reader("Description"))
            End If

            Dim _Executor As String = String.Empty
            If Not IsDBNull(reader("Executor")) Then
                _Executor = CStr(reader("Executor"))
            End If

            Dim _emailExecutor As String = String.Empty
            If Not IsDBNull(reader("emailExecutor")) Then
                _emailExecutor = CStr(reader("emailExecutor"))
            End If

            Dim _DateExecution As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("DateExecution")) Then
                If IsDate(reader("DateExecution")) Then
                    _DateExecution = reader("DateExecution")
                End If
            End If

            Dim _ExecutorComment As String = String.Empty
            If Not IsDBNull(reader("ExecutorComment")) Then
                _ExecutorComment = CStr(reader("ExecutorComment"))
            End If

            Dim _TicketStatus As Integer = 0
            If Not IsDBNull(reader("TicketStatus")) Then
                _TicketStatus = CInt(reader("TicketStatus").ToString)
            End If

            Dim _TicketType As Integer = 0
            If Not IsDBNull(reader("TicketType")) Then
                _TicketType = CInt(reader("TicketType").ToString)
            End If


            Dim elementName As String = String.Empty
            If Not IsDBNull(reader("elementName")) Then
                elementName = CStr(reader("elementName"))
            End If

            Dim elementId As Integer = 0
            If Not IsDBNull(reader("elementId")) Then
                elementId = CInt(reader("elementId").ToString)
            End If

            Dim _obj As hs_Tickets = New hs_Tickets(_TickedId, _
                                                    _hsId, _
                                                    _TicketTitle, _
                                                    _Requester, _
                                                    _emailRequester, _
                                                    _DateRequest, _
                                                    _Description, _
                                                    _Executor, _
                                                    _emailExecutor, _
                                                    _DateExecution, _
                                                    _ExecutorComment, _
                                                    _TicketStatus, _
                                                    _TicketType, _
                                                    elementName, _
                                                    elementId)
            Return _obj
        End Function

        Public Overrides Function hs_Tickets_Add(hsId As Integer, _
                                                 TicketTitle As String, _
                                                 Requester As String, _
                                                 emailRequester As String, _
                                                 Description As String, _
                                                 Executor As String, _
                                                 emailExecutor As String, _
                                                 UserName As String, _
                                                 TicketType As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Tickets]"
            sqlString += " ("
            sqlString += " [hsId],"
            sqlString += " [TicketTitle],"
            sqlString += " [Requester],"
            sqlString += " [emailRequester],"
            sqlString += " [DateRequest],"
            sqlString += " [Description],"
            sqlString += " [Executor],"
            sqlString += " [emailExecutor],"
            sqlString += " [UserName],"
            sqlString += " [TicketType]"
            sqlString += ")"
            sqlString += "     VALUES"
            sqlString += "("
            sqlString += " @hsId,"
            sqlString += " @TicketTitle,"
            sqlString += " @Requester,"
            sqlString += " @emailRequester,"
            sqlString += " getdate(),"
            sqlString += " @Description,"
            sqlString += " @Executor,"
            sqlString += " @emailExecutor,"
            sqlString += " @UserName,"
            sqlString += " @TicketType"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@TicketTitle", SqlDbType.NVarChar, 50, ParameterDirection.Input, TicketTitle)
            AddParamToSQLCmd(sqlCmd, "@Requester", SqlDbType.NVarChar, 50, ParameterDirection.Input, Requester)
            AddParamToSQLCmd(sqlCmd, "@emailRequester", SqlDbType.NVarChar, 255, ParameterDirection.Input, emailRequester)
            AddParamToSQLCmd(sqlCmd, "@Description", SqlDbType.NText, 0, ParameterDirection.Input, Description)
            AddParamToSQLCmd(sqlCmd, "@Executor", SqlDbType.NVarChar, 50, ParameterDirection.Input, Executor)
            AddParamToSQLCmd(sqlCmd, "@emailExecutor", SqlDbType.NVarChar, 255, ParameterDirection.Input, emailExecutor)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@TicketType", SqlDbType.Int, 0, ParameterDirection.Input, TicketType)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_AddBySystem(hsId As Integer, _
                                                         TicketTitle As String, _
                                                         Requester As String, _
                                                         emailRequester As String, _
                                                         Description As String, _
                                                         Executor As String, _
                                                         emailExecutor As String, _
                                                         UserName As String, _
                                                         elementName As String, _
                                                         elementId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Tickets]"
            sqlString += " ("
            sqlString += " [hsId],"
            sqlString += " [TicketTitle],"
            sqlString += " [Requester],"
            sqlString += " [emailRequester],"
            sqlString += " [DateRequest],"
            sqlString += " [Description],"
            sqlString += " [Executor],"
            sqlString += " [emailExecutor],"
            sqlString += " [UserName],"
            sqlString += " [TicketType]"
            sqlString += " [elementName]"
            sqlString += " [elementId]"
            sqlString += ")"
            sqlString += "     VALUES"
            sqlString += "("
            sqlString += " @hsId,"
            sqlString += " @TicketTitle,"
            sqlString += " @Requester,"
            sqlString += " @emailRequester,"
            sqlString += " getdate(),"
            sqlString += " @Description,"
            sqlString += " @Executor,"
            sqlString += " @emailExecutor,"
            sqlString += " @UserName,"
            sqlString += " @TicketType"
            sqlString += " @elementName"
            sqlString += " @elementId"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@TicketTitle", SqlDbType.NVarChar, 50, ParameterDirection.Input, TicketTitle)
            AddParamToSQLCmd(sqlCmd, "@Requester", SqlDbType.NVarChar, 50, ParameterDirection.Input, Requester)
            AddParamToSQLCmd(sqlCmd, "@emailRequester", SqlDbType.NVarChar, 255, ParameterDirection.Input, emailRequester)
            AddParamToSQLCmd(sqlCmd, "@Description", SqlDbType.NText, 0, ParameterDirection.Input, Description)
            AddParamToSQLCmd(sqlCmd, "@Executor", SqlDbType.NVarChar, 50, ParameterDirection.Input, Executor)
            AddParamToSQLCmd(sqlCmd, "@emailExecutor", SqlDbType.NVarChar, 255, ParameterDirection.Input, emailExecutor)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@TicketType", SqlDbType.Int, 0, ParameterDirection.Input, 1)
            AddParamToSQLCmd(sqlCmd, "@elementName", SqlDbType.NVarChar, 0, ParameterDirection.Input, elementName)
            AddParamToSQLCmd(sqlCmd, "@elementId", SqlDbType.Int, 0, ParameterDirection.Input, elementId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_AddBySystem " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_Del(TicketId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Tickets]"
            sqlString += " where TicketId=@TicketId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@TicketId", SqlDbType.Int, 0, ParameterDirection.Input, TicketId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_getTotOpen(hsid As Integer) As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "select count([TicketId]) as tot from hs_Tickets"
            sqlString += " where hsId=@hsId"
            sqlString += "  and TicketStatus<=1"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsid)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("tot")) Then
                        retVal = CInt(reader("tot"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("hs_Tickets_getTotOpen " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_List(hsId As Integer, TicketStatus As Integer, searchString As String) As List(Of hs_Tickets)
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_Tickets"
            sqlString += " where hsId=@hsId"
            sqlString += "   and TicketStatus=@TicketStatus"
            If Not String.IsNullOrEmpty(searchString) Then
                sqlString += "	 and (TicketTitle like @SearchString or"
                sqlString += "	     Requester like @SearchString or"
                sqlString += "	     Executor like @SearchString)"
            End If

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@TicketStatus", SqlDbType.Int, 0, ParameterDirection.Input, TicketStatus)
            If Not String.IsNullOrEmpty(searchString) Then
                AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(searchString) & "%")
            End If

            Dim _List As New List(Of hs_Tickets)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_Tickets(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Tickets_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Tickets_ListPaged(hsId As Integer, TicketStatus As Integer, searchString As String, RowNumber As Integer) As List(Of hs_Tickets)
            Dim sqlString As String = String.Empty
            sqlString = "select RowNumber ,hs_Tickets.*"
            sqlString += " from"
            sqlString += " (select row_number() over(order by hs_Tickets.TicketId) as rownumber ,hs_Tickets.* "
            sqlString += "from hs_Tickets"
            sqlString += " where hsId=@hsId"
            sqlString += "   and TicketStatus=@TicketStatus"
            If Not String.IsNullOrEmpty(searchString) Then
                sqlString += "	 and (TicketTitle like @SearchString or"
                sqlString += "	     Requester like @SearchString or"
                sqlString += "	     Executor like @SearchString)"
            End If
            sqlString += " ) as hs_Tickets"
            sqlString += " where RowNumber betWeen @first and @last"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@TicketStatus", SqlDbType.Int, 0, ParameterDirection.Input, TicketStatus)
            If Not String.IsNullOrEmpty(searchString) Then
                AddParamToSQLCmd(sqlCmd, "SearchString", SqlDbType.NVarChar, 50, ParameterDirection.Input, "%" & RTrim(searchString) & "%")
            End If
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, RowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, RowNumber + 100)

            Dim _List As New List(Of hs_Tickets)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_Tickets(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Tickets_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Tickets_Read(TicketId As Integer) As hs_Tickets
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_Tickets"
            sqlString += " where TicketId=@TicketId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@TicketId", SqlDbType.Int, 0, ParameterDirection.Input, TicketId)

            Dim _List As New List(Of hs_Tickets)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Tickets(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Tickets_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Tickets_readByElement(hsId As Integer, elementName As String, elementId As Integer) As hs_Tickets
            Dim sqlString As String = String.Empty
            sqlString = "select top 1 * from hs_Tickets"
            sqlString += " where hsId=@hsId"
            sqlString += "   and elementName=@elementName"
            sqlString += "   and elementId=@elementId"
            sqlString += "   and TicketStatus<5"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@elementName", SqlDbType.NVarChar, 0, ParameterDirection.Input, elementName)
            AddParamToSQLCmd(sqlCmd, "@elementId", SqlDbType.Int, 0, ParameterDirection.Input, elementId)

            Dim _List As New List(Of hs_Tickets)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Tickets(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Tickets_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Tickets_Update(TicketId As Integer, _
                                                    TicketTitle As String, _
                                                    Requester As String, _
                                                    emailRequester As String, _
                                                    Description As String, _
                                                    Executor As String, _
                                                    emailExecutor As String, _
                                                    UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Tickets]"
            sqlString += " set TicketTitle=@TicketTitle"
            sqlString += " , Requester=@Requester"
            sqlString += " , emailRequester=@emailRequester"
            sqlString += " , Description=@Description"
            sqlString += " , Executor=@Executor"
            sqlString += " , emailExecutor=@emailExecutor"
            sqlString += " , UserName=@UserName"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where TicketId=@TicketId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@TicketId", SqlDbType.Int, 0, ParameterDirection.Input, TicketId)
            AddParamToSQLCmd(sqlCmd, "@TicketTitle", SqlDbType.NVarChar, 50, ParameterDirection.Input, TicketTitle)
            AddParamToSQLCmd(sqlCmd, "@Requester", SqlDbType.NVarChar, 50, ParameterDirection.Input, Requester)
            AddParamToSQLCmd(sqlCmd, "@emailRequester", SqlDbType.NVarChar, 255, ParameterDirection.Input, emailRequester)
            AddParamToSQLCmd(sqlCmd, "@Description", SqlDbType.NText, 0, ParameterDirection.Input, Description)
            AddParamToSQLCmd(sqlCmd, "@Executor", SqlDbType.NVarChar, 50, ParameterDirection.Input, Executor)
            AddParamToSQLCmd(sqlCmd, "@emailExecutor", SqlDbType.NVarChar, 255, ParameterDirection.Input, emailExecutor)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_ChangeStatus(TicketId As Integer, DateExecution As Date, ExecutorComment As String, TicketStatus As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Tickets]"
            sqlString += " set DateExecution=@DateExecution"
            sqlString += " , ExecutorComment=@ExecutorComment"
            sqlString += " , TicketStatus=@TicketStatus"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where TicketId=@TicketId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@TicketId", SqlDbType.Int, 0, ParameterDirection.Input, TicketId)
            AddParamToSQLCmd(sqlCmd, "@DateExecution", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, DateExecution)
            AddParamToSQLCmd(sqlCmd, "@TicketStatus", SqlDbType.TinyInt, 0, ParameterDirection.Input, TicketStatus)
            AddParamToSQLCmd(sqlCmd, "@ExecutorComment", SqlDbType.NText, 0, ParameterDirection.Input, ExecutorComment)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_ChangeStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_setElement(TicketId As Integer, elementName As String, elementId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Tickets]"
            sqlString += " set elementName=@elementName"
            sqlString += " , elementId=@elementId"
            sqlString += "   where TicketId=@TicketId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@TicketId", SqlDbType.Int, 0, ParameterDirection.Input, TicketId)
            AddParamToSQLCmd(sqlCmd, "@elementName", SqlDbType.NVarChar, 0, ParameterDirection.Input, elementName)
            AddParamToSQLCmd(sqlCmd, "@elementId", SqlDbType.Int, 0, ParameterDirection.Input, elementId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_setElement " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_Tickets_Executors"
        Private Function preparerecord_hs_Tickets_Executors(reader As SqlDataReader) As hs_Tickets_Executors
            Dim _Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                _Id = CInt(reader("Id").ToString)
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _IdContatto As Integer = 0
            If Not IsDBNull(reader("IdContatto")) Then
                _IdContatto = CInt(reader("IdContatto").ToString)
            End If

            Dim _emailaddress As String = String.Empty
            If Not IsDBNull(reader("emailaddress")) Then
                _emailaddress = CStr(reader("emailaddress"))
            End If

            Dim Nome As String = String.Empty
            If Not IsDBNull(reader("Nome")) Then
                Nome = CStr(reader("Nome"))
            End If

            Dim Descrizione As String = String.Empty
            If Not IsDBNull(reader("Descrizione")) Then
                Descrizione = CStr(reader("Descrizione"))
            End If

            Dim _obj As hs_Tickets_Executors = New hs_Tickets_Executors(_Id, _hsId, _IdContatto, _emailaddress, Nome, Descrizione)
            Return _obj
        End Function

        Public Overrides Function hs_Tickets_Executors_Add(hsId As Integer, IdContatto As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_Tickets_Executors]"
            sqlString += "(hsId,"
            sqlString += "[IdContatto]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@IdContatto"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@IdContatto", SqlDbType.Int, 0, ParameterDirection.Input, IdContatto)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Executors_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_Executors_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Tickets_Executors]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Executors_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_Executors_List(hsId As Integer) As List(Of hs_Tickets_Executors)
            Dim sqlString As String = String.Empty
            sqlString = "select hs_Tickets_Executors.* "
            sqlString += " ,Impianti_Contatti.emailaddress"
            sqlString += " ,Impianti_Contatti.Nome"
            sqlString += " ,Impianti_Contatti.Descrizione"
            sqlString += " from hs_Tickets_Executors"
            sqlString += " Inner join Impianti_Contatti on Impianti_Contatti.IdContatto=hs_Tickets_Executors.IdContatto"
            sqlString += " where hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of hs_Tickets_Executors)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_Tickets_Executors(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Tickets_Executors_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Tickets_Executors_Read(Id As Integer) As hs_Tickets_Executors
            Dim sqlString As String = String.Empty
            sqlString = "select hs_Tickets_Executors.* "
            sqlString += " ,Impianti_Contatti.emailaddress"
            sqlString += " ,Impianti_Contatti.Nome"
            sqlString += " ,Impianti_Contatti.Descrizione"
            sqlString += " from hs_Tickets_Executors"
            sqlString += " Inner join Impianti_Contatti on Impianti_Contatti.IdContatto=hs_Tickets_Executors.IdContatto"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of hs_Tickets_Executors)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Tickets_Executors(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Tickets_Executors_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function
#End Region
#Region "hs_Tickets_Requesters"
        Private Function preparerecord_hs_Tickets_Requesters(reader As SqlDataReader) As hs_Tickets_Requesters
            Dim _Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                _Id = CInt(reader("Id").ToString)
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _IdContatto As Integer = 0
            If Not IsDBNull(reader("IdContatto")) Then
                _IdContatto = CInt(reader("IdContatto").ToString)
            End If

            Dim _emailaddress As String = String.Empty
            If Not IsDBNull(reader("emailaddress")) Then
                _emailaddress = CStr(reader("emailaddress"))
            End If

            Dim Nome As String = String.Empty
            If Not IsDBNull(reader("Nome")) Then
                Nome = CStr(reader("Nome"))
            End If

            Dim Descrizione As String = String.Empty
            If Not IsDBNull(reader("Descrizione")) Then
                Descrizione = CStr(reader("Descrizione"))
            End If

            Dim _obj As hs_Tickets_Requesters = New hs_Tickets_Requesters(_Id, _hsId, _IdContatto, _emailaddress, Nome, Descrizione)
            Return _obj
        End Function

        Public Overrides Function hs_Tickets_Requesters_Add(hsId As Integer, IdContatto As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_Tickets_Requesters]"
            sqlString += "(hsId,"
            sqlString += "[IdContatto]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@IdContatto"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@IdContatto", SqlDbType.Int, 0, ParameterDirection.Input, IdContatto)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Requesters_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_Requesters_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Tickets_Requesters]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Tickets_Requesters_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Tickets_Requesters_List(hsId As Integer) As List(Of hs_Tickets_Requesters)
            Dim sqlString As String = String.Empty
            sqlString = "select hs_Tickets_Requesters.* "
            sqlString += " ,Impianti_Contatti.emailaddress"
            sqlString += " ,Impianti_Contatti.Nome"
            sqlString += " ,Impianti_Contatti.Descrizione"
            sqlString += " from hs_Tickets_Requesters"
            sqlString += " Inner join Impianti_Contatti on Impianti_Contatti.IdContatto=hs_Tickets_Requesters.IdContatto"
            sqlString += " where hs_Tickets_Requesters.hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of hs_Tickets_Requesters)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_Tickets_Requesters(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Tickets_Requesters_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Tickets_Requesters_Read(Id As Integer) As hs_Tickets_Requesters
            Dim sqlString As String = String.Empty
            sqlString = "select hs_Tickets_Requesters.* "
            sqlString += " ,Impianti_Contatti.emailaddress"
            sqlString += " ,Impianti_Contatti.Nome"
            sqlString += " ,Impianti_Contatti.Descrizione"
            sqlString += " from hs_Tickets_Requesters"
            sqlString += " Inner join Impianti_Contatti on Impianti_Contatti.IdContatto=hs_Tickets_Requesters.IdContatto"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of hs_Tickets_Requesters)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Tickets_Requesters(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Tickets_Requesters_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "hs_UserAlert"
        Private Function preparerecord_hs_UserAlert(reader As SqlDataReader) As hs_UserAlert
            Dim _Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                _Id = CInt(reader("Id").ToString)
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _IdContatto As Integer = 0
            If Not IsDBNull(reader("IdContatto")) Then
                _IdContatto = CInt(reader("IdContatto").ToString)
            End If

            Dim _emailaddress As String = String.Empty
            If Not IsDBNull(reader("emailaddress")) Then
                _emailaddress = CStr(reader("emailaddress"))
            End If
            Dim _obj As hs_UserAlert = New hs_UserAlert(_Id, _hsId, _IdContatto, _emailaddress)
            Return _obj
        End Function

        Public Overrides Function hs_UserAlert_Add(hsId As Integer, IdContatto As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_UserAlert]"
            sqlString += "(hsId,"
            sqlString += "[IdContatto]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@IdContatto"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@IdContatto", SqlDbType.Int, 0, ParameterDirection.Input, IdContatto)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_UserAlert_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_UserAlert_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_UserAlert]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_UserAlert_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_UserAlert_List(hsId As Integer) As List(Of hs_UserAlert)
            Dim sqlString As String = String.Empty
            sqlString = "select hs_UserAlert.* "
            sqlString += " ,Impianti_Contatti.emailaddress"
            sqlString += " from hs_UserAlert"
            sqlString += " Inner join Impianti_Contatti on Impianti_Contatti.IdContatto=hs_UserAlert.IdContatto"
            sqlString += " where hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of hs_UserAlert)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_UserAlert(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_UserAlert_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_UserAlert_Read(Id As Integer) As hs_UserAlert
            Dim sqlString As String = String.Empty
            sqlString = "select hs_UserAlert.* "
            sqlString += " ,Impianti_Contatti.emailaddress"
            sqlString += " from hs_UserAlert"
            sqlString += " Inner join Impianti_Contatti on Impianti_Contatti.IdContatto=hs_UserAlert.IdContatto"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of hs_UserAlert)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_UserAlert(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_UserAlert_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function
#End Region

#Region "HeatingSystem"
        Private Function prepareRecord_HeatingSystem(reader As SqlDataReader) As HeatingSystem
            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId"))
            End If

            Dim _IdImpianto As String = String.Empty
            If Not IsDBNull(reader("IdImpianto")) Then
                _IdImpianto = reader("IdImpianto").ToString
            End If

            Dim _Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                _Descr = CStr(reader("Descr"))
            End If

            Dim _Indirizzo As String = String.Empty
            If Not IsDBNull(reader("Indirizzo")) Then
                _Indirizzo = CStr(reader("Indirizzo"))
            End If

            Dim _Latitude As Decimal = 0
            If Not IsDBNull(reader("Latitude")) Then
                _Latitude = CDec(reader("Latitude"))
            End If

            Dim _Longitude As Decimal = 0
            If Not IsDBNull(reader("Longitude")) Then
                _Longitude = CDec(reader("Longitude"))
            End If

            Dim _AltSLM As Integer = 0
            If Not IsDBNull(reader("AltSLM")) Then
                _AltSLM = CDec(reader("AltSLM"))
            End If

            Dim _MaintenanceMode As Boolean
            If Not IsDBNull(reader("MaintenanceMode")) Then
                _MaintenanceMode = CBool(reader("MaintenanceMode"))
            End If

            Dim _VPNConnectionId As Integer = 0
            If Not IsDBNull(reader("VPNConnectionId")) Then
                _VPNConnectionId = CInt(reader("VPNConnectionId"))
            End If

            Dim _MapId As Integer = 0
            If Not IsDBNull(reader("MapId")) Then
                _MapId = CInt(reader("MapId").ToString)
            End If

            Dim _TempExt As Decimal = 0
            If Not IsDBNull(reader("TempExt")) Then
                _TempExt = CDec(reader("TempExt"))
            End If

            Dim _stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                _stato = CInt(reader("stato").ToString)
            End If

            Dim _Note As String = String.Empty
            If Not IsDBNull(reader("Note")) Then
                _Note = CStr(reader("Note"))
            End If

            Dim _lastRec As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("lastRec")) Then
                If IsDate(reader("lastRec")) Then
                    _lastRec = reader("lastRec")
                End If
            End If

            Dim _isOnline As Boolean
            If Not IsDBNull(reader("isOnline")) Then
                _isOnline = CBool(reader("isOnline"))
            End If

            Dim _isEnabled As Boolean
            If Not IsDBNull(reader("testMode")) Then
                _isEnabled = CBool(reader("testMode"))
            End If

            Dim _IwMonitoringId As Integer = 0
            If Not IsDBNull(reader("IwMonitoringId")) Then
                _IwMonitoringId = CInt(reader("IwMonitoringId").ToString)
            End If

            Dim _IwMonitoringDes As String = String.Empty
            If Not IsDBNull(reader("IwMonitoringDes")) Then
                _IwMonitoringDes = CStr(reader("IwMonitoringDes"))
            End If

            Dim _DesImpianto As String = String.Empty
            If Not IsDBNull(reader("DesImpianto")) Then
                _DesImpianto = CStr(reader("DesImpianto"))
            End If


            Dim _obj As HeatingSystem = New HeatingSystem(_hsId, _
                                                          _IdImpianto, _
                                                          _Descr, _
                                                          _Indirizzo, _
                                                          _Latitude, _
                                                          _Longitude, _
                                                          _AltSLM, _
                                                          _MaintenanceMode, _
                                                          _VPNConnectionId, _
                                                          _MapId, _
                                                          _TempExt, _
                                                          _stato, _
                                                          _Note, _
                                                          _lastRec, _
                                                          _isOnline, _
                                                          _isEnabled, _
                                                          _IwMonitoringId, _
                                                          _IwMonitoringDes, _
                                                          _DesImpianto)
            Return _obj
        End Function

        Public Overrides Function HeatingSystem_Add(IdImpianto As String, _
                                                    Descr As String, _
                                                    Indirizzo As String, _
                                                    Latitude As Decimal, _
                                                    Longitude As Decimal, _
                                                    AltSLM As Integer, _
                                                    UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [HeatingSystem]"
            sqlString += "           ([IdImpianto]"
            sqlString += "           ,[Descr]"
            sqlString += "           ,[Indirizzo]"
            sqlString += "           ,[Latitude]"
            sqlString += "           ,[Longitude]"
            sqlString += "           ,[AltSLM]"
            sqlString += "           ,[UserName])"
            sqlString += "     VALUES"
            sqlString += "           (@IdImpianto"
            sqlString += "           ,@Descr"
            sqlString += "           ,@Indirizzo"
            sqlString += "           ,@Latitude"
            sqlString += "           ,@Longitude"
            sqlString += "           ,@AltSLM"
            sqlString += "           ,@UserName)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@Indirizzo", SqlDbType.NVarChar, 255, ParameterDirection.Input, Indirizzo)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Longitude)
            AddParamToSQLCmd(sqlCmd, "@AltSLM", SqlDbType.Decimal, 0, ParameterDirection.Input, AltSLM)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_Del(hsId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [HeatingSystem]"
            sqlString += "where hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("HeatingSystem_Del " & ex.Message)
                    retVal = False
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_getTotMap(IdImpianto As String) As Integer
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "select count([MapId])  as totMap from HeatingSystem"
            sqlString += " where IdImpianto=@IdImpianto"
            sqlString += " and isnull(MapId,0)>0"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    If Not IsDBNull(reader("totMap")) Then
                        retVal = CInt(reader("totMap"))
                    End If
                End If
            Catch ex As Exception
                scriviLog("HeatingSystem_getTotMap " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_List(IdImpianto As String) As List(Of HeatingSystem)
            Dim sqlString As String = String.Empty
            sqlString = "select HeatingSystem.*,Impianti.DesImpianto"
            sqlString += " from HeatingSystem"
            sqlString += " inner join Impianti on Impianti.IdImpianto=HeatingSystem.IdImpianto"
            sqlString += " where HeatingSystem.IdImpianto = @IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)


            Dim _list As New List(Of HeatingSystem)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_HeatingSystem(reader))
                Loop
            Catch ex As Exception
                scriviLog("HeatingSystem_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function HeatingSystem_ListEnabled(IdImpianto As String) As List(Of HeatingSystem)
            Dim sqlString As String = String.Empty
            sqlString = "select HeatingSystem.*,Impianti.DesImpianto"
            sqlString += " from HeatingSystem"
            sqlString += " inner join Impianti on Impianti.IdImpianto=HeatingSystem.IdImpianto"
            sqlString += " where HeatingSystem.IdImpianto = @IdImpianto"
            sqlString += "   and HeatingSystem.testMode=1"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim MyGuid As Guid = New Guid(IdImpianto)
            AddParamToSQLCmd(sqlCmd, "@IdImpianto", SqlDbType.UniqueIdentifier, 0, ParameterDirection.Input, MyGuid)


            Dim _list As New List(Of HeatingSystem)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_HeatingSystem(reader))
                Loop
            Catch ex As Exception
                scriviLog("HeatingSystem_ListEnabled " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function HeatingSystem_ListAll() As List(Of HeatingSystem)
            Dim sqlString As String = String.Empty
            sqlString = "select HeatingSystem.*,Impianti.DesImpianto"
            sqlString += " from HeatingSystem"
            sqlString += " inner join Impianti on Impianti.IdImpianto=HeatingSystem.IdImpianto"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _list As New List(Of HeatingSystem)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_HeatingSystem(reader))
                Loop
            Catch ex As Exception
                scriviLog("HeatingSystem_ListAll " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function HeatingSystem_Read(hsId As Integer) As HeatingSystem
            Dim sqlString As String = String.Empty
            sqlString = "select HeatingSystem.*,Impianti.DesImpianto"
            sqlString += " from HeatingSystem"
            sqlString += " inner join Impianti on Impianti.IdImpianto=HeatingSystem.IdImpianto"
            sqlString += " where HeatingSystem.hsId = @hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)


            Dim _list As New List(Of HeatingSystem)
            Dim connection As New SqlConnection(DataAccess.SCP)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(prepareRecord_HeatingSystem(reader))
                End If
            Catch ex As Exception
                scriviLog("HeatingSystem_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function HeatingSystem_setMaintenanceMode(hsId As Integer, MaintenanceMode As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set MaintenanceMode=@MaintenanceMode"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@MaintenanceMode", SqlDbType.Bit, 0, ParameterDirection.Input, MaintenanceMode)
            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setMaintenanceMode " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_setisOnline(hsId As Integer, isOnline As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set isOnline=isOnline"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@isOnline", SqlDbType.Bit, 0, ParameterDirection.Input, isOnline)
            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setisOnline " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_setIwMonitoring(hsId As Integer, IwMonitoringId As Integer, IwMonitoringDes As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set IwMonitoringId=@IwMonitoringId"
            sqlString += " ,IwMonitoringDes=@IwMonitoringDes"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@IwMonitoringId", SqlDbType.Int, 0, ParameterDirection.Input, IwMonitoringId)
            AddParamToSQLCmd(sqlCmd, "@IwMonitoringDes", SqlDbType.NChar, 50, ParameterDirection.Input, IwMonitoringDes)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setIwMonitoring " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_setlastRec(hsId As Integer, lastRec As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set lastRec=@lastRec"
            sqlString += " , isOnline=1"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@lastRec", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, lastRec)
            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setlastRec " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_setNote(hsId As Integer, Note As String, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set Note=@Note"
            sqlString += " , UserName=@UserName"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Note", SqlDbType.NText, 0, ParameterDirection.Input, Note)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setNote " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_setStatus(hsId As Integer, stato As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set stato=@stato"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_setIsEnabled(hsId As Integer, isEnabled As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set testMode=@isEnabled"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@isEnabled", SqlDbType.Int, 0, ParameterDirection.Input, isEnabled)
            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_setIsEnabled " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function HeatingSystem_Update(hsId As Integer, _
                                                       Descr As String, _
                                                       Indirizzo As String, _
                                                       Latitude As Decimal, _
                                                       Longitude As Decimal, _
                                                       AltSLM As Integer, _
                                                       VPNConnectionId As Integer, _
                                                       MapId As Integer, _
                                                       UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update HeatingSystem"
            sqlString += " set Descr=@Descr"
            sqlString += " ,Indirizzo=@Indirizzo"
            sqlString += " ,Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += " ,AltSLM=@AltSLM"
            sqlString += " ,VPNConnectionId=@VPNConnectionId"
            sqlString += " ,MapId=@MapId"
            sqlString += " , UserName=@UserName"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += " where [hsId]=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@Indirizzo", SqlDbType.NVarChar, 255, ParameterDirection.Input, Indirizzo)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 0, ParameterDirection.Input, Longitude)
            AddParamToSQLCmd(sqlCmd, "@AltSLM", SqlDbType.Decimal, 0, ParameterDirection.Input, AltSLM)
            AddParamToSQLCmd(sqlCmd, "@VPNConnectionId", SqlDbType.Int, 0, ParameterDirection.Input, VPNConnectionId)
            AddParamToSQLCmd(sqlCmd, "@MapId", SqlDbType.Int, 0, ParameterDirection.Input, MapId)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.scp)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    Throw New Exception("HeatingSystem_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "hs_Controller"
        Private Function preparerecord_hs_Controller(reader As SqlDataReader) As hs_Controller
            Dim _ControllerId As Integer = 0
            If Not IsDBNull(reader("ControllerId")) Then
                _ControllerId = CInt(reader("ControllerId").ToString)
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _ControllerDescr As String = String.Empty
            If Not IsDBNull(reader("ControllerDescr")) Then
                _ControllerDescr = reader("ControllerDescr")
            End If

            Dim _NoteInterne As String = String.Empty
            If Not IsDBNull(reader("NoteInterne")) Then
                _NoteInterne = reader("NoteInterne")
            End If

            Dim _obj As hs_Controller = New hs_Controller(_ControllerId, _hsId, _ControllerDescr, _NoteInterne)
            Return _obj
        End Function

        Public Overrides Function hs_Controller_Add(hsId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_Controller]"
            sqlString += "(hsId,"
            sqlString += "[ControllerDescr],"
            sqlString += "[NoteInterne]"
            sqlString += "[UserName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@ControllerDescr,"
            sqlString += "@NoteInterne"
            sqlString += "@UserName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@ControllerDescr", SqlDbType.NVarChar, 50, ParameterDirection.Input, ControllerDescr)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Controller_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Controller_Del(ControllerId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Controller]"
            sqlString += " where ControllerId=@ControllerId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ControllerId", SqlDbType.Int, 0, ParameterDirection.Input, ControllerId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Controller_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Controller_List(hsId As Integer) As List(Of hs_Controller)
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_Controller"
            sqlString += " where hsId=@hsId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of hs_Controller)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_Controller(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Controller_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Controller_Read(ControllerId As Integer) As hs_Controller
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_Controller"
            sqlString += " where ControllerId=@ControllerId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ControllerId", SqlDbType.Int, 0, ParameterDirection.Input, ControllerId)

            Dim _List As New List(Of hs_Controller)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Controller(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Controller_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Controller_Update(ControllerId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Controller]"
            sqlString += " set ControllerDescr=@ControllerDescr"
            sqlString += " , NoteInterne=@NoteInterne"
            sqlString += " , UserName=@UserName"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where ControllerId=@ControllerId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ControllerId", SqlDbType.Int, 0, ParameterDirection.Input, ControllerId)
            AddParamToSQLCmd(sqlCmd, "@ControllerDescr", SqlDbType.NVarChar, 50, ParameterDirection.Input, ControllerDescr)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Controller_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_ControllerDetail"
        Private Function preparerecord_hs_ControllerDetail(reader As SqlDataReader) As hs_ControllerDetail
            Dim _id As Integer = 0
            If Not IsDBNull(reader("id")) Then
                _id = CInt(reader("id").ToString)
            End If

            Dim _ControllerId As Integer = 0
            If Not IsDBNull(reader("ControllerId")) Then
                _ControllerId = CInt(reader("ControllerId").ToString)
            End If

            Dim _Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                _Descr = reader("Descr")
            End If

            Dim _NoteInterne As String = String.Empty
            If Not IsDBNull(reader("NoteInterne")) Then
                _NoteInterne = reader("NoteInterne")
            End If

            Dim _qta As Integer = 0
            If Not IsDBNull(reader("qta")) Then
                _qta = CInt(reader("qta").ToString)
            End If

            Dim _obj As hs_ControllerDetail = New hs_ControllerDetail(_id, _ControllerId, _Descr, _NoteInterne, _qta)
            Return _obj
        End Function

        Public Overrides Function hs_ControllerDetail_Add(ControllerId As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_ControllerDetail]"
            sqlString += "(ControllerId,"
            sqlString += "[Descr],"
            sqlString += "[NoteInterne],"
            sqlString += "[qta],"
            sqlString += "[UserName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@ControllerId,"
            sqlString += "@Descr,"
            sqlString += "@NoteInterne,"
            sqlString += "@qta,"
            sqlString += "@UserName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ControllerId", SqlDbType.Int, 0, ParameterDirection.Input, ControllerId)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)
            AddParamToSQLCmd(sqlCmd, "@qta", SqlDbType.Int, 0, ParameterDirection.Input, qta)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ControllerDetail_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_ControllerDetail_Del(id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_ControllerDetail]"
            sqlString += " where id=@id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@id", SqlDbType.Int, 0, ParameterDirection.Input, id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ControllerDetail_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_ControllerDetail_List(ControllerId As Integer) As List(Of hs_ControllerDetail)
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_ControllerDetail"
            sqlString += " where ControllerId=@ControllerId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ControllerId", SqlDbType.Int, 0, ParameterDirection.Input, ControllerId)

            Dim _List As New List(Of hs_ControllerDetail)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_ControllerDetail(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_ControllerDetail_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_ControllerDetail_Read(id As Integer) As hs_ControllerDetail
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_ControllerDetail"
            sqlString += " where id=@id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@id", SqlDbType.Int, 0, ParameterDirection.Input, id)


            Dim _List As New List(Of hs_ControllerDetail)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_ControllerDetail(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_ControllerDetail_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_ControllerDetail_Update(id As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_ControllerDetail]"
            sqlString += " set Descr=@Descr"
            sqlString += " , NoteInterne=@NoteInterne"
            sqlString += " , qta=@qta"
            sqlString += " , UserName=@UserName"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where id=@id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@id", SqlDbType.Int, 0, ParameterDirection.Input, id)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)
            AddParamToSQLCmd(sqlCmd, "@qta", SqlDbType.Int, 0, ParameterDirection.Input, qta)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_ControllerDetail_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "Lux"
        Private Function prepareRecord_Lux(reader As SqlDataReader) As Lux
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim lastReceived As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("lastReceived")) Then
                If IsDate(reader("lastReceived")) Then
                    lastReceived = reader("lastReceived")
                End If
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim Latitude As Decimal = 0
            If Not IsDBNull(reader("Latitude")) Then
                Latitude = CDec(reader("Latitude"))
            End If

            Dim Longitude As Decimal = 0
            If Not IsDBNull(reader("Longitude")) Then
                Longitude = CDec(reader("Longitude"))
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim WorkingTimeCounter As Decimal = 0
            If Not IsDBNull(reader("WorkingTimeCounter")) Then
                WorkingTimeCounter = CDec(reader("WorkingTimeCounter").ToString)
            End If

            Dim PowerOnCycleCounter As Decimal = 0
            If Not IsDBNull(reader("PowerOnCycleCounter")) Then
                PowerOnCycleCounter = CDec(reader("PowerOnCycleCounter").ToString)
            End If

            Dim LightON As Boolean = False
            If Not IsDBNull(reader("LightON")) Then
                LightON = CBool(reader("LightON"))
            End If

            Dim _obj As Lux = New Lux(Id, hsId, Cod, Descr, lastReceived, marcamodello, installationDate, Latitude, Longitude, stato, WorkingTimeCounter, PowerOnCycleCounter, LightON)
            Return _obj

        End Function

        Public Overrides Function Lux_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Lux]"
            sqlString += "(hsId,"
            sqlString += "[Cod],"
            sqlString += "[Descr],"
            sqlString += "[UserName],"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@Cod,"
            sqlString += "@Descr,"
            sqlString += "@UserName,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Lux]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_List(hsId As Integer) As List(Of Lux)
            Dim sqlString As String = String.Empty
            sqlString = "select Lux.*"
            sqlString += " FROM Lux "
            sqlString += " where Lux.hsId=@hsId"
            sqlString += " order by dbo.CodSort(Cod)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of Lux)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_Lux(reader))
                Loop
            Catch ex As Exception
                scriviLog("Lux_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Lux_Read(Id As Integer) As Lux
            Dim sqlString As String = String.Empty
            sqlString = "select Lux.*"
            sqlString += " FROM Lux "
            sqlString += " where Lux.Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of Lux)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Lux(reader))
                End If
            Catch ex As Exception
                scriviLog("Lux_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Lux_ReadByCod(hsId As Integer, Cod As String) As Lux
            Dim sqlString As String = String.Empty
            sqlString = "select Lux.*"
            sqlString += " FROM Lux "
            sqlString += " where Lux.Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _List As New List(Of Lux)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Lux(reader))
                End If
            Catch ex As Exception
                scriviLog("Lux_ReadByCod " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Lux_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux]"
            sqlString += " set Cod=@Cod"
            sqlString += " , Descr=@Descr"
            sqlString += " , UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux]"
            sqlString += " set stato=@stato"
            sqlString += "    ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_setStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_setValue(hsId As Integer, Cod As String, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux]"
            sqlString += "   SET [WorkingTimeCounter] = @WorkingTimeCounter"
            sqlString += "      ,[PowerOnCycleCounter] = @PowerOnCycleCounter"
            sqlString += "      ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)

            AddParamToSQLCmd(sqlCmd, "@WorkingTimeCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, WorkingTimeCounter)
            AddParamToSQLCmd(sqlCmd, "@PowerOnCycleCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, PowerOnCycleCounter)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_setValue " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux]"
            sqlString += " set Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += "   where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Longitude)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_setGeoLocation " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_setLightByCod(hsId As Integer, Cod As String, LightON As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux]"
            sqlString += " set LightON=@LightON"
            sqlString += "   where hsId=@hsId"
            sqlString += "     and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@LightON", SqlDbType.Bit, 0, ParameterDirection.Input, LightON)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_setLightByCod " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "Lux_replacement_history"
        Private Function prepareRecord_Lux_replacement_history(reader As SqlDataReader) As Lux_replacement_history
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim ParentId As Integer = 0
            If Not IsDBNull(reader("ParentId")) Then
                ParentId = CInt(reader("ParentId").ToString)
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim note As String = String.Empty
            If Not IsDBNull(reader("note")) Then
                note = CStr(reader("note"))
            End If

            Dim userName As String = String.Empty
            If Not IsDBNull(reader("userName")) Then
                userName = CStr(reader("userName"))
            End If

            Dim _obj As Lux_replacement_history = New Lux_replacement_history(Id, ParentId, marcamodello, installationDate, note, userName)
            Return _obj
        End Function

        Public Overrides Function Lux_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Lux_replacement_history]"
            sqlString += "(ParentId,"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate],"
            sqlString += "[note],"
            sqlString += "[userName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@ParentId,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate,"
            sqlString += "@note,"
            sqlString += "@userName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_replacement_history_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_replacement_history_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Lux_replacement_history]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_replacement_history_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Lux_replacement_history_List(ParentId As Integer) As List(Of Lux_replacement_history)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM Lux_replacement_history "
            sqlString += " where ParentId=@ParentId"
            sqlString += " order by installationDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)

            Dim _List As New List(Of Lux_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_Lux_replacement_history(reader))
                Loop
            Catch ex As Exception
                scriviLog("Lux_replacement_history_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Lux_replacement_history_Read(Id As Integer) As Lux_replacement_history
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM Lux_replacement_history "
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of Lux_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Lux_replacement_history(reader))
                End If
            Catch ex As Exception
                scriviLog("Lux_replacement_history_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Lux_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux_replacement_history]"
            sqlString += " set UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , note=@note"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_replacement_history_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "log_Lux"
        Private Function prepareRecord_log_Lux(reader As SqlDataReader) As log_Lux
            Dim LogId As Integer = 0
            If Not IsDBNull(reader("LogId")) Then
                LogId = CInt(reader("LogId").ToString)
            End If

            Dim dtLog As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("dtLog")) Then
                If IsDate(reader("dtLog")) Then
                    dtLog = reader("dtLog")
                End If
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim WorkingTimeCounter As Decimal
            If Not IsDBNull(reader("WorkingTimeCounter")) Then
                WorkingTimeCounter = CDec(reader("WorkingTimeCounter").ToString)
            End If

            Dim PowerOnCycleCounter As Decimal
            If Not IsDBNull(reader("PowerOnCycleCounter")) Then
                PowerOnCycleCounter = CDec(reader("PowerOnCycleCounter").ToString)
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim LightON As Boolean = False
            If Not IsDBNull(reader("LightON")) Then
                LightON = CBool(reader("LightON"))
            End If

            Dim _obj As log_Lux = New log_Lux(LogId, dtLog, hsId, Cod, Descr, stato, WorkingTimeCounter, PowerOnCycleCounter, LightON)

            Return _obj
        End Function

        Public Overrides Function log_Lux_Add(hsId As Integer, Cod As String, Descr As String, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, stato As Integer, LightON As Boolean, dtLog As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [Lux]"
            sqlString += "          ([dtLog]"
            sqlString += "           ,[hsId]"
            sqlString += "           ,[Cod]"
            sqlString += "           ,[Descr]"
            sqlString += "           ,[WorkingTimeCounter]"
            sqlString += "           ,[PowerOnCycleCounter]"
            sqlString += "           ,[LightON]"
            sqlString += "           ,[stato])"
            sqlString += "     VALUES"
            sqlString += "           (@dtLog"
            sqlString += "           ,@hsId"
            sqlString += "           ,@Cod"
            sqlString += "           ,@Descr"
            sqlString += "           ,@WorkingTimeCounter"
            sqlString += "           ,@PowerOnCycleCounter"
            sqlString += "           ,@LightON"
            sqlString += "           ,@stato)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@dtLog", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, dtLog)
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 100, ParameterDirection.Input, Descr)

            AddParamToSQLCmd(sqlCmd, "@WorkingTimeCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, WorkingTimeCounter)
            AddParamToSQLCmd(sqlCmd, "@PowerOnCycleCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, PowerOnCycleCounter)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            AddParamToSQLCmd(sqlCmd, "@LightON", SqlDbType.Bit, 0, ParameterDirection.Input, LightON)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_Lux_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function log_Lux_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Lux)
            Dim sqlString As String = String.Empty
            sqlString = "select * from Lux"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and dtLog between @fromDate and @toDate"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of log_Lux)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_Lux(reader))
                    ' _list.Add(New log_cymt100)
                Loop
            Catch ex As Exception
                scriviLog("log_Lux_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Lux_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Lux)
            Dim sqlString As String = String.Empty

            sqlString = "select RowNumber, Lux.*"
            sqlString += " from"
            sqlString += " (select row_number() over(order by Lux.dtLog) as RowNumber"
            sqlString += " ,Lux.*"
            sqlString += " from Lux"
            sqlString += " where Lux.hsId=@hsId"
            sqlString += " and Lux.Cod=@Cod"
            sqlString += " and Lux.dtLog between @fromDate and @toDate) as Lux"
            sqlString += " where RowNumber betWeen @first and @last"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of log_Lux)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_Lux(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_Lux_ListPaged " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Lux_logNotSent(hsId As Integer, Cod As String) As log_Lux
            Dim sqlString As String = String.Empty
            sqlString = "select top 1 * from Lux"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and [sent2Innowatio]=0"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _list As New List(Of log_Lux)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(prepareRecord_log_Lux(reader))
                End If
            Catch ex As Exception
                scriviLog("log_Lux_logNotSent " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Lux_setIsSent(Logid As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Lux]"
            sqlString += " set [sent2Innowatio]=1"
            sqlString += " where Logid=@Logid"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Logid", SqlDbType.Int, 0, ParameterDirection.Input, Logid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_Lux_setIsSent " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "LuxM"
        Private Function prepareRecord_LuxM(reader As SqlDataReader) As LuxM
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim lastReceived As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("lastReceived")) Then
                If IsDate(reader("lastReceived")) Then
                    lastReceived = reader("lastReceived")
                End If
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim Latitude As Decimal = 0
            If Not IsDBNull(reader("Latitude")) Then
                Latitude = CDec(reader("Latitude"))
            End If

            Dim Longitude As Decimal = 0
            If Not IsDBNull(reader("Longitude")) Then
                Longitude = CDec(reader("Longitude"))
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim Voltage As Decimal
            If Not IsDBNull(reader("Voltage")) Then
                Voltage = CDec(reader("Voltage").ToString)
            End If

            Dim Curr As Decimal = 0
            If Not IsDBNull(reader("Curr")) Then
                Curr = CDec(reader("Curr").ToString)
            End If

            Dim EnergyCounter As Decimal = 0
            If Not IsDBNull(reader("EnergyCounter")) Then
                EnergyCounter = CDec(reader("EnergyCounter").ToString)
            End If

            Dim WorkingTimeCounter As Decimal = 0
            If Not IsDBNull(reader("WorkingTimeCounter")) Then
                WorkingTimeCounter = CDec(reader("WorkingTimeCounter").ToString)
            End If

            Dim PowerOnCycleCounter As Decimal = 0
            If Not IsDBNull(reader("PowerOnCycleCounter")) Then
                PowerOnCycleCounter = CDec(reader("PowerOnCycleCounter").ToString)
            End If

            Dim Temp As Decimal = 0
            If Not IsDBNull(reader("Temp")) Then
                Temp = CDec(reader("Temp").ToString)
            End If

            Dim LightON As Boolean = False
            If Not IsDBNull(reader("LightON")) Then
                LightON = CBool(reader("LightON"))
            End If

            Dim CurrentMode As Integer = 0
            If Not IsDBNull(reader("CurrentMode")) Then
                CurrentMode = CInt(reader("CurrentMode").ToString)
            End If

            Dim forcedOn As Boolean = False
            If Not IsDBNull(reader("forcedOn")) Then
                forcedOn = CBool(reader("forcedOn"))
            End If

            Dim forcedOff As Boolean = False
            If Not IsDBNull(reader("forcedOff")) Then
                forcedOff = CBool(reader("forcedOff"))
            End If

            Dim isManual As Boolean = False
            If Not IsDBNull(reader("isManual")) Then
                isManual = CBool(reader("isManual"))
            End If

            Dim _obj As LuxM = New LuxM(Id, hsId, Cod, Descr, lastReceived, marcamodello, installationDate, Latitude, Longitude, stato,
                                        Voltage, Curr, EnergyCounter, WorkingTimeCounter, PowerOnCycleCounter, Temp, LightON,
                                        CurrentMode, forcedOn, forcedOff, isManual)
            Return _obj

        End Function

        Public Overrides Function LuxM_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [LuxM]"
            sqlString += "(hsId,"
            sqlString += "[Cod],"
            sqlString += "[Descr],"
            sqlString += "[UserName],"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@Cod,"
            sqlString += "@Descr,"
            sqlString += "@UserName,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [LuxM]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_List(hsId As Integer) As List(Of LuxM)
            Dim sqlString As String = String.Empty
            sqlString = "select LuxM.*"
            sqlString += " FROM LuxM "
            sqlString += " where LuxM.hsId=@hsId"
            sqlString += " order by dbo.CodSort(Cod)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of LuxM)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_LuxM(reader))
                Loop
            Catch ex As Exception
                scriviLog("LuxM_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function LuxM_Read(Id As Integer) As LuxM
            Dim sqlString As String = String.Empty
            sqlString = "select LuxM.*"
            sqlString += " FROM LuxM "
            sqlString += " where LuxM.Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of LuxM)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_LuxM(reader))
                End If
            Catch ex As Exception
                scriviLog("LuxM_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function LuxM_ReadByCod(hsId As Integer, Cod As String) As LuxM
            Dim sqlString As String = String.Empty
            sqlString = "select LuxM.*"
            sqlString += " FROM LuxM "
            sqlString += " where LuxM.Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _List As New List(Of LuxM)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_LuxM(reader))
                End If
            Catch ex As Exception
                scriviLog("LuxM_ReadByCod " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function LuxM_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += " set Cod=@Cod"
            sqlString += " , Descr=@Descr"
            sqlString += " , UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += " set stato=@stato"
            sqlString += "    ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_setStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_setValue(hsId As Integer, Cod As String, Voltage As Decimal, Curr As Decimal, EnergyCounter As Decimal, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, Temp As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += "   SET [Voltage] = @Voltage"
            sqlString += "      ,[Curr] = @Curr"
            sqlString += "      ,[EnergyCounter] = @EnergyCounter"
            sqlString += "      ,[WorkingTimeCounter] = @WorkingTimeCounter"
            sqlString += "      ,[PowerOnCycleCounter] = @PowerOnCycleCounter"
            sqlString += "      ,[Temp] = @Temp"
            sqlString += "      ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)

            AddParamToSQLCmd(sqlCmd, "@Voltage", SqlDbType.Decimal, 0, ParameterDirection.Input, Voltage)
            AddParamToSQLCmd(sqlCmd, "@Curr", SqlDbType.Decimal, 0, ParameterDirection.Input, Curr)
            AddParamToSQLCmd(sqlCmd, "@EnergyCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, EnergyCounter)
            AddParamToSQLCmd(sqlCmd, "@WorkingTimeCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, WorkingTimeCounter)
            AddParamToSQLCmd(sqlCmd, "@PowerOnCycleCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, PowerOnCycleCounter)
            AddParamToSQLCmd(sqlCmd, "@Temp", SqlDbType.Decimal, 0, ParameterDirection.Input, Temp)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_setValue " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += " set Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += "   where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Longitude)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_setGeoLocation " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_setCurrentMode(Id As Integer, currentMode As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += " set CurrentMode=@CurrentMode"
            sqlString += "    ,lastReceived=@lastReceived"
            sqlString += "   where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@CurrentMode", SqlDbType.Int, 0, ParameterDirection.Input, currentMode)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Lux_setCurrentMode " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_setLightByCod(hsId As Integer, Cod As String, LightON As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += " set LightON=@LightON"
            sqlString += "   where hsId=@hsId"
            sqlString += "     and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@LightON", SqlDbType.Bit, 0, ParameterDirection.Input, LightON)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_setLightByCod " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "LuxM_replacement_history"
        Private Function prepareRecord_LuxM_replacement_history(reader As SqlDataReader) As LuxM_replacement_history
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim ParentId As Integer = 0
            If Not IsDBNull(reader("ParentId")) Then
                ParentId = CInt(reader("ParentId").ToString)
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim note As String = String.Empty
            If Not IsDBNull(reader("note")) Then
                note = CStr(reader("note"))
            End If

            Dim userName As String = String.Empty
            If Not IsDBNull(reader("userName")) Then
                userName = CStr(reader("userName"))
            End If

            Dim _obj As LuxM_replacement_history = New LuxM_replacement_history(Id, ParentId, marcamodello, installationDate, note, userName)
            Return _obj
        End Function

        Public Overrides Function LuxM_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [LuxM_replacement_history]"
            sqlString += "(ParentId,"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate],"
            sqlString += "[note],"
            sqlString += "[userName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@ParentId,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate,"
            sqlString += "@note,"
            sqlString += "@userName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_replacement_history_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_replacement_history_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [LuxM_replacement_history]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_replacement_history_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function LuxM_replacement_history_List(ParentId As Integer) As List(Of LuxM_replacement_history)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM LuxM_replacement_history "
            sqlString += " where ParentId=@ParentId"
            sqlString += " order by installationDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)

            Dim _List As New List(Of LuxM_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_LuxM_replacement_history(reader))
                Loop
            Catch ex As Exception
                scriviLog("LuxM_replacement_history_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function LuxM_replacement_history_Read(Id As Integer) As LuxM_replacement_history
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM LuxM_replacement_history "
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of LuxM_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_LuxM_replacement_history(reader))
                End If
            Catch ex As Exception
                scriviLog("LuxM_replacement_history_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function LuxM_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM_replacement_history]"
            sqlString += " set UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , note=@note"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("LuxM_replacement_history_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "log_LuxM"
        Private Function prepareRecord_log_LuxM(reader As SqlDataReader) As log_LuxM
            Dim LogId As Integer = 0
            If Not IsDBNull(reader("LogId")) Then
                LogId = CInt(reader("LogId").ToString)
            End If

            Dim dtLog As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("dtLog")) Then
                If IsDate(reader("dtLog")) Then
                    dtLog = reader("dtLog")
                End If
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim Voltage As Decimal
            If Not IsDBNull(reader("Voltage")) Then
                Voltage = CDec(reader("Voltage").ToString)
            End If

            Dim Curr As Decimal
            If Not IsDBNull(reader("Curr")) Then
                Curr = CDec(reader("Curr").ToString)
            End If

            Dim EnergyCounter As Decimal = 0
            If Not IsDBNull(reader("EnergyCounter")) Then
                EnergyCounter = CDec(reader("EnergyCounter").ToString)
            End If

            Dim WorkingTimeCounter As Decimal
            If Not IsDBNull(reader("WorkingTimeCounter")) Then
                WorkingTimeCounter = CDec(reader("WorkingTimeCounter").ToString)
            End If

            Dim PowerOnCycleCounter As Decimal
            If Not IsDBNull(reader("PowerOnCycleCounter")) Then
                PowerOnCycleCounter = CDec(reader("PowerOnCycleCounter").ToString)
            End If

            Dim Temp As Decimal
            If Not IsDBNull(reader("Temp")) Then
                Temp = CDec(reader("Temp").ToString)
            End If

            Dim LightON As Boolean = False
            If Not IsDBNull(reader("LightON")) Then
                LightON = CBool(reader("LightON"))
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim _obj As log_LuxM = New log_LuxM(LogId, dtLog, hsId, Cod, Descr, stato, Voltage, Curr, EnergyCounter, WorkingTimeCounter, PowerOnCycleCounter, Temp, LightON)

            Return _obj
        End Function

        Public Overrides Function log_LuxM_Add(hsId As Integer, Cod As String, Descr As String, Voltage As Decimal, Curr As Decimal, EnergyCounter As Decimal, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, Temp As Decimal, stato As Integer, LightON As Boolean, dtLog As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [LuxM]"
            sqlString += "          ([dtLog]"
            sqlString += "           ,[hsId]"
            sqlString += "           ,[Cod]"
            sqlString += "           ,[Descr]"
            sqlString += "           ,[Voltage]"
            sqlString += "           ,[Curr]"
            sqlString += "           ,[EnergyCounter]"
            sqlString += "           ,[WorkingTimeCounter]"
            sqlString += "           ,[PowerOnCycleCounter]"
            sqlString += "           ,[Temp]"
            sqlString += "           ,[LightON]"
            sqlString += "           ,[stato])"
            sqlString += "     VALUES"
            sqlString += "           (@dtLog"
            sqlString += "           ,@hsId"
            sqlString += "           ,@Cod"
            sqlString += "           ,@Descr"
            sqlString += "           ,@Voltage"
            sqlString += "           ,@Curr"
            sqlString += "           ,@EnergyCounter"
            sqlString += "           ,@WorkingTimeCounter"
            sqlString += "           ,@PowerOnCycleCounter"
            sqlString += "           ,@Temp"
            sqlString += "           ,@LightON"
            sqlString += "           ,@stato)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@dtLog", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, dtLog)
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 100, ParameterDirection.Input, Descr)

            AddParamToSQLCmd(sqlCmd, "@Voltage", SqlDbType.Decimal, 0, ParameterDirection.Input, Voltage)
            AddParamToSQLCmd(sqlCmd, "@Curr", SqlDbType.Decimal, 0, ParameterDirection.Input, Curr)
            AddParamToSQLCmd(sqlCmd, "@EnergyCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, EnergyCounter)
            AddParamToSQLCmd(sqlCmd, "@WorkingTimeCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, WorkingTimeCounter)
            AddParamToSQLCmd(sqlCmd, "@PowerOnCycleCounter", SqlDbType.Decimal, 0, ParameterDirection.Input, PowerOnCycleCounter)
            AddParamToSQLCmd(sqlCmd, "@Temp", SqlDbType.Decimal, 0, ParameterDirection.Input, Temp)
            AddParamToSQLCmd(sqlCmd, "@LightON", SqlDbType.Bit, 0, ParameterDirection.Input, LightON)

            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_LuxM_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function log_LuxM_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_LuxM)
            Dim sqlString As String = String.Empty
            sqlString = "select * from LuxM"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and dtLog between @fromDate and @toDate"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of log_LuxM)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_LuxM(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_LuxM_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_LuxM_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_LuxM)
            Dim sqlString As String = String.Empty

            sqlString = "select RowNumber, LuxM.*"
            sqlString += " from"
            sqlString += " (select row_number() over(order by LuxM.dtLog) as RowNumber"
            sqlString += " ,LuxM.*"
            sqlString += " from LuxM"
            sqlString += " where LuxM.hsId=@hsId"
            sqlString += " and LuxM.Cod=@Cod"
            sqlString += " and LuxM.dtLog between @fromDate and @toDate) as LuxM"
            sqlString += " where RowNumber betWeen @first and @last"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of log_LuxM)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_LuxM(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_LuxM_ListPaged " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_LuxM_logNotSent(hsId As Integer, Cod As String) As log_LuxM
            Dim sqlString As String = String.Empty
            sqlString = "select top 1 * from LuxM"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and [sent2Innowatio]=0"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _list As New List(Of log_LuxM)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(prepareRecord_log_LuxM(reader))
                End If
            Catch ex As Exception
                scriviLog("log_Lux_logNotSent " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_LuxM_setIsSent(Logid As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [LuxM]"
            sqlString += " set [sent2Innowatio]=1"
            sqlString += " where Logid=@Logid"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Logid", SqlDbType.Int, 0, ParameterDirection.Input, Logid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_LuxM_setIsSent " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "Psg"
        Private Function prepareRecord_Psg(reader As SqlDataReader) As Psg
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim lastReceived As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("lastReceived")) Then
                If IsDate(reader("lastReceived")) Then
                    lastReceived = reader("lastReceived")
                End If
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim Latitude As Decimal = 0
            If Not IsDBNull(reader("Latitude")) Then
                Latitude = CDec(reader("Latitude"))
            End If

            Dim Longitude As Decimal = 0
            If Not IsDBNull(reader("Longitude")) Then
                Longitude = CDec(reader("Longitude"))
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim currentValue As Integer = 0
            If Not IsDBNull(reader("currentValue")) Then
                currentValue = CInt(reader("currentValue").ToString)
            End If

            Dim _obj As SCP.BLL.Psg = New SCP.BLL.Psg(Id, hsId, Cod, Descr, lastReceived, marcamodello, installationDate, Latitude, Longitude, stato, currentValue)

            ' Dim _obj As Psg = New Psg
            Return _obj

        End Function

        Public Overrides Function Psg_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Psg]"
            sqlString += "(hsId,"
            sqlString += "[Cod],"
            sqlString += "[Descr],"
            sqlString += "[UserName],"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@Cod,"
            sqlString += "@Descr,"
            sqlString += "@UserName,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Psg]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_List(hsId As Integer) As List(Of Psg)
            Dim sqlString As String = String.Empty
            sqlString = "select Psg.*"
            sqlString += " FROM Psg "
            sqlString += " where Psg.hsId=@hsId"
            sqlString += " order by dbo.CodSort(Cod)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of Psg)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    'Dim _obj As SCP.BLL.Psg = prepareRecord_Psg(reader)
                    _List.Add(prepareRecord_Psg(reader))
                Loop
            Catch ex As Exception
                scriviLog("Psg_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Psg_Read(Id As Integer) As Psg
            Dim sqlString As String = String.Empty
            sqlString = "select Psg.*"
            sqlString += " FROM Psg "
            sqlString += " where Psg.Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of Psg)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Psg(reader))
                End If
            Catch ex As Exception
                scriviLog("Psg_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Psg_ReadByCod(hsId As Integer, Cod As String) As Psg
            Dim sqlString As String = String.Empty
            sqlString = "select Psg.*"
            sqlString += " FROM Psg "
            sqlString += " where Psg.Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _List As New List(Of Psg)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Psg(reader))
                End If
            Catch ex As Exception
                scriviLog("Psg_ReadByCod " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Psg_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Psg]"
            sqlString += " set Cod=@Cod"
            sqlString += " , Descr=@Descr"
            sqlString += " , UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Psg]"
            sqlString += " set stato=@stato"
            sqlString += "    ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_setStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_setValue(hsId As Integer, Cod As String, currentValue As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Psg]"
            sqlString += "   SET [currentValue] = @currentValue"
            sqlString += "      ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)

            AddParamToSQLCmd(sqlCmd, "@currentValue", SqlDbType.Int, 0, ParameterDirection.Input, currentValue)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_setValue " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Psg]"
            sqlString += " set Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += "   where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Longitude)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_setGeoLocation " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "Psg_replacement_history"
        Private Function prepareRecord_Psg_replacement_history(reader As SqlDataReader) As Psg_replacement_history
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim ParentId As Integer = 0
            If Not IsDBNull(reader("ParentId")) Then
                ParentId = CInt(reader("ParentId").ToString)
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim note As String = String.Empty
            If Not IsDBNull(reader("note")) Then
                note = CStr(reader("note"))
            End If

            Dim userName As String = String.Empty
            If Not IsDBNull(reader("userName")) Then
                userName = CStr(reader("userName"))
            End If

            Dim _obj As Psg_replacement_history = New Psg_replacement_history(Id, ParentId, marcamodello, installationDate, note, userName)
            Return _obj
        End Function

        Public Overrides Function Psg_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Psg_replacement_history]"
            sqlString += "(ParentId,"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate],"
            sqlString += "[note],"
            sqlString += "[userName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@ParentId,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate,"
            sqlString += "@note,"
            sqlString += "@userName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_replacement_history_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_replacement_history_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Psg_replacement_history]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_replacement_history_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Psg_replacement_history_List(ParentId As Integer) As List(Of Psg_replacement_history)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM Psg_replacement_history "
            sqlString += " where ParentId=@ParentId"
            sqlString += " order by installationDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)

            Dim _List As New List(Of Psg_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_Psg_replacement_history(reader))
                Loop
            Catch ex As Exception
                scriviLog("Psg_replacement_history_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Psg_replacement_history_Read(Id As Integer) As Psg_replacement_history
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM Psg_replacement_history "
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of Psg_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Psg_replacement_history(reader))
                End If
            Catch ex As Exception
                scriviLog("Psg_replacement_history_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Psg_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Psg_replacement_history]"
            sqlString += " set UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , note=@note"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Psg_replacement_history_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "log_Psg"
        Private Function prepareRecord_log_Psg(reader As SqlDataReader) As log_Psg
            Dim LogId As Integer = 0
            If Not IsDBNull(reader("LogId")) Then
                LogId = CInt(reader("LogId").ToString)
            End If

            Dim dtLog As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("dtLog")) Then
                If IsDate(reader("dtLog")) Then
                    dtLog = reader("dtLog")
                End If
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim currentValue As Integer
            If Not IsDBNull(reader("currentValue")) Then
                currentValue = CInt(reader("currentValue").ToString)
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim _obj As log_Psg = New log_Psg(LogId, dtLog, hsId, Cod, Descr, stato, currentValue)

            Return _obj
        End Function

        Public Overrides Function log_Psg_Add(hsId As Integer, Cod As String, Descr As String, currentValue As Integer, stato As Integer, dtLog As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [Psg]"
            sqlString += "          ([dtLog]"
            sqlString += "           ,[hsId]"
            sqlString += "           ,[Cod]"
            sqlString += "           ,[Descr]"
            sqlString += "           ,[currentValue]"
            sqlString += "           ,[stato])"
            sqlString += "     VALUES"
            sqlString += "           (@dtLog"
            sqlString += "           ,@hsId"
            sqlString += "           ,@Cod"
            sqlString += "           ,@Descr"
            sqlString += "           ,@currentValue"
            sqlString += "           ,@stato)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@dtLog", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, dtLog)
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 100, ParameterDirection.Input, Descr)

            AddParamToSQLCmd(sqlCmd, "@currentValue", SqlDbType.Int, 0, ParameterDirection.Input, currentValue)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_Psg_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function log_Psg_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Psg)
            Dim sqlString As String = String.Empty
            sqlString = "select * from Psg"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and dtLog between @fromDate and @toDate"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of log_Psg)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_Psg(reader))
                    ' _list.Add(New log_cymt100)
                Loop
            Catch ex As Exception
                scriviLog("log_Psg_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Psg_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Psg)
            Dim sqlString As String = String.Empty

            sqlString = "select RowNumber, Psg.*"
            sqlString += " from"
            sqlString += " (select row_number() over(order by Psg.dtLog) as RowNumber"
            sqlString += " ,Psg.*"
            sqlString += " from Psg"
            sqlString += " where Psg.hsId=@hsId"
            sqlString += " and Psg.Cod=@Cod"
            sqlString += " and Psg.dtLog between @fromDate and @toDate) as Psg"
            sqlString += " where RowNumber betWeen @first and @last"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of log_Psg)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_Psg(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_Psg_ListPaged " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Psg_logNotSent(hsId As Integer, Cod As String) As log_Psg
            Dim sqlString As String = String.Empty
            sqlString = "select top 1 * from Psg"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and [sent2Innowatio]=0"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _list As New List(Of log_Psg)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(prepareRecord_log_Psg(reader))
                End If
            Catch ex As Exception
                scriviLog("log_Psg_logNotSent " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Psg_setIsSent(Logid As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Psg]"
            sqlString += " set [sent2Innowatio]=1"
            sqlString += " where Logid=@Logid"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Logid", SqlDbType.Int, 0, ParameterDirection.Input, Logid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_Psg_setIsSent " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "Zrel"
        Private Function prepareRecord_Zrel(reader As SqlDataReader) As Zrel
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim lastReceived As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("lastReceived")) Then
                If IsDate(reader("lastReceived")) Then
                    lastReceived = reader("lastReceived")
                End If
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim Latitude As Decimal = 0
            If Not IsDBNull(reader("Latitude")) Then
                Latitude = CDec(reader("Latitude"))
            End If

            Dim Longitude As Decimal = 0
            If Not IsDBNull(reader("Longitude")) Then
                Longitude = CDec(reader("Longitude"))
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If


            Dim LQI As Integer = 0
            If Not IsDBNull(reader("LQI")) Then
                LQI = CDec(reader("LQI"))
            End If

            Dim Temperature As Decimal = 0
            If Not IsDBNull(reader("Temperature")) Then
                Temperature = CDec(reader("Temperature"))
            End If

            Dim MeshParentId As Integer = 0
            If Not IsDBNull(reader("MeshParentId")) Then
                MeshParentId = CInt(reader("MeshParentId").ToString)
            End If

            Dim CurrentId As Integer = 0
            If Not IsDBNull(reader("CurrentId")) Then
                CurrentId = CInt(reader("CurrentId").ToString)
            End If

            Dim _obj As SCP.BLL.Zrel = New SCP.BLL.Zrel(Id, hsId, Cod, Descr, lastReceived, marcamodello, installationDate, Latitude, Longitude, stato, LQI, Temperature, MeshParentId, CurrentId)

            ' Dim _obj As Psg = New Psg
            Return _obj

        End Function

        Public Overrides Function Zrel_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Zrel]"
            sqlString += "(hsId,"
            sqlString += "[Cod],"
            sqlString += "[Descr],"
            sqlString += "[UserName],"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@Cod,"
            sqlString += "@Descr,"
            sqlString += "@UserName,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Zrel]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_List(hsId As Integer) As List(Of Zrel)
            Dim sqlString As String = String.Empty
            sqlString = "select Zrel.*"
            sqlString += " FROM Zrel "
            sqlString += " where Zrel.hsId=@hsId"
            sqlString += " order by dbo.CodSort(Cod)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of Zrel)

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_Zrel(reader))
                Loop
            Catch ex As Exception
                scriviLog("Zrel_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Zrel_Read(Id As Integer) As Zrel
            Dim sqlString As String = String.Empty
            sqlString = "select Zrel.*"
            sqlString += " FROM Zrel "
            sqlString += " where Zrel.Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of Zrel)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Zrel(reader))
                End If
            Catch ex As Exception
                scriviLog("Zrel_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Zrel_ReadByCod(hsId As Integer, Cod As String) As Zrel
            Dim sqlString As String = String.Empty
            sqlString = "select Zrel.*"
            sqlString += " FROM Zrel "
            sqlString += " where Zrel.Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _List As New List(Of Zrel)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Zrel(reader))
                End If
            Catch ex As Exception
                scriviLog("Zrel_ReadByCod " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Zrel_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Zrel]"
            sqlString += " set Cod=@Cod"
            sqlString += " , Descr=@Descr"
            sqlString += " , UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, UCase(Cod))
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, Descr)
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Zrel]"
            sqlString += " set stato=@stato"
            sqlString += "    ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_setStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_setValue(hsId As Integer, Cod As String, LQI As Integer, Temperature As Decimal, MeshParentId As Decimal, CurrentId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Zrel]"
            sqlString += "   SET [LQI] = @LQI"
            sqlString += "      ,Temperature=@Temperature"
            sqlString += "      ,MeshParentId=@MeshParentId"
            sqlString += "      ,CurrentId=@CurrentId"
            sqlString += "      ,lastReceived=@lastReceived"
            sqlString += "   where hsId=@hsId"
            sqlString += "   and Cod=@Cod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)

            AddParamToSQLCmd(sqlCmd, "@LQI", SqlDbType.Int, 0, ParameterDirection.Input, LQI)
            AddParamToSQLCmd(sqlCmd, "@Temperature", SqlDbType.Decimal, 0, ParameterDirection.Input, Temperature)
            AddParamToSQLCmd(sqlCmd, "@MeshParentId", SqlDbType.Decimal, 0, ParameterDirection.Input, MeshParentId)
            AddParamToSQLCmd(sqlCmd, "@CurrentId", SqlDbType.Int, 0, ParameterDirection.Input, CurrentId)
            AddParamToSQLCmd(sqlCmd, "@lastReceived", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, Now)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_setValue " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Zrel]"
            sqlString += " set Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += "   where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Longitude)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_setGeoLocation " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "Zrel_replacement_history"
        Private Function prepareRecord_Zrel_replacement_history(reader As SqlDataReader) As Zrel_replacement_history
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim ParentId As Integer = 0
            If Not IsDBNull(reader("ParentId")) Then
                ParentId = CInt(reader("ParentId").ToString)
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim note As String = String.Empty
            If Not IsDBNull(reader("note")) Then
                note = CStr(reader("note"))
            End If

            Dim userName As String = String.Empty
            If Not IsDBNull(reader("userName")) Then
                userName = CStr(reader("userName"))
            End If

            Dim _obj As Zrel_replacement_history = New Zrel_replacement_history(Id, ParentId, marcamodello, installationDate, note, userName)
            Return _obj
        End Function

        Public Overrides Function Zrel_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [Zrel_replacement_history]"
            sqlString += "(ParentId,"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate],"
            sqlString += "[note],"
            sqlString += "[userName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@ParentId,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate,"
            sqlString += "@note,"
            sqlString += "@userName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_replacement_history_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_replacement_history_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [Zrel_replacement_history]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_replacement_history_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function Zrel_replacement_history_List(ParentId As Integer) As List(Of Zrel_replacement_history)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM Zrel_replacement_history "
            sqlString += " where ParentId=@ParentId"
            sqlString += " order by installationDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)

            Dim _List As New List(Of Zrel_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_Zrel_replacement_history(reader))
                Loop
            Catch ex As Exception
                scriviLog("Zrel_replacement_history_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Zrel_replacement_history_Read(Id As Integer) As Zrel_replacement_history
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM Zrel_replacement_history "
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of Zrel_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_Zrel_replacement_history(reader))
                End If
            Catch ex As Exception
                scriviLog("Zrel_replacement_history_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function Zrel_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Zrel_replacement_history]"
            sqlString += " set UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , note=@note"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("Zrel_replacement_history_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "log_Zrel"
        Private Function prepareRecord_log_Zrel(reader As SqlDataReader) As log_Zrel
            Dim LogId As Integer = 0
            If Not IsDBNull(reader("LogId")) Then
                LogId = CInt(reader("LogId").ToString)
            End If

            Dim dtLog As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("dtLog")) Then
                If IsDate(reader("dtLog")) Then
                    dtLog = reader("dtLog")
                End If
            End If

            Dim hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                hsId = CInt(reader("hsId").ToString)
            End If

            Dim Cod As String = 0
            If Not IsDBNull(reader("Cod")) Then
                Cod = reader("Cod").ToString
            End If

            Dim Descr As String = String.Empty
            If Not IsDBNull(reader("Descr")) Then
                Descr = CStr(reader("Descr"))
            End If

            Dim stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                stato = CInt(reader("stato").ToString)
            End If

            Dim LQI As Integer = 0
            If Not IsDBNull(reader("LQI")) Then
                LQI = CDec(reader("LQI"))
            End If

            Dim Temperature As Decimal = 0
            If Not IsDBNull(reader("Temperature")) Then
                Temperature = CDec(reader("Temperature"))
            End If

            Dim MeshParentId As Integer = 0
            If Not IsDBNull(reader("MeshParentId")) Then
                MeshParentId = CInt(reader("MeshParentId").ToString)
            End If

            Dim CurrentId As Integer = 0
            If Not IsDBNull(reader("CurrentId")) Then
                CurrentId = CInt(reader("CurrentId").ToString)
            End If

            Dim _obj As log_Zrel = New log_Zrel(LogId, dtLog, hsId, Cod, Descr, stato, LQI, Temperature, MeshParentId, CurrentId)

            Return _obj
        End Function

        Public Overrides Function log_Zrel_Add(hsId As Integer, Cod As String, Descr As String, LQI As Integer, Temperature As Decimal, MeshParentId As Integer, CurrentId As Integer, stato As Integer, dtLog As Date) As Boolean

            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [Zrel]"
            sqlString += "          ([dtLog]"
            sqlString += "           ,[hsId]"
            sqlString += "           ,[Cod]"
            sqlString += "           ,[Descr]"
            sqlString += "           ,[LQI]"
            sqlString += "           ,[Temperature]"
            sqlString += "           ,[MeshParentId]"
            sqlString += "           ,[CurrentId]"
            sqlString += "           ,[stato])"
            sqlString += "     VALUES"
            sqlString += "           (@dtLog"
            sqlString += "           ,@hsId"
            sqlString += "           ,@Cod"
            sqlString += "           ,@Descr"
            sqlString += "           ,@LQI"
            sqlString += "           ,@Temperature"
            sqlString += "           ,@MeshParentId"
            sqlString += "           ,@CurrentId"
            sqlString += "           ,@stato)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@dtLog", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, dtLog)
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@Descr", SqlDbType.NVarChar, 100, ParameterDirection.Input, Descr)

            AddParamToSQLCmd(sqlCmd, "@LQI", SqlDbType.Int, 0, ParameterDirection.Input, LQI)
            AddParamToSQLCmd(sqlCmd, "@Temperature", SqlDbType.Decimal, 0, ParameterDirection.Input, Temperature)
            AddParamToSQLCmd(sqlCmd, "@MeshParentId", SqlDbType.Decimal, 0, ParameterDirection.Input, MeshParentId)
            AddParamToSQLCmd(sqlCmd, "@CurrentId", SqlDbType.Decimal, 0, ParameterDirection.Input, CurrentId)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_Zrel_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function log_Zrel_List(hsId As Integer, Cod As String, fromDate As Date, toDate As Date) As List(Of log_Zrel)
            Dim sqlString As String = String.Empty
            sqlString = "select * from Zrel"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and dtLog between @fromDate and @toDate"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of log_Zrel)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_Zrel(reader))
                    ' _list.Add(New log_cymt100)
                Loop
            Catch ex As Exception
                scriviLog("log_Zrel_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Zrel_ListPaged(hsId As Integer, Cod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_Zrel)
            Dim sqlString As String = String.Empty

            sqlString = "select RowNumber, Zrel.*"
            sqlString += " from"
            sqlString += " (select row_number() over(order by Zrel.dtLog) as RowNumber"
            sqlString += " ,Zrel.*"
            sqlString += " from Zrel"
            sqlString += " where Zrel.hsId=@hsId"
            sqlString += " and Zrel.Cod=@Cod"
            sqlString += " and Zrel.dtLog between @fromDate and @toDate) as Zrel"
            sqlString += " where RowNumber betWeen @first and @last"
            sqlString += " order by dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of log_Zrel)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(prepareRecord_log_Zrel(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_Zrel_ListPaged " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Zrel_logNotSent(hsId As Integer, Cod As String) As log_Zrel
            Dim sqlString As String = String.Empty
            sqlString = "select top 1 * from Zrel"
            sqlString += " where hsId=@hsId"
            sqlString += " and Cod=@Cod"
            sqlString += " and [sent2Innowatio]=0"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@Cod", SqlDbType.NVarChar, 10, ParameterDirection.Input, Cod)

            Dim _list As New List(Of log_Zrel)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(prepareRecord_log_Zrel(reader))
                End If
            Catch ex As Exception
                scriviLog("log_Zrel_logNotSent " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_Zrel_setIsSent(Logid As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [Zrel]"
            sqlString += " set [sent2Innowatio]=1"
            sqlString += " where Logid=@Logid"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Logid", SqlDbType.Int, 0, ParameterDirection.Input, Logid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_Zrel_setIsSent " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "hs_Cron"
        Private Function preparerecord_hs_cron(reader As SqlDataReader) As hs_Cron
            Dim _CronId As Integer = 0
            If Not IsDBNull(reader("CronId")) Then
                _CronId = CInt(reader("CronId").ToString)
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _SetPoint As Decimal = 0
            If Not IsDBNull(reader("SetPoint")) Then
                _SetPoint = CDec(reader("SetPoint").ToString)
            End If

            Dim _isOn As Boolean = False
            If Not IsDBNull(reader("isOn")) Then
                _isOn = CDec(reader("isOn").ToString)
            End If

            Dim _TLocal As Decimal = 0
            If Not IsDBNull(reader("TLocal")) Then
                _TLocal = CDec(reader("TLocal").ToString)
            End If

            Dim _CronCod As String = String.Empty
            If Not IsDBNull(reader("CronCod")) Then
                _CronCod = reader("CronCod").ToString
            End If

            Dim _CronDescr As String = String.Empty
            If Not IsDBNull(reader("CronDescr")) Then
                _CronDescr = CStr(reader("CronDescr"))
            End If

            Dim _stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                _stato = CInt(reader("stato").ToString)
            End If

            Dim _NoteInterne As String = String.Empty
            If Not IsDBNull(reader("NoteInterne")) Then
                _NoteInterne = CStr(reader("NoteInterne"))
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim Latitude As Decimal = 0
            If Not IsDBNull(reader("Latitude")) Then
                Latitude = CDec(reader("Latitude"))
            End If

            Dim Longitude As Decimal = 0
            If Not IsDBNull(reader("Longitude")) Then
                Longitude = CDec(reader("Longitude"))
            End If


            Dim CronType As Integer = 0
            If Not IsDBNull(reader("CronType")) Then
                CronType = CInt(reader("CronType"))
            End If

            Dim _obj As hs_Cron = New hs_Cron(_CronId, _hsId, _SetPoint, _isOn, _TLocal, _CronCod, _CronDescr, _stato, _NoteInterne, marcamodello, installationDate, Latitude, Longitude, CronType)
            Return _obj
        End Function

        Public Overrides Function hs_Cron_Add(hsId As Integer, CronCod As String, CronDescr As String, NoteInterne As String, UserName As String, marcamodello As String, installationDate As Date, CronType As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_Cron]"
            sqlString += "(hsId,"
            sqlString += "[CronCod],"
            sqlString += "[CronDescr],"
            sqlString += "[NoteInterne],"
            sqlString += "[UserName],"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate],"
            sqlString += "[CronType]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@CronCod,"
            sqlString += "@CronDescr,"
            sqlString += "@NoteInterne,"
            sqlString += "@UserName,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate,"
            sqlString += "@CronType"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)
            AddParamToSQLCmd(sqlCmd, "@CronDescr", SqlDbType.NVarChar, 50, ParameterDirection.Input, CronDescr)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@CronType", SqlDbType.Int, 0, ParameterDirection.Input, CronType)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Del(CronId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Cron]"
            sqlString += " where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_List(hsId As Integer) As List(Of hs_Cron)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron.*"
            sqlString += " FROM hs_Cron"
            sqlString += " where hs_Cron.hsId=@hsId and CronType=0"
            sqlString += " order by dbo.CodSort(CronCod)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)

            Dim _List As New List(Of hs_Cron)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_cron(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_ListOnOff(hsid As Integer) As List(Of hs_Cron)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron.*"
            sqlString += " FROM hs_Cron"
            sqlString += " where hs_Cron.hsId=@hsId and CronType=1"
            sqlString += " order by dbo.CodSort(CronCod)"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsid)

            Dim _List As New List(Of hs_Cron)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecord_hs_cron(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Read(CronId As Integer) As hs_Cron
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron.*"
            sqlString += " FROM hs_Cron"
            sqlString += " where hs_Cron.CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)

            Dim _List As New List(Of hs_Cron)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_cron(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_ReadByCronCod(hsId As Integer, CronCod As String) As hs_Cron
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron.*"
            sqlString += " FROM hs_Cron"
            sqlString += " where hs_Cron.hsId=@hsId"
            sqlString += "   and hs_Cron.CronCod=@CronCod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)

            Dim _List As New List(Of hs_Cron)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_cron(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_ReadByCronCod " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Update(CronId As Integer, CronCod As String, CronDescr As String, NoteInterne As String, UserName As String, marcamodello As String, installationDate As Date, CronType As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron]"
            sqlString += " set CronCod=@CronCod"
            sqlString += " , CronDescr=@CronDescr"
            sqlString += " , NoteInterne=@NoteInterne"
            sqlString += " , UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , CronType=@CronType"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)
            AddParamToSQLCmd(sqlCmd, "@CronDescr", SqlDbType.NVarChar, 50, ParameterDirection.Input, CronDescr)
            AddParamToSQLCmd(sqlCmd, "@NoteInterne", SqlDbType.NText, 0, ParameterDirection.Input, NoteInterne)
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@CronType", SqlDbType.Int, 0, ParameterDirection.Input, CronType)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_UpdateMarcamodello(Id As Integer, marcamodello As String, installationDate As Date, UserName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron]"
            sqlString += " set UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@UserName", SqlDbType.NVarChar, 256, ParameterDirection.Input, UserName)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_UpdateMarcamodello " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_setStatus(hsId As Integer, CronCod As String, SetPoint As Decimal, stato As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron]"
            sqlString += " set SetPoint=@SetPoint"
            sqlString += " , stato=@stato"
            sqlString += " where hsId=@hsId"
            sqlString += "   and CronCod=@CronCod"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NChar, 10, ParameterDirection.Input, CronCod)
            AddParamToSQLCmd(sqlCmd, "@SetPoint", SqlDbType.Decimal, 0, ParameterDirection.Input, SetPoint)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_setStatus " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_setGeoLocation(CronId As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron]"
            sqlString += " set Latitude=@Latitude"
            sqlString += " ,Longitude=@Longitude"
            sqlString += "   where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@Latitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Latitude)
            AddParamToSQLCmd(sqlCmd, "@Longitude", SqlDbType.Decimal, 10, ParameterDirection.Input, Longitude)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cal_setGeoLocation " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_Cron_replacement_history"
        Private Function prepareRecord_hs_Cron_replacement_history(reader As SqlDataReader) As hs_Cron_replacement_history
            Dim Id As Integer = 0
            If Not IsDBNull(reader("Id")) Then
                Id = CInt(reader("Id").ToString)
            End If

            Dim ParentId As Integer = 0
            If Not IsDBNull(reader("ParentId")) Then
                ParentId = CInt(reader("ParentId").ToString)
            End If

            Dim marcamodello As String = String.Empty
            If Not IsDBNull(reader("marcamodello")) Then
                marcamodello = CStr(reader("marcamodello"))
            End If

            Dim installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("installationDate")) Then
                If IsDate(reader("installationDate")) Then
                    installationDate = reader("installationDate")
                End If
            End If

            Dim note As String = String.Empty
            If Not IsDBNull(reader("note")) Then
                note = CStr(reader("note"))
            End If

            Dim userName As String = String.Empty
            If Not IsDBNull(reader("userName")) Then
                userName = CStr(reader("userName"))
            End If

            Dim _obj As hs_Cron_replacement_history = New hs_Cron_replacement_history(Id, ParentId, marcamodello, installationDate, note, userName)
            Return _obj
        End Function

        Public Overrides Function hs_Cron_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_Cron_replacement_history]"
            sqlString += "(ParentId,"
            sqlString += "[marcamodello],"
            sqlString += "[installationDate],"
            sqlString += "[note],"
            sqlString += "[userName]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@ParentId,"
            sqlString += "@marcamodello,"
            sqlString += "@installationDate,"
            sqlString += "@note,"
            sqlString += "@userName"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_replacement_history_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_replacement_history_Del(Id As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Cron_replacement_history]"
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_replacement_history_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_replacement_history_List(ParentId As Integer) As List(Of hs_Cron_replacement_history)
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM hs_Cron_replacement_history "
            sqlString += " where ParentId=@ParentId"
            sqlString += " order by installationDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@ParentId", SqlDbType.Int, 0, ParameterDirection.Input, ParentId)

            Dim _List As New List(Of hs_Cron_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_hs_Cron_replacement_history(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_replacement_history_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_replacement_history_Read(Id As Integer) As hs_Cron_replacement_history
            Dim sqlString As String = String.Empty
            sqlString = "select *"
            sqlString += " FROM hs_Cron_replacement_history "
            sqlString += " where Id=@Id"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)

            Dim _List As New List(Of hs_Cron_replacement_history)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_hs_Cron_replacement_history(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_replacement_history_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As Date, note As String, userName As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_replacement_history]"
            sqlString += " set UserName=@UserName"
            sqlString += " , marcamodello=@marcamodello"
            sqlString += " , installationDate=@installationDate"
            sqlString += " , note=@note"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += "   where Id=@Id"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@Id", SqlDbType.Int, 0, ParameterDirection.Input, Id)
            AddParamToSQLCmd(sqlCmd, "@marcamodello", SqlDbType.NVarChar, 0, ParameterDirection.Input, marcamodello)
            AddParamToSQLCmd(sqlCmd, "@installationDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, installationDate)
            AddParamToSQLCmd(sqlCmd, "@note", SqlDbType.NText, 0, ParameterDirection.Input, note)
            AddParamToSQLCmd(sqlCmd, "@userName", SqlDbType.NVarChar, 0, ParameterDirection.Input, userName)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_replacement_history_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_Cron_Profile"
        Private Function preparerecordhs_Cron_Profile(reader As SqlDataReader) As hs_Cron_Profile
            Dim CronId As Integer = 0
            If Not IsDBNull(reader("CronId")) Then
                CronId = CInt(reader("CronId").ToString)
            End If

            Dim ProfileY As Integer = 0
            If Not IsDBNull(reader("ProfileY")) Then
                ProfileY = CInt(reader("ProfileY").ToString)
            End If

            Dim ProfileNr As Integer = 0
            If Not IsDBNull(reader("ProfileNr")) Then
                ProfileNr = CInt(reader("ProfileNr").ToString)
            End If

            Dim descr As String = String.Empty
            If Not IsDBNull(reader("realdescr")) Then
                descr = CStr(reader("realdescr"))
            End If

            Dim ProfileData As Decimal() = New Decimal(95) {}
            If Not IsDBNull(reader("ProfileData")) Then
                Dim _ProfileData As String = reader("ProfileData")
                Dim ar() As String = _ProfileData.Split(";")
                For x = 0 To 95
                    ProfileData(x) = CDec(ar(x))
                Next
            End If
            Dim _obj As hs_Cron_Profile = New hs_Cron_Profile(CronId, ProfileY, ProfileNr, descr, ProfileData)
            Return _obj
        End Function

        Public Overrides Function hs_Cron_Profile_Add(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData() As Decimal) As Boolean
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Cron_Profile]"
            sqlString += "           ([CronId]"
            sqlString += "           ,[ProfileY]"
            sqlString += "           ,[ProfileNr]"
            sqlString += "           ,[descr]"
            sqlString += "           ,[ProfileData]"
            sqlString += ")"
            sqlString += "     VALUES"
            sqlString += "           (@CronId"
            sqlString += "           ,@ProfileY"
            sqlString += "           ,@ProfileNr"
            sqlString += "           ,@descr"
            sqlString += "           ,@ProfileData"
            sqlString += ")"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileY", SqlDbType.Int, 0, ParameterDirection.Input, ProfileY)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)
            AddParamToSQLCmd(sqlCmd, "@descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, descr)
            Dim _profile As String = String.Empty
            For x = 0 To 95
                _profile += ProfileData(x).ToString & ";"
            Next
            AddParamToSQLCmd(sqlCmd, "@ProfileData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)
            'sqlCmd.Parameters.AddWithValue("@ProfileData", ProfileData.ToString)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_currentProfile_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Profile_Clear(CronId As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Cron_Profile]"
            sqlString += " where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Clear " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Profile_List(CronId As Integer, ProfileY As Integer) As List(Of hs_Cron_Profile)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron_Profile.*"
            sqlString += "  , hs_Cron_Profile_Descr.descr as realdescr"
            sqlString += " FROM hs_Cron_Profile"
            sqlString += " LEFT OUTER JOIN hs_Cron_Profile_Descr ON hs_cron_Profile.CronId = hs_Cron_Profile_Descr.CronId AND hs_cron_Profile.ProfileNr = hs_Cron_Profile_Descr.ProfileNr"
            sqlString += " where hs_Cron_Profile.CronId=@CronId"
            sqlString += "   and hs_Cron_Profile.ProfileY=@ProfileY"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileY", SqlDbType.Int, 0, ParameterDirection.Input, ProfileY)

            Dim _List As New List(Of hs_Cron_Profile)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecordhs_Cron_Profile(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Read(CronId As Integer, ProfileY As Integer, ProfileNr As Integer) As hs_Cron_Profile
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron_Profile.*"
            sqlString += "  , hs_Cron_Profile_Descr.descr as realdescr"
            sqlString += " FROM hs_Cron_Profile"
            sqlString += " LEFT OUTER JOIN hs_Cron_Profile_Descr ON hs_cron_Profile.CronId = hs_Cron_Profile_Descr.CronId AND hs_cron_Profile.ProfileNr = hs_Cron_Profile_Descr.ProfileNr"
            sqlString += " where hs_Cron_Profile.CronId=@CronId"
            sqlString += "   and hs_Cron_Profile.ProfileY=@ProfileY"
            sqlString += "   and hs_Cron_Profile.ProfileNr=@ProfileNr"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileY", SqlDbType.Int, 0, ParameterDirection.Input, ProfileY)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)

            Dim _List As New List(Of hs_Cron_Profile)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecordhs_Cron_Profile(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Update(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData() As Decimal) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_Profile]"
            sqlString += " set ProfileData=@ProfileData"
            sqlString += " , descr=@descr"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += " where CronId=@CronId"
            sqlString += "   and ProfileY=@ProfileY"
            sqlString += "   and ProfileNr=@ProfileNr"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileY", SqlDbType.Int, 0, ParameterDirection.Input, ProfileY)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)
            AddParamToSQLCmd(sqlCmd, "@descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, descr)
            Dim _profile As String = String.Empty
            For x = 0 To 95
                _profile += ProfileData(x).ToString & ";"
            Next
            AddParamToSQLCmd(sqlCmd, "@ProfileData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_Cron_Profile_Descr"
        Private Function preparerecordhs_Cron_Profile_Descr(reader As SqlDataReader) As hs_Cron_Profile_Descr
            Dim CronId As Integer = 0
            If Not IsDBNull(reader("CronId")) Then
                CronId = CInt(reader("CronId").ToString)
            End If

            Dim ProfileNr As Integer = 0
            If Not IsDBNull(reader("ProfileNr")) Then
                ProfileNr = CInt(reader("ProfileNr").ToString)
            End If

            Dim descr As String = String.Empty
            If Not IsDBNull(reader("descr")) Then
                descr = CStr(reader("descr"))
            End If

            Dim _obj As hs_Cron_Profile_Descr = New hs_Cron_Profile_Descr(CronId, ProfileNr, descr)
            Return _obj
        End Function

        Public Overrides Function hs_Cron_Profile_Descr_Add(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Cron_Profile_Descr]"
            sqlString += "           ([CronId]"
            sqlString += "           ,[ProfileNr]"
            sqlString += "           ,[descr]"
            sqlString += ")"
            sqlString += "     VALUES"
            sqlString += "           (@CronId"
            sqlString += "           ,@ProfileNr"
            sqlString += "           ,@descr"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)
            AddParamToSQLCmd(sqlCmd, "@descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, descr)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Descr_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Profile_Descr_Del(CronId As Integer, ProfileNr As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Cron_Profile_Descr]"
            sqlString += " where CronId=@CronId"
            sqlString += "   and ProfileNr=@ProfileNr"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Descr_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Profile_Descr_List(CronId As Integer) As List(Of hs_Cron_Profile_Descr)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT *"
            sqlString += " FROM hs_Cron_Profile_Descr"
            sqlString += " where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)

            Dim _List As New List(Of hs_Cron_Profile_Descr)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(preparerecordhs_Cron_Profile_Descr(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_Descr_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Descr_Read(CronId As Integer, ProfileNr As Integer) As hs_Cron_Profile_Descr
            Dim sqlString As String = String.Empty
            sqlString = "SELECT *"
            sqlString += " FROM hs_Cron_Profile_Descr"
            sqlString += " where CronId=@CronId"
            sqlString += "   and ProfileNr=@ProfileNr"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)

            Dim _List As New List(Of hs_Cron_Profile_Descr)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecordhs_Cron_Profile_Descr(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_Descr_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Descr_Update(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_Profile_Descr]"
            sqlString += " set descr=@descr"
            sqlString += " where CronId=@CronId"
            sqlString += "   and ProfileNr=@ProfileNr"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)
            AddParamToSQLCmd(sqlCmd, "@descr", SqlDbType.NVarChar, 50, ParameterDirection.Input, descr)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Descr_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_Cron_Profile_Tasks"
        Private Function prepareRecord_hs_Cron_Profile_Tasks(reader As SqlDataReader) As hs_Cron_Profile_Tasks
            Dim CronId As Integer = 0
            If Not IsDBNull(reader("CronId")) Then
                CronId = CInt(reader("CronId").ToString)
            End If

            Dim TaskId As Integer = 0
            If Not IsDBNull(reader("TaskId")) Then
                TaskId = CInt(reader("TaskId").ToString)
            End If

            Dim ProfileNr As Integer = 0
            If Not IsDBNull(reader("ProfileNr")) Then
                ProfileNr = CInt(reader("ProfileNr").ToString)
            End If

            Dim Subject As String = String.Empty
            If Not IsDBNull(reader("Subject")) Then
                Subject = CStr(reader("Subject").ToString)
            End If

            Dim StartDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("StartDate")) Then
                If IsDate(reader("StartDate")) Then
                    StartDate = reader("StartDate")
                End If
            End If

            Dim EndDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("EndDate")) Then
                If IsDate(reader("EndDate")) Then
                    EndDate = reader("EndDate")
                End If
            End If

            Dim RecurrencePattern As String = String.Empty
            If Not IsDBNull(reader("RecurrencePattern")) Then
                RecurrencePattern = CStr(reader("RecurrencePattern").ToString)
            End If

            Dim ExceptionAppointments As String = String.Empty
            If Not IsDBNull(reader("ExceptionAppointments")) Then
                ExceptionAppointments = CStr(reader("ExceptionAppointments").ToString)
            End If

            Dim yearsRepeatable As Boolean = False
            If Not IsDBNull(reader("yearsRepeatable")) Then
                yearsRepeatable = CBool(reader("yearsRepeatable"))
            End If

            Dim CronCod As String = String.Empty
            If Not IsDBNull(reader("CronCod")) Then
                CronCod = CStr(reader("CronCod").ToString)
            End If

            Dim CronDescr As String = String.Empty
            If Not IsDBNull(reader("CronDescr")) Then
                CronDescr = CStr(reader("CronDescr").ToString)
            End If

            Dim _obj As hs_Cron_Profile_Tasks = New hs_Cron_Profile_Tasks(TaskId, CronId, ProfileNr, Subject, StartDate, EndDate, RecurrencePattern, ExceptionAppointments, yearsRepeatable, CronCod, CronDescr)
            Return _obj
        End Function

        Public Overrides Function hs_Cron_Profile_Tasks_Add(CronId As Integer,
                                                            ProfileNr As Integer,
                                                            Subject As String,
                                                            StartDate As Date,
                                                            EndDate As Date,
                                                            RecurrencePattern As String,
                                                            ExceptionAppointments As String,
                                                            yearsRepeatable As Boolean) As Boolean
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Cron_Profile_Tasks]"
            sqlString += "           ([CronId]"
            sqlString += "           ,[ProfileNr]"
            sqlString += "           ,[Subject]"
            sqlString += "           ,[StartDate]"
            sqlString += "           ,[EndDate]"
            sqlString += "           ,[RecurrencePattern]"
            sqlString += "           ,[ExceptionAppointments]"
            sqlString += "           ,[yearsRepeatable]"
            sqlString += ")"
            sqlString += "     VALUES"
            sqlString += "           (@CronId"
            sqlString += "           ,@ProfileNr"
            sqlString += "           ,@Subject"
            sqlString += "           ,@StartDate"
            sqlString += "           ,@EndDate"
            sqlString += "           ,@RecurrencePattern"
            sqlString += "           ,@ExceptionAppointments"
            sqlString += "           ,@yearsRepeatable"
            sqlString += ")"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)

            AddParamToSQLCmd(sqlCmd, "@Subject", SqlDbType.NVarChar, 100, ParameterDirection.Input, UCase(Subject))
            AddParamToSQLCmd(sqlCmd, "@StartDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EndDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, EndDate)
            AddParamToSQLCmd(sqlCmd, "@RecurrencePattern", SqlDbType.NVarChar, 255, ParameterDirection.Input, RecurrencePattern)
            AddParamToSQLCmd(sqlCmd, "@ExceptionAppointments", SqlDbType.Text, 0, ParameterDirection.Input, ExceptionAppointments)
            AddParamToSQLCmd(sqlCmd, "@yearsRepeatable", SqlDbType.Bit, 0, ParameterDirection.Input, yearsRepeatable)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Tasks_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Profile_Tasks_Del(CronId As Integer, ProfileNr As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Cron_Profile_Tasks]"
            sqlString += " where CronId=@CronId"
            sqlString += "   and ProfileNr=@ProfileNr"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Tasks_Del " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Profile_Tasks_List(CronId As Integer) As List(Of hs_Cron_Profile_Tasks)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron_Profile_Tasks.*, hs_Cron.CronCod, hs_Cron.CronDescr"
            sqlString += " FROM hs_Cron_Profile_Tasks "
            sqlString += " INNER JOIN hs_Cron ON hs_Cron_Profile_Tasks.CronId = hs_Cron.CronId"
            sqlString += " where hs_Cron_Profile_Tasks.CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)

            Dim _List As New List(Of hs_Cron_Profile_Tasks)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_hs_Cron_Profile_Tasks(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_Tasks_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Tasks_ListAll() As List(Of hs_Cron_Profile_Tasks)
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron_Profile_Tasks.*, hs_Cron.CronCod, hs_Cron.CronDescr"
            sqlString += " FROM hs_Cron_Profile_Tasks "
            sqlString += " INNER JOIN hs_Cron ON hs_Cron_Profile_Tasks.CronId = hs_Cron.CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            Dim _List As New List(Of hs_Cron_Profile_Tasks)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection


            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                Do While reader.Read
                    _List.Add(prepareRecord_hs_Cron_Profile_Tasks(reader))
                Loop
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_Tasks_ListAll " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Tasks_Read(TaskId As Integer) As hs_Cron_Profile_Tasks
            Dim sqlString As String = String.Empty
            sqlString = "SELECT hs_Cron_Profile_Tasks.*, hs_Cron.CronCod, hs_Cron.CronDescr"
            sqlString += " FROM hs_Cron_Profile_Tasks "
            sqlString += " INNER JOIN hs_Cron ON hs_Cron_Profile_Tasks.CronId = hs_Cron.CronId"
            sqlString += " where hs_Cron_Profile_Tasks.TaskId=@TaskId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@TaskId", SqlDbType.Int, 0, ParameterDirection.Input, TaskId)

            Dim _List As New List(Of hs_Cron_Profile_Tasks)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(prepareRecord_hs_Cron_Profile_Tasks(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_Profile_Tasks_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Profile_Tasks_Update(TaskId As Integer,
                                                               ProfileNr As Integer,
                                                               Subject As String,
                                                               StartDate As Date,
                                                               EndDate As Date,
                                                               RecurrencePattern As String,
                                                               ExceptionAppointments As String,
                                                               yearsRepeatable As Boolean) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_Profile_Tasks]"
            sqlString += " set Subject=@Subject"
            sqlString += "   , ProfileNr=@ProfileNr"
            sqlString += "   ,StartDate=@StartDate"
            sqlString += "   ,EndDate=@EndDate"
            sqlString += "   ,RecurrencePattern=@RecurrencePattern"
            sqlString += "   ,ExceptionAppointments=@ExceptionAppointments"
            sqlString += "   ,yearsRepeatable=@yearsRepeatable"
            sqlString += " where TaskId=@TaskId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@TaskId", SqlDbType.Int, 0, ParameterDirection.Input, TaskId)
            AddParamToSQLCmd(sqlCmd, "@ProfileNr", SqlDbType.Int, 0, ParameterDirection.Input, ProfileNr)
            AddParamToSQLCmd(sqlCmd, "@Subject", SqlDbType.NVarChar, 100, ParameterDirection.Input, UCase(Subject))
            AddParamToSQLCmd(sqlCmd, "@StartDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EndDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, EndDate)
            AddParamToSQLCmd(sqlCmd, "@RecurrencePattern", SqlDbType.NVarChar, 255, ParameterDirection.Input, RecurrencePattern)
            AddParamToSQLCmd(sqlCmd, "@ExceptionAppointments", SqlDbType.Text, 0, ParameterDirection.Input, ExceptionAppointments)
            AddParamToSQLCmd(sqlCmd, "@yearsRepeatable", SqlDbType.Bit, 0, ParameterDirection.Input, yearsRepeatable)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Profile_Tasks_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
#Region "hs_Cron_Calendar"
        Private Function preparerecord_hs_Cron_Calendar(reader As SqlDataReader) As hs_Cron_Calendar
            Dim CronId As Integer = 0
            If Not IsDBNull(reader("CronId")) Then
                CronId = CInt(reader("CronId").ToString)
            End If

            Dim Calyear As Integer = 0
            If Not IsDBNull(reader("Calyear")) Then
                Calyear = CInt(reader("Calyear").ToString)
            End If

            Dim Calmonth As Integer = 0
            If Not IsDBNull(reader("Calmonth")) Then
                Calmonth = CInt(reader("Calmonth").ToString)
            End If

            Dim RealMonthData As Integer() = New Integer(31) {}
            If Not IsDBNull(reader("RealMonthData")) Then
                Dim _monthData As String = reader("RealMonthData")
                Dim ar() As String = _monthData.Split(";")
                For x = 0 To 31
                    RealMonthData(x) = CInt(ar(x))
                Next
            End If

            Dim DesiredMonthData As Integer() = New Integer(31) {}
            If Not IsDBNull(reader("DesiredMonthData")) Then
                Dim _monthData As String = reader("DesiredMonthData")
                Dim ar() As String = _monthData.Split(";")
                For x = 0 To 31
                    DesiredMonthData(x) = CInt(ar(x))
                Next
            End If

            Dim TasksForDesired As Integer() = New Integer(31) {}
            If Not IsDBNull(reader("TasksForDesired")) Then
                Dim _monthData As String = reader("TasksForDesired")
                Dim ar() As String = _monthData.Split(";")
                For x = 0 To 31
                    TasksForDesired(x) = CInt(ar(x))
                Next
            End If

            Dim LastSend As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("LastSend")) Then
                If IsDate(reader("LastSend")) Then
                    LastSend = reader("LastSend")
                End If
            End If

            Dim _obj As hs_Cron_Calendar = New hs_Cron_Calendar(CronId, Calyear, Calmonth, RealMonthData, DesiredMonthData, TasksForDesired, LastSend)
            Return _obj
        End Function

        Public Overrides Function hs_Cron_Calendar_Add(CronId As Integer, Calyear As Integer, Calmonth As Integer, monthData() As Integer) As Boolean
            Dim retVal As Integer = 0
            Dim sqlString As String = String.Empty
            sqlString = "INSERT INTO [hs_Cron_Calendar]"
            sqlString += "           ([CronId]"
            sqlString += "           ,[Calyear]"
            sqlString += "           ,[Calmonth]"
            sqlString += "           ,[RealMonthData]"
            sqlString += "           ,[DesiredMonthData]"
            sqlString += ")"
            sqlString += "     VALUES"
            sqlString += "           (@CronId"
            sqlString += "           ,@Calyear"
            sqlString += "           ,@Calmonth"
            sqlString += "           ,@RealMonthData"
            sqlString += "           ,@DesiredMonthData"
            sqlString += ")"


            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@Calyear", SqlDbType.Int, 0, ParameterDirection.Input, Calyear)
            AddParamToSQLCmd(sqlCmd, "@Calmonth", SqlDbType.Int, 0, ParameterDirection.Input, Calmonth)
            Dim _profile As String = String.Empty
            For x = 0 To 31
                _profile += monthData(x).ToString & ";"
            Next

            Dim TasksForDesired As String = String.Empty
            Dim profile_task_list As List(Of hs_Cron_Profile_Tasks) = hs_Cron_Profile_Tasks.ListByMonth(CronId, Calyear, Calmonth)
            If Not profile_task_list Is Nothing Then
                Dim x As Integer = 0
                For Each _task As hs_Cron_Profile_Tasks In profile_task_list
                    TasksForDesired += _task.ProfileNr & ";"
                    x = x + 1
                    If x > 31 Then Exit For
                Next
                profile_task_list = Nothing
                AddParamToSQLCmd(sqlCmd, "@DesiredMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, TasksForDesired)
            Else
                AddParamToSQLCmd(sqlCmd, "@DesiredMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)
            End If


            AddParamToSQLCmd(sqlCmd, "@RealMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)


            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Calendar_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Calendar_Clear(CronId As Object) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "delete from [hs_Cron_Calendar]"
            sqlString += " where CronId=@CronId"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Calendar_Clear " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Calendar_Read(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_Cron_Calendar
            Dim sqlString As String = String.Empty
            sqlString = "SELECT *"
            sqlString += " FROM hs_Cron_Calendar"
            sqlString += " where CronId=@CronId"
            sqlString += "   and Calyear=@Calyear"
            sqlString += "   and Calmonth=@Calmonth"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@Calyear", SqlDbType.Int, 0, ParameterDirection.Input, Calyear)
            AddParamToSQLCmd(sqlCmd, "@Calmonth", SqlDbType.Int, 0, ParameterDirection.Input, Calmonth)

            Dim _List As New List(Of hs_Cron_Calendar)()

            Dim Connection As New SqlConnection(DataAccess.SCP)
            Connection.Open()
            sqlCmd.Connection() = Connection

            Dim reader As SqlDataReader = sqlCmd.ExecuteReader()
            Try
                If reader.Read Then
                    _List.Add(preparerecord_hs_Cron_Calendar(reader))
                End If
            Catch ex As Exception
                scriviLog("hs_Cron_Calendar_Read " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not Connection Is Nothing) Then
                Connection.Close()
                Connection = Nothing
            End If
            If _List.Count > 0 Then
                Return _List(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function hs_Cron_Calendar_Update(CronId As Integer, Calyear As Integer, Calmonth As Integer, TasksForDesired As Integer(), DesiredMonthData As Integer()) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_Calendar]"
            sqlString += " set DesiredMonthData=@DesiredMonthData"
            sqlString += " , TasksForDesired=@TasksForDesired"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += " where CronId=@CronId"
            sqlString += "   and Calyear=@Calyear"
            sqlString += "   and Calmonth=@Calmonth"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@Calyear", SqlDbType.Int, 0, ParameterDirection.Input, Calyear)
            AddParamToSQLCmd(sqlCmd, "@Calmonth", SqlDbType.Int, 0, ParameterDirection.Input, Calmonth)
            Dim str_TasksForDesired As String = String.Empty
            Dim str_DesiredMonthData As String = String.Empty
            For x = 0 To 31
                str_TasksForDesired += TasksForDesired(x).ToString & ";"
                str_DesiredMonthData += DesiredMonthData(x).ToString & ";"
            Next
            AddParamToSQLCmd(sqlCmd, "@TasksForDesired", SqlDbType.NVarChar, 0, ParameterDirection.Input, str_TasksForDesired)
            AddParamToSQLCmd(sqlCmd, "@DesiredMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, str_DesiredMonthData)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Calendar_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Calendar_UpdateDesired(CronId As Integer, Calyear As Integer, Calmonth As Integer, DesiredMonthData() As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_Calendar]"
            sqlString += " set DesiredMonthData=@DesiredMonthData"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += " where CronId=@CronId"
            sqlString += "   and Calyear=@Calyear"
            sqlString += "   and Calmonth=@Calmonth"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@Calyear", SqlDbType.Int, 0, ParameterDirection.Input, Calyear)
            AddParamToSQLCmd(sqlCmd, "@Calmonth", SqlDbType.Int, 0, ParameterDirection.Input, Calmonth)
            Dim _profile As String = String.Empty
            For x = 0 To 31
                _profile += DesiredMonthData(x).ToString & ";"
            Next
            AddParamToSQLCmd(sqlCmd, "@RealMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)
            AddParamToSQLCmd(sqlCmd, "@DesiredMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Calendar_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function hs_Cron_Calendar_UpdateReal(CronId As Integer, Calyear As Integer, Calmonth As Integer, RealMonthData() As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron_Calendar]"
            sqlString += " set RealMonthData=@RealMonthData"
            ' sqlString += " , DesiredMonthData=@DesiredMonthData"
            sqlString += " , LastUpdate=LastUpdate+1"
            sqlString += " , dtUpdate=getdate()"
            sqlString += " where CronId=@CronId"
            sqlString += "   and Calyear=@Calyear"
            sqlString += "   and Calmonth=@Calmonth"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@CronId", SqlDbType.Int, 0, ParameterDirection.Input, CronId)
            AddParamToSQLCmd(sqlCmd, "@Calyear", SqlDbType.Int, 0, ParameterDirection.Input, Calyear)
            AddParamToSQLCmd(sqlCmd, "@Calmonth", SqlDbType.Int, 0, ParameterDirection.Input, Calmonth)
            Dim _profile As String = String.Empty
            For x = 0 To 31
                _profile += RealMonthData(x).ToString & ";"
            Next
            AddParamToSQLCmd(sqlCmd, "@RealMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)
            ' AddParamToSQLCmd(sqlCmd, "@DesiredMonthData", SqlDbType.NVarChar, 0, ParameterDirection.Input, _profile)

            Using cn As SqlConnection = New SqlConnection(DataAccess.SCP)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("hs_Cron_Calendar_Update " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region

#Region "log_hs_Cron"
        Private Function preparerecord_log_hs_Cron(reader As SqlDataReader) As log_hs_Cron
            Dim _LogId As Integer = 0
            If Not IsDBNull(reader("LogId")) Then
                _LogId = CInt(reader("LogId").ToString)
            End If

            Dim _dtLog As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
            If Not IsDBNull(reader("dtLog")) Then
                If IsDate(reader("dtLog")) Then
                    _dtLog = reader("dtLog")
                End If
            End If

            Dim _hsId As Integer = 0
            If Not IsDBNull(reader("hsId")) Then
                _hsId = CInt(reader("hsId").ToString)
            End If

            Dim _SetPoint As Decimal = 0
            If Not IsDBNull(reader("SetPoint")) Then
                _SetPoint = CDec(reader("SetPoint").ToString)
            End If

            Dim _CronCod As String = String.Empty
            If Not IsDBNull(reader("CronCod")) Then
                _CronCod = reader("CronCod").ToString
            End If

            Dim _CronDescr As String = String.Empty
            If Not IsDBNull(reader("CronDescr")) Then
                _CronDescr = CStr(reader("CronDescr"))
            End If

            Dim _stato As Integer = 0
            If Not IsDBNull(reader("stato")) Then
                _stato = CInt(reader("stato").ToString)
            End If

            Dim _obj As log_hs_Cron = New log_hs_Cron(_LogId, _dtLog, _hsId, _SetPoint, _CronCod, _CronDescr, _stato)
            Return _obj
        End Function

        Public Overrides Function log_hs_Cron_Add(hsId As Integer, SetPoint As Decimal, CronCod As String, CronDescr As String, stato As Integer, dtLog As Date) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "insert into [hs_Cron]"
            sqlString += "(hsId,"
            sqlString += "[SetPoint],"
            sqlString += "[CronCod],"
            sqlString += "[CronDescr],"
            sqlString += "[stato],"
            sqlString += "[dtLog]"
            sqlString += ")"
            sqlString += " values("
            sqlString += "@hsId,"
            sqlString += "@SetPoint,"
            sqlString += "@CronCod,"
            sqlString += "@CronDescr,"
            sqlString += "@stato,"
            sqlString += "@dtLog"
            sqlString += ")"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString

            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)
            AddParamToSQLCmd(sqlCmd, "@CronDescr", SqlDbType.NVarChar, 50, ParameterDirection.Input, CronDescr)
            AddParamToSQLCmd(sqlCmd, "@SetPoint", SqlDbType.Decimal, 0, ParameterDirection.Input, SetPoint)
            AddParamToSQLCmd(sqlCmd, "@stato", SqlDbType.Int, 0, ParameterDirection.Input, stato)
            AddParamToSQLCmd(sqlCmd, "@dtLog", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, dtLog)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_hs_Cron_Add " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function

        Public Overrides Function log_hs_Cron_List(hsId As Integer, CronCod As String, fromDate As Date, toDate As Date) As List(Of log_hs_Cron)
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_Cron"
            sqlString += " where hsId=@hsId"
            sqlString += " and CronCod=@CronCod"
            sqlString += " and dtLog between @fromDate and @toDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of log_hs_Cron)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(preparerecord_log_hs_Cron(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_hs_Cron_List " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_hs_Cron_ListPaged(hsId As Integer, CronCod As String, fromDate As Date, toDate As Date, rowNumber As Integer) As List(Of log_hs_Cron)
            Dim sqlString As String = String.Empty

            sqlString = "select RowNumber, hs_Cron.*"
            sqlString += " from"
            sqlString += " (select row_number() over(order by hs_Cron.LogId) as RowNumber"
            sqlString += " ,hs_Cron.*"
            sqlString += " from hs_Cron"
            sqlString += " where hs_Cron.hsId=@hsId"
            sqlString += " and hs_Cron.CronCod=@CronCod"
            sqlString += " and hs_Cron.dtLog between @fromDate and @toDate) as hs_Cron"
            sqlString += " where RowNumber betWeen @first and @last"
            sqlString += " order by hs_Cron.dtLog"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)
            AddParamToSQLCmd(sqlCmd, "@first", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 1)
            AddParamToSQLCmd(sqlCmd, "@last", SqlDbType.Real, 0, ParameterDirection.Input, rowNumber + 100)

            Dim _list As New List(Of log_hs_Cron)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(preparerecord_log_hs_Cron(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_hs_Cron_ListPaged " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_hs_Cron_ListAll(hsId As Integer, fromDate As Date, toDate As Date) As List(Of log_hs_Cron)
            Dim sqlString As String = String.Empty
            sqlString = "select * from hs_Cron"
            sqlString += " where hsId=@hsId"
            sqlString += " and dtLog between @fromDate and @toDate"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@fromDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, fromDate)
            AddParamToSQLCmd(sqlCmd, "@toDate", SqlDbType.SmallDateTime, 0, ParameterDirection.Input, toDate)

            Dim _list As New List(Of log_hs_Cron)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                Do While reader.Read
                    _list.Add(preparerecord_log_hs_Cron(reader))
                Loop
            Catch ex As Exception
                scriviLog("log_hs_Cron_ListAll " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_hs_Cron_logNotSent(hsId As Integer, CronCod As String) As log_hs_Cron
            Dim sqlString As String = String.Empty
            sqlString = "select top 1 * from hs_Cron"
            sqlString += " where hsId=@hsId"
            sqlString += " and CronCod=@CronCod"
            sqlString += " and [sent2Innowatio]=0"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@hsId", SqlDbType.Int, 0, ParameterDirection.Input, hsId)
            AddParamToSQLCmd(sqlCmd, "@CronCod", SqlDbType.NVarChar, 10, ParameterDirection.Input, CronCod)

            Dim _list As New List(Of log_hs_Cron)
            Dim connection As New SqlConnection(DataAccess.HS_LOG)
            connection.Open()
            sqlCmd.Connection() = connection
            Dim reader As SqlDataReader = sqlCmd.ExecuteReader
            Try
                If reader.Read Then
                    _list.Add(preparerecord_log_hs_Cron(reader))
                End If
            Catch ex As Exception
                scriviLog("log_hs_Cron_logNotSent " & ex.Message)
            End Try
            If (Not reader Is Nothing) Then
                reader.Close()
            End If
            If (Not connection Is Nothing) Then
                connection.Close()
                connection = Nothing
            End If
            If _list.Count > 0 Then
                Return _list(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function log_hs_Cron_setIsSent(Logid As Integer) As Boolean
            Dim retVal As Boolean = False
            Dim sqlString As String = String.Empty
            sqlString = "update [hs_Cron]"
            sqlString += " set [sent2Innowatio]=1"
            sqlString += " where Logid=@Logid"

            Dim sqlCmd As SqlCommand = New SqlCommand()
            sqlCmd.CommandText = sqlString
            AddParamToSQLCmd(sqlCmd, "@Logid", SqlDbType.Int, 0, ParameterDirection.Input, Logid)

            Using cn As SqlConnection = New SqlConnection(DataAccess.HS_LOG)
                Try
                    sqlCmd.Connection = cn
                    cn.Open()
                    sqlCmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    scriviLog("log_hs_Cron_setIsSent " & ex.Message)
                    Return retVal
                End Try
                retVal = True
            End Using
            Return retVal
        End Function
#End Region
    End Class
End Namespace
