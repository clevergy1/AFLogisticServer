Imports System.ServiceModel
Imports System.ServiceModel.Activation
Imports System.ServiceModel.Web
Imports SCP.BLL

Imports System.ServiceModel.Channels
Imports System.Security.Cryptography
Imports System.IO

Imports System.Net.Mail
Imports System.Threading
Imports System.Net

<ServiceContract(Namespace:="")>
<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
Public Class S1

#Region "aspnetroles"
    <OperationContract()>
    <WebGet()>
    Public Function aspnetroles_Add(RoleName As String) As Boolean
        If String.IsNullOrEmpty(RoleName) Then
            Return False
        End If
        Return SCP.BLL.aspnetroles.Add(RoleName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetroles_Del(RoleName As String) As Boolean
        If String.IsNullOrEmpty(RoleName) Then
            Return False
        End If
        Return SCP.BLL.aspnetroles.Del(RoleName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetroles_List() As List(Of aspnetroles)
        Return SCP.BLL.aspnetroles.List
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetroles_ListActive(IdImpianto As String) As List(Of aspnetroles)
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        Return SCP.BLL.aspnetroles.ListActive(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetroles_Update(RoleId As String, RoleName As String) As Boolean
        If String.IsNullOrEmpty(RoleId) Then
            Return False
        End If
        If String.IsNullOrEmpty(RoleName) Then
            Return False
        End If
        Return SCP.BLL.aspnetroles.Update(RoleId, RoleName)
    End Function
#End Region

#Region "aspnetusers"
    Private Function RandomNumber(ByVal min As Integer, ByVal max As Integer) As Integer
        Dim random As New Random()
        Return random.Next(min, max)
    End Function 'RandomNumber
    Private Function RandomString(ByVal size As Integer, ByVal lowerCase As Boolean) As String
        Dim builder As New StringBuilder()
        Dim random As New Random()
        Dim ch As Char
        Dim i As Integer
        For i = 0 To size - 1
            ch = Convert.ToChar(Convert.ToInt32((26 * random.NextDouble() + 65)))
            builder.Append(ch)
        Next i
        If lowerCase Then
            Return builder.ToString().ToLower()
        End If
        Return builder.ToString()
    End Function 'RandomString
    Public Function CreatePassword() As String
        Dim builder As New StringBuilder()
        builder.Append(RandomString(4, True))
        builder.Append(RandomNumber(1000, 9999))
        builder.Append(RandomString(2, False))
        Return builder.ToString()
    End Function 'GetPassword

    Private Function loadModule(moduleName As String) As String
        Dim fullpath As String = HttpContext.Current.Request.MapPath("~\Modules\" & moduleName)
        Dim sr As StreamReader = New StreamReader(fullpath)
        Dim retVal As String = sr.ReadToEnd()
        sr.Close()
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_AddUser(a) As Boolean
        Dim retVal As Boolean = False

        Dim username As String = String.Empty
        Dim password As String = String.Empty
        Dim rolename As String = String.Empty
        Dim comment As String = String.Empty
        Dim email As String = String.Empty

        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            password = ar0(1)
            rolename = ar0(2)
            comment = ar0(3)
            email = ar0(4)
            retVal = SCP.BLL.aspnetUsers.AddUser(username, password, rolename, comment, email)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_UpdUser(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Dim comment As String = String.Empty
        Dim email As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            comment = ar0(1)
            email = ar0(2)
            retVal = SCP.BLL.aspnetUsers.UpdUser(username, comment, email)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_approveUser(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            retVal = SCP.BLL.aspnetUsers.approveUser(username)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_disapproveUser(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            retVal = SCP.BLL.aspnetUsers.disapproveUser(username)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_lockUser(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            retVal = SCP.BLL.aspnetUsers.lockUser(username)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_UnlockUser(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            retVal = SCP.BLL.aspnetUsers.UnlockUser(username)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_changeUserPass(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Dim oldPass As String = String.Empty
        Dim newPass As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            oldPass = ar0(1)
            newPass = ar0(2)
            retVal = SCP.BLL.aspnetUsers.changeUserPass(username, oldPass, newPass)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetuser_emailNewPass(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Dim lang As String = String.Empty
        Dim resetpwd As String = String.Empty
        Dim newpwd As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            lang = ar0(1)
            resetpwd = SCP.BLL.aspnetUsers.resetP(username)
        Catch ex As Exception
            retVal = False
        End Try
        newpwd = CreatePassword()
        If SCP.BLL.aspnetUsers.changeUserPass(username, resetpwd, newpwd) Then
            Dim moduleName As String = "Parking_forgotPwdEmail_" & lang & ".html"
            Dim mailContent As String = loadModule(moduleName)
            mailContent = Replace(mailContent, "@!@pwd@!@", newpwd)

            Dim Message As New System.Net.Mail.MailMessage
            Message.From = New MailAddress("noReply@clevergy.it")
            Message.To.Add(username)
            Message.Subject = "Password reset"
            Message.Body = mailContent
            Message.IsBodyHtml = True

            ' Replace SmtpMail.SmtpServer = server with the following:
            Dim smtp As New SmtpClient("smtp.mandrillapp.com")
            smtp.Port = 587
            smtp.EnableSsl = True
            smtp.Credentials = New System.Net.NetworkCredential("clv.emcs.00@gmail.com", "gr4SmRnLxNn6hRFyW8lq-Q")
            Try
                smtp.Send(Message)
                retVal = True
            Catch ex As Exception
                retVal = False
            End Try
        Else
            retVal = False
        End If
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_removeUser(a As String) As Boolean
        Dim retVal As Boolean = False
        Dim username As String = String.Empty
        Try
            Dim decrypted = DecryptStringAES(a)
            Dim ar0 As String() = decrypted.Split(";")
            username = ar0(0)
            retVal = SCP.BLL.aspnetUsers.removeUser(username)
        Catch ex As Exception
            retVal = False
        End Try
        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_GetActiveUsersByRoleName(RoleName As String) As List(Of aspnetUsers)
        If String.IsNullOrEmpty(RoleName) Then
            Return Nothing
        End If
        Return SCP.BLL.aspnetUsers.GetActiveUsersByRoleName(RoleName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_GetTotActiveUsersByRoleName(RoleName As String) As Integer
        If String.IsNullOrEmpty(RoleName) Then
            Return 0
        End If
        Return SCP.BLL.aspnetUsers.GetTotActiveUsersByRoleName(RoleName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_GetUsersByRoleName(RoleName As String, searchString As String, IdImpianto As String) As List(Of aspnetUsers)
        If String.IsNullOrEmpty(RoleName) Then
            Return Nothing
        End If
        Return SCP.BLL.aspnetUsers.GetUsersByRoleName(RoleName, searchString, IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_GetUsersAll() As List(Of aspnetUsers)
        Return SCP.BLL.aspnetUsers.GetUsersAll()
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_GetUserByUserName(UserName As String) As aspnetUsers
        If String.IsNullOrEmpty(UserName) Then
            Return Nothing
        End If
        Return SCP.BLL.aspnetUsers.GetUserByUserName(UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_userok(UserId As String) As Boolean
        Return SCP.BLL.aspnetUsers.userok(UserId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_whoIsOnline() As List(Of String)
        Return SCP.BLL.aspnetUsers.whoIsOnline
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_getUserO(UserId As String) As String
        Return SCP.BLL.aspnetUsers.getUserO(UserId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_getUserS(UserId As String) As String
        Return SCP.BLL.aspnetUsers.getUserS(UserId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function aspnetusers_getUserA(UserId As String) As String
        Return SCP.BLL.aspnetUsers.getUserA(UserId)
    End Function
#End Region

#Region "aspnetusersstat"
    <OperationContract()>
    <WebGet()>
    Public Function aspnetusersstat_Read() As aspnetusersstat
        Return SCP.BLL.aspnetusersstat.Read
    End Function
#End Region

#Region "dbActivityMonitor"
    <OperationContract()>
    <WebGet()>
    Public Function dbActivityMonitor_List() As List(Of dbActivityMonitor)
        Return SCP.BLL.dbActivityMonitor.List
    End Function
#End Region

#Region "dbstats"
    <OperationContract()>
    <WebGet()>
    Public Function dbstats_Read() As dbstats
        Return SCP.BLL.dbstats.Read
    End Function
#End Region

#Region "decryption"
    Public Shared Function DecryptStringAES(a As String) As String
        Dim keybytes = Encoding.UTF8.GetBytes("7061737323313233")
        Dim iv = Encoding.UTF8.GetBytes("7061737323313233")
        'DECRYPT FROM CRIPTOJS

        Try
            Dim encrypted = Convert.FromBase64String(a)
            Dim decriptedFromJavascript = DecryptStringFromBytes(encrypted, keybytes, iv)
            Return decriptedFromJavascript
        Catch ex As Exception
            Return String.Empty
        End Try

        'Return String.Format("roundtrip reuslt:{0}{1}Javascript result:{2}", roundtrip, Environment.NewLine, decriptedFromJavascript)
    End Function

    Private Shared Function DecryptStringFromBytes(cipherText As Byte(), key As Byte(), iv As Byte()) As String
        ' Check arguments.
        If cipherText Is Nothing OrElse cipherText.Length <= 0 Then
            Throw New ArgumentNullException("cipherText")
        End If
        If key Is Nothing OrElse key.Length <= 0 Then
            Throw New ArgumentNullException("key")
        End If
        If iv Is Nothing OrElse iv.Length <= 0 Then
            Throw New ArgumentNullException("key")
        End If

        ' Declare the string used to hold
        ' the decrypted text.
        Dim plaintext As String = Nothing

        ' Create an RijndaelManaged object
        ' with the specified key and IV.
        Using rijAlg = New RijndaelManaged()
            'Settings
            rijAlg.Mode = CipherMode.CBC
            'rijAlg.Padding = PaddingMode.PKCS7
            rijAlg.Padding = PaddingMode.None
            rijAlg.FeedbackSize = 128

            rijAlg.Key = key
            rijAlg.IV = iv

            ' Create a decrytor to perform the stream transform.
            Dim decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV)

            ' Create the streams used for decryption.
            Using msDecrypt = New MemoryStream(cipherText)
                Using csDecrypt = New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)
                    Using srDecrypt = New StreamReader(csDecrypt)
                        ' Read the decrypted bytes from the decrypting stream
                        ' and place them in a string.
                        plaintext = srDecrypt.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using

        Return plaintext
    End Function

    Private Shared Function EncryptStringToBytes(plainText As String, key As Byte(), iv As Byte()) As Byte()
        ' Check arguments.
        If plainText Is Nothing OrElse plainText.Length <= 0 Then
            Throw New ArgumentNullException("plainText")
        End If
        If key Is Nothing OrElse key.Length <= 0 Then
            Throw New ArgumentNullException("key")
        End If
        If iv Is Nothing OrElse iv.Length <= 0 Then
            Throw New ArgumentNullException("key")
        End If
        Dim encrypted As Byte()
        ' Create a RijndaelManaged object
        ' with the specified key and IV.
        Using rijAlg = New RijndaelManaged()
            rijAlg.Mode = CipherMode.CBC
            rijAlg.Padding = PaddingMode.PKCS7
            rijAlg.FeedbackSize = 128

            rijAlg.Key = key
            rijAlg.IV = iv

            ' Create a decrytor to perform the stream transform.
            Dim encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV)

            ' Create the streams used for encryption.
            Using msEncrypt = New MemoryStream()
                Using csEncrypt = New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                    Using swEncrypt = New StreamWriter(csEncrypt)
                        'Write all data to the stream.
                        swEncrypt.Write(plainText)
                    End Using
                    encrypted = msEncrypt.ToArray()
                End Using
            End Using
        End Using

        ' Return the encrypted bytes from the memory stream.
        Return encrypted
    End Function
#End Region

#Region "end user enroll"
    <OperationContract()>
    <WebGet()>
    Public Function logIn(a As String) As String
        Dim retval As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")
        Dim ts = TimeSpan.FromMilliseconds(Convert.ToDouble(ar0(2)))
        ''non deve esserci differenza maggiore di due secondi tra l'orario passato
        ''e quello attuale.
        'If Now.Second - ts.Seconds < 2 Then
        '    If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
        '        retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        '    End If
        'End If
        If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
            retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        End If
        Return retval
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function register(a As String) As String
        Dim retval As Boolean = False
        Dim ProviderUserKey As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")

        Try
            Dim _newuser As MembershipUser = Membership.CreateUser(ar0(0), ar0(1))
            If Not _newuser Is Nothing Then
                _newuser.IsApproved = True
                Roles.AddUserToRole(_newuser.UserName, "users")
                retval = True
                ProviderUserKey = _newuser.ProviderUserKey.ToString()
                _newuser = Nothing


            End If
        Catch ex As Exception
            retval = False
        End Try

        If retval = False Then
            Return "0;"
        Else
            Return "1;" & ProviderUserKey
        End If
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function logA(a As String) As String
        Dim retval As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")
        Dim ts = TimeSpan.FromMilliseconds(Convert.ToDouble(ar0(2)))
        ''non deve esserci differenza maggiore di 10 secondi tra l'orario passato
        ''e quello attuale.
        'If Now.Second - ts.Seconds < 11 Then
        '    If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
        '        If Roles.IsUserInRole(ar0(0), "Administrators") Then
        '            retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        '        End If
        '    End If
        'End If
        If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
            If Roles.IsUserInRole(ar0(0), "Administrators") Then
                retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
            End If
        End If
        Return retval
    End Function

    <OperationContract()>
     <WebGet()>
    Public Function logO(a As String) As String
        Dim retval As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")
        Dim ts = TimeSpan.FromMilliseconds(Convert.ToDouble(ar0(2)))

        ''non deve esserci differenza maggiore di 10 secondi tra l'orario passato
        ''e quello attuale.
        'If Now.Second - ts.Seconds < 11 Then
        '    'If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
        '    '    If Roles.IsUserInRole(ar0(0), "Operators") Then
        '    '        retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        '    '    End If
        '    'End If
        'End If
        If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
            If Roles.IsUserInRole(ar0(0), "Operators") Then
                retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
            End If
        End If

        Return retval
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function logU(a As String) As String
        Dim retval As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")
        Dim ts = TimeSpan.FromMilliseconds(Convert.ToDouble(ar0(2)))

        ''non deve esserci differenza maggiore di 10 secondi tra l'orario passato
        ''e quello attuale.
        'If Now.Second - ts.Seconds < 11 Then
        '    If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
        '        If Roles.IsUserInRole(ar0(0), "Endusers") Then
        '            retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        '        End If
        '    End If
        'End If

        If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
            If Roles.IsUserInRole(ar0(0), "Endusers") Then
                retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
            End If
        End If

        Return retval
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function logS(a As String) As String
        Dim retval As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")
        Dim ts = TimeSpan.FromMilliseconds(Convert.ToDouble(ar0(2)))

        ''non deve esserci differenza maggiore di 10 secondi tra l'orario passato
        ''e quello attuale.
        'If Now.Second - ts.Seconds < 11 Then
        '    If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
        '        If Roles.IsUserInRole(ar0(0), "Supervisors") Then
        '            retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        '        End If
        '    End If
        'End If
        If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
            If Roles.IsUserInRole(ar0(0), "Supervisors") Then
                retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
            End If
        End If
        Return retval
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function logM(a As String) As String
        Dim retval As String = String.Empty
        Dim decrypted = DecryptStringAES(a)
        Dim ar0 As String() = decrypted.Split(";")
        Dim ts = TimeSpan.FromMilliseconds(Convert.ToDouble(ar0(2)))

        ''non deve esserci differenza maggiore di 10 secondi tra l'orario passato
        ''e quello attuale.
        'If Now.Second - ts.Seconds < 11 Then
        '    If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
        '        If Roles.IsUserInRole(ar0(0), "Maintainers") Then
        '            retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
        '        End If
        '    End If
        'End If
        If Membership.ValidateUser(ar0(0), ar0(1)) = True Then
            If Roles.IsUserInRole(ar0(0), "Maintainers") Then
                retval = Membership.GetUser(ar0(0)).ProviderUserKey.ToString()
            End If
        End If
        Return retval
    End Function
#End Region

#Region "UserMenu"
    <OperationContract()>
    <WebGet()>
    Public Function getUserMenu(username As String) As List(Of SCP.BLL.UserMenu)
        Dim retVal As New List(Of SCP.BLL.UserMenu)

        Dim _list As List(Of SCP.BLL.UserMenu) = SCP.BLL.UserMenu.List(username)
        If Not _list Is Nothing Then

            Dim _usermenuList = From c In _list Where c.IdVocePadre = 0 Select c
            If Not _usermenuList Is Nothing Then
                If _usermenuList.Count > 0 Then
                    For Each c In _usermenuList
                        Dim fathermenu As New SCP.BLL.UserMenu
                        fathermenu.IdVoceMenu = c.IdVoceMenu
                        fathermenu.IdVocePadre = c.IdVocePadre
                        fathermenu.IdAzione = c.IdAzione
                        fathermenu.DescrVoce = c.DescrVoce
                        fathermenu.URLAzione = c.URLAzione
                        fathermenu.mainPage = c.mainPage
                        retVal.Add(fathermenu)

                        Dim _VociMenu = From o In _list Where o.IdVocePadre = fathermenu.IdVoceMenu Select o
                        If Not _VociMenu Is Nothing Then
                            If _VociMenu.Count > 0 Then
                                For Each o In _VociMenu
                                    Dim vocemenu As New SCP.BLL.UserMenu
                                    vocemenu.IdVoceMenu = o.IdVoceMenu
                                    vocemenu.IdVocePadre = o.IdVocePadre
                                    vocemenu.IdAzione = o.IdAzione
                                    vocemenu.DescrVoce = o.DescrVoce
                                    vocemenu.URLAzione = o.URLAzione
                                    vocemenu.mainPage = String.Empty
                                    retVal.Add(vocemenu)
                                Next
                            End If
                            _VociMenu = Nothing
                        End If
                    Next
                End If
                _usermenuList = Nothing
            End If

        End If
        Return retVal
    End Function
    <OperationContract()>
    <WebGet()>
    Public Function UserMenu_ListFather(UserName As String) As List(Of SCP.BLL.UserMenu)
        If String.IsNullOrEmpty(UserName) Then
            Return Nothing
        End If
        Return SCP.BLL.UserMenu.ListFather(UserName)
    End Function
    <OperationContract()>
    <WebGet()>
    Public Function UserMenu_ListChildren(UserName As String, IdVocePadre As Integer) As List(Of SCP.BLL.UserMenu)
        If String.IsNullOrEmpty(UserName) Then
            Return Nothing
        End If
        If IdVocePadre = 0 Then
            Return Nothing
        End If
        Return SCP.BLL.UserMenu.ListChildren(UserName, IdVocePadre)
    End Function
#End Region
#Region "menuUtente"
    <OperationContract()>
    <WebGet()>
    Public Function menuUtente_Read(username As String) As List(Of SCP.BLL.menuUtente)
        If String.IsNullOrEmpty(username) Then
            Return Nothing
        End If

        Return SCP.BLL.menuUtente.Read(username)
    End Function
#End Region

#Region "Impianti"
    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Add(DesImpianto As String, _
                               Indirizzo As String, _
                               Latitude As Decimal, _
                               Longitude As Decimal, _
                               AltSLM As Integer, _
                               IsActive As Boolean) As Boolean
        If String.IsNullOrEmpty(DesImpianto) Then
            Return False
        End If
        Return SCP.BLL.Impianti.Add(DesImpianto, Indirizzo, Latitude, Longitude, AltSLM, IsActive)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_List(searchString As String) As List(Of SCP.BLL.Impianti)
        Return SCP.BLL.Impianti.List(searchString)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_ListByUser(UserId As String, searchString As String) As List(Of Impianti)
        If String.IsNullOrEmpty(UserId) Then
            Return Nothing
        End If
        Return SCP.BLL.Impianti.ListByUser(UserId, searchString)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_ListPaged(id As Integer, searchString As String) As List(Of Impianti)
        Return SCP.BLL.Impianti.ListPaged(id, searchString)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Read(IdImpianto As String) As Impianti
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        Return SCP.BLL.Impianti.Read(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Upd(IdImpianto As String, _
                               DesImpianto As String, _
                               Indirizzo As String, _
                               Latitude As Decimal, _
                               Longitude As Decimal, _
                               AltSLM As Integer, _
                               IsActive As Boolean) As Boolean

        If String.IsNullOrEmpty(IdImpianto) Then
            Return False
        End If
        If String.IsNullOrEmpty(DesImpianto) Then
            Return False
        End If
        Return SCP.BLL.Impianti.Upd(IdImpianto, DesImpianto, Indirizzo, Latitude, Longitude, AltSLM, IsActive)
    End Function
#End Region

#Region "Impianti_Contatti"
    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Contatti_Add(IdImpianto As String, _
                                          Descrizione As String, _
                                          Indirizzo As String, _
                                          Nome As String, _
                                          TelFisso As String, _
                                          TelMobile As String, _
                                          emailaddress As String) As Boolean
        If String.IsNullOrEmpty(IdImpianto) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descrizione) Then
            Return False
        End If
        Return SCP.BLL.Impianti_Contatti.Add(IdImpianto, Descrizione, Indirizzo, Nome, TelFisso, TelMobile, emailaddress)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Contatti_Del(IdContatto As Integer) As Boolean
        If IdContatto <= 0 Then
            Return False
        End If
        Return SCP.BLL.Impianti_Contatti.Del(IdContatto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Contatti_List(IdImpianto As String) As IList(Of Impianti_Contatti)
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        Return SCP.BLL.Impianti_Contatti.List(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Contatti_Read(IdContatto As Integer) As Impianti_Contatti
        If IdContatto <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Impianti_Contatti.Read(IdContatto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_Contatti_Update(IdContatto As Integer, _
                                             Descrizione As String, _
                                             Indirizzo As String, _
                                             Nome As String, _
                                             TelFisso As String, _
                                             TelMobile As String, _
                                             emailaddress As String) As Boolean
        If IdContatto <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Descrizione) Then
            Return False
        End If
        Return SCP.BLL.Impianti_Contatti.Update(IdContatto, Descrizione, Indirizzo, Nome, TelFisso, TelMobile, emailaddress)
    End Function
#End Region

#Region "Impianti_RemoteConnections"
    <OperationContract()>
    <WebGet()>
    Public Function Impianti_RemoteConnections_Add(IdImpianto As String, Descr As String, remoteaddress As String, connectionType As Integer, NoteInterne As String) As Boolean
        If String.IsNullOrEmpty(IdImpianto) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim IdAddress As Integer = SCP.BLL.Impianti_RemoteConnections.getLastId() + 1
        Return SCP.BLL.Impianti_RemoteConnections.Add(IdImpianto, IdAddress, Descr, remoteaddress, connectionType, NoteInterne)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_RemoteConnections_Del(IdImpianto As String, IdAddress As Integer) As Boolean
        If String.IsNullOrEmpty(IdImpianto) Then
            Return False
        End If
        If IdAddress <= 0 Then
            Return False
        End If
        Return SCP.BLL.Impianti_RemoteConnections.Del(IdImpianto, IdAddress)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_RemoteConnections_List(IdImpianto As String) As List(Of Impianti_RemoteConnections)
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        Return SCP.BLL.Impianti_RemoteConnections.List(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Impianti_RemoteConnections_Read(IdImpianto As String, IdAddress As Integer) As Impianti_RemoteConnections
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        If IdAddress <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Impianti_RemoteConnections.Read(IdImpianto, IdAddress)
    End Function

    <OperationContract()>
   <WebGet()>
    Public Function Impianti_RemoteConnections_Upd(IdImpianto As String, IdAddress As Integer, Descr As String, remoteaddress As String, connectionType As Integer, NoteInterne As String) As Boolean
        If String.IsNullOrEmpty(IdImpianto) Then
            Return False
        End If
        If IdAddress <= 0 Then
            Return False
        End If
        Return SCP.BLL.Impianti_RemoteConnections.Upd(IdImpianto, IdAddress, Descr, remoteaddress, connectionType, NoteInterne)
    End Function
#End Region

#Region "hs_Docs"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Docs_Add(hsId As Integer, _
                                DocName As String, _
                                Creator As String) As Boolean
        If String.IsNullOrEmpty(DocName) Then
            Return False
        End If

        Return SCP.BLL.hs_Docs.Add(hsId, DocName, Creator)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Docs_Del(IdDoc As Integer) As Boolean
        If IdDoc <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Docs.Del(IdDoc)
    End Function

    <OperationContract()>
   <WebGet()>
    Public Function hs_Docs_List(hsId As Integer) As List(Of hs_Docs)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Docs.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Docs_Read(IdDoc As Integer) As hs_Docs
        If IdDoc <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Docs.Read(IdDoc)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Docs_getBinaryData(IdDoc As Integer) As Boolean
        If IdDoc <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Docs.getBinaryData(IdDoc)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_DocsgetTot(hsId As Integer) As Integer
        If hsId <= 0 Then
            Return 0
        End If
        Return SCP.BLL.hs_Docs.getTot(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Docs_setBinaryData(IdDoc As Integer, BinaryData As String, UserName As String) As Boolean
        If IdDoc <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(BinaryData) Then
            Return False
        End If
        Return SCP.BLL.hs_Docs.setBinaryData(IdDoc, BinaryData, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Docs_Update(IdDoc As Integer, DocName As String, UserName As String) As Boolean
        If IdDoc <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(DocName) Then
            Return False
        End If
        Return SCP.BLL.hs_Docs.Update(IdDoc, DocName, UserName)
    End Function
#End Region

#Region "hs_ErrorCodes"
    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorCodes_Add(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
        If String.IsNullOrEmpty(elementCode) Then
            Return False
        End If
        If errorCode <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(DescIT) Or String.IsNullOrEmpty(DescEN) Then
            Return False
        End If
        Return SCP.BLL.hs_ErrorCodes.Add(elementCode, errorCode, errorLevel, DescIT, DescEN)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorCodes_Del(elementCode As String, errorCode As Integer) As Boolean
        If String.IsNullOrEmpty(elementCode) Then
            Return False
        End If
        If errorCode <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_ErrorCodes.Del(elementCode, elementCode)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorCodes_List(elementCode As String) As List(Of hs_ErrorCodes)
        If String.IsNullOrEmpty(elementCode) Then
            Return Nothing
        End If
        Return SCP.BLL.hs_ErrorCodes.List(elementCode)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorCodes_Read(elementCode As String, errorCode As Integer) As hs_ErrorCodes
        If String.IsNullOrEmpty(elementCode) Then
            Return Nothing
        End If
        If errorCode <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_ErrorCodes.Read(elementCode, errorCode)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorCodes_Update(elementCode As String, errorCode As Integer, errorLevel As Integer, DescIT As String, DescEN As String) As Boolean
        If String.IsNullOrEmpty(elementCode) Then
            Return False
        End If
        If errorCode <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(DescIT) Or String.IsNullOrEmpty(DescEN) Then
            Return False
        End If
        Return SCP.BLL.hs_ErrorCodes.Update(elementCode, errorCode, errorLevel, DescIT, DescEN)
    End Function
#End Region

#Region "hs_ErrorLog"
    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorLog_List(hsId As Integer, fromDate As String, toDate As String) As List(Of hs_ErrorLog)
        If hsId <= 0 Then
            Return Nothing
        End If
        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)
        Return SCP.BLL.hs_ErrorLog.List(hsId, _fromDate, _toDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorLog_ListAll(hsId As Integer, rowNumber As Integer) As List(Of hs_ErrorLog)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_ErrorLog.ListAll(hsId, rowNumber)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ErrorLog_ListByElement(hsId As Integer, hselement As String, rowNumber As Integer) As List(Of hs_ErrorLog)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(hselement) Then
            Return Nothing
        End If
        Return SCP.BLL.hs_ErrorLog.ListByElement(hsId, hselement, rowNumber)
    End Function
#End Region

#Region "hs_Tickets"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Add(hsId As Integer, _
                                   TicketTitle As String, _
                                   Requester As String, _
                                   emailRequester As String, _
                                   Description As String, _
                                   Executor As String, _
                                   emailExecutor As String, _
                                   UserName As String, _
                                   TicketType As Integer) As Boolean
        If String.IsNullOrEmpty(TicketTitle) Then
            Return False
        End If
        If String.IsNullOrEmpty(Requester) Then
            Return False
        End If
        'If String.IsNullOrEmpty(emailRequester) Then
        '    Return False
        'End If
        If String.IsNullOrEmpty(Description) Then
            Return False
        End If
        If String.IsNullOrEmpty(Executor) Then
            Return False
        End If
        If String.IsNullOrEmpty(emailExecutor) Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets.Add(hsId, TicketTitle, Requester, emailRequester, Description, Executor, emailExecutor, UserName, TicketType)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Del(TicketId As Integer) As Boolean
        If TicketId <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets.Del(TicketId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_List(hsId As Integer, TicketStatus As Integer, searchString As String) As List(Of hs_Tickets)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets.List(hsId, TicketStatus, searchString)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_ListPaged(hsId As Integer, TicketStatus As Integer, searchString As String, RowNumber As Integer) As List(Of hs_Tickets)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets.ListPaged(hsId, TicketStatus, searchString, RowNumber)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_getTotOpen(hsId As Integer) As Integer
        If hsId <= 0 Then
            Return 0
        End If
        Return SCP.BLL.hs_Tickets.getTotOpen(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Read(TicketId As Integer) As hs_Tickets
        If TicketId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets.Read(TicketId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Update(TicketId As Integer, _
                                      TicketTitle As String, _
                                      Requester As String, _
                                      emailRequester As String, _
                                      Description As String, _
                                      Executor As String, _
                                      emailExecutor As String, _
                                      UserName As String) As Boolean
        If TicketId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(TicketTitle) Then
            Return False
        End If
        If String.IsNullOrEmpty(Requester) Then
            Return False
        End If
        'If String.IsNullOrEmpty(emailRequester) Then
        '    Return False
        'End If
        If String.IsNullOrEmpty(Description) Then
            Return False
        End If
        If String.IsNullOrEmpty(Executor) Then
            Return False
        End If
        If String.IsNullOrEmpty(emailExecutor) Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets.Update(TicketId, TicketTitle, Requester, emailRequester, Description, Executor, emailExecutor, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_ChangeStatus(TicketId As Integer, DateExecution As String, ExecutorComment As String, TicketStatus As Integer) As Boolean
        If TicketId <= 0 Then
            Return False
        End If

        'richiesta chiusura, la data di esecuzione è impostata ad oggi
        If TicketStatus = 5 Or TicketStatus = 6 Then
            DateExecution = FormatDateTime(Now, DateFormat.GeneralDate)
        End If

        If String.IsNullOrEmpty(DateExecution) Then
            DateExecution = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        End If

        Dim _DateExecution As Date = CDate(DateExecution)
        Return SCP.BLL.hs_Tickets.ChangeStatus(TicketId, _DateExecution, ExecutorComment, TicketStatus)
    End Function
#End Region
#Region "hs_Tickets_Executors"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Executors_Add(hsId As Integer, IdContatto As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If IdContatto <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets_Executors.Add(hsId, IdContatto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Executors_Del(Id As Integer) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets_Executors.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Executors_List(hsId As Integer) As List(Of hs_Tickets_Executors)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets_Executors.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Executors_Read(Id As Integer) As hs_Tickets_Executors
        If Id <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets_Executors.Read(Id)
    End Function
#End Region
#Region "hs_Tickets_Requesters"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Requesters_Add(hsId As Integer, IdContatto As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If IdContatto <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets_Requesters.Add(hsId, IdContatto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Requesters_Del(Id As Integer) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Tickets_Requesters.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Requesters_List(hsId As Integer) As List(Of hs_Tickets_Requesters)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets_Requesters.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Tickets_Requesters_Read(Id As Integer) As hs_Tickets_Requesters
        If Id <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Tickets_Requesters.Read(Id)
    End Function
#End Region

#Region "HeatingSystem"
    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_Add(IdImpianto As String, Descr As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Integer, UserName As String) As Boolean
        If String.IsNullOrEmpty(IdImpianto) Then
            Return False
        End If
        If String.IsNullOrEmpty(Indirizzo) Then
            Return False
        End If
        Return SCP.BLL.HeatingSystem.Add(IdImpianto, Descr, Indirizzo, Latitude, Longitude, AltSLM, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_Del(hsId As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        Return SCP.BLL.HeatingSystem.Del(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_getTotMap(IdImpianto As String) As Integer
        If String.IsNullOrEmpty(IdImpianto) Then Return 0
        Return SCP.BLL.HeatingSystem.getTotMap(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_List(IdImpianto As String) As List(Of HeatingSystem)
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        Return SCP.BLL.HeatingSystem.List(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_ListEnabled(IdImpianto As String) As List(Of HeatingSystem)
        If String.IsNullOrEmpty(IdImpianto) Then
            Return Nothing
        End If
        Return SCP.BLL.HeatingSystem.ListEnabled(IdImpianto)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_ListAll() As List(Of HeatingSystem)
        Return SCP.BLL.HeatingSystem.ListAll()
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_Read(hsId As Integer) As HeatingSystem
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.HeatingSystem.Read(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_setMaintenanceMode(hsId As Integer, MaintenanceMode As Boolean) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        Return SCP.BLL.HeatingSystem.setMaintenanceMode(hsId, MaintenanceMode)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_setIwMonitoring(hsId As Integer, IwMonitoringId As Integer, IwMonitoringDes As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        Return SCP.BLL.HeatingSystem.setIwMonitoring(hsId, IwMonitoringId, IwMonitoringDes)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_setIsEnabled(hsId As Integer, isEnabled As Boolean) As Boolean
        If hsId <= 0 Then Return False
        Return SCP.BLL.HeatingSystem.setIsEnabled(hsId, isEnabled)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_setNote(hsId As Integer, Note As String, UserName As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        Return SCP.BLL.HeatingSystem.setNote(hsId, Note, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_Update(hsId As Integer, Descr As String, Indirizzo As String, Latitude As Decimal, Longitude As Decimal, AltSLM As Integer, VPNConnectionId As Integer, MapId As Integer, UserName As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Indirizzo) Then
            Return False
        End If
        Return SCP.BLL.HeatingSystem.Update(hsId, Descr, Indirizzo, Latitude, Longitude, AltSLM, VPNConnectionId, MapId, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_resetSystemStatus(hsId As Integer) As Boolean
        If hsId <= 0 Then Return False
        Return SCP.BLL.HeatingSystem.resetSystemStatus(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function HeatingSystem_requestLog(hsId As Integer) As Boolean
        If hsId <= 0 Then Return False
        Return SCP.BLL.HeatingSystem.requestLog(hsId)
    End Function
#End Region

#Region "hs_Elem"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Elem_Read(hsId As Integer) As hs_Elem
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Elem.Read(hsId)
    End Function
#End Region

#Region "tbhsElem"
    <OperationContract()>
    <WebGet()>
    Public Function tbhsElem_List() As List(Of tbhsElem)
        Return SCP.BLL.tbhsElem.List
    End Function
#End Region


#Region "hs_Controller"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Controller_Add(hsId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Controller.Add(hsId, ControllerDescr, NoteInterne, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Controller_Del(ControllerId As Integer) As Boolean
        If ControllerId <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Controller.Del(ControllerId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Controller_List(hsId As Integer) As List(Of hs_Controller)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Controller.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Controller_Read(ControllerId As Integer) As hs_Controller
        If ControllerId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Controller.Read(ControllerId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Controller_Update(ControllerId As Integer, ControllerDescr As String, NoteInterne As String, UserName As String) As Boolean
        If ControllerId <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Controller.Update(ControllerId, ControllerDescr, NoteInterne, UserName)
    End Function
#End Region

#Region "hs_ControllerDetail"
    <OperationContract()>
    <WebGet()>
    Public Function hs_ControllerDetail_Add(ControllerId As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
        If ControllerId <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_ControllerDetail.Add(ControllerId, Descr, NoteInterne, qta, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ControllerDetail_Del(id As Integer) As Boolean
        If id <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_ControllerDetail.Del(id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ControllerDetail_List(ControllerId As Integer) As List(Of hs_ControllerDetail)
        If ControllerId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_ControllerDetail.List(ControllerId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ControllerDetail_Read(id As Integer) As hs_ControllerDetail
        If id <= 0 Then
            Return Nothing
        End If

        Return SCP.BLL.hs_ControllerDetail.Read(id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_ControllerDetail_Update(id As Integer, Descr As String, NoteInterne As String, qta As Integer, UserName As String) As Boolean
        If id <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_ControllerDetail.Update(id, Descr, NoteInterne, qta, UserName)
    End Function
#End Region

#Region "Lux"
    <OperationContract()>
    <WebGet()>
    Public Function Lux_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = installationDate
        End If
        Return SCP.BLL.Lux.Add(hsId, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_Del(Id As Integer) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Return SCP.BLL.Lux.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_List(hsId As Integer) As List(Of Lux)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Lux.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_Read(Id As Integer) As Lux
        If Id <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Lux.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_ReadByCod(hsId As Integer, Cod As String) As Lux
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If
        Return SCP.BLL.Lux.ReadByCod(hsId, Cod)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_ReadByAmb(hsId As Integer, IdAmbiente As Integer) As List(Of Lux)
        If IdAmbiente <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Lux.ReadByAmb(hsId, IdAmbiente)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If Id <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = installationDate
        End If
        Return SCP.BLL.Lux.Update(Id, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        Return SCP.BLL.Lux.setStatus(hsId, Cod, stato)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_setValue(hsId As Integer, Cod As String, LightON As Boolean, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, CurrentMode As Integer, forcedOn As Boolean, forcedOff As Boolean) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If

        Return SCP.BLL.Lux.setValue(hsId, Cod, LightON, WorkingTimeCounter, PowerOnCycleCounter, CurrentMode, forcedOn, forcedOff)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.Lux.setGeoLocation(Id, Latitude, Longitude)
    End Function
#End Region
#Region "Lux_replacement_history"
    <OperationContract()>
    <WebGet()>
    Public Function Lux_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If ParentId <= 0 Then Return False

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.Lux_replacement_history.Add(ParentId, marcamodello, _installationDate, note, userName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_replacement_history_Del(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.Lux_replacement_history.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_replacement_history_List(ParentId As Integer) As List(Of Lux_replacement_history)
        If ParentId <= 0 Then Return Nothing
        Return SCP.BLL.Lux_replacement_history.List(ParentId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_replacement_history_Read(Id As Integer) As Lux_replacement_history
        If Id <= 0 Then Return Nothing
        Return SCP.BLL.Lux_replacement_history.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Lux_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If Id <= 0 Then Return False
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.Lux_replacement_history.Update(Id, marcamodello, _installationDate, note, userName)
    End Function
#End Region
#Region "log_Lux"
    <OperationContract()>
    <WebGet()>
    Public Function log_Lux_List(hsId As Integer, Cod As String, fromDate As String, toDate As String) As List(Of log_Lux)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_Lux.List(hsId, Cod, _fromDate, _toDate)
    End Function
    <OperationContract()>
    <WebGet()>
    Public Function log_LuxList_Paged(hsId As Integer, Cod As String, fromDate As String, toDate As String, rowNumber As Integer) As List(Of log_Lux)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_Lux.ListPaged(hsId, Cod, _fromDate, _toDate, rowNumber)
    End Function
#End Region

#Region "LuxM"
    <OperationContract()>
    <WebGet()>
    Public Function LuxM_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Integer
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = installationDate
        End If
        Return SCP.BLL.LuxM.Add(hsId, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_Del(Id As Integer) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Return SCP.BLL.LuxM.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_List(hsId As Integer) As List(Of LuxM)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.LuxM.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_Read(Id As Integer) As LuxM
        If Id <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.LuxM.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_ReadByCod(hsId As Integer, Cod As String) As LuxM
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If
        Return SCP.BLL.LuxM.ReadByCod(hsId, Cod)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If Id <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = installationDate
        End If
        Return SCP.BLL.LuxM.Update(Id, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        Return SCP.BLL.LuxM.setStatus(hsId, Cod, stato)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_setValue(hsId As Integer, Cod As String, Voltage As Decimal, Curr As Decimal, EnergyCounter As Decimal, WorkingTimeCounter As Decimal, PowerOnCycleCounter As Decimal, Temp As Decimal) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If

        Return SCP.BLL.LuxM.setValue(hsId, Cod, Voltage, Curr, EnergyCounter, WorkingTimeCounter, PowerOnCycleCounter, Temp)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.LuxM.setGeoLocation(Id, Latitude, Longitude)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_cmd_LightOn(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.LuxM.cmd_LightOn(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_cmd_LightOff(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.LuxM.cmd_LightOff(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_cmd_RestoreWorkingMode(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.LuxM.cmd_RestoreWorkingMode(Id)
    End Function
#End Region
#Region "LuxM_replacement_history"
    <OperationContract()>
    <WebGet()>
    Public Function LuxM_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If ParentId <= 0 Then Return False

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.LuxM_replacement_history.Add(ParentId, marcamodello, _installationDate, note, userName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_replacement_history_Del(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.LuxM_replacement_history.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_replacement_history_List(ParentId As Integer) As List(Of LuxM_replacement_history)
        If ParentId <= 0 Then Return Nothing
        Return SCP.BLL.LuxM_replacement_history.List(ParentId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_replacement_history_Read(Id As Integer) As LuxM_replacement_history
        If Id <= 0 Then Return Nothing
        Return SCP.BLL.LuxM_replacement_history.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function LuxM_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If Id <= 0 Then Return False
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.LuxM_replacement_history.Update(Id, marcamodello, _installationDate, note, userName)
    End Function
#End Region
#Region "log_LuxM"
    <OperationContract()>
    <WebGet()>
    Public Function log_LuxM_List(hsId As Integer, Cod As String, fromDate As String, toDate As String) As List(Of log_LuxM)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_LuxM.List(hsId, Cod, _fromDate, _toDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function log_LuxM_ListPaged(hsId As Integer, Cod As String, fromDate As String, toDate As String, rowNumber As Integer) As List(Of log_LuxM)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_LuxM.ListPaged(hsId, Cod, _fromDate, _toDate, rowNumber)
    End Function
#End Region

#Region "Psg"
    <OperationContract()>
    <WebGet()>
    Public Function Psg_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If

        Return SCP.BLL.Psg.Add(hsId, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_Del(Id As Integer) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Return SCP.BLL.Psg.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_List(hsId As Integer) As List(Of SCP.BLL.Psg)
        If hsId <= 0 Then
            Return Nothing
        End If

        'Dim retVal As List(Of Psg) = SCP.BLL.Psg.List(hsId)
        'Return retVal
        Return SCP.BLL.Psg.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_Read(Id As Integer) As Psg
        If Id <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Psg.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_ReadByCod(hsId As Integer, Cod As String) As Psg
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If
        Return SCP.BLL.Psg.ReadByCod(hsId, Cod)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function psg_ReadByLux(hsId As Integer, luxCod As String) As List(Of Psg)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(luxCod) Then
            Return Nothing
        End If

        Return SCP.BLL.Psg.ReadByLux(hsId, luxCod)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If Id <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = installationDate
        End If
        Return SCP.BLL.Psg.Update(Id, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        Return SCP.BLL.Psg.setStatus(hsId, Cod, stato)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_setValue(hsId As Integer, Cod As String, currentValue As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If

        Return SCP.BLL.Psg.setValue(hsId, Cod, currentValue)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.Psg.setGeoLocation(Id, Latitude, Longitude)
    End Function
#End Region
#Region "Psg_replacement_history"
    <OperationContract()>
    <WebGet()>
    Public Function Psg_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If ParentId <= 0 Then Return False

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.Psg_replacement_history.Add(ParentId, marcamodello, _installationDate, note, userName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_replacement_history_Del(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.Psg_replacement_history.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_replacement_history_List(ParentId As Integer) As List(Of Psg_replacement_history)
        If ParentId <= 0 Then Return Nothing
        Return SCP.BLL.Psg_replacement_history.List(ParentId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_replacement_history_Read(Id As Integer) As Psg_replacement_history
        If Id <= 0 Then Return Nothing
        Return SCP.BLL.Psg_replacement_history.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Psg_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If Id <= 0 Then Return False
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.Psg_replacement_history.Update(Id, marcamodello, _installationDate, note, userName)
    End Function
#End Region
#Region "log_Psg"
    <OperationContract()>
    <WebGet()>
    Public Function log_Psg_List(hsId As Integer, Cod As String, fromDate As String, toDate As String) As List(Of log_Psg)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_Psg.List(hsId, Cod, _fromDate, _toDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function log_Psg_ListPaged(hsId As Integer, Cod As String, fromDate As String, toDate As String, rowNumber As Integer) As List(Of log_Psg)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_Psg.ListPaged(hsId, Cod, _fromDate, _toDate, rowNumber)
    End Function

#End Region

#Region "Zrel"
    <OperationContract()>
    <WebGet()>
    Public Function Zrel_Add(hsId As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If

        Return SCP.BLL.Zrel.Add(hsId, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_Del(Id As Integer) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Return SCP.BLL.Zrel.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_List(hsId As Integer) As List(Of SCP.BLL.Zrel)
        If hsId <= 0 Then
            Return Nothing
        End If

        'Dim retVal As List(Of Psg) = SCP.BLL.Psg.List(hsId)
        'Return retVal
        Return SCP.BLL.Zrel.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_Read(Id As Integer) As Zrel
        If Id <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Zrel.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_ReadByCod(hsId As Integer, Cod As String) As Zrel
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If
        Return SCP.BLL.Zrel.ReadByCod(hsId, Cod)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_Update(Id As Integer, Cod As String, Descr As String, UserName As String, marcamodello As String, installationDate As String) As Boolean
        If Id <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        If String.IsNullOrEmpty(Descr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = installationDate
        End If
        Return SCP.BLL.Zrel.Update(Id, Cod, Descr, UserName, marcamodello, _installationDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_setStatus(hsId As Integer, Cod As String, stato As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        Return SCP.BLL.Zrel.setStatus(hsId, Cod, stato)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_setValue(hsId As Integer, Cod As String, LQI As Integer, Temperature As Decimal, MeshParentId As Decimal, CurrentId As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return False
        End If
        Return SCP.BLL.Zrel.setValue(hsId, Cod, LQI, Temperature, MeshParentId, CurrentId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_setGeoLocation(Id As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.Zrel.setGeoLocation(Id, Latitude, Longitude)
    End Function
#End Region
#Region "Zrel_replacement_history"
    <OperationContract()>
    <WebGet()>
    Public Function Zrel_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If ParentId <= 0 Then Return False

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.Zrel_replacement_history.Add(ParentId, marcamodello, _installationDate, note, userName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_replacement_history_Del(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.Zrel_replacement_history.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_replacement_history_List(ParentId As Integer) As List(Of Zrel_replacement_history)
        If ParentId <= 0 Then Return Nothing
        Return SCP.BLL.Zrel_replacement_history.List(ParentId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_replacement_history_Read(Id As Integer) As Zrel_replacement_history
        If Id <= 0 Then Return Nothing
        Return SCP.BLL.Zrel_replacement_history.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Zrel_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If Id <= 0 Then Return False
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.Zrel_replacement_history.Update(Id, marcamodello, _installationDate, note, userName)
    End Function
#End Region
#Region "log_Zrel"
    <OperationContract()>
    <WebGet()>
    Public Function log_Zrel_List(hsId As Integer, Cod As String, fromDate As String, toDate As String) As List(Of log_Zrel)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If
        Return SCP.BLL.log_Zrel.List(hsId, Cod, fromDate, toDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function log_Zrel_ListPaged(hsId As Integer, Cod As String, fromDate As String, toDate As String, rowNumber As Integer) As List(Of log_Zrel)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(Cod) Then
            Return Nothing
        End If
        Return SCP.BLL.log_Zrel.ListPaged(hsId, Cod, fromDate, toDate, rowNumber)
    End Function

#End Region

#Region "hs_Cron"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_Add(hsId As Integer, CronCod As String, CronDescr As String, NoteInterne As String, UserName As String, marcamodello As String, installationDate As String, CronType As Integer) As Boolean
        If hsId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(CronCod) Then
            Return False
        End If
        If String.IsNullOrEmpty(CronDescr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        Return SCP.BLL.hs_Cron.Add(hsId, UCase(CronCod), CronDescr, NoteInterne, UserName, marcamodello, _installationDate, CronType)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_Del(CronId As Integer) As Boolean
        If CronId <= 0 Then
            Return False
        End If
        Return SCP.BLL.hs_Cron.Del(CronId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_List(hsId As Integer) As List(Of hs_Cron)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Cron.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_ListOnOff(hsid As Integer) As List(Of hs_Cron)
        If hsid <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Cron.ListOnOff(hsid)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_Read(CronId As Integer) As hs_Cron
        If CronId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Cron.Read(CronId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_ReadByCronCod(hsId As Integer, CronCod As String) As hs_Cron
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(CronCod) Then
            Return Nothing
        End If
        Return SCP.BLL.hs_Cron.ReadByCronCod(hsId, UCase(CronCod))
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_Update(CronId As Integer, CronCod As String, CronDescr As String, NoteInterne As String, UserName As String, marcamodello As String, installationDate As String, CronType As Integer) As Boolean
        If CronId <= 0 Then
            Return False
        End If
        If String.IsNullOrEmpty(CronCod) Then
            Return False
        End If
        If String.IsNullOrEmpty(CronDescr) Then
            Return False
        End If
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        Return SCP.BLL.hs_Cron.Update(CronId, UCase(CronCod), CronDescr, NoteInterne, UserName, marcamodello, _installationDate, CronType)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_UpdateMarcamodello(Id As Integer, marcamodello As String, installationDate As String, UserName As String) As Boolean
        If Id <= 0 Then
            Return False
        End If
        Dim _installationDate As Date = CDate(installationDate)
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.hs_Cron.UpdateMarcamodello(Id, marcamodello, _installationDate, UserName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_setGeoLocation(CronId As Integer, Latitude As Decimal, Longitude As Decimal) As Boolean
        If CronId <= 0 Then Return False
        Return SCP.BLL.hs_Cron.setGeoLocation(CronId, Latitude, Longitude)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_Restart(CronId As Integer) As Boolean
        If CronId <= 0 Then Return False
        Return SCP.BLL.hs_Cron.Restart(CronId)
    End Function
#End Region
#Region "hs_Cron_replacement_history"
    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_replacement_history_Add(ParentId As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If ParentId <= 0 Then Return False

        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.hs_Cron_replacement_history.Add(ParentId, marcamodello, _installationDate, note, userName)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_replacement_history_Del(Id As Integer) As Boolean
        If Id <= 0 Then Return False
        Return SCP.BLL.hs_Cron_replacement_history.Del(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_replacement_history_List(ParentId As Integer) As List(Of hs_Cron_replacement_history)
        If ParentId <= 0 Then Return Nothing
        Return SCP.BLL.hs_Cron_replacement_history.List(ParentId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_replacement_history_Read(Id As Integer) As hs_Cron_replacement_history
        If Id <= 0 Then Return Nothing
        Return SCP.BLL.hs_Cron_replacement_history.Read(Id)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_Cron_replacement_history_Update(Id As Integer, marcamodello As String, installationDate As String, note As String, userName As String) As Boolean
        If Id <= 0 Then Return False
        Dim _installationDate As Date = FormatDateTime("01/01/1900", DateFormat.GeneralDate)
        If IsDate(installationDate) Then
            _installationDate = CDate(installationDate)
        End If
        If _installationDate.Year <= 1900 Then Return False
        Return SCP.BLL.hs_Cron_replacement_history.Update(Id, marcamodello, _installationDate, note, userName)
    End Function
#End Region
#Region "hs_amb_Profile (cronotermostati profili correnti)"
    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_get(hsid As Integer, CronCod As String, ProfileYear As Integer, ProfileNr As Integer) As hs_amb_Profile
        Return SCP.BLL.hs_amb_Profile.getProfile(hsid, CronCod, ProfileYear, ProfileNr)
        'Dim retVal As hs_amb_Profile = Nothing

        'If hsid <= 0 Then Return Nothing
        'If String.IsNullOrEmpty(CronCod) Then Return Nothing

        'Dim CronId As Integer = 0
        'Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.ReadByCronCod(hsid, CronCod)
        'If Not m_hs_Cron Is Nothing Then
        '    CronId = m_hs_Cron.CronId
        '    m_hs_Cron = Nothing
        'End If
        'If CronId <= 0 Then Return Nothing

        'Dim connectionType As Integer = 0
        'Dim remoteAddress As String = String.Empty
        'Dim remotePort As Integer = 0
        'Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsid)
        'If Not hs Is Nothing Then
        '    If hs.VPNConnectionId > 0 Then
        '        Dim remoteConnection As SCP.BLL.Impianti_RemoteConnections = SCP.BLL.Impianti_RemoteConnections.Read(hs.IdImpianto, hs.VPNConnectionId)
        '        If Not remoteConnection Is Nothing Then
        '            If remoteConnection.connectionType = 2 Or remoteConnection.connectionType = 3 Then
        '                connectionType = remoteConnection.connectionType
        '                remoteAddress = remoteConnection.remoteAddress
        '                remotePort = remoteConnection.remotePort
        '            End If
        '            remoteConnection = Nothing
        '        End If
        '    End If
        '    hs = Nothing
        'End If
        'If remotePort <= 0 Then Return Nothing

        'Dim i As Int16 = 0
        'Dim s As String = String.Empty

        'Dim sendObj As String = "0C00"
        'Dim sendIdx As String = Replace(CronCod, "CRON", String.Empty) & "00"
        'Dim sendCmd As String = "0100"
        'Dim sendLen As String = "0300"
        'Dim sendData As String = String.Empty

        'i = Convert.ToInt16(ProfileYear)
        's = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
        'sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(ProfileNr, 16).PadLeft(2, "0")

        'Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

        'Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
        'Dim _wr As New hsdsm
        'm_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
        '_wr = Nothing

        'If m_gmDsm.returnValue = True Then

        '    Dim retCode As String = Mid(m_gmDsm.returnData, 9, 4)
        '    If retCode = "0000" Then

        '        Dim lBytes As Byte() = Array.CreateInstance(GetType(Byte), 2)
        '        lBytes(0) = CInt("&H" & Mid(m_gmDsm.returnData, 13, 2))
        '        lBytes(1) = CInt("&H" & Mid(m_gmDsm.returnData, 15, 2))
        '        Dim dataLen As Integer = BitConverter.ToInt16(lBytes, 0)

        '        If dataLen > 0 Then
        '            Dim ar() As String = New String(95) {}
        '            Dim data2Elabor As String = Mid(m_gmDsm.returnData, 17)
        '            Dim ii As Integer = 0
        '            For x As Integer = 0 To (dataLen * 2) - 1 Step 4
        '                ar(ii) = Mid(data2Elabor, x + 1, 4)
        '                ii = ii + 1
        '            Next
        '            retVal = New hs_amb_Profile
        '            retVal.CronId = CronId
        '            retVal.ProfileY = ProfileYear
        '            retVal.ProfileNr = ProfileNr

        '            Dim profileData() As Decimal = New Decimal(95) {}
        '            For x As Integer = 0 To 95
        '                Dim tset As Decimal = 0
        '                lBytes(0) = CInt("&H" & Mid(ar(x), 1, 2))
        '                lBytes(1) = CInt("&H" & Mid(ar(x), 3, 2))
        '                tset = BitConverter.ToInt16(lBytes, 0) / 10
        '                profileData(x) = tset
        '            Next
        '            retVal.ProfileData = profileData

        '            Dim obj As SCP.BLL.hs_amb_Profile = SCP.BLL.hs_amb_Profile.Read(retVal.CronId, retVal.ProfileY, retVal.ProfileNr)
        '            If Not obj Is Nothing Then
        '                SCP.BLL.hs_amb_Profile.Update(retVal.CronId, retVal.ProfileY, retVal.ProfileNr, obj.descr, retVal.ProfileData)
        '                retVal.descr = obj.descr
        '                obj = Nothing
        '            Else
        '                SCP.BLL.hs_amb_Profile.Add(retVal.CronId, retVal.ProfileY, retVal.ProfileNr, "profilo " & retVal.ProfileNr.ToString, retVal.ProfileData)
        '                retVal.descr = "profilo " & retVal.ProfileNr.ToString
        '            End If


        '        End If ' datalen >0

        '    End If ' recCode OK

        'End If


        'Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_List(CronId As Integer, ProfileY As Integer) As List(Of hs_amb_Profile)
        If CronId <= 0 Then Return Nothing
        If ProfileY <= 1900 Then Return Nothing
        Return SCP.BLL.hs_amb_Profile.List(CronId, ProfileY)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Read(CronId As Integer, ProfileYear As Integer, ProfileNr As Integer) As hs_amb_Profile
        Return SCP.BLL.hs_amb_Profile.Read(CronId, ProfileYear, ProfileNr)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_set(hsid As Integer, CronCod As String, ProfileYear As Integer, ProfileNr As Integer, strProfileData As String) As Boolean
        Dim retVal As Boolean = False

        If hsid <= 0 Then Return False
        If String.IsNullOrEmpty(CronCod) Then Return False

        Dim CronId As Integer = 0
        Dim m_hs_Cron As SCP.BLL.hs_Cron = SCP.BLL.hs_Cron.ReadByCronCod(hsid, CronCod)
        If Not m_hs_Cron Is Nothing Then
            CronId = m_hs_Cron.CronId
            m_hs_Cron = Nothing
        End If
        If CronId <= 0 Then Return False

        Dim connectionType As Integer = 0
        Dim remoteAddress As String = String.Empty
        Dim remotePort As Integer = 0
        Dim hs As SCP.BLL.HeatingSystem = SCP.BLL.HeatingSystem.Read(hsid)
        If Not hs Is Nothing Then
            If hs.VPNConnectionId > 0 Then
                Dim remoteConnection As SCP.BLL.Impianti_RemoteConnections = SCP.BLL.Impianti_RemoteConnections.Read(hs.IdImpianto, hs.VPNConnectionId)
                If Not remoteConnection Is Nothing Then
                    If remoteConnection.connectionType = 2 Or remoteConnection.connectionType = 3 Then
                        connectionType = remoteConnection.connectionType
                        remoteAddress = remoteConnection.remoteAddress
                        remotePort = remoteConnection.remotePort
                    End If
                    remoteConnection = Nothing
                End If
            End If
            hs = Nothing
        End If
        If remotePort <= 0 Then Return False


        Dim ProfileData As Decimal() = New Decimal(95) {}
        Dim ar() As String = strProfileData.Split(";")
        For y As Integer = 0 To 95
            ProfileData(y) = CDec(ar(y))
        Next

        Dim i As Int16 = 0
        Dim s As String = String.Empty

        Dim sendObj As String = "0C00"
        ' Dim sendIdx As String = Replace(CronCod, "CRON", String.Empty) & "00"

        Dim inx As Integer = CInt(Replace(CronCod, "CRON", String.Empty))
        Dim sendIdx As String = UCase(Convert.ToString(inx, 16).PadLeft(2, "0")) & "00" ' Replace(CronCod, "CRON", String.Empty) & "00"


        Dim sendCmd As String = "0200"
        Dim sendLen As String = "C300"
        Dim sendData As String = String.Empty

        i = Convert.ToInt16(ProfileYear)
        s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
        sendData = Mid(s, 3, 2) & Mid(s, 1, 2) & Convert.ToString(ProfileNr, 16).PadLeft(2, "0")

        Dim lBytes As Byte() = Array.CreateInstance(GetType(Byte), 2)

        For x As Integer = 0 To 95
            i = Convert.ToInt16(ProfileData(x))
            s = UCase(Convert.ToString(i, 16)).PadLeft(4, "0")
            Dim TsetHex As String = Mid(s, 3, 2) & Mid(s, 1, 2)
            sendData += TsetHex
        Next

        Dim sendString As String = sendObj & sendIdx & sendCmd & sendLen & sendData

        Dim m_gmDsm As gmDsmResponse = New gmDsmResponse
        Dim _wr As New hsdsm
        m_gmDsm = _wr.send(connectionType, remoteAddress, remotePort, sendString)
        _wr = Nothing

        If m_gmDsm.returnValue = True Then
            If m_gmDsm.returnData.Length >= 16 Then
                Dim responseObj As String = Mid(m_gmDsm.returnData, 1, 4)
                Dim responseIdx As String = Mid(m_gmDsm.returnData, 5, 4)
                Dim responseCod As String = Mid(m_gmDsm.returnData, 9, 4)
                If responseCod = "0000" Then
                    retVal = True
                End If
            End If
        End If

        Return retVal
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Update(CronId As Integer, ProfileY As Integer, ProfileNr As Integer, descr As String, ProfileData As Decimal()) As Boolean
        If CronId <= 0 Then Return False
        If ProfileY <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        If ProfileData.Length <= 0 Then Return False
        Return SCP.BLL.hs_amb_Profile.Update(CronId, ProfileY, ProfileNr, descr, ProfileData)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_setProfileNow(CronId As Integer, ProfileNr As Integer) As Boolean
        If CronId <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        Return SCP.BLL.hs_amb_Profile.setProfileNow(CronId, ProfileNr)
    End Function
#End Region
#Region "hs_amb_Profile_Descr"
    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Descr_Add(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
        If CronId <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        If String.IsNullOrEmpty(descr) Then Return False
        Return SCP.BLL.hs_amb_Profile_Descr.Add(CronId, ProfileNr, descr)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Descr_Del(CronId As Integer, ProfileNr As Integer) As Boolean
        If CronId <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        Return SCP.BLL.hs_amb_Profile_Descr.Del(CronId, ProfileNr)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Descr_List(CronId As Integer) As List(Of hs_amb_Profile_Descr)
        If CronId <= 0 Then Return Nothing
        Return SCP.BLL.hs_amb_Profile_Descr.List(CronId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Descr_Read(CronId As Integer, ProfileNr As Integer) As hs_amb_Profile_Descr
        If CronId <= 0 Then Return Nothing
        If ProfileNr < 0 Then Return Nothing
        Return SCP.BLL.hs_amb_Profile_Descr.Read(CronId, ProfileNr)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Descr_Update(CronId As Integer, ProfileNr As Integer, descr As String) As Boolean
        If CronId <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        If String.IsNullOrEmpty(descr) Then Return False
        Return SCP.BLL.hs_amb_Profile_Descr.Update(CronId, ProfileNr, descr)
    End Function
#End Region
#Region "hs_amb_Profile_Tasks"
    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_Add(CronId As Integer,
                                              ProfileNr As Integer,
                                              Subject As String,
                                              StartDate As String,
                                              EndDate As String,
                                              chkMonday As Boolean,
                                              chkTuesday As Boolean,
                                              chkWednesday As Boolean,
                                              chkThursday As Boolean,
                                              chkFriday As Boolean,
                                              chkSaturday As Boolean,
                                              chkSunday As Boolean,
                                              yearsRepeatable As Boolean) As Boolean
        If CronId <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        If String.IsNullOrEmpty(Subject) Then Return False
        Return SCP.BLL.hs_amb_Profile_Tasks.Add(CronId, ProfileNr, Subject, StartDate, EndDate, chkMonday, chkTuesday, chkWednesday, chkThursday, chkFriday, chkSaturday, chkSunday, yearsRepeatable)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_Del(CronId As Integer, ProfileNr As Integer) As Boolean
        If CronId <= 0 Then Return False
        If ProfileNr < 0 Then Return False
        Return SCP.BLL.hs_amb_Profile_Tasks.Del(CronId, ProfileNr)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_List(CronId As Integer) As List(Of hs_amb_Profile_Tasks)
        If CronId <= 0 Then Return Nothing
        Return SCP.BLL.hs_amb_Profile_Tasks.List(CronId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_ListAll() As List(Of hs_amb_Profile_Tasks)
        Return SCP.BLL.hs_amb_Profile_Tasks.ListAll()
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_ListByMonth(CronId As Integer, calYear As Integer, calMonth As Integer) As List(Of hs_amb_Profile_Tasks)
        If CronId <= 0 Then Return Nothing
        Return SCP.BLL.hs_amb_Profile_Tasks.ListByMonth(CronId, calYear, calMonth)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_ListByDates(CronId As Integer, startDate As String, endDate As String) As List(Of hs_amb_Profile_Tasks)
        If CronId <= 0 Then Return Nothing

        Dim _startDate As Date = CDate(startDate)
        Dim a As Date = CDate(endDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _endDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.hs_amb_Profile_Tasks.ListByDates(CronId, _startDate, _endDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_ListCurrent(CronId As Integer, selDate As String) As List(Of hs_amb_Profile_Tasks)
        If CronId <= 0 Then Return Nothing
        Dim _selDate As Date = Now
        If Not String.IsNullOrEmpty(selDate) Then
            _selDate = FormatDateTime(selDate, DateFormat.GeneralDate) ' CDate(selDate)
        End If
        Return SCP.BLL.hs_amb_Profile_Tasks.ListCurrent(CronId, _selDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_Read(TaskId As Integer) As hs_amb_Profile_Tasks
        If TaskId <= 0 Then Return Nothing
        Return SCP.BLL.hs_amb_Profile_Tasks.Read(TaskId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Profile_Tasks_Update(TaskId As Integer,
                                                 ProfileNr As Integer,
                                                 Subject As String,
                                                 StartDate As String,
                                                 EndDate As String,
                                                 chkMonday As Boolean,
                                                 chkTuesday As Boolean,
                                                 chkWednesday As Boolean,
                                                 chkThursday As Boolean,
                                                 chkFriday As Boolean,
                                                 chkSaturday As Boolean,
                                                 chkSunday As Boolean,
                                                 yearsRepeatable As Boolean) As Boolean
        If TaskId <= 0 Then Return Nothing
        If ProfileNr < 0 Then Return Nothing
        If String.IsNullOrEmpty(Subject) Then Return Nothing
        Return SCP.BLL.hs_amb_Profile_Tasks.Update(TaskId, ProfileNr, Subject, StartDate, EndDate, chkMonday, chkTuesday, chkWednesday, chkThursday, chkFriday, chkSaturday, chkSunday, yearsRepeatable)
    End Function
#End Region
#Region "hs_amb_Calendar"
    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Calendar_Add(CronId As Integer, Calyear As Integer, Calmonth As Integer, strmonthData As String) As Boolean
        If CronId <= 0 Then Return False
        If Calyear <= 0 Then Return False
        If Calmonth <= 0 Then Return False
        If Calmonth > 12 Then Return False
        Dim monthData As Integer() = New Integer(32) {}
        Dim ar() As String = strmonthData.Split(";")
        For x As Integer = 0 To 31
            monthData(x) = CInt(ar(x))
        Next
        Return SCP.BLL.hs_amb_Calendar.Add(CronId, Calyear, Calmonth, monthData)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Calendar_Read(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
        If CronId <= 0 Then Return Nothing
        If Calyear <= 0 Then Return Nothing
        If Calmonth <= 0 Then Return Nothing
        If Calmonth > 12 Then Return Nothing
        Return SCP.BLL.hs_amb_Calendar.Read(CronId, Calyear, Calmonth)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_CalendarReadFromDB(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
        If CronId <= 0 Then Return Nothing
        If Calyear <= 0 Then Return Nothing
        If Calmonth <= 0 Then Return Nothing
        If Calmonth > 12 Then Return Nothing
        Return SCP.BLL.hs_amb_Calendar.Read(CronId, Calyear, Calmonth)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Calendar_UpdateDesired(CronId As Integer, Calyear As Integer, Calmonth As Integer, strmonthData As String) As Boolean
        If CronId <= 0 Then Return False
        If Calyear <= 0 Then Return False
        If Calmonth <= 0 Then Return False
        If Calmonth > 12 Then Return False
        Dim DesiredMonthData As Integer() = New Integer(31) {}
        Dim ar() As String = strmonthData.Split(";")
        For x As Integer = 0 To ar.Length - 2
            DesiredMonthData(x) = CInt(ar(x))
        Next
        Return SCP.BLL.hs_amb_Calendar.UpdateDesired(CronId, Calyear, Calmonth, DesiredMonthData)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Calendar_get(CronId As Integer, Calyear As Integer, Calmonth As Integer) As hs_amb_Calendar
        Return SCP.BLL.hs_amb_Calendar.CalendarGet(CronId, Calyear, Calmonth)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function hs_amb_Calendar_Set(CronId As Integer, Calyear As Integer, Calmonth As Integer) As Boolean
        Return SCP.BLL.hs_amb_Calendar.CalendarSet(CronId, Calyear, Calmonth)
    End Function

#End Region
#Region "log_hs_Cron"
    <OperationContract()>
    <WebGet()>
    Public Function log_hs_Cron_List(hsId As Integer, CronCod As String, fromDate As String, toDate As String) As List(Of log_hs_Cron)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(CronCod) Then
            Return Nothing
        End If
        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)
        Return SCP.BLL.log_hs_Cron.List(hsId, CronCod, _fromDate, _toDate)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function log_hs_Cron_ListPaged(hsId As Integer, CronCod As String, fromDate As String, toDate As String, rowNumber As Integer) As List(Of log_hs_Cron)
        If hsId <= 0 Then
            Return Nothing
        End If
        If String.IsNullOrEmpty(CronCod) Then
            Return Nothing
        End If

        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)

        Return SCP.BLL.log_hs_Cron.ListPaged(hsId, CronCod, _fromDate, _toDate, rowNumber)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function log_hs_Cron_ListAll(hsId As Integer, fromDate As String, toDate As String) As List(Of log_hs_Cron)
        If hsId <= 0 Then
            Return Nothing
        End If
        Dim _fromDate As Date = CDate(fromDate)
        Dim a As Date = CDate(toDate)
        Dim b As Date = DateAdd(DateInterval.Day, 1, a)
        Dim _toDate As Date = DateAdd(DateInterval.Minute, -1, b)
        Return SCP.BLL.log_hs_Cron.ListAll(hsId, _fromDate, _toDate)
    End Function
#End Region

#Region "Ambienti"
    <OperationContract()>
    <WebGet()>
    Public Function Ambienti_List(hsId As Integer) As List(Of Ambienti)
        If hsId <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Ambienti.List(hsId)
    End Function

    <OperationContract()>
    <WebGet()>
    Public Function Ambienti_Read(IdAmbiente As Integer) As Ambienti
        If IdAmbiente <= 0 Then
            Return Nothing
        End If
        Return SCP.BLL.Ambienti.Read(IdAmbiente)
    End Function
#End Region

End Class
