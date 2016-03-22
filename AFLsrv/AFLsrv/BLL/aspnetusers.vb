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
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports SCP.DAL

Namespace SCP.BLL
    Public Class aspnetUsers
        Private _UserName As String
        Private _Email As String
        Private _Password As String
        Private _Comment As String
        Private _RoleName As String
        Private _IsApproved As Boolean
        Private _CreateDate As Date
        Private _LastLoginDate As Date
        Private _LastActivityDate As Date
        Private _LastPasswordChangedDate As Date
        Private _UserId As String
        Private _IsLockedOut As Boolean
        Private _LastLockoutDate As Date

        Private _totImpianti As Integer

#Region "constructor"
        Private Shared _getUsersByRoleName As List(Of aspnetUsers)

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, False, FormatDateTime("01/01/1900", DateFormat.GeneralDate), FormatDateTime("01/01/1900", DateFormat.GeneralDate), FormatDateTime("01/01/1900", DateFormat.GeneralDate), FormatDateTime("01/01/1900", DateFormat.GeneralDate), String.Empty, False, FormatDateTime("01/01/1900", DateFormat.GeneralDate), 0)
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="UserName"></param>
        ''' <param name="Email"></param>
        ''' <param name="Password"></param>
        ''' <param name="Comment"></param>
        ''' <param name="IsApproved"></param>
        ''' <param name="CreateDate"></param>
        ''' <param name="LastLoginDate"></param>
        ''' <param name="LastActivityDate"></param>
        ''' <param name="LastPasswordChangedDate"></param>
        ''' <param name="UserId"></param>
        ''' <param name="IsLockedOut"></param>
        ''' <param name="LastLockoutDate"></param>
        ''' <remarks></remarks>
        Public Sub New(UserName As String, _
                       Email As String, _
                       Password As String, _
                       Comment As String, _
                       RoleName As String, _
                       IsApproved As Boolean, _
                       CreateDate As Date, _
                       LastLoginDate As Date, _
                       LastActivityDate As Date, _
                       LastPasswordChangedDate As Date, _
                       UserId As String, _
                       IsLockedOut As Boolean, _
                       LastLockoutDate As Date, _
                       totImpianti As Integer)
            Me._UserName = UserName
            Me._Email = Email
            Me._Password = Password
            Me._Comment = Comment
            Me._RoleName = RoleName
            Me._IsApproved = IsApproved
            Me._CreateDate = CreateDate
            Me._LastLoginDate = LastLoginDate
            Me._LastActivityDate = LastActivityDate
            Me._LastPasswordChangedDate = LastPasswordChangedDate
            Me._UserId = UserId
            Me._IsLockedOut = IsLockedOut
            Me._LastLockoutDate = LastLockoutDate

            Me._totImpianti = totImpianti
        End Sub


#End Region

#Region "methods"
        Public Shared Function AddUser(username As String, password As String, rolename As String, comment As String, email As String) As Boolean
            Dim retVal As Boolean = False
            Try

                Dim result As MembershipUser = Membership.CreateUser(username, password)
                If Not result Is Nothing Then
                    result.IsApproved = True
                    result.Comment = comment
                    result.Email = email
                    Roles.AddUserToRole(result.UserName, rolename)
                    Membership.UpdateUser(result)
                    result = Nothing
                    retVal = True
                Else
                    retVal = False
                End If
            Catch ex As Exception
                retVal = False
            End Try
            Return retVal
        End Function

        Public Shared Function UpdUser(username As String, comment As String, email As String) As Boolean
            Dim retVal As Boolean = False
            Try
                Dim result As MembershipUser = Membership.GetUser(username)
                If Not result Is Nothing Then
                    result.Comment = comment
                    result.Email = email
                    Membership.UpdateUser(result)
                    result = Nothing
                    retVal = True
                End If
            Catch ex As Exception
                retVal = False
            End Try
            Return retVal
        End Function

        Public Shared Function approveUser(username As String) As Boolean
            Dim retVal As Boolean = False
            Try
                Dim userInfo As MembershipUser = Membership.GetUser(username)
                Dim UserId As String = userInfo.ProviderUserKey.ToString
                retVal = DataAccessHelper.GetDataAccess.aspnetUsers_approveUser(UserId)
            Catch ex As Exception
                retVal = False
            End Try
            Return retVal
        End Function

        Public Shared Function disapproveUser(username As String) As Boolean
            Dim retVal As Boolean = False
            Try
                Dim userInfo As MembershipUser = Membership.GetUser(username)
                Dim UserId As String = userInfo.ProviderUserKey.ToString
                retVal = DataAccessHelper.GetDataAccess.aspnetUsers_disapproveUser(UserId)
            Catch ex As Exception
                retVal = False
            End Try
            Return retVal
        End Function

        Public Shared Function lockUser(username As String) As Boolean
            Dim retval As Boolean = False
            Try
                Dim userInfo As MembershipUser = Membership.GetUser(username)
                Dim UserId As String = userInfo.ProviderUserKey.ToString
                userInfo = Nothing
                retval = DataAccessHelper.GetDataAccess.aspnetUsers_lockUser(UserId)
            Catch ex As Exception
                retval = False
            End Try
            Return retval
        End Function

        Public Shared Function UnlockUser(username As String) As Boolean
            Dim retval As Boolean = False
            Try
                Dim userInfo As MembershipUser = Membership.GetUser(username)
                userInfo.UnlockUser()
                userInfo = Nothing
                retval = True
            Catch ex As Exception
                retval = False
            End Try
            Return retval
        End Function

        Public Shared Function changeUserPass(username As String, oldPass As String, newPass As String) As Boolean
            Dim retVal As Boolean = False
            Try
                Dim result As MembershipUser = Membership.GetUser(username)
                If Not result Is Nothing Then
                    If result.ChangePassword(oldPass, newPass) = True Then
                        retVal = True
                    End If
                End If
            Catch ex As Exception
                retVal = False
            End Try
            Return retVal
        End Function

        Public Shared Function removeUser(username As String) As Boolean
            Dim retVal As Boolean = False

            Dim _user As aspnetUsers = aspnetUsers.GetUserByUserName(username)
            Try
                retVal = Membership.DeleteUser(username)
                If retVal = True Then
                    If Not _user Is Nothing Then
                        'Impianti_Users.Clear(_user.UserId)
                        'Impianti_UserDevices.Clear(_user.UserId)
                    End If
                End If
            Catch ex As Exception
                retVal = False
            End Try
            Return retVal
        End Function

        Public Shared Function resetP(username As String) As String
            Dim retVal As String = String.Empty
            Dim userInfo As MembershipUser = Membership.GetUser(username)
            If Not userInfo Is Nothing Then
                retVal = userInfo.ResetPassword()
            End If
            userInfo = Nothing
            Return retVal
        End Function

        Public Shared Function GetActiveUsersByRoleName(RoleName As String) As List(Of aspnetUsers)
            If String.IsNullOrEmpty(RoleName) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess().aspnetUsers_GetActiveUsersByRoleName(RoleName)
        End Function

        Public Shared Function GetTotActiveUsersByRoleName(RoleName As String) As Integer
            If String.IsNullOrEmpty(RoleName) Then
                Return 0
            End If
            Return DataAccessHelper.GetDataAccess.aspnetUsers_GetTotActiveUsersByRoleName(RoleName)
        End Function

        Public Shared Function GetUsersByRoleName(RoleName As String, searchString As String, IdImpianto As String) As List(Of aspnetUsers)

            If String.IsNullOrEmpty(RoleName) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess().aspnetUsers_GetUsersByRoleName(RoleName, searchString, IdImpianto)

        End Function

        Public Shared Function GetUsersAll() As List(Of aspnetUsers)
            Return DataAccessHelper.GetDataAccess().aspnetUsers_GetUsersAll()
        End Function

        Public Shared Function GetUserByUserName(UserName As String) As aspnetUsers
            If String.IsNullOrEmpty(UserName) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess().aspnetUsers_GetUserByUserName(UserName)
        End Function

        Public Shared Function List(id As Integer, SearchString As String) As List(Of aspnetUsers)
            Return DataAccessHelper.GetDataAccess.aspnetUsers_List(id, SearchString)
        End Function

        Public Shared Function userok(UserId As String) As Boolean
            Dim retVal As Boolean = False
            Dim _user As aspnetUsers = DataAccessHelper.GetDataAccess.aspnetUsers_Read(UserId)
            If Not _user Is Nothing Then
                If _user.IsApproved = True And _user.IsLockedOut = False Then
                    retVal = True
                End If
            End If
            Return retVal
        End Function

        Public Shared Function whoIsOnline() As List(Of String)
            Return DataAccessHelper.GetDataAccess.aspnetUsers_whoIsOnline
        End Function

        Public Shared Function getUserO(UserId As String) As String
            Dim retVal As String = String.Empty
            Dim _user As aspnetUsers = DataAccessHelper.GetDataAccess.aspnetUsers_Read(UserId)
            If Not _user Is Nothing Then
                If _user.IsApproved = True And _user.IsLockedOut = False Then
                    If _user.RoleName = "Operators" Then
                        retVal = _user.UserName
                    End If
                End If
            End If
            Return retVal
        End Function

        Public Shared Function getUserS(UserId As String) As String
            Dim retVal As String = String.Empty
            Dim _user As aspnetUsers = DataAccessHelper.GetDataAccess.aspnetUsers_Read(UserId)
            If Not _user Is Nothing Then
                If _user.IsApproved = True And _user.IsLockedOut = False Then
                    If _user.RoleName = "Supervisors" Then
                        retVal = _user.UserName
                    End If
                End If
            End If
            Return retVal
        End Function

        Public Shared Function getUserA(UserId As String) As String
            Dim retVal As String = String.Empty
            Dim _user As aspnetUsers = DataAccessHelper.GetDataAccess.aspnetUsers_Read(UserId)
            If Not _user Is Nothing Then
                If _user.IsApproved = True And _user.IsLockedOut = False Then
                    If _user.RoleName = "Administrators" Then
                        retVal = _user.UserName
                    End If
                End If
            End If
            Return retVal
        End Function
#End Region

#Region "properties"
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property UserName() As String
            Get
                If String.IsNullOrEmpty(Me._UserName) Then
                    Return String.Empty
                Else
                    Return Me._UserName
                End If
            End Get
            Set(value As String)
                Me._UserName = value
            End Set
        End Property

        Public Property Email() As String
            Get
                If String.IsNullOrEmpty(Me._Email) Then
                    Return String.Empty
                Else
                    Return Me._Email
                End If
            End Get
            Set(value As String)
                Me._Email = value
            End Set
        End Property

        Public Property Password() As String
            Get
                If String.IsNullOrEmpty(Me._Password) Then
                    Return String.Empty
                Else
                    Return Me._Password
                End If
            End Get
            Set(value As String)
                Me._Password = value
            End Set
        End Property

        Public Property Comment() As String
            Get
                If String.IsNullOrEmpty(Me._Comment) Then
                    Return String.Empty
                Else
                    Return Me._Comment
                End If
            End Get
            Set(value As String)
                Me._Comment = value
            End Set
        End Property

        Public Property RoleName() As String
            Get
                Return Me._RoleName
            End Get
            Set(value As String)
                Me._RoleName = value
            End Set
        End Property

        Public Property IsApproved() As Boolean
            Get
                Return Me._IsApproved
            End Get
            Set(value As Boolean)
                Me._IsApproved = value
            End Set
        End Property

        Public Property CreateDate() As Date
            Get
                Return Me._CreateDate
            End Get
            Set(value As Date)
                Me._CreateDate = value
            End Set
        End Property

        Public Property LastLoginDate() As Date
            Get
                Return Me._LastLoginDate
            End Get
            Set(value As Date)
                Me._LastLoginDate = value
            End Set
        End Property

        Public Property LastActivityDate() As Date
            Get
                Return Me._LastActivityDate
            End Get
            Set(value As Date)
                Me._LastActivityDate = value
            End Set
        End Property

        Public Property LastPasswordChangedDate() As Date
            Get
                Return Me._LastPasswordChangedDate
            End Get
            Set(value As Date)
                Me._LastPasswordChangedDate = value
            End Set
        End Property

        Public Property UserId() As String
            Get
                If String.IsNullOrEmpty(Me._UserId) Then
                    Return String.Empty
                Else
                    Return Me._UserId
                End If
            End Get
            Set(value As String)
                Me._UserId = value
            End Set
        End Property

        Public Property IsLockedOut() As Boolean
            Get
                Return Me._IsLockedOut
            End Get
            Set(value As Boolean)
                Me._IsLockedOut = value
            End Set
        End Property

        Public Property LastLockoutDate() As Date
            Get
                Return Me._LastLockoutDate
            End Get
            Set(value As Date)
                Me._LastLockoutDate = value
            End Set
        End Property

        Public Property totImpianti As Integer
            Get
                Return Me._totImpianti
            End Get
            Set(value As Integer)
                Me._totImpianti = value
            End Set
        End Property
#End Region

#Region "Password"
        ' Define default min and max password lengths.
        Private Shared DEFAULT_MIN_PASSWORD_LENGTH As Integer = 8
        Private Shared DEFAULT_MAX_PASSWORD_LENGTH As Integer = 16

        ' Define supported password characters divided into groups.
        ' You can add (or remove) characters to (from) these groups.
        Private Shared PASSWORD_CHARS_LCASE As String = "abcdefgijkmnopqrstwxyz"
        Private Shared PASSWORD_CHARS_UCASE As String = "ABCDEFGHJKLMNPQRSTWXYZ"
        Private Shared PASSWORD_CHARS_NUMERIC As String = "23456789"
        Private Shared PASSWORD_CHARS_SPECIAL As String = "*$-+?_&=!%{}/"

        Public Shared Function GeneratePwd() As String
            GeneratePwd = Generate(DEFAULT_MIN_PASSWORD_LENGTH, _
                                DEFAULT_MAX_PASSWORD_LENGTH, True)
        End Function

        Public Shared Function GeneratePwd(length As Integer) As String
            GeneratePwd = Generate(length, length, True)
        End Function

        Public Shared Function GeneratePwd(minLenght As Integer, maxLenght As Integer) As String
            GeneratePwd = Generate(minLenght, maxLenght, True)
        End Function

        Public Shared Function GenerateLogin() As String
            GenerateLogin = Generate(DEFAULT_MIN_PASSWORD_LENGTH, DEFAULT_MAX_PASSWORD_LENGTH, False)
        End Function

        Public Shared Function GenerateLogin(length As Integer) As String
            GenerateLogin = Generate(length, length, False)
        End Function

        Public Shared Function GenerateLogin(minLenght As Integer, maxLenght As Integer) As String
            GenerateLogin = Generate(minLenght, maxLenght, False)
        End Function

        Public Shared Function Generate(minLength As Integer, _
        maxLength As Integer, isSimple As Boolean) As String

            ' Make sure that input parameters are valid.
            If (minLength <= 0 Or maxLength <= 0 Or minLength > maxLength) Then
                Generate = Nothing
            End If

            ' Create a local array containing supported password characters
            ' grouped by types. You can remove character groups from this
            ' array, but doing so will weaken the password strength.
            Dim charGroups As Char()()
            If isSimple Then
                charGroups = New Char()() _
                { _
                    PASSWORD_CHARS_LCASE.ToCharArray(), _
                    PASSWORD_CHARS_UCASE.ToCharArray(), _
                    PASSWORD_CHARS_NUMERIC.ToCharArray(), _
                    PASSWORD_CHARS_SPECIAL.ToCharArray() _
                }
            Else
                charGroups = New Char()() _
                { _
                    PASSWORD_CHARS_LCASE.ToCharArray(), _
                    PASSWORD_CHARS_NUMERIC.ToCharArray() _
                }
            End If

            ' Use this array to track the number of unused characters in each
            ' character group.
            Dim charsLeftInGroup As Integer() = New Integer(charGroups.Length - 1) {}

            ' Initially, all characters in each group are not used.
            Dim I As Integer
            For I = 0 To charsLeftInGroup.Length - 1
                charsLeftInGroup(I) = charGroups(I).Length
            Next

            ' Use this array to track (iterate through) unused character groups.
            Dim leftGroupsOrder As Integer() = New Integer(charGroups.Length - 1) {}

            ' Initially, all character groups are not used.
            For I = 0 To leftGroupsOrder.Length - 1
                leftGroupsOrder(I) = I
            Next

            ' Because we cannot use the default randomizer, which is based on the
            ' current time (it will produce the same "random" number within a
            ' second), we will use a random number generator to seed the
            ' randomizer.

            ' Use a 4-byte array to fill it with random bytes and convert it then
            ' to an integer value.
            Dim randomBytes As Byte() = New Byte(3) {}

            ' Generate 4 random bytes.
            Dim rng As RNGCryptoServiceProvider = New RNGCryptoServiceProvider()

            rng.GetBytes(randomBytes)

            ' Convert 4 bytes into a 32-bit integer value.
            Dim seed As Integer = ((randomBytes(0) And &H7F) << 24 Or _
                                    randomBytes(1) << 16 Or _
                                    randomBytes(2) << 8 Or _
                                    randomBytes(3))

            ' Now, this is real randomization.
            Dim random As Random = New Random(seed)

            ' This array will hold password characters.
            Dim password As Char() = Nothing

            ' Allocate appropriate memory for the password.
            If (minLength < maxLength) Then
                password = New Char(random.Next(minLength - 1, maxLength)) {}
            Else
                password = New Char(minLength - 1) {}
            End If

            ' Index of the next character to be added to password.
            Dim nextCharIdx As Integer

            ' Index of the next character group to be processed.
            Dim nextGroupIdx As Integer

            ' Index which will be used to track not processed character groups.
            Dim nextLeftGroupsOrderIdx As Integer

            ' Index of the last non-processed character in a group.
            Dim lastCharIdx As Integer

            ' Index of the last non-processed group.
            Dim lastLeftGroupsOrderIdx As Integer = leftGroupsOrder.Length - 1

            ' Generate password characters one at a time.
            For I = 0 To password.Length - 1

                ' If only one character group remained unprocessed, process it;
                ' otherwise, pick a random character group from the unprocessed
                ' group list. To allow a special character to appear in the
                ' first position, increment the second parameter of the Next
                ' function call by one, i.e. lastLeftGroupsOrderIdx + 1.
                If (lastLeftGroupsOrderIdx = 0) Then
                    nextLeftGroupsOrderIdx = 0
                Else
                    nextLeftGroupsOrderIdx = random.Next(0, lastLeftGroupsOrderIdx)
                End If

                ' Get the actual index of the character group, from which we will
                ' pick the next character.
                nextGroupIdx = leftGroupsOrder(nextLeftGroupsOrderIdx)

                ' Get the index of the last unprocessed characters in this group.
                lastCharIdx = charsLeftInGroup(nextGroupIdx) - 1

                ' If only one unprocessed character is left, pick it; otherwise,
                ' get a random character from the unused character list.
                If (lastCharIdx = 0) Then
                    nextCharIdx = 0
                Else
                    nextCharIdx = random.Next(0, lastCharIdx + 1)
                End If

                ' Add this character to the password.
                password(I) = charGroups(nextGroupIdx)(nextCharIdx)

                ' If we processed the last character in this group, start over.
                If (lastCharIdx = 0) Then
                    charsLeftInGroup(nextGroupIdx) = _
                                    charGroups(nextGroupIdx).Length
                    ' There are more unprocessed characters left.
                Else
                    ' Swap processed character with the last unprocessed character
                    ' so that we don't pick it until we process all characters in
                    ' this group.
                    If (lastCharIdx <> nextCharIdx) Then
                        Dim temp As Char = charGroups(nextGroupIdx)(lastCharIdx)
                        charGroups(nextGroupIdx)(lastCharIdx) = _
                                    charGroups(nextGroupIdx)(nextCharIdx)
                        charGroups(nextGroupIdx)(nextCharIdx) = temp
                    End If

                    ' Decrement the number of unprocessed characters in
                    ' this group.
                    charsLeftInGroup(nextGroupIdx) = _
                               charsLeftInGroup(nextGroupIdx) - 1
                End If

                ' If we processed the last group, start all over.
                If (lastLeftGroupsOrderIdx = 0) Then
                    lastLeftGroupsOrderIdx = leftGroupsOrder.Length - 1
                    ' There are more unprocessed groups left.
                Else
                    ' Swap processed group with the last unprocessed group
                    ' so that we don't pick it until we process all groups.
                    If (lastLeftGroupsOrderIdx <> nextLeftGroupsOrderIdx) Then
                        Dim temp As Integer = _
                                    leftGroupsOrder(lastLeftGroupsOrderIdx)
                        leftGroupsOrder(lastLeftGroupsOrderIdx) = _
                                    leftGroupsOrder(nextLeftGroupsOrderIdx)
                        leftGroupsOrder(nextLeftGroupsOrderIdx) = temp
                    End If

                    ' Decrement the number of unprocessed groups.
                    lastLeftGroupsOrderIdx = lastLeftGroupsOrderIdx - 1
                End If
            Next

            ' Convert password characters into a string and return the result.
            Generate = New String(password)
        End Function

#End Region

    End Class
End Namespace
