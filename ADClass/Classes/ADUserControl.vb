Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Imports System.Reflection

Public Class ADUserControl

    Const _AdminUserName As String = "Company\Username"
    Const _AdminPassword As String = "PutPasswordHere"
    Const _ServerAddress As String = "1.1.1.1"
    Const _ADDomainName As String = "dc.company.com"
    Const _ADContainer As String = "OU=users,DC=company-us,DC=com"

    Private _ADContext As PrincipalContext
    Private _ADSearcher As PrincipalSearcher
    Private _ADGroup As GroupPrincipal
    Private _ADUser As UserPrincipal

    Public Sub New()

        _ADContext = New PrincipalContext(ContextType.Domain, _ADDomainName, _ADContainer, _AdminUserName, _AdminPassword)
        _ADUser = New UserPrincipal(_ADContext)

    End Sub

    Function getListOfWebPortalADUsers() As List(Of String)

        Try
            Dim webPortalUserList As New List(Of String)

            _ADUser = New UserPrincipal(_ADContext)
            _ADSearcher = New PrincipalSearcher(_ADUser)

            For Each userObj As UserPrincipal In _ADSearcher.FindAll()
                'Debug.WriteLine(userObj.DisplayName)
                webPortalUserList.Add(userObj.SamAccountName)
            Next

            Return webPortalUserList
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Return Nothing
        End Try

    End Function

    Public Function updateADEmailAddress(username As String, emailAddress As String) As Boolean

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, username)

            _ADUser.EmailAddress = emailAddress
            _ADUser.Save()

            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error: " + ex.ToString)

            'MessageBox.Show("Error updating email address for AD User (" + username + ") : " + ex.ToString)

            Return False
        End Try

    End Function

    Public Function deleteADUser(username As String) As Boolean

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, username)

            Dim de As DirectoryEntry = DirectCast(_ADUser.GetUnderlyingObject(), DirectoryEntry)

            de.DeleteTree()
            Debug.WriteLine("Deleted AD user: " + username)

            Return True
        Catch ex As Exception
            'MessageBox.Show("AD User not found, we cannot delete. Exiting." + vbNewLine + ex.ToString)
            Return False
        End Try

    End Function
    Function getGroupListForUser(username As String) As List(Of String)

        Dim returnGroupList As New List(Of String)

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, username)

            Dim de As DirectoryEntry = DirectCast(_ADUser.GetUnderlyingObject(), DirectoryEntry)

            'Get our membership string for ALL values, to get group membership from later
            For Each membershipString As String In de.Properties("memberOf")

                Debug.WriteLine("MembershipString: " + membershipString)
                Dim splitMembershipString As String() = membershipString.Split(CChar(","))

                If splitMembershipString(0).Substring(0, 2) = "CN" Then
                    'Debug.WriteLine("We have a group! Save it!")
                    Dim savedGroup As String = splitMembershipString(0).Substring(3, splitMembershipString(0).Length - 3)
                    'Debug.WriteLine("Our saved group name: " + savedGroup)
                    returnGroupList.Add(savedGroup)
                End If

            Next

            Debug.WriteLine("Done getting groups")

            Return returnGroupList
        Catch Ex As Exception
            'MessageBox.Show("Error getting groups for user." + Ex.ToString)
            Return returnGroupList
        End Try

    End Function

    Function removeUserFromGroup(userName As String, groupName As String) As Boolean

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, userName)

            Dim group As GroupPrincipal = GroupPrincipal.FindByIdentity(_ADContext, IdentityType.Name, groupName)

            If group.Members.Contains(_ADUser) Then
                Debug.WriteLine("Group " + groupName + " contains user: " + userName)
                group.Members.Remove(_ADUser)
                group.Save()
                Debug.WriteLine("Removed User: " + _ADUser.Name + " from: " + group.Name)
            End If

            Return True
        Catch ex As Exception
            'MessageBox.Show("Error Removing User: " + userName + " from group: " + groupName + vbNewLine + vbNewLine + ex.ToString)
            Return False
        End Try

    End Function

    Function addUserToGroup(userName As String, groupName As String) As Boolean

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, userName)
            Debug.WriteLine("Got our contect and user.")

            Dim group As GroupPrincipal = GroupPrincipal.FindByIdentity(_ADContext, IdentityType.Name, groupName)

            If Not group.Members.Contains(_ADUser) Then
                Debug.WriteLine("Group " + groupName + " does not contain user: " + userName)
                group.Members.Add(_ADUser)
                group.Save()
                Debug.WriteLine("Added User: " + _ADUser.Name + " to: " + groupName)
            Else
                Debug.WriteLine("User already exists in group.")
            End If

            Return True
        Catch ex As Exception
            'MessageBox.Show("Error Adding User: " + userName + " to group: " + groupName + vbNewLine + vbNewLine + ex.ToString)
            Return False
        End Try

        Return True

    End Function

    Function setADPassword(username As String, password As String, emailIncluded As Boolean, emailAddress As String) As Boolean

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, username)

            _ADUser.SetPassword(password)
            _ADUser.Save()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Function getADUser(username As String) As UserPrincipal

        Try
            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, username)

            Return _ADUser
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Function getUsersInGroup(groupName As String) As GroupPrincipal

        Try
            Dim Group As GroupPrincipal = GroupPrincipal.FindByIdentity(_ADContext, groupName)

            Return Group
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Function newPortalADUser(username As String, password As String, firstname As String, lastname As String, customerEmail As String, groupList As List(Of String)) As Boolean

        Try

            _ADUser = UserPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, username)

            If _ADUser Is Nothing Then
                System.Diagnostics.Debug.WriteLine("No user found, good to go")

                _ADUser = New UserPrincipal(_ADContext, username, password, True)

                With (_ADUser)
                    .Name = username
                    .DisplayName = firstname + " " + lastname
                    .UserPrincipalName = username & "@company.com"
                    .Enabled = True
                    .PasswordNotRequired = False
                    .PasswordNeverExpires = False
                    .UserCannotChangePassword = False
                    .GivenName = Nothing
                    .MiddleName = Nothing
                    .Surname = Nothing
                    .EmailAddress = customerEmail
                    .SetPassword(password)
                End With
                _ADUser.Save()

                '(Re)Set account lockout as needed
                If _ADUser.IsAccountLockedOut Then
                    _ADUser.UnlockAccount()
                    _ADUser.RefreshExpiredPassword()
                    _ADUser.Save()
                End If

                Dim usergroups As List(Of String) = GetUserGroups() 'Get AD Groups

                'Set user groups as needed
                For Each groupName As String In usergroups
                    If groupList.Contains(groupName) Then

                        _ADGroup = GroupPrincipal.FindByIdentity(_ADContext, IdentityType.SamAccountName, groupName)
                        If Not _ADGroup.GetMembers.Contains(_ADUser) Then
                            _ADGroup.Members.Add(_ADUser) : _ADGroup.Save()
                        End If

                    End If
                Next groupName

                ' Validate login credentials
                If _ADContext.ValidateCredentials(username, password) Then
                    Return True
                Else
                    'MessageBox.Show("Error during AD user creation process.")
                    Return False
                End If
            Else
                System.Diagnostics.Debug.WriteLine("Existing user found with Username. Exiting.")
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function GetUserGroups() As List(Of String)

        Dim UserGroups As New List(Of String)
        Try
            Using _ADGroups As New GroupPrincipal(_ADContext)
                _ADSearcher = New PrincipalSearcher(_ADGroups)
                For Each group As GroupPrincipal In _ADSearcher.FindAll()
                    UserGroups.Add(group.SamAccountName)
                Next group
            End Using

            UserGroups.Sort()

            Return UserGroups
        Catch ex As Exception
            UserGroups = New List(Of String)
            Return UserGroups
        End Try

    End Function

    Public Function getAccountsNotLoggedInForXDays(numOfDays As Integer, getAllAccounts As Boolean) As Dictionary(Of String, Date)

        Try
            Dim returnDictionary As New Dictionary(Of String, Date)

            Dim _ADUsers As New UserPrincipal(_ADContext)
            _ADSearcher = New PrincipalSearcher(_ADUsers)
            For Each userObj As UserPrincipal In _ADSearcher.FindAll()

                If Not userObj.LastLogon Is Nothing Then

                    If getAllAccounts = False Then
                        If DateTime.Today.Subtract(DateTime.Parse(userObj.LastLogon.ToString)).Days >= numOfDays Then

                            returnDictionary.Add(userObj.SamAccountName, DateTime.Parse(userObj.LastLogon.ToString))
                        End If
                    Else
                        returnDictionary.Add(userObj.SamAccountName, DateTime.Parse(userObj.LastLogon.ToString))
                    End If

                End If
            Next

            Return returnDictionary
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw New Exception("There was an error gathering accounts not logged in to for " + numOfDays.ToString + " days." + vbNewLine + ex.ToString)
        End Try

    End Function

    Public Function getPasswordExpirationInfo(onlyExpiredAccounts As Boolean) As Dictionary(Of String, Date)

        Try
            Dim listOfExpiredPasswordUsers As New Dictionary(Of String, Date)

            'Have to do a LDAP lookup at the domain level to get the max password age in days
            Dim maxPwdAge As Integer = 0

            Dim ent As DirectoryEntry = New DirectoryEntry("LDAP://" + "DC=subdomain,DC=company,DC=com")
            ent.AuthenticationType = AuthenticationTypes.Secure

            Dim ds As DirectorySearcher = New DirectorySearcher(ent)
            Dim filter As String = "(maxPwdAge=*)"
            ds.Filter = filter

            Dim results As SearchResult = ds.FindOne()
            If Not results Is Nothing Then
                Dim pwdAge As String = results.Properties("maxPwdAge")(0).ToString

                Dim convertedInt64 As Int64 = Int64.Parse(pwdAge)
                Dim tempMaxPwdAge As Double = convertedInt64 / -864000000000

                maxPwdAge = CInt(tempMaxPwdAge)
            End If

            'Now that we have the max password age in days, we can check it against last password set attribute to find out if the password is expired
            _ADUser = New UserPrincipal(_ADContext)
            _ADSearcher = New PrincipalSearcher(_ADUser)

            For Each userObj As UserPrincipal In _ADSearcher.FindAll()
                If Not userObj.LastPasswordSet Is Nothing Then
                    Dim passwordExpireDate As Date = CDate(userObj.LastPasswordSet.ToString).AddDays(maxPwdAge)

                    If onlyExpiredAccounts = True Then
                        If DateTime.Today >= passwordExpireDate Then
                            listOfExpiredPasswordUsers.Add(userObj.SamAccountName, passwordExpireDate)
                        End If
                    Else
                        listOfExpiredPasswordUsers.Add(userObj.SamAccountName, passwordExpireDate)
                    End If

                End If
            Next

            Return listOfExpiredPasswordUsers
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw New Exception("There was an error gathering users with expired passwords" + vbNewLine + ex.ToString)
        End Try

    End Function


End Class
