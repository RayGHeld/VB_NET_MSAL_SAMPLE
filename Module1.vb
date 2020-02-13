Imports Microsoft.Identity.Client

Module Module1

    Private _accessToken As String = String.Empty
    Private Const client_id As String = "28a00d08-ba05-4015-a269-9d0546e850b9" '<-- enter the client_id guid here
    Private Const tenant_id As String = "organizations" '<-- enter either your tenant id here
    Private authority As String = $"https://login.microsoftonline.com/{tenant_id}"
    Private scopes As New List(Of String)

    Sub Main()

        scopes.Add($"{client_id}/.default")

        'Console.WriteLine("Starting Synchronous Sample...")

        'SyncSample()

        'Console.WriteLine($"{Environment.NewLine}End Synchronous Sample...{Environment.NewLine}Start Asynchronous Sample...")


        'Console.WriteLine($"{Environment.NewLine}End Asynchronous Sample.{Environment.NewLine}Press any key to close...")
        Dim userInput As String = String.Empty
        Do Until userInput.ToUpper.Trim() = "EXIT"
            Console.Clear()
            AsyncSample()
            Console.WriteLine($"{Environment.NewLine}Type Exit and press Enter to quit, or just press Enter to try again...")
            userInput = Console.ReadLine()
        Loop


    End Sub

#Region "Synchronous Code"


    Private Sub SyncSample()
        If Login() Then
            Console.WriteLine(_accessToken)
        End If
    End Sub

    Private Function Login() As Boolean
        _accessToken = String.Empty

        Dim publicClientApp As IPublicClientApplication
        publicClientApp = PublicClientApplicationBuilder.Create(client_id).WithAuthority(authority).Build()

        Dim accounts As IEnumerable(Of IAccount) = publicClientApp.GetAccountsAsync().Result()
        Dim firstAccount As IAccount = accounts.FirstOrDefault()
        Dim authResult As AuthenticationResult

        Try
            authResult = publicClientApp.AcquireTokenSilent(scopes, firstAccount).ExecuteAsync().Result()
            _accessToken = authResult.AccessToken
        Catch e As MsalUiRequiredException
            Try
                authResult = publicClientApp.AcquireTokenInteractive(scopes).ExecuteAsync().Result()
                _accessToken = authResult.AccessToken
            Catch ex As Exception
                Console.WriteLine($"Auth Exception: {ex.Message}")
            End Try
        Catch ex As Exception
            Console.WriteLine($"Auth Exception: {ex.Message}")
        End Try


        Return _accessToken <> String.Empty

    End Function

#End Region

#Region "Asynchronous Code"

    Private Sub AsyncSample()

        Dim task As Task(Of Boolean) = LoginTask()

        If task.Result() Then
            Console.WriteLine(_accessToken)
        End If
    End Sub

    Private Async Function LoginTask() As Task(Of Boolean)
        _accessToken = String.Empty

        Dim publicClientApp As IPublicClientApplication
        publicClientApp = PublicClientApplicationBuilder.Create(client_id).WithAuthority(authority).Build()

        Dim accounts As IEnumerable(Of IAccount) = Await publicClientApp.GetAccountsAsync()
        Dim firstAccount As IAccount = accounts.FirstOrDefault()
        Dim authResult As AuthenticationResult

        Dim tryInteractive As Boolean = False

        Try
            authResult = Await publicClientApp.AcquireTokenSilent(scopes, firstAccount).ExecuteAsync()
            _accessToken = authResult.AccessToken
        Catch e As MsalUiRequiredException
            tryInteractive = True
        End Try

        If tryInteractive Then
            Try
                authResult = Await publicClientApp.AcquireTokenInteractive(scopes).ExecuteAsync()
                _accessToken = authResult.AccessToken
            Catch ex As Exception
                Console.WriteLine($"Auth Exception: {ex.Message}")
            End Try
        End If

        ' Console.WriteLine($"Id Token: {authResult.IdToken}")

        Return _accessToken <> String.Empty

    End Function

#End Region

End Module