# Connect-ExchangeOnline latest versions (>3.7.x) returns "A window handle must be configured" error when launching from Powershell ISE


Here is the complete error:

```output
PS C:\WINDOWS\system32> Connect-ExchangeOnline -UserPrincipalName $Login
Error Acquiring Token:
A window handle must be configured. See https://aka.ms/msal-net-wam#parent-window-handles
A window handle must be configured. See https://aka.ms/msal-net-wam#parent-window-handles
At C:\Program 
Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\3.9.0\netFramework\ExchangeOnlineManagement.psm1:766 char:21
+                     throw $_.Exception.InnerException;
+                     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : OperationStopped: (:) [], MsalClientException
    + FullyQualifiedErrorId : A window handle must be configured. See https://aka.ms/msal-net-wam#parent-window-handles
```

To solve this I used the information from [David Homer's website](https://david-homer.blogspot.com/2025/01/exchange-online-management-powershell.html) :

```powershell
# Login string
$Login = "your_User_Principal_Name@Your_Domain.ext"

# Routine from https://david-homer.blogspot.com/2025/01/exchange-online-management-powershell.html
$msalPath = [System.IO.Path]::GetDirectoryName((Get-Module ExchangeOnlineManagement).Path);
Add-Type -Path "$msalPath\Microsoft.IdentityModel.Abstractions.dll";
Add-Type -Path "$msalPath\Microsoft.Identity.Client.dll";
[Microsoft.Identity.Client.IPublicClientApplication] $application = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create("fb78d390-0c51-40cd-8e17-fdbfab77341b").WithDefaultRedirectUri().Build();
$result = $application.AcquireTokenInteractive([string[]]"https://outlook.office365.com/.default").ExecuteAsync().Result;
Connect-ExchangeOnline -AccessToken $result.AccessToken -UserPrincipalName $login;
```


