<#
.SYNOPSIS
    This script is used for hybrid scenario when a list of different modules are required to connect to O365.
    This script will verify if the modules are installed and connected you to the right online modalities per choice.
    The credential are taken from the Credential Manager after so there is no need to enter credential every time.

.PREREQUISITIES
    Make sure you create a Credential Reposityry Generic Credentials in order to include it under the $TeanantCredentialKey
    parameter (the name of the credentials).

.NOTES  
    Version                   : 0.1
    Rights Required           : Local Admin, WinRM Service Running, Modules Installed
    Authors                   : Guy Bachar
    Last Update               : 20-May-2015
    Twitter/Blog              : @GuyBachar, http://guybachar.wordpress.com

.REFRENCES
    https://support.microsoft.com/en-us/kb/2955287
    https://technet.microsoft.com/en-us/library/jj984289%28v=exchg.150%29.aspx
    https://technet.microsoft.com/en-us/library/dn568015.aspx
    https://technet.microsoft.com/en-us/library/dn362795%28v=ocs.15%29.aspx?f=255&MSPPError=-2147217396


.VERSION
    0.1 - Initial Version for connecting Online resources

#>

PARAM
(
    [Parameter(Mandatory=$True) ][ValidateNotNullorEmpty()][string] $TenantCredentialKey,
    [Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $Office365Online,    
    [Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $ExchangeOnline,
    [Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $ExchangeOnlineWithProxy,
    [Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $LyneOnline,
    [Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $CheckVersions,
    [Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $Office365withEXO
    #[Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $AzureOnline   
)

######################################################################
# Common functions
######################################################################
#
# Run a cmdlet silently without throwing error
#
function RunCmdletSilently(
    [string] $cmdlet,
    [object] $parameters,
    [object] $pipelineObjects = $null,
    [ref] $output = [ref]$null)
{
    $parameters["ErrorAction"] = "Continue"
    $Global:Error.Clear()
    
    $result = $null
    try
    {
        if ($null -ne $pipelineObjects)
        {
            $result = ($pipelineObjects | &$cmdlet @parameters)
        }
        else
        {
            $result = (&$cmdlet @parameters)
        }
    
        if ($Global:Error.Count -eq 0)
        {
            $output.Value = $result
            return $true
        }
    }
    catch
    {
    }

    return $false
}

#
# Office 365 functions
#
function ConnecttoExchangeOnline(
    [System.Management.Automation.PSCredential] $credential)
{
    $session = $null
    
    $parameters = @{
        "ConfigurationName" = "Microsoft.Exchange";
        "ConnectionURI" = $ExchangeOnlineUrl
        "AllowRedirection" = $true;
        "SessionOption" = (New-PSSessionOption -SkipCACheck -SkipCNCheck);
        "Credential" = $credential;
        "Authentication" = "Basic";
    }
        
    if (RunCmdletSilently "New-PSSession" $parameters $null ([ref]$session))
    {
        return $session
    }

    return $null
}

#
# Verifying Azure Module existence
#
function Get-WindowsAzurePowerShellVersion
{
[CmdletBinding()]
Param ()
 
## - Section to query local system for Windows Azure PowerShell version already installed:
$AzurePowerShellExists = (Get-Module -ListAvailable | Where-Object{ $_.Name -eq 'Azure' }) `
| Select Version, Name, Author | Format-Table -AutoSize;

if ($AzurePowerShellExists -ne $null) {
Write-Host "`r`nWindows Azure PowerShell Installed version: " -ForegroundColor 'Yellow';
$AzurePowerShellExists}
else {Write-Host "`r`nWindows Azure PowerShell module is not Installed" -ForegroundColor 'Red';}

 
## - Section to query web Platform installer for the latest available Windows Azure PowerShell version:
Write-Host "Windows Azure PowerShell available download version (http://azure.microsoft.com/en-us/downloads): " -ForegroundColor 'Green';
[reflection.assembly]::LoadWithPartialName("Microsoft.Web.PlatformInstaller") | Out-Null;
$ProductManager = New-Object Microsoft.Web.PlatformInstaller.ProductManager;
$ProductManager.Load(); $ProductManager.Products `
| Where-object{
($_.Title -match "Microsoft Azure PowerShell") `
-and ($_.Author -eq 'Microsoft Corporation')
} `
| Select-Object Version, Title, Published, Author | Format-Table -AutoSize;
};

#
# Verifying Lync Online Module existence
#
function Get-LyncOnlinePowerShellVersion
{
[CmdletBinding()]
Param ()
 
## - Section to query local system for Windows Azure PowerShell version already installed:
$LyncPowerShellExists = (Get-Module -ListAvailable | Where-Object{ $_.Name -eq 'LyncOnlineConnector' }) `
| Select Version, Name, Author | Format-Table -AutoSize;

if ($LyncPowerShellExists -ne $null) {
Write-Host "`r`nLync Online PowerShell Installed version: " -ForegroundColor 'Yellow';
$LyncPowerShellExists
Import-Module LyncOnlineConnector}
else {Write-Host "`r`nLync Online PowerShell module is not Installed (https://www.microsoft.com/en-us/download/details.aspx?id=39366) " -ForegroundColor 'Red';}

};


######################################################################
# Parameters
######################################################################
#
# Downloads URL's
#
$DownloadAzureURL = "http://azure.microsoft.com/en-us/downloads/"
$DownloadLyncURL = "https://www.microsoft.com/en-us/download/details.aspx?id=39366"

#
# Office 365 to connect with
#
#$ExchangeOnlineUrl = "https://pilot.outlook.com/PowerShell-LiveID/"
$ExchangeOnlineUrl = "https://outlook.office365.com/powershell-liveid/"



######################################################################
# Script Start
######################################################################

#
# Verifying Administrator Elevation
#
#region Verifying Administrator Elevation
Write-Host ""
Write-Host "`nVERIFYING USER PERMISSIONS" -BackgroundColor Green -ForegroundColor Black
Start-Sleep -Seconds 2
#Verify if the Script is running under Admin privileges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
  [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
  Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!"
  Write-Host 
  Break
}
#endregion


if ($CheckVersions.IsPresent)
{
Write-Host "VERIFYING MODULES INSTALLATIONS" -BackgroundColor Green -ForegroundColor Black
Get-WindowsAzurePowerShellVersion
Get-LyncOnlinePowerShellVersion
}

######################################################################
# Verifying Credentials existence
######################################################################
Write-Host "`nLOADING CREDENTIALS FROM CREDENTIAL MANAGER" -BackgroundColor Green -ForegroundColor Black


######################################################################
# API to load credential from generic credential store
######################################################################
$CredManager = @"
using System;
using System.Net;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace SyncSiteMailbox
{
    /// <summary>
    /// </summary>
    public class CredManager
    {
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CredReadW")]
        public static extern bool CredRead([MarshalAs(UnmanagedType.LPWStr)] string target, [MarshalAs(UnmanagedType.I4)] CRED_TYPE type, UInt32 flags, [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(CredentialMarshaler))] out Credential cred);
        
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto, EntryPoint = "CredFree")]
        public static extern void CredFree(IntPtr buffer);

        /// <summary>
        /// </summary>
        public enum CRED_TYPE : uint
        {
            /// <summary>
            /// </summary>
            CRED_TYPE_GENERIC = 1,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_PASSWORD = 2,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_CERTIFICATE = 3,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_VISIBLE_PASSWORD = 4,

            /// <summary>
            /// </summary>
            CRED_TYPE_MAXIMUM = 5, // Maximum supported cred type
        }
        
        /// <summary>
        /// </summary>
        public enum CRED_PERSIST : uint
        {
            /// <summary>
            /// </summary>
            CRED_PERSIST_SESSION = 1,

            /// <summary>
            /// </summary>
            CRED_PERSIST_LOCAL_MACHINE = 2,

            /// <summary>
            /// </summary>
            CRED_PERSIST_ENTERPRISE = 3
        }
        
        /// <summary>
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct CREDENTIAL
        {
            internal UInt32 flags;
            internal CRED_TYPE type;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string targetName;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string comment;
            internal System.Runtime.InteropServices.ComTypes.FILETIME lastWritten;
            internal UInt32 credentialBlobSize;
            internal IntPtr credentialBlob;
            internal CRED_PERSIST persist;
            internal UInt32 attributeCount;
            internal IntPtr credAttribute;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string targetAlias;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string userName;
        }
        
        /// <summary>
        /// Credential
        /// </summary>
        public class Credential
        {
            private SecureString secureString = null;

            /// <summary>
            /// </summary>
            internal Credential(CREDENTIAL cred)
            {
                this.credential = cred;
                unsafe
                {
                    this.secureString = new SecureString((char*)this.credential.credentialBlob.ToPointer(), (int)this.credential.credentialBlobSize / sizeof(char));
                }                
            }

            /// <summary>
            /// </summary>
            public string UserName
            {
                get { return this.credential.userName; }
            }

            /// <summary>
            /// </summary>
            public SecureString Password
            {
                get
                {
                    return this.secureString;
                }
            }

            /// <summary>
            /// </summary>
            internal CREDENTIAL Struct
            {
                get { return this.credential; }
            }

            private CREDENTIAL credential;
        }

        internal class CredentialMarshaler : ICustomMarshaler
        {
            public void CleanUpManagedData(object ManagedObj)
            {
                // Nothing to do since all data can be garbage collected.
            }

            public void CleanUpNativeData(IntPtr pNativeData)
            {
                if (pNativeData == IntPtr.Zero)
                {
                    return;
                }
                CredFree(pNativeData);
            }

            public int GetNativeDataSize()
            {
                return Marshal.SizeOf(typeof(CREDENTIAL));
            }

            public IntPtr MarshalManagedToNative(object obj)
            {
                Credential cred = (Credential)obj;
                if (cred == null)
                {
                    return IntPtr.Zero;
                }

                IntPtr nativeData = Marshal.AllocCoTaskMem(this.GetNativeDataSize());
                Marshal.StructureToPtr(cred.Struct, nativeData, false);

                return nativeData;
            }

            public object MarshalNativeToManaged(IntPtr pNativeData)
            {
                if (pNativeData == IntPtr.Zero)
                {
                    return null;
                }
                CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(pNativeData, typeof(CREDENTIAL));
                return new Credential(cred);
            }

            public static ICustomMarshaler GetInstance(string cookie)
            {
                return new CredentialMarshaler();
            }
        }    
        

        /// <summary>
        /// ReadCredential
        /// </summary>
        /// <param name="credentialKey"></param>
        /// <returns></returns>
        public static NetworkCredential ReadCredential(string credentialKey)
        {
            Credential credential;
            CredRead(credentialKey, CRED_TYPE.CRED_TYPE_GENERIC, 0, out credential);
            return credential != null ? new NetworkCredential(credential.UserName, credential.Password) : null;
        }
    }
}
"@

######################################################################
# Load credential APIs
######################################################################
$CredManagerType = $null
try
{
    $CredManagerType = [SyncSiteMailbox.CredManager]
}
catch [Exception]
{
}

if($null -eq $CredManagerType)
{
    $compilerParameters = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters
    $compilerParameters.CompilerOptions = "/unsafe"
    [void]$compilerParameters.ReferencedAssemblies.Add("System.dll")
    Add-Type $CredManager -CompilerParameters $compilerParameters
    $CredManagerType = [SyncSiteMailbox.CredManager]
}

######################################################################
# Load tenant credential from generic credential store
######################################################################
$TenantCredential = $null #Primary

#Write-Host "Load tenant credential is from generic credential store."
try
{
    $credential = $CredManagerType::ReadCredential($TenantCredentialKey)
    if ($null -ne $credential)
    {
        $TenantCredential = New-Object System.Management.Automation.PSCredential ($credential.UserName, $credential.SecurePassword);
    }
}
catch [Exception]
{
    $TenantCredential = $null
    $errorMessage = $_.Exception.Message
    Write-Host "Tenant credential cannot be loaded correctly: $errorMessage."
}

if ($null -eq $TenantCredential)
{
    Write-Host "Tenant credential cannot be loaded please ensure you have configured in credential manager correctly."
}



######################################################################
# Connect to MSOnline with MSO prefix
######################################################################
if ($Office365Online.IsPresent)
    {
    Write-Host "`nCONNECTING TO OFFICE 365 ONLINE" -BackgroundColor Green -ForegroundColor Black
    $params = @{
        "Name" = "MSOnline";
        "Prefix" = "MSO"
    }

    if ( -not (RunCmdletSilently "Import-Module" $params))
    {
        Write-Host "Microsoft Online Service Module is not installed, please install it from http://onlinehelp.microsoft.com/en-us/office365-enterprises/ff652560.aspx."
    }

    $params = @{
        "Credential" = $TenantCredential
    }

    if( -not (RunCmdletSilently "Connect-MSOMsolService" $params))
    {
        Write-Host "Microsoft Online Service cannot be connected, please ensure $($TenantCredential.UserName) has read only permission."
    }
}


######################################################################
# Connect to Office 365 with EXO prefix
######################################################################
if ($Office365withEXO.IsPresent)
{
        Write-Host "`nCONNECTING TO OFFICE 365 WITH EXCHANGE ONLINE" -BackgroundColor Green -ForegroundColor Black
        $EXOSession = ConnecttoExchangeOnline $TenantCredential
        if ($null -eq $EXOSession)
            {
        Write-Host "Office 365 cannot be connected by $($TenantCredential.UserName) because of $Error."
        }

        try
        {
            $params = @{
                "Session" = $EXOSession;
                "Prefix" = "EXO";
                "AllowClobber" = $true
            }

    
            if ( -not (RunCmdletSilently "Import-PSSession" $params))
            {
                Write-Host "Office 365 session cannot be imported to current PowerShell window."
            }
    
            Write-Host "Office 365 is connected successfully by $($TenantCredential.UserName)."
        }


        finally
        {
            ######################################################################
            # Disconnect Office 365 session
            ######################################################################
            $params = @{
                "Session" = $EXOSession
            }
    
            if ( -not (RunCmdletSilently "Remove-PSSession" $params))
            {
                Write-Host "Office 365 cannot be disconnected because of $Error." $false
            }
        }
}

######################################################################
# Connect to Exchange Online
######################################################################
if ($ExchangeOnline.IsPresent)
{
    Write-Host "`nCONNECTING TO EXCHANGE ONLINE" -BackgroundColor Green -ForegroundColor Black

    #Verify if Proxy is needed
    if ($ExchangeOnlineWithProxy.IsPresent){
        Write-Output "Using IE Proxy Settings"
        $ProxySettings = New-PSSessionOption -ProxyAccessType IEConfig
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $TenantCredential -Authentication Basic -AllowRedirection -SessionOption $ProxySettings
        }
    else{
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $TenantCredential -Authentication Basic -AllowRedirection
        }
        Import-PSSession $Session -AllowClobber
}

######################################################################
# Connect to Lync Online
######################################################################
if ($LyneOnline.IsPresent)
{
    Write-Host "`nCONNECTING TO LYNC ONLINE" -BackgroundColor Green -ForegroundColor Black
    $DefaultDomain = Get-MsolDomain | Where {$_.IsDefault -eq $True}
    $CSSession = New-CsOnlineSession -Credential $TenantCredential -OverrideAdminDomain $DefaultDomain.Name
    Import-PSSession $CSSession -AllowClobber 
}