param
(
[Parameter (Mandatory = $true)]
[string]$User,
[Parameter (Mandatory = $true)]
[string]$Password,
[Parameter (Mandatory = $true)]
[string]$ExchangeServer
)

$LogFile = "ExportAliasDLsAllLog";
$OutputFile ="ExportAliasDLsAll";
##usage:
##.\ExportAliasDLsAll.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10

#Enable-ADOptionalFeature 'Recycle Bin Feature' -scope ForestOrConfigurationSet -Target 'ad.lab' -Verbose
# Get-ADObject -filter {isdeleted -eq $true -and name -ne 'Deleted Objects'} -includeDeletedObjects -property *
# $o= Get-ADObject -filter {isdeleted -eq $true -and name -ne 'Deleted Objects' -and objectclass -eq 'group'} -includeDeletedObjects -property *
#By default, if an object has been deleted, it can be recovered within a 180 days interval. 
#This value is specified in the msDS-DeletedObjectLifetime attribute. However, if you want to change this value, you can use the following command:


#region Functions

function LogInfo($msg)
{
    Write-Host $msg -f Yellow
    LogToFile($msg);
}
function LogProgress($msg)
{
    Write-Host $msg -f Cyan
    LogToFile($msg);
}
function LogError($msg)
{
    Write-Host $msg -f Red
    LogToFile("Error: $msg");
}
function LogSuccess($msg)
{
    Write-Host $msg -f Green
    LogToFile($msg);
}

function LogToFile($msg)
{
    $time =get-date;
    "[$time] $msg" | Out-File -FilePath $LogFile -Append;
}
function ConnectToExchange
{
    $ExchangeSession = Get-PSSession |?{$_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.State -eq 'Opened'};
    if($ExchangeSession)
    {
        ##If an existing session already exists in PowerShell console, use that one, instead of creating a new one;
        LogInfo "`nUsing existing exchange powershell session connected to computer [$($ExchangeSession.ComputerName)]`n";
        return;
    }

    RemovePSSession

    LogInfo "Connecing to Exchange server using [User: $User] [Password: $Password] and [Exchange Server: $ExchangeServer]"

    $EncodedPwd = ConvertTo-SecureString $Password -AsPlainText -Force;
    $UserCredential= New-Object PSCredential($User, $EncodedPwd);
    $PSUrl = "https://$ExchangeServer/PowerShell/";

    $PSSessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -MaximumRedirection 5;
    $Error.Clear();
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PSUrl -Authentication Basic -Credential $UserCredential -SessionOption $PSSessionOption -ErrorAction:SilentlyContinue;

    if($Error.Count -eq 0)
    {
        LogInfo "`nImporting exchange powershell session";
        Import-PSSession $Session -DisableNameChecking
    }
    else
    {
        LogError "Could not load exchange powershell session. $Error";
    }
}

function RemovePSSession
{
    LogInfo "`nRemoving PS Sessions";
    Get-PSSession | Remove-PSSession;
}
function GetScriptPath
{     Split-Path $myInvocation.ScriptName 
}

#endregion Functions End

$scriptPath = GetScriptPath;
$sTime = get-date;
$timeStr = $sTime.ToString("dd-MM-yyyy-hh");

$LogFolder = "$scriptPath\Logs";
$LogFile   = "$LogFolder\$LogFile-$timeStr.log";

$OutputFolder = "$scriptPath\Output";
$OutputFile = "$OutputFolder\$OutputFile-$timeStr.csv";

$AllDLsOutputFolder = "$OutputFolder\AllDLs";

##\\Logs folder
if(-not (Test-Path $LogFolder))
{
    New-item $LogFolder -ItemType Directory | Out-Null
}

##\\Output folder
if(-not (Test-Path $OutputFolder))
{
    New-item $OutputFolder -ItemType Directory | Out-Null 
}

##\\Output\AllDLs folder
if(-not (Test-Path $AllDLsOutputFolder))
{
    New-item $AllDLsOutputFolder -ItemType Directory | Out-Null 
}


if(-not (Test-Path $LogFile))
{
    New-item $LogFile -ItemType File | Out-Null
}


$Error.Clear();

$w = (Get-Host).UI.RawUI.WindowSize.Width/2;
$top = "$("*" * $w)"
$bot = "$("*" * $w)"


LogProgress $top
LogProgress "Script started at [$sTime ]";
LogProgress $bot

ConnectToExchange;

$DLsToExport = @();

if($Error.Count -eq 0 )
{
    LogProgress "`nGetting all groups";

    $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
    $Error.Clear();
    if($Error.Count -eq 0 )
    {
        if($DLs  -ne $null -and $DLs.Count -gt 0)
        {    
            LogSuccess "`nFound $($DLs.Count) groups`n";

            if($DLs.Count -gt 0)
            {            
                $DLs |Select Name, Alias,PrimarySmtpAddress| Export-CSV $OutputFile -NoTypeInformation;    

                LogSuccess "`nExported $($DLs.Count) records to [$OutputFile]`n";         
            }
            else
            {
                LogSuccess "`nNo DL found`n";
            }
        }
        else
        {
            LogError "`nRetrieved 0 groups.`n";
        }
    }
    else
    {
        LogError "`nCould not get groups. $Error`n";
    }
}


##Cleannig session
#RemovePSSession