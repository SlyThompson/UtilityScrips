param
(
[Parameter(Mandatory = $true)]
[string]$User,
[Parameter(Mandatory = $true)]
[string]$Password,
[Parameter(Mandatory = $true)]
[string]$ExchangeServer,
[Parameter(Mandatory = $true)]
[string]$Mailbox,
[Parameter(Mandatory = $true)]
[string]$Received,
[Parameter(Mandatory = $true)]
[string]$Sent

)

$LogFile =""; #Log file name would be ScriptNameLog($time).log
$OutputFile ="";  #Output CSV file name would be ScriptName($time).csv

##usage:
##.\ExportPFPermissions.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13

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
    $UserCredential= New-Object System.Management.Automation.PSCredential($User, $EncodedPwd);
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

##Taking script name from PS environment
$scriptName = $MyInvocation.MyCommand.Name
$scriptName  =$scriptName.Replace(".ps1","");

$scriptPath = GetScriptPath;
$sTime = get-date;
$timeStr = $sTime.ToString("dd-MM-yyyy-hh");

$LogFolder = "$scriptPath\Logs";
$LogFile   = "$LogFolder\$($scriptName)Log-$timeStr.log";

$OutputFolder = "$scriptPath\Output";
$OutputFile = "$OutputFolder\$scriptName-$timeStr.csv";


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

$PFsToExport = @{};

if($Error.Count -eq 0 )
{
    LogProgress "`nVerifying mailbox [$Mailbox]";

	$mailboxObject = Get-Mailbox $Mailbox -ErrorAction:SilentlyContinue;
    
	$Error.Clear();

    $Mailboxes = @{};
    if($mailboxObject)
    {    
        LogSuccess "`nFound mailbox $($mailboxObject.DisplayName)";

        LogProgress "`nSearch mailbox [$Mailbox] for Received [$Received] and Sent [$Sent]";

        $mailboxObject | Search-mailbox -SearchQuery {(Received -lt $Received) -or (Sent -lt $Sent)} -DeleteContent -Force

        if($Error.Count -eq 0)
        {
            LogSuccess "Command executed successfully.";
        }
        else
        {
            LogError "`nError. $Error`n";
        }
        
    }
    else
    {
        LogError "`nCould not find $mailbox. $Error`n";
    }
}


LogProgress "`nScript ended at [$(get-date)]`n";
##Cleannig session
#RemovePSSession