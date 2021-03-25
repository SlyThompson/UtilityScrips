param
(
[Parameter (Mandatory = $true)]
[string]$User,
[Parameter (Mandatory = $true)]
[string]$Password,
[Parameter (Mandatory = $true)]
[string]$ExchangeServer,
[Parameter (Mandatory = $false)]
[string]$NumberOfDays =30
)

$LogFile = "ExportSharedMailboxLoginsLog";
$OutputFile ="ExportSharedMailboxLogins";
##usage:

##To export all mailboxes who have not logged on since last 30 days [30 is the default period here in script]
##.\ExportSharedMailboxLogins.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10

##To export all mailboxes who have not logged on since last 60 days 
##.\ExportSharedMailboxLogins.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10 -NumberOfDays 60


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

#$AllDLsOutputFolder = "$OutputFolder\AllDLs";

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

$MailboxesDic =@{};

if($Error.Count -eq 0 )
{
    LogProgress "`nGetting shared mailbox list";

    $time = (get-date).AddDays(-$NumberOfDays);    

    $Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails Shared -ErrorAction:SilentlyContinue 
    
    if($Error.Count -eq 0 )
    {
        foreach($mbx in $Mailboxes)
        {
            $MailboxesDic.Add($mbx.ExchangeGuid, $mbx.PrimarySmtpAddress);
        }

        LogProgress "`nGetting user logons";    
       
        $Error.Clear();

        $Stats = $Mailboxes| Get-mailboxStatistics | ? {$_.LastLogonTime -lt $time}|Select DisplayName, LastLogonTime,ItemCount, MailboxGuid ; 
        
        if($Error.Count -eq 0)
        {     
            $count = 1;
            if($Stats.Count -ne $null)
            {
                $count = $Stats.Count;
            }

            LogSuccess "`nFound ($Count) mailbox stats`n";
            if($Count -gt 0)
            {         
                $Stats | Select DisplayName,@{Label = "EmailAddress";Expression = { $MailboxesDic[$_.MailboxGuid]}} ,LastLogonTime, ItemCount |Export-Csv $OutputFile -NoTypeInformation;
            }
        }    
        else
        {
            LogError "$Error";
        }
    }
    else
    {
        LogError "$Error";
    }
}
else
{
    LogError "$Error";
}


LogInfo "Script ended at [$(Get-Date)]";
##Cleannig session
#RemovePSSession