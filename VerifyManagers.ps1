param
(
[Parameter (Mandatory = $true)]
[string]$User,
[Parameter (Mandatory = $true)]
[string]$Password,
[Parameter (Mandatory = $true)]
[string]$ExchangeServer
)

$LogFile = "VerifyManagersLog";
$OutputFile ="DLsWithNoManager";
##usage:
##.\VerifyManagers.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10

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

$scriptPath = GetScriptPath;
$sTime = get-date;
$timeStr = $sTime.ToString("dd-MM-yyyy-hh");

$LogFolder = "$scriptPath\Logs";
$LogFile   = "$LogFolder\$LogFile-$timeStr.log";

$OutputFolder = "$scriptPath\Output";
$OutputFile = "$OutputFolder\$OutputFile-$timeStr.csv";

if(-not (Test-Path $LogFolder))
{
    New-item $LogFolder -ItemType Directory; 
}

if(-not (Test-Path $OutputFolder))
{
    New-item $OutputFolder -ItemType Directory; 
}


if(-not (Test-Path $LogFile))
{
    New-item $LogFile -ItemType File; 
}

#endregion


$Error.Clear();

$w = (Get-Host).UI.RawUI.WindowSize.Width/2;
$top = "$("*" * $w)"
$bot = "$("*" * $w)"


LogProgress $top
LogProgress "Script started at [$sTime ]";
LogProgress $bot

ConnectToExchange;

$DLsWithNoManagers = @();
if($Error.Count -eq 0 )
{
    LogProgress "`nGetting all groups";

    $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
    $Error.Clear();

    if($DLs  -ne $null -and $DLs.Count -gt 0)
    {    
        LogSuccess "`nFound $($DLs.Count) groups`n";

        ##Iterating over $csv rows
        $count =0;
        foreach($DL  in $DLs)
        {
            $count++;    

            $Identity = $DL.Alias;
            LogProgress "$count# Processing group [$Identity ($($DL.PrimarySmtpAddress))]";               
                
            if($DL.ManagedBy -eq $null)
            {
                 LogInfo "Manager is not set for group [$Identity]";
                 $DLsWithNoManagers+= $DL;
            }      
        }

        if($DLsWithNoManagers.Count -gt 0)
        {
            $DLsWithNoManagers |Select Name, DisplayName, PrimarySmtpAddress, DistinguishedName| Export-CSV $OutputFile -NoTypeInformation;    
            
            LogSuccess "`nExported $($DLsWithNoManagers.Count) records to [$OutputFile]`n";         
        }
        else
        {
            LogSuccess "`nNo DL found withougt manager`n";
        }
    }
    else
    {
        LogError "`nCould not get groups. $Error`n";
    }
}


##Cleannig session
#RemovePSSession