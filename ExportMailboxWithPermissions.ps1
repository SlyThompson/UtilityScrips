param
(
[Parameter(Mandatory = $true)]
[string]$User,
[Parameter(Mandatory = $true)]
[string]$Password,
[Parameter(Mandatory = $true)]
[string]$ExchangeServer,
[Parameter(Mandatory = $false)]
[string]$EmailAddress,
[Parameter(Mandatory = $false)]
[string]$CsvFile
)

$LogFile = "ExportMailboxWithPermissionsLog";
$OutputFile ="ExportMailboxWithPermissions";
##usage:
##.\ExportMailboxWithPermissions.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'mailboxes.csv'
##.\ExportMailboxWithPermissions.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -EmailAddress test1@sp3.local

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

$MailboxesToExport = @();

if($Error.Count -eq 0 )
{    

    if(-not [string]::IsNullOrEmpty($EmailAddress))
    {
        LogInfo "Working for single user '$EmailAddress'";
        $mailbox = Get-Mailbox $EmailAddress -ErrorAction:SilentlyContinue;    
        if($mailbox -ne $null)
        {
                $mailboxObj = New-Object PSObject;
                $mailboxObj | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias;
                $mailboxObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName;
                $mailboxObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress;
                
                
                $Delegates = "";
                $permissions = Get-MailboxPermission $mailbox.PrimarySmtpAddress;
                foreach($permission in $permissions)
                {
                    if(! $permission.User.Contains(" ") -and $permission.AccessRights.Contains("FullAccess"))
                    {
                        Write-Host "Processing possible delegate $($permission.User) with permission $($permission.AccessRights)";
                        $delegate = Get-Mailbox $permission.User -ErrorAction:SilentlyContinue;    

                        if($delegate -ne $null)
                        {
                            $Delegates+= $delegate.PrimarySmtpAddress.Tostring()+";";
                        }
                    }
                }

                $Delegates= $Delegates.Trim(';');
                $mailboxObj | Add-Member -MemberType NoteProperty -Name Delegates -Value $Delegates;

                $MailboxesToExport+= $mailboxObj;   
        }
    }
    elseif((-not [string]::IsNullOrEmpty($CsvFile)) -and (test-path $CsvFile))
    {
        LogProgress "`nGetting all mailboxes from CSV";
        $csv = Import-Csv $CsvFile;
        $count =0;
        foreach($row in $csv)
        {
            if([string]::IsNullOrEmpty($row.PrimarySmtpAddress))
            {
                continue;
            }

            $count++;
            LogProgress "[$count] Getting mailbox info and permissions for user [$($row.PrimarySmtpAddress)]";

            $mailbox = Get-Mailbox $row.PrimarySmtpAddress -ErrorAction:SilentlyContinue;    
            if($mailbox -ne $null)
            {
                $mailboxObj = New-Object PSObject;
                $mailboxObj | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias;
                $mailboxObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName;
                $mailboxObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress;
                
                
                $Delegates = "";
                $permissions = Get-MailboxPermission $mailbox.PrimarySmtpAddress;
                foreach($permission in $permissions)
                {
                    if(! $permission.User.Contains(" ") -and $permission.AccessRights.Contains("FullAccess"))
                    {
                        Write-Host "Processing possible delegate $($permission.User) with permission $($permission.AccessRights)";
                        $delegate = Get-Mailbox $permission.User -ErrorAction:SilentlyContinue;    

                        if($delegate -ne $null)
                        {
                            $Delegates+= $delegate.PrimarySmtpAddress.Tostring()+";";
                        }
                    }
                }

                $Delegates= $Delegates.Trim(';');

                $mailboxObj | Add-Member -MemberType NoteProperty -Name Delegates -Value $Delegates;

                $MailboxesToExport+= $mailboxObj;   
            }              
        }
    }
    else
    {
        LogError "Please input either -EmailAddress or -CsvFile `n";
    }

    if($MailboxesToExport.Count -gt 0)
    {
        $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

        LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
    }
}
else
{
    LogError "$Error`n";
}

LogInfo "Script ended at [$(Get-Date)]";

##Cleannig session
#RemovePSSession