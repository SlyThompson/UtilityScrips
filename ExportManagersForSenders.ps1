param
(
[Parameter(Mandatory = $true)]
[string]$User,
[Parameter(Mandatory = $true)]
[string]$Password,
[Parameter(Mandatory = $true)]
[string]$ExchangeServer,
[Parameter(Mandatory = $true)]
[string]$CsvFile
)

$LogFile =""; #Log file name would be ScriptNameLog($time).log
$OutputFile ="";  #Output CSV file name would be ScriptName($time).csv
# D:\Projects\Powershell\Ex\DG ManagedBy (Sylvester)\Output\ExportMailboxWithPermissions-18-04-2020-04.csv

##usage:
##.\ExportManagersForSenders.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'delegates.csv'

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

$MailboxesToExport = @();

if($Error.Count -eq 0 )
{    
    if((test-path $CsvFile))
    {
        LogProgress "`nGetting all delegate mailboxes from CSV [$csvFile]";
        $rawCSV = Import-Csv $CsvFile;
        $csv = $rawCSV |Group-Object -Property Sender

        $count =0;
        $total = 1;
        if($csv.Count -ne $null)
        {
            $total = $csv.Count;
        }

        if($total -gt 0)
        {
            $Managers = @{};        
        
            $count = 0;
            foreach($row in $csv)
            {
                $user = $row.Name.ToLower();

                if([string]::IsNullOrEmpty($user))
                {
                    continue;
                }            
                           
                $count++;

                LogProgress "`n[$count\$Total] Processing sender [$user]";   
                    
                $mailboxObj = $null                    
                $userObj = Get-User $user -ErrorAction:SilentlyContinue;
                    
                if($userObj -ne $null)
                {              
                     $managerDn  = $userObj.Manager;
                     $managerUser = $null;

                                
                     if(! [string]::IsNullOrEmpty($managerDn))
                     {
                          if($Managers.ContainsKey($managerDn.ToLower()))
                          {                   
                              $existingMailbox = $Managers[$managerDn.ToLower()];
                        
                              Read-Host "Found existing";
                              $mailboxObj = New-Object PSObject -Property @{
                                      UserEmail       = $User                                            
                                      ManagerName       = $existingMailbox.ManagerName
                                      ManagerEmail = $existingMailbox.ManagerEmail                                            
                              }
                          }
                          else
                          {
                              $mailboxObj = New-Object PSObject;
                              $mailboxObj | Add-Member -MemberType NoteProperty -Name UserEmail -Value $User; 
                                        
                              LogInfo "Getting manager [$managerDn] details";
                              $managerUser = Get-Mailbox $managerDn -ErrorAction:SilentlyContinue;                        
                              if($managerUser)
                              {
                                   $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.DisplayName; 
                                   $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.PrimarySmtpAddress;                             
                              }
                              else
                              {
                                   $managerUser = Get-User $managerDn -ErrorAction:SilentlyContinue;
                                   $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.DisplayName; 
                                   $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.WindowsEmailAddress;                             
                              }

                              $Managers.Add($User.ToLower(), $mailboxObj);                        
                          }
                     }
                     else
                     {
                        LogInfo "Skipping user [$user] due to no manager set.";
                     }
                 }                    
                 else
                 {
                        LogInfo "Sender [$user] is not a User object.";
                 }
                 if($mailboxObj -ne $null)
                 {                    
                        $MailboxesToExport+= $mailboxObj;
                 }

                 Read-Host "$count";                                       
            }   
        
        
            if($MailboxesToExport.Count -gt 0)
            {
                $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

                LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
            }    
        }
        else
        {
            LogSuccess "`nNo record found in csv file";
        }        
    }
    else
    {
        LogError "`nCsvFile doesn't exist [$csvFile]`n";
    }    
}
else
{
    LogError "`n$Error";
}

LogInfo "`nScript ended at [$(Get-Date)]";

##Cleannig session
#RemovePSSession