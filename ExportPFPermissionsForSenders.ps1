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
##.\ExportPFPermissionsForSenders.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'ExportManagersForSenders-05-06-2020-02.csv'

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

$Mailboxes = @{};
$MailboxesToExport = @();

if($Error.Count -eq 0 )
{    
    if((test-path $CsvFile))
    {
        LogProgress "`nGetting all delegate mailboxes from CSV [$csvFile]";
        $csv = Import-Csv $CsvFile;
        #$csv = $rawCSV |Group-Object -Property UserEmail

        $count =0;
        $total = 1;
        if($csv.Count -ne $null)
        {
            $total = $csv.Count;
        }

        if($total -gt 0)
        {       
            $count = 0;
            foreach($row in $csv)
            {
                #So here, using email address of the user, we have to verify that user exists in Exchange Server\AD.
                #We also have to get the Canonical name of the user which is in format domain.ad/ou/user_identity
                #Because, PF permission lists the user in canonical format, so for matching, we need canonical user name for each user email address

                $user = $row.UserEmail;

                if([string]::IsNullOrEmpty($user))
                {
                    continue;
                }            
                           
                $count++;

                LogProgress "`n[$count\$Total] Processing user [$user]";   
                    
                $mailboxObj = $null                    
                $userObj = Get-User $user -ErrorAction:SilentlyContinue;
                    
                if($userObj -ne $null)
                {       
                     $Mailboxes.Add($userObj.Identity,$row); #$mailboxObj);      
                }                    
                else
                {
                        LogInfo "Sender [$user] is not a User object.";
                }
                 
                #Read-Host "$count";                                       
            }   
        
        
            if($Mailboxes.Count -gt 0)
            {
                LogSuccess "`nRetrieved $($Mailboxes.Count) user records from csv.`n";                 

                $AllPFs = Get-publicFolder \ -recurse
                $AllPFs = $Allpfs | ?{$_.Name -ne 'IPM_SUBTREE'}

                foreach($pf in $AllPFs)
                {
                    $pfIdentity = $pf.Identity.ToString();
                    Write-Host "Processing public folder '$pfIdentity'";
                    
                    #$permissions = Get-publicFolderClientPermission $pfIdentity  | Where-Object { -not($_.User.IsDefault -eq $true -or $_.User.IsAnonymous -eq $true)}
                    
                    $permissions = Get-publicFolderClientPermission $pfIdentity  | ? {$_.User -ne "Anonymous" -and $_.User -ne "Default"}

                    foreach($perm in $permissions)
                    {
                        $permUser = $perm.User.ToString();

                        ##Check if user canonical name exists for existing users (which exists in input csv file)
                        if($Mailboxes.ContainsKey($permUser))
                        {
                            #If current permission user exists in csv file, then Get the user email address, and manager info using user's canonical name

                            $existingMailbox = $Mailboxes[$permUser];

                            $mailboxObj = New-Object PSObject;
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name UserEmail -Value $existingMailbox.UserEmail; 
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name PublicFolder -Value $perm.Identity.ToString(); 
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name Permissions -Value  $perm.AccessRights;

                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $existingMailbox.ManagerName; 
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value $existingMailbox.ManagerEmail; 

                            $MailboxesToExport+= $mailboxObj;
                        }
                        else
                        {
                            #Write-Host "Skipping $permUser";
                        }
                    }
                    
                }

                if($MailboxesToExport.Count -gt 0)
                {
                    $MailboxesToExport | Sort-Object UserEmail | Export-CSV $OutputFile -NoTypeInformation;    

                    LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
                }
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