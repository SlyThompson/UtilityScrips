param
(
[Parameter (Mandatory = $false)]
[string]$User,
[Parameter (Mandatory = $false)]
[string]$Password,
[Parameter (Mandatory = $false)]
[string]$ADServer
)

$LogFile = "ExportDeletedDLsLog";
$OutputFile ="ExportDeletedDLs";
##usage:
##.\ExportDeletedDLs.ps1
##.\ExportDeletedDLs.ps1 -User ad\administrator -Password ok -ADServer 192.168.10.10


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

$AllDLsOutputFolder = "$OutputFolder\DeletedDLs";

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

$Error.Clear();

LogInfo "`nImporting ActiveDirectory module";

Import-Module ActiveDirectory

$ADParameters =@{};

$ADParameters.Add("IncludeDeletedObjects", $true);

if(![string]::IsNullOrEmpty($User) -and ![string]::IsNullOrEmpty($Password))
{
    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force;
    $Credentials = New-Object PSCredential $User, $SecurePassword ;

    $ADParameters.Add("Credential",$Credentials);
}

if($ADServer)
{
    $ADParameters.Add("Server",$ADServer);
}

$DLsToExport = @();

if($Error.Count -eq 0 )
{
    LogProgress "`nGetting all deleted groups";

    $DLs =  Get-ADObject -filter {isdeleted -eq $true -and objectclass -eq 'group' -and msExchRecipientDisplayType -eq 1} -property * @ADParameters
    $Error.Clear();
    if($Error.Count -eq 0 )
    {
        if($DLs  -ne $null -and $DLs.Count -gt 0)
        {    
            LogSuccess "`nFound $($DLs.Count) deleted groups";
            
            $DLs |Select @{Label = "Name";Expression = {($_.Name -split "`n")[0]}}, @{Label = "OU";Expression = {$_.LastKnownParent}},  @{Label = "DeletionTime";Expression = {$_.WhenChanged}}, @{Label = "EmailAddress";Expression = {$_.Mail}},@{Label = "Manager";Expression = {$_.ManagedBy}}  | Export-CSV $OutputFile -NoTypeInformation;    

            LogSuccess "`nExported $($DLs.Count) records to [$OutputFile]";                     

            foreach($DL in $DLs)
            {
                ##If DL has more than 0 members then export them
                if($DL.Member.Count -gt 0)                
                {
                    $Identity = $DL.Mail;
                 
                    $path = "$AllDLsOutputFolder\$Identity.csv";                

                    if(Test-Path $path)
                    {
                        Remove-Item $path -Confirm:$false -ErrorAction:SilentlyContinue;
                    }

                    $members = $DL.Member -split "`n";

                    "member" |Out-File $path -Append;

                    LogProgress "`nExporting members of DL '$Identity'";
                    
                    foreach($member in $members)
                    {
                        $member |Out-File $path -Append;    
                    }
                }                
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