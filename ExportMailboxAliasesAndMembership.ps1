param
(
[Parameter(Mandatory = $true)]
[string]$User,
[Parameter(Mandatory = $true)]
[string]$Password,
[Parameter(Mandatory = $true)]
[string]$ExchangeServer
)

$LogFile =""; #Log file name would be ScriptNameLog($time).log
$OutputFile ="";  #Output CSV file name would be ScriptName($time).csv
$OutputFileGroupMemberships ="";  #Output CSV file name would be ScriptName($time).csv
##usage:
##.\ExportMailboxAliasesAndMembership.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13

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


$OutputFileGroupMemberships = "$OutputFolder\GroupMemberships-$timeStr.csv";

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
$GroupMemberships = @();

if($Error.Count -eq 0 )
{
    LogProgress "`nGetting mailbox list.";

	#$PFs = Get-PublicFolder \ -recurse -ResultSize Unlimited -ErrorAction:SilentlyContinue;
    $Mailboxes =  Get-Mailbox  |?{$_.RecipientTypeDetails -ne "DiscoveryMailbox"}
    
	$Error.Clear();

    if($Mailboxes.Count -gt 0)
    {    
        LogSuccess "`nFound $($Mailboxes.Count) mailboxes.`n";
        #Get-ADUser $user | Get-ADPrincipalGroupMembership | select -Expand Distinguishedname | Get-DistributionGroup -EA 0
        $AllGroups = Get-DistributionGroup -resultSize Unlimited;

        foreach($mailbox in $Mailboxes)
        {   
            $memberOfGroup= $false;     
            $Username = $mailbox.PrimarySmtpAddress
            
            $DistributionGroups=  $AllGroups| where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains $Username}
            
            $count = 1;
            if($DistributionGroups.Count -ne $null)
            {
                $count = $DistributionGroups.Count;
				$DistributionGroups = $DistributionGroups | ?{$_.Name -ne $null -and $_.Name -ne ''};
            }
            if($count -gt 0)
            {
                $memberOfGroup= $true;
				foreach($dg in $DistributionGroups)
                {
				    if($dg.Name -eq $null)
					{
						continue;
					}
					#Write-Host "User:$($mailbox.PrimarySmtpAddress) Group:$($dg.PrimarySmtpAddress)";
                	$memObj = New-Object PSObject;
                	$memObj | Add-Member -MemberType NoteProperty -Name UserName -Value $mailbox.DisplayName; 
                	$memObj | Add-Member -MemberType NoteProperty -Name UserEmailAddress -Value $mailbox.PrimarySmtpAddress
                	$memObj | Add-Member -MemberType NoteProperty -Name MailboxType -Value $mailbox.RecipientTypeDetails; 
                	$memObj | Add-Member -MemberType NoteProperty -Name GroupName -Value $dg.Name;
                	$memObj | Add-Member -MemberType NoteProperty -Name GroupEmailAddress -Value $dg.PrimarySmtpAddress; 
                
                	$GroupMemberships += $memObj;                
            	}
            }            
            else
            {
                $DistributionGroups= $null;
            }
            
            foreach($ea in $mailbox.EmailAddresses)
            {				
                $mbxObj = New-Object PSObject;
                $mbxObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName; 
                $mbxObj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $ea.SmtpAddress
                $mbxObj | Add-Member -MemberType NoteProperty -Name IsPrimaryEmailAddress -Value $ea.IsPrimaryAddress;
                $mbxObj | Add-Member -MemberType NoteProperty -Name MailboxType -Value $mailbox.RecipientTypeDetails; 
                $mbxObj | Add-Member -MemberType NoteProperty -Name IsMemberOfAnyGroup -Value $memberOfGroup; 
                $MailboxesToExport += $mbxObj;
            }                
        }

        if($MailboxesToExport.Count -gt 0)
        {		
            $MailboxesToExport | Export-Csv $OutputFile -NoTypeInformation
            LogSuccess "Exported $($MailboxesToExport.Count) email addresses to file [$OutputFile]";
        }
        if($GroupMemberships.Count -gt 0)
        {
            $GroupMemberships | Export-Csv $OutputFileGroupMemberships -NoTypeInformation
            LogSuccess "Exported $($GroupMemberships.Count) user group memberships to file [$OutputFileGroupMemberships]";
        }
    }
    else
    {
        LogError "`nCould not get maibx list. $Error`n";
    }
}


##Cleannig session
#RemovePSSession