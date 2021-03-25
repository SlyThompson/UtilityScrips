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
$OutputFileGroupMemberships ="";  #Output CSV file name would be ScriptName($time).csv
##usage:
##.\ExportMailboxAliasesAndMembership-v2.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 '.\Output\EmailAddresses.csv'

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
    if(test-path $CsvFile)
    {        
        LogProgress "`nGetting email addresses from CSV [$csvFile]";
        $csv = Import-Csv $CsvFile;
        $count =0;
        $total = 1;
        if($csv.Count -ne $null)
        {
            $total = $csv.Count;
        }
        LogSuccess "`nFound $total email addresses in csv file.`n";

        if($total -gt 0)
        {            
            $AllGroups = Get-DistributionGroup -resultSize Unlimited;
            $count = 0;
            foreach($row in $csv)
            {
                $emailAddress = $row.EmailAddress.ToLower();
                if([string]::IsNullOrEmpty($row) -or [string]::IsNullOrEmpty($emailAddress))
                {
                    continue;
                }            
                
                $count++;

                LogProgress "`n[$count\$Total] Processing email address [$emailAddress]";                
            
                $mailbox =$null;
                $mailbox = Get-Mailbox $emailAddress -ErrorAction:SilentlyContinue;
                if($mailbox -ne $null)
                {
                    LogSuccess "Mailbox found for email address [$emailAddress] of type '$($mailbox.RecipientTypeDetails)'";

                    $memberOfGroup= $false;     
                    $Username = $mailbox.PrimarySmtpAddress
            
                    $DistributionGroups=  $AllGroups| where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains $Username}
            
                    $DLCount = 1;
                    if($DistributionGroups.Count -ne $null)
                    {
                        $DLCount = $DistributionGroups.Count;
				        #$DistributionGroups = $DistributionGroups | ?{$_.Name -ne $null -and $_.Name -ne ''};
                    }
                    if($DLCount -gt 0)
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
                else
                {
                    LogError "Mailbox could not be found for email address [$emailAddress].";# $Error"; $Error.Clear();
                }
            }
        
            if($MailboxesToExport.Count -gt 0)
            {		
                $MailboxesToExport | Export-Csv $OutputFile -NoTypeInformation
                LogProgress "`nExported $($MailboxesToExport.Count) email addresses to file [$OutputFile]";
            }
            if($GroupMemberships.Count -gt 0)
            {
                $GroupMemberships | Export-Csv $OutputFileGroupMemberships -NoTypeInformation
                LogProgress "`nExported $($GroupMemberships.Count) user group memberships to file [$OutputFileGroupMemberships]";
            }
        }    
        else
        {
            LogError "`nCould not get any data from csv file. $Error`n";
        }
    }
    else
    {
        LogError "`nCsv file doesn't exist [$CsvFile]`n";
    }    
}

##Cleannig session
#RemovePSSession