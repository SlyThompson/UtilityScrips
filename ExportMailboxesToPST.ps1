param
(
[Parameter(Mandatory = $true)]
[string]$User,
[Parameter(Mandatory = $true)]
[string]$Password,
[Parameter(Mandatory = $true)]
[string]$ExchangeServer,
[Parameter(Mandatory = $true)]
[string]$PSTFolderPath,
[Parameter(Mandatory = $false)]##Input either a single email address or a CSV File
[string]$EmailAddress,
[Parameter(Mandatory = $false)]##Input either a single email address or a CSV File
[string]$CsvFile, ##Output of ExportMailboxStats.ps1 script
[Parameter(Mandatory = $false)]
[int]$NumberOfDays = 1095
)

$LogFile = "";
$OutputFile ="";
##usage:
##.\ExportMailboxesToPST.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 NumberOfDays 1095 -CsvFile 'mailboxes.csv' ##Output of ExportMailboxStats.ps1 script
##.\ExportMailboxesToPST.ps1 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 NumberOfDays 1095 -EmailAddress administrator@sp3.local -PSTFolderPath \\192.168.10.13\PSTs

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

if( [string]::IsNullOrEmpty($PSTFolderPath) -or (-not (Test-Path $PSTFolderPath)))
{
	LogError "PST folder path [$PSTFolderPath] doesnt exist.";
	exit;
}
else
{
	if(!$PSTFolderPath.EndsWith("\"))
	{
		$PSTFolderPath+="\";
	}
}
$Error.Clear();

$w = (Get-Host).UI.RawUI.WindowSize.Width/2;
$top = "$("*" * $w)"
$bot = "$("*" * $w)"

LogProgress $top
LogProgress "Script started at [$sTime ]";
LogProgress $bot

ConnectToExchange;

$MailboxesToExport = @{}; ##Dictionary
$ExportRequests = @();##Array
if($Error.Count -eq 0 )
{    
    $rangeTime = (get-date).AddDays(-$NumberOfDays);
    $range =  $rangeTime.ToString("dddd, dd MMMM yyyy 00:00:00");    
    $rangeTime = [System.DateTime]::Parse($range);      
    if(-not [string]::IsNullOrEmpty($EmailAddress))
    {
        LogProgress "Creating export request for single user '$EmailAddress', for emails received\sent after '[$range]'";
		$mailbox =$null;
        $mailbox = Get-Mailbox $EmailAddress -ErrorAction:SilentlyContinue;    
        if($mailbox)
        {
			$identity =$mailbox.Identity;
			$existing = Get-MailboxExportRequest |? {$_.Mailbox -eq $identity};
			if($existing)
			{
				if($existing.Status -eq 'Completed')
				{
					LogInfo "Removing existing request for user '$EmailAddress'";
					Remove-MailboxExportRequest $existing.RequestGuid -Confirm:$false;
					LogInfo "Creating export request now";
				}
			}
					
			$exportRequest =$null;
			
		    $pstFilePath = "$PSTFolderPath\$EmailAddress.pst"
			$exportRequest = New-MailboxExportRequest $EmailAddress -FilePath $pstFilePath -ContentFilter {Received -gt $rangeTime};
			$identity = $exportRequest.RequestGuid;
			
			if($exportRequest)
			{
				LogSuccess "Mailbox export request initiated successfully for [$EmailAddress] with Id [$identity] at path [$($exportRequest.FilePath)]"; 
			    $existingRequest = Get-MailboxExportRequestStatistics $identity;
				do
				{
					LogInfo "Mailbox [$EmailAddress] export request is $($existingRequest.Status) PercentComplete [$($existingRequest.PercentComplete)] TimeTaken [$($existingRequest.OverallDuration)]"; 
					Sleep -Seconds 5;				
					$existingRequest = Get-MailboxExportRequestStatistics  $identity
				}
			    while($existingRequest.Status -eq 'InProgress' -or $existingRequest.Status -eq 'Queued')
				
				if($existingRequest.Status -eq 'Completed')
				{
					LogSuccess "Mailbox export request completed successfully for [$EmailAddress] at path [$($exportRequest.FilePath)]"; 
				}
				else
				{
					LogError "Mailbox export request ended with status [$($existingRequest.Status)] for mailbox [$EmailAddress]"; 
				}
			}
			else
			{
				LogError "Export request cannot be created. $Error";
			}			
        }
		else
		{
			LogError "Mailbox '$EmailAddress' doesn't exist.";
		}
    }
    elseif(-not [string]::IsNullOrEmpty($CsvFile))
    {
		if(test-path $CsvFile)
		{
        	LogProgress "`nGetting all mailboxes from CSV [$CsvFile]";
        	$csv = Import-Csv $CsvFile;
        	$count =0;
			$total = $csv.Count
		
        	foreach($row in $csv)
        	{
            	if([string]::IsNullOrEmpty($row.PrimarySmtpAddress))
            	{
                	continue;
            	}

				$EmailAddress = $row.PrimarySmtpAddress;
            	$count++;
            
				LogProgress "`n[$count\$total] Creating export request for single user '$EmailAddress', for emails received\sent after '[$range]'";
            
				$mailbox = Get-Mailbox $row.PrimarySmtpAddress -ErrorAction:SilentlyContinue;    
			
            	if($mailbox -ne $null)
            	{          
                	$exportRequest =$null;
					$identity =$mailbox.Identity;
			    	$pstFilePath = "$PSTFolderPath\$EmailAddress.pst"
					$existing =$null;
					$existing = Get-MailboxExportRequest |? {$_.Mailbox -eq $identity};
					if($existing)
					{
						if($existing.Status -eq 'Completed')
						{
							LogInfo "Removing existing request for user '$EmailAddress'";
							Remove-MailboxExportRequest $existing.RequestGuid -Confirm:$false;
							LogProgress "Creating export request now";
						}
						elseif($existing.Status -eq 'Queued' -or $existing.Status -eq 'InProgress')
						{
							LogInfo "Mailbox already has export request with status: $($existing.Status )";
							$mailboxObj = New-Object PSObject;
    	            		$mailboxObj | Add-Member -MemberType NoteProperty -Name Mailbox -Value $EmailAddress;
        	        		$mailboxObj | Add-Member -MemberType NoteProperty -Name Status -Value $existingRequest.Status;
            	    		$mailboxObj | Add-Member -MemberType NoteProperty -Name SizeInKB -Value $row.SizeInKB;
							$mailboxObj | Add-Member -MemberType NoteProperty -Name ItemCount -Value $row.ItemCount
							$mailboxObj | Add-Member -MemberType NoteProperty -Name FilePath -Value $existingRequest.FilePath;
					
							$ExportRequests += $mailboxObj;
							continue;
						}
					}
					
					$exportRequest = New-MailboxExportRequest $EmailAddress -FilePath $pstFilePath -ContentFilter {Received -gt $rangeTime};
					$identity = $exportRequest.RequestGuid;
			
					if($exportRequest)
					{
						#$MailboxesToExport.Add($mailbox.Identity , $EmailAddress);  
						LogSuccess "Mailbox export request initiated successfully for [$EmailAddress] at path [$($exportRequest.FilePath)]"; 						
						LogSuccess "[Mailbox Size: $($row.SizeInKB) KB]  and [Item Count: $($row.ItemCount)]";
					    #LogInfo "Mailbox [$EmailAddress] export request is $($existingRequest.Status) PercentComplete [$($existingRequest.PercentComplete)] TimeTaken [$($existingRequest.OverallDuration)]"; 
					
						$existingRequest = Get-MailboxExportRequestStatistics $exportRequest.RequestGuid #| ? {$_.RequestGuid -eq $identity};					
					
						$mailboxObj = New-Object PSObject;
    	            	$mailboxObj | Add-Member -MemberType NoteProperty -Name Mailbox -Value $EmailAddress;
        	        	$mailboxObj | Add-Member -MemberType NoteProperty -Name Status -Value $existingRequest.Status;
            	    	$mailboxObj | Add-Member -MemberType NoteProperty -Name SizeInKB -Value $row.SizeInKB;
						$mailboxObj | Add-Member -MemberType NoteProperty -Name ItemCount -Value $row.ItemCount
						$mailboxObj | Add-Member -MemberType NoteProperty -Name FilePath -Value $existingRequest.FilePath;
					
						$ExportRequests += $mailboxObj;
					}
					else
					{
						LogError "Export request cannot be created. $Error";
					}			                
            	}              
			}
			
			if($ExportRequests.Count -gt 0)
			{			
				$ExportRequests | Export-CSV $OutputFile -NoTypeInformation;    

        		LogSuccess "`nExported $($ExportRequests.Count) records to [$OutputFile]`n";  
			}
        }		
		else
		{
			LogError "Csv file [$CsvFile] doesn't exist";
		}		
    }
    else
    {
        LogError "Please input either -EmailAddress or -CsvFile `n";
		LogError $EmailAddress;
		LogError $CsvFile
    } 
}
else
{
    LogError "$Error`n";
}

LogInfo "Script ended at [$(Get-Date)]";

##Cleannig session
#RemovePSSession