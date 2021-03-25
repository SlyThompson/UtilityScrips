param
(
[Parameter(Mandatory = $true)]
[string]$User,
[Parameter(Mandatory = $true)]
[string]$Password,
[Parameter(Mandatory = $false)]
[string]$ExchangeServer ="outlook.office365.com",
[Parameter(Mandatory = $true)]##Input either a single email address or a CSV File
[string]$CsvFile, ##Output of ExportMailboxStats.ps1 script
[Parameter(Mandatory = $true)]##Input either a single email address or a CSV File
[bool]$CreateMembership =$true ##Output of ExportMailboxStats.ps1 script

)

$LogFile = "";
$OutputFile ="";
##usage:
##.\CreateDLsO365.ps1 -User AD\administrator -Password ok

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
    #$PSUrl = "https://$ExchangeServer/PowerShell/";
    $PSUrl = "https://$ExchangeServer/PowerShell-LiveID?PSVersion=2.0"

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

$AllDLsOutputFolder = "$OutputFolder\AllDLs";
##\\Logs folder
if(-not (Test-Path $LogFolder))
{
    New-item $LogFolder -ItemType Directory | Out-Null
}

if(-not (Test-Path $LogFile))
{
    New-item $LogFile -ItemType File | Out-Null
}


##\\Logs folder
if(-not (Test-Path $AllDLsOutputFolder))
{
    LogError "Folder '$AllDLsOutputFolder' doesn't exist";
}
else
{
    $CsvFiles = Get-ChildItem $AllDLsOutputFolder -Filter *.csv

    if($CsvFiles.Count -gt 0)
    {
        $Error.Clear();

        $w = (Get-Host).UI.RawUI.WindowSize.Width/2;
        $top = "$("*" * $w)"
        $bot = "$("*" * $w)"


        LogProgress $top
        LogProgress "Script started at [$sTime ]";
        LogProgress $bot

        ConnectToExchange;

        $DLsToExport = @();

        if(-not [string]::IsNullOrEmpty($CsvFile))
        {
		    if(test-path $CsvFile)
		    {
        
                $csv = Import-Csv $CsvFile;
                $total = $csv.Count	
                LogProgress "`nProcessing '$total' DLs from CSV [$CsvFile]";
        	
            	$count =0;			
		
            	foreach($row in $csv)
            	{
                	if([string]::IsNullOrEmpty($row.PrimarySmtpAddress))
            	    {
                	    continue;
                	}

	    			$EmailAddress = $row.PrimarySmtpAddress;
                    $Original=  $EmailAddress;
                    
                    #Comment following line for production";
                    
                    $Name = $row.Name;
                    $DisplayName = $row.DisplayName;

                    $ManagerDisplayName = $row.ManagerDisplayName;
                    $ManagerEmail = $row.ManagerEmail;
                    $DLCsvFilePath = "$AllDLsOutputFolder\$Original.csv";
                	$count++;
            
                    #Write-Host $DLCsvFilePath;

                    if(Test-Path $DLCsvFilePath)
                    {                    
	    			    LogProgress "`n[$count\$total] Creating DL as [Name= '$Name'] [Email= $EmailAddress] [Manager= $ManagerEmail]";
                        $dl =$null;
			    	    $dl = Get-DistributionGroup 	$EmailAddress -ErrorAction:SilentlyContinue;    
			
            	        if($dl)
                	    {          
                    	    LogInfo "DL already exist";	                
                	    }     
                        else
                        {
                            Write-Host $Error;
                            $dl = New-DistributionGroup -Name $Name -DisplayName $DisplayName -PrimarySmtpAddress $EmailAddress;
                            if($dl)
                            {
                                LogSuccess "DL created successfully";                                                
                            }
                            else
                            {
                                LogError "DL could not be created. $Error"; $Error.Clear();
                            }
                        }
                    
                        if($dl)
                        {

                            if(-not [string]::IsNullOrEmpty($ManagerEmail))
                            {
                                LogProgress "Setting manager [Name ='ManagerDisplayName', Email= $ManagerEmail] to DL '$EmailAddress'";

                                Set-DistributionGroup $dl.PrimarySmtpAddress -ManagedBy $ManagerEmail -ErrorAction:SilentlyContinue; 
                                if($Error.Count -gt 0)
                                {
                                    LogError "Manager could not be set. $Error"; $Error.Clear();
                                }
                                else
                                {
                                    LogSuccess "Manager updated successfully";
                                }
                            }                        
                        }
                        if($CreateMembership)
                        {
                            $members = Import-Csv $DLCsvFilePath;
                            if($members.Count -gt 0)
                            {
                                $membersCount = $members.Count;
                                $count2 =0 ;
                                
                                LogInfo "Found $($membersCount) members in csv file [$DLCsvFilePath]";

                                foreach($member in $members)
                                {
                                    $MemberEmailAddress = $member.PrimarySmtpAddress;                                                                        
                                    $MemberName = $member.Name;                                    
                                    $count2++;

                                    if(-not[string]::IsNullOrEmpty($MemberEmailAddress))
                                    {                              

                                        LogProgress "`n[$count2\$membersCount] adding user [Name= '$MemberName', Email= $MemberEmailAddress] as the member to group '$EmailAddress'";

                                        Add-DistributionGroupMember $EmailAddress -Member $MemberEmailAddress -ErrorAction:SilentlyContinue; 
                                        if($Error.Count -gt 0)
                                        {
                                            LogError "Member could not be added. $Error"; $Error.Clear();
                                        }
                                        else
                                        {
                                             LogSuccess "Member added successfully";
                                        }
                                    }
                                    else
                                    {
                                        LogProgress "`n[$count2\$membersCount] Skipping '$MemberName' due to empty email address";
                                    }

                                    #Read-Host $count2;  
                                }

                            }
                        }
                        
                        #Read-Host $count;   
                    }
                    else
                    {
                        LogProgress "[$count\$total] Group csv doesnt exist. [$DLCsvFilePath]";
                    }                     
                       
			    }##foreach loop		
            }		
		    else
		    {
			    LogError "Csv file [$CsvFile] doesn't exist";
		    }		
        }
        else
        {
            LogError "Please input -CsvFile `n";    	    
        } 
    }
    else
    {
        LogError "`nFolder '$AllDLsOutputFolder' doesn't have any csv file, which contains group memberships";
    }
}
##Cleannig session
#RemovePSSession