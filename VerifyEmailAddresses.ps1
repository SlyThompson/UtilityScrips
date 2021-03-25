param
(
[Parameter (Mandatory = $true)]
[string]$User,
[Parameter (Mandatory = $true)]
[string]$Password,
[Parameter (Mandatory = $true)]
[string]$ExchangeServer,

[Parameter (Mandatory = $true)]
[ValidateScript({
            if( -Not ($_ | Test-Path) ){
                throw "File does not exist"
            }
            return $true
        })]
[string]$CsvFile
)

$LogFile =""; #Log file name would be ScriptNameLog($time).log
$OutputFile ="";
##usage:
##.\VerifyEmailAddresses.ps1 -User Sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile "EmailAddresses.csv"

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
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PSUrl -Authentication Basic -Credential $UserCredential -SessionOption $PSSessionOption

    if($Error.Count -eq 0)
    {
        LogInfo "`nImporting existing exchange powershell session";
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

if(-not (Test-Path $LogFolder))
{
    New-item $LogFolder -ItemType Directory; 
}

if(-not (Test-Path $LogFile))
{
    New-item $LogFile -ItemType File; 
}

#endregion

##Taking script name from PS environment
$scriptName = $MyInvocation.MyCommand.Name
$scriptName  =$scriptName.Replace(".ps1","");

$scriptPath = GetScriptPath;
$sTime = get-date;
$timeStr = $sTime.ToString("dd-MM-yyyy-hh");

$LogFolder = "$scriptPath\Logs";
$LogFile   = "$LogFolder\$($scriptName)Log-$timeStr.log";

$OutputFolder = "$scriptPath\Output";
$OutputFile = "$OutputFolder\$($scriptName)-$timeStr.csv";

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

$csv = $CsvFile; #"$PSScriptRoot\DistributionListWithAssignedOwners.csv"

ConnectToExchange;

if($Error.Count -eq 0 )
{
    LogProgress "`nProcessing csv [$csv]";

    ##Importing csv data to a variable
    $csvRows = Import-Csv $csv

    $Error.Clear();

    if($csvRows -ne $null -and $csvRows.Count -gt 0)
    {    
        LogSuccess "`nFound $($csvRows.Count) rows in csv file`n";

        ##Iterating over $csv rows
        $count =0;

        $ObjectsToExport = @();        

        foreach($row in $csvRows) 
        {
            $count++;                     
                      
            if(-not [string]::IsNullOrEmpty($row.EmailAddress.Trim()))
            {                

                $Identity = $row.EmailAddress.Trim()
                
                LogProgress "`n$count# Processing email address '$Identity '";               
                LogInfo "Verifying email address ...";
                
                $PF =$null;
                $PF = Get-Recipient $Identity -ErrorAction:SilentlyContinue; 
                $isAlias = $false;
                $isValid =$false;
                $forwarding = $null;
                $type= "";

                if($PF)               
                {
                    $isValid = $true;
                    
                    $primary = $PF.PrimarySmtpAddress.ToString();
                    
                    if($primary.Tolower() -ne $Identity.Tolower())
                    {
                        $isAlias =$true;
                    }   

                    $type =$PF.RecipientTypeDetails.ToString();

                    if($type.Contains("Mailbox"))
                    {
                        $fwding =$null;
                        LogInfo  "Getting mail forwarding ...";
                        $fwding = Get-Mailbox $Identity |Select ForwardingAddress
                        if($fwding.ForwardingAddress)
                        {
                            $forwarding = Get-Mailbox $fwding.ForwardingAddress |Select PrimarySmtpAddress
                            LogProgress "Forwarding is set to '$($forwarding.PrimarySmtpAddress)'";
                        }
                        else
                        {
                            LogProgress "Forwarding not set";
                        }                        
                    }
                    else
                    {
                        LogSuccess "Recipient '$Identity' is not a mailbox, instead '$type'";
                    }
                }
                else
                {
                    LogError "Email [$Identity] doesn't exist in Exchange server";    
                }                
                
                $groupObj = New-Object PSObject;
                
                $groupObj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $Identity;
                $groupObj | Add-Member -MemberType NoteProperty -Name IsValidEmailAddress -Value $isValid
                $groupObj | Add-Member -MemberType NoteProperty -Name IsAlias -Value $isAlias;               
                
                $groupObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $PF.PrimarySmtpAddress;
                $groupObj | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $forwarding.PrimarySmtpAddress;
                $groupObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $PF.DisplayName;
                $groupObj | Add-Member -MemberType NoteProperty -Name Type -Value $type;
                

                $ObjectsToExport += $groupObj;                           
                #read-host "E";
            }    
            else
            {
                LogInfo "Skipping csv row# $count, due to empty manager SmtpAddress";
            } 
            
                 
        }  
        
        if($ObjectsToExport.Count -gt 0)
        {            
            $ObjectsToExport | Export-CSV $OutputFile -NoTypeInformation;    

            LogSuccess "`nExported $($ObjectsToExport.Count) records to [$OutputFile]`n";         
        }
        else
        {
            LogSuccess "`nNo DL found`n";
        }     
    }
    else
    {
        LogError "`nNo data found in csv file";
    }
}

##Cleannig session
#RemovePSSession