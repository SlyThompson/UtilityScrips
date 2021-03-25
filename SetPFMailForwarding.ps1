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
[string]$CsvFile,

[Parameter (Mandatory = $true)]
[string]$EmailAddressToForward
)

$LogFile =""; #Log file name would be ScriptNameLog($time).log

##usage:
##.\SetPFMailForwarding.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile "MailEnabledPFs.csv" -EmailAddressToForward  test1@sp3.local;

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


##\\Logs folder
if(-not (Test-Path $LogFolder))
{
    New-item $LogFolder -ItemType Directory | Out-Null
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
        
        LogInfo "Verifying EmailAddressToForward ($EmailAddressToForward) identity ...";
        
        $recipient = $null;
        $recipient= Get-Recipient $EmailAddressToForward -ErrorAction:SilentlyContinue;
        if($recipient)
        {
            LogSuccess "$EmailAddressToForward is a '$($recipient.RecipientTypeDetails)'";
        }
        else
        {
            LogError "$EmailAddressToForward' is not a valid recipient inside current exchange server. $Error `n";
            exit;
        }        

        foreach($row in $csvRows) 
        {
            $count++;                     
                      
            if(-not [string]::IsNullOrEmpty($row.PublicFolder))
            {                

                $Identity = $row.PublicFolder
                
                LogProgress "`n$count# Setting mail forwarding for PublicFolder [$Identity)] to email address '$EmailAddressToForward'";               
                LogInfo "Verifying public folder identity ...";
                
                $PF =$null;
                $PF = Get-MailPublicFolder $Identity -ErrorAction:SilentlyContinue; 
                
                if($PF)               
                {
                    LogInfo  "Setting mail forwarding ...";

                    Set-MailPublicFolder $Identity -ForwardingAddress $EmailAddressToForward -ErrorAction:SilentlyContinue;
                    if($Error.Count -eq 0)
                    {
                        LogSuccess "Mail forwarding set successfully for PF [$Identity].";                        

                    }
                    else
                    {
                        LogError "Mail forwarding failed for PF [$Identity].. $Error";
                        $Error.Clear();
                    }                    
                }
                else
                {
                    LogError "PF [$Identity] is not a Mail Enabled public folder" ;                                        
                }                
                
                #read-host "E";
            }    
            else
            {
                LogInfo "Skipping csv row# $count, due to empty manager SmtpAddress";
            }      
        }       
    }
    else
    {
        LogError "`nNo data found in csv file";
    }
}

##Cleannig session
#RemovePSSession