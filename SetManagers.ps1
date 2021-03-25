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

$LogFile = "SetManagersLog";
##usage:
##.\SetManagers.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10 -CsvFile "C:\Scripts\File.csv";

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

$Error.Clear();

$w = (Get-Host).UI.RawUI.WindowSize.Width/2;
$top = "$("*" * $w)"
$bot = "$("*" * $w)"


LogProgress $top
LogProgress "Script started at [$sTime ]";
LogProgress $bot

$csv = $CsvFilePath; #"$PSScriptRoot\DistributionListWithAssignedOwners.csv"

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

        foreach($row in $csvRows) 
        {
            $count++;                     
                      
            if(-not [string]::IsNullOrEmpty($row.SmtpAddress))
            {                

                $Identity = $row.Distro_Name
                
                LogProgress "`n$count# Processing group [$Identity)] with manager [ $($row.NewManager) ($($row.SMTPAddress))]";
               
                $var =$null;
                $DL = Get-DistributionGroup $Identity ;#-ErrorAction:SilentlyContinue; #-ErrorAction:SilentlyContinue;

                #Write-Host $Error -ForegroundColor Magenta;
                if($Error.Count -gt 0)#$var -ne $null)# 
                {
                    LogToFile "DL [$Identity] could not be found on exchange server. $Error" ;                    
                }
                else
                {
                    LogSuccess  "DL [$Identity] found.";

                    $manager = Get-User $row.SMTPAddress -ErrorAction:SilentlyContinue;
                    if($manager -ne $null)
                    {
                        LogProgress "Setting group manager";
                        Set-DistributionGroup $Identity -ManagedBy $row.SMTPAddress -ErrorAction:SilentlyContinue;
                        if($Error.Count -eq 0)
                        {
                            LogSuccess "Manager set successfully.";
                        }
                        else
                        {
                            LogError "Error: $Error"; 
                        }
                    }
                    else
                    {
                        LogError "Manager user doesn't exist. [ $($row.NewManager) ($($row.SMTPAddress))]. $Error";
                    }                    
                }

                $Error.Clear();
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
RemovePSSession