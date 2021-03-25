param
(
[Parameter (Mandatory = $true)]
[string]$User,
[Parameter (Mandatory = $true)]
[string]$Password,
[Parameter (Mandatory = $true)]
[string]$ExchangeServer,

[Parameter (Mandatory = $false)]
[string]$CsvFilePath
)

$LogFile = "RemoveDLWithNoManager";
##usage:
##.\RemoveDLWithNoManager.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10 -CsvFilePath "C:\Scripts\File.csv";

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
    ##Importing csv data to a variable

    $csvRows =@();
    if([string]::IsNullOrEmpty($CsvFilePath))
    {
        do
        {
            $group =Read-Host "`nEnter DL Name or Email Address to be removed ";
        }
        while(!$group);

        $groupObj = New-Object PSObject;
        $groupObj | Add-Member -MemberType NoteProperty -Name Distro_Name -Value $group;

        $csvRows+= $groupObj;
        
    }
    else
    {
       if(Test-Path $CsvFilePath)
       {
            LogProgress "`nProcessing csv [$csv]";
            $csvRows = Import-Csv $csv
            LogSuccess "`nFound $($csvRows.Count) rows in csv file`n";
       }
       else
       {
           LogError "CSV file doesnt exist.";
           exit;
       }        
    }

    $Error.Clear();

    if($csvRows -ne $null -and $csvRows.Count -gt 0)
    {
        ##Iterating over $csv rows
        $count =0;

        foreach($row in $csvRows) 
        {
            $count++;                     
            $Identity = $row.Distro_Name
                      
            if(-not [string]::IsNullOrEmpty($Identity))
            {                                
                
                LogProgress "`n$count# Processing group [$Identity]";
               
                $var =$null;
                $DL = Get-DistributionGroup $Identity ;#-ErrorAction:SilentlyContinue; #-ErrorAction:SilentlyContinue;

                
                if($Error.Count -gt 0)#$var -ne $null)# 
                {
                    LogToFile "DL [$Identity] could not be found on exchange server. $Error" ;                    
                }
                else
                {
                    LogSuccess  "DL [$Identity] found.";

                    $manager = $DL.ManagedBy;
                    if($manager -eq $null)
                    {
                        LogProgress "Group manager is not set. Removing DL.";

                        Remove-DistributionGroup $Identity -Confirm:$false;
                        
                        if($Error.Count -eq 0)
                        {
                            LogSuccess "DL removed successfully.";
                        }
                        else
                        {
                            LogError "Error: $Error"; 
                        }

                        #read-host "E";
                    }
                    else
                    {
                        LogSuccess "Manager [$manager] is set.";
                    }                    
                }

                $Error.Clear();
                #
            }    
            else
            {
                LogInfo "`nSkipping csv row# $count, due to empty Distro_Name";
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