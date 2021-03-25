#region Functions

$Commands = "Get-User, Get-Mailbox, Get-PublicFolder, Get-PublicFolderClientPermission";
$global:LogFile =""; 

$ExchangeSession = $null;
$WarningPreference = "SilentlyContinue";
$ModuleName = "ExchangeModule";

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
    "[$time] $msg" | Out-File -FilePath $global:LogFile -Append;
}

function ConnectToExchange
{
    $ExchangeSession = Get-PSSession |?{$_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.State -eq 'Opened'};
    if($ExchangeSession)
    {
        ##If an existing session already exists in PowerShell console, use that one, instead of creating a new one;
        Write-Host "`nUsing existing exchange powershell session connected to computer [$($ExchangeSession.ComputerName)]`n" ;
        return;
    }

    RemovePSSession

    #return;

    Write-Host "Connecing to Exchange server using [User: $User] [Password: $Password] and [Exchange Server: $ExchangeServer]" -f Yellow

    $EncodedPwd = ConvertTo-SecureString $Password -AsPlainText -Force;
    $UserCredential= New-Object System.Management.Automation.PSCredential($User, $EncodedPwd);

    $PSUrl = "https://$ExchangeServer/PowerShell/";

    $PSSessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -MaximumRedirection 5;
    $Error.Clear();
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PSUrl -Authentication Basic -Credential $UserCredential -SessionOption $PSSessionOption -ErrorAction:SilentlyContinue;

    if($Error.Count -eq 0)
    {
        Write-Host "`nImporting exchange powershell session" -f Yellow
       
        #Import-PSSession $ExchangeSession  -AllowClobber #-Module "ExchangePSSession"  #-DisableNameChecking 
        #Import-Module "ExchangePSSession" -Global
        #$ss = Import-PSSession $ExchangeSession  # -CommandName $Commands   -Module "CustomExchangeModule"
        Import-Module (Import-PSSession $ExchangeSession )  -Global 
        
    }
    else
    {
        Write-Host "Could not load exchange powershell session. $Error" -f Red;
    }
}

function RemovePSSession
{
    #Write-Host "`nRemoving PS Sessions";
    Get-PSSession | Remove-PSSession;
}
function GetScriptPath
{     
    Split-Path $myInvocation.ScriptName 
}


function InitializeLoggging($OutputFileName)
{

        $scriptPath = GetScriptPath;
        $sTime = get-date;
        $timeStr = $sTime.ToString("dd-MM-yyyy-hh");

        $LogFolder = "$scriptPath\Logs";
        $global:LogFile   = "$LogFolder\$($ModuleName)Log-$timeStr.log";

        $OutputFolder = "$scriptPath\Output";
        $OutputFile = "$OutputFolder\$OutputFileName-$timeStr.csv";

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

        if(-not (Test-Path $global:LogFile))
        {
            New-item $global:LogFile -ItemType File | Out-Null
        }

        return $OutputFile;
}

##CsvFile: Input csv should be Output of ExportManagersForSenders script
function ExportPFPermissionsForSenders
{
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

    process 
    {    
        #Log file name would be ScriptNameLog($time).log
        $OutputFile ="";  #Output CSV file name would be ScriptName($time).csv
        ##usage:
        ## ExportPFPermissionsForSenders -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'ExportManagersForSenders-05-06-2020-02.csv'

        ##Taking script name from PS environment
        $scriptName = $MyInvocation.MyCommand.Name
        
        $OutputFile = InitializeLoggging -OutputFileName $scriptName
        $Error.Clear();        

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

                Read-Host "Enter"
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
    }
}


##CsvFile: Input csv should be Output of SearchEmailTrackingLogs script
function ExportManagersForSenders
{
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

    process 
    {   
        ##usage:
        ## ExportManagersForSenders -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'senders.csv'
        
        ##Taking function name from PS environment 
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();

        ConnectToExchange;

        $MailboxesToExport = @();

        if($Error.Count -eq 0 )
        {    
            if((test-path $CsvFile))
            {
                LogProgress "Getting all delegate mailboxes from CSV [$csvFile]";
                $rawCSV = Import-Csv $CsvFile;
                $csv = $rawCSV |Group-Object -Property Sender
                    
                $count =0;
                $total = 1;
                if($csv.Count -ne $null)
                {
                    $total = $csv.Count;
                }

                if($total -gt 0)
                {
                    $Managers = @{};        
        
                    $count = 0;
                    foreach($row in $csv)
                    {
                        $user = $row.Name.ToLower();

                         if([string]::IsNullOrEmpty($user))
                         {
                             continue;
                         }            
                               
                         $count++;

                         LogProgress "[$count\$Total] Processing sender [$user]";   
                    
                         $mailboxObj = $null                    
                         $userObj = Get-User $user -ErrorAction:SilentlyContinue;
                    
                         if($userObj -ne $null)
                         {              
                            $managerDn  = $userObj.Manager;
                            $managerUser = $null;
                                    
                            if(! [string]::IsNullOrEmpty($managerDn))
                            {
                                if($Managers.ContainsKey($managerDn.ToLower()))
                                {                   
                                      $existingMailbox = $Managers[$managerDn.ToLower()];
                            
                                      Read-Host "Found existing";
                                      $mailboxObj = New-Object PSObject -Property @{
                                          UserEmail       = $User                                            
                                          ManagerName       = $existingMailbox.ManagerName
                                          ManagerEmail = $existingMailbox.ManagerEmail                                            
                                      }
                                }
                                else
                                {
                                    $mailboxObj = New-Object PSObject;
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name UserEmail -Value $User; 
                                        
                                    LogInfo "Getting manager [$managerDn] details";
                                    $managerUser = Get-Mailbox $managerDn -ErrorAction:SilentlyContinue;                        
                                    if($managerUser)
                                    {
                                           $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.DisplayName; 
                                           $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.PrimarySmtpAddress;                             
                                    }
                                    else
                                    {
                                       $managerUser = Get-User $managerDn -ErrorAction:SilentlyContinue;
                                       $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.DisplayName; 
                                       $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.WindowsEmailAddress;                             
                                    }

                                    $Managers.Add($User.ToLower(), $mailboxObj);                        
                                }
                            }
                            else
                            {
                                LogInfo "Skipping user [$user] due to no manager set.";
                            }
                         }                    
                         else
                         {
                             LogInfo "Sender [$user] is not a User object.";
                         }
                     
                         if($mailboxObj -ne $null)
                         {                    
                             $MailboxesToExport+= $mailboxObj;
                         }

                         #Read-Host "$count";                                       
                    }           
            
                    if($MailboxesToExport.Count -gt 0)
                    {
                         $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

                        LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
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
     }
 }

##CsvFile: Input csv should be Output of ExportMailboxWithPermissions script
function ExportDelegatesAndManagers
{
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

    process 
    {   
        ##usage:
        ## ExportDelegatesAndManagers2 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'delegates.csv'
        
        ##Taking function name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $MailboxesToExport = @();

        if($Error.Count -eq 0 )
        {    
            if((test-path $CsvFile))
            {
                LogProgress "`nGetting all delegate mailboxes from CSV [$csvFile]";
                $csv = Import-Csv $CsvFile;
                $count =0;

                $Delegates = @();

                foreach($row in $csv)
                {
                    $user = $row.PrimarySmtpAddress.ToLower();

                    if([string]::IsNullOrEmpty($row.Delegates))
                    {
                        continue;
                    }

                    #Write-Host $row.Delegates;

                    $DelegatesArray = $row.Delegates.Split(";", [StringSplitOptions]::RemoveEmptyEntries);
            
                    foreach($Delegate in $DelegatesArray)
                    {              
                        if(! $Delegates.Contains($Delegate.ToLower()))
                        {
                            $Delegates+=$Delegate.Tolower();                
                        }
              
                    }                      
                }

                if($Delegates.Count -gt 0)
                {
                    LogSuccess "Total $($Delegates.Count) delegates found";
                    $Total = $Delegates.Count;
                    $count = 0;
                    foreach($Delegate in $Delegates)
                    {
                        $count++;

                        LogProgress "`n[$count\$Total] Processing delegate $Delegate";

                        $mailbox= $null;
                        $mailbox = Get-Mailbox $Delegate -ErrorAction:SilentlyContinue;    

                        if($mailbox)
                        {
                            $mailboxObj = New-Object PSObject;
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name DelegateAlias -Value $mailbox.Alias;
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name DelegateDisplayName -Value $mailbox.DisplayName;
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name DelegateEmail -Value $mailbox.PrimarySmtpAddress; 
                            $mailboxObj | Add-Member -MemberType NoteProperty -Name UserEmail -Value $Delegate.PrimarySmtpAddress; 
                
                            LogInfo "Getting delegate [$($mailbox.Identity)] details";
                            $delegateUser = Get-User $mailbox.Identity ;
                        
                            $managerDn  = $delegateUser.Manager;
                            $managerUser = $null;

                            if(! [string]::IsNullOrEmpty($managerDn))
                            {
                                LogInfo "Getting manager [$managerDn] details";
                                $managerUser = Get-Mailbox $managerDn;                        
                                if($managerUser)
                                {
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.Name; 
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.PrimarySmtpAddress;                             
                                }
                                else
                                {
                                    $managerUser = Get-User $managerDn;                        
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.Name; 
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.WindowsEmailAddress;                             
                                }

                                $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerPath -Value $managerDn; 
                            }
                    

                            $MailboxesToExport+= $mailboxObj;
                        }
                        else
                        {
                            LogInfo "Mailbox could not be found for delegate '$Delegate'";
                        }
                    }

                    if($MailboxesToExport.Count -gt 0)
                    {
                        $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

                        LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
                    }           
                }   
                else
                {
                    LogSuccess "`nNo delegates found in csv file";
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
    }
}
##CsvFile: Input csv should be Output of ExportMailboxWithPermissions script
function ExportDelegatesAndManagers2
{
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

    process 
    {   
        ##usage:
        ## ExportDelegatesAndManagers2 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'delegates.csv'

        ##Taking function name from PS environment  

        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();
        

        ConnectToExchange;

        $MailboxesToExport = @();

        if($Error.Count -eq 0 )
        {    
            if((test-path $CsvFile))
            {
                LogProgress "`nGetting all delegate mailboxes from CSV [$csvFile]";
                $csv = Import-Csv $CsvFile;
                $count =0;
                $total = 1;
                if($csv.Count -ne $null)
                {
                    $total = $csv.Count;
                }

                if($total -gt 0)
                {
                    $Delegates = @{};            
        
                    $count = 0;
                    foreach($row in $csv)
                    {
                        $user = $row.PrimarySmtpAddress.ToLower();

                        if([string]::IsNullOrEmpty($row.Delegates))
                        {
                            continue;
                        }            
                
                        $DelegatesArray = $row.Delegates.Split(";", [StringSplitOptions]::RemoveEmptyEntries);
            
                        foreach($delegate in $DelegatesArray)
                        {                
                            $count++;

                            LogProgress "`n[$count\$Total] Processing delegate [$delegate] for User [$User]";   
                    
                            $mailboxObj = $null

                            if($Delegates.ContainsKey($delegate.ToLower()))
                            {
                                $existingMailbox = $Delegates[$delegate.ToLower()];
                                
                                $mailboxObj = New-Object PSObject -Property @{
                                    UserEmail       = $User
                                    DelegateAlias             = $existingMailbox.DelegateAlias
                                    DelegateDisplayName       = $existingMailbox.DelegateDisplayName
                                    DelegateEmail     = $existingMailbox.DelegateEmail                            
                                    ManagerName       = $existingMailbox.ManagerName
                                    ManagerEmail = $existingMailbox.ManagerEmail
                                    ManagerPath =$existingMailbox.ManagerPath
                                }
                            }
                            else
                            {
                                $mailbox= $null;
                                $mailbox = Get-Mailbox $delegate -ErrorAction:SilentlyContinue;    
                                if($mailbox)
                                {
                                    $mailboxObj = New-Object PSObject;
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name UserEmail -Value $User; 
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name DelegateAlias -Value $mailbox.Alias;
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name DelegateDisplayName -Value $mailbox.DisplayName;
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name DelegateEmail -Value $mailbox.PrimarySmtpAddress;                            
                                        
                
                                    LogInfo "Getting delegate [$($mailbox.Identity)] details";
                                    $delegateUser = Get-User $mailbox.Identity ;
                        
                                    $managerDn  = $delegateUser.Manager;
                                    $managerUser = $null;

                                    if(! [string]::IsNullOrEmpty($managerDn))
                                    {
                                        LogInfo "Getting manager [$managerDn] details";
                                        $managerUser = Get-Mailbox $managerDn;                        
                                        if($managerUser)
                                        {
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.Name; 
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.PrimarySmtpAddress;                             
                                        }
                                        else
                                        {
                                            $managerUser = Get-User $managerDn;                        
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.Name; 
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.WindowsEmailAddress;                             
                                        }

                                        $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerPath -Value $managerDn; 
                                    }
                    
                                    $Delegates.Add($delegate.ToLower(), $mailboxObj);                        
                                }
                                else
                                {
                                    LogInfo "Mailbox could not be found for delegate '$delegate'";
                                }
                            }

                            if($mailboxObj -ne $null)
                            {                    
                                $MailboxesToExport+= $mailboxObj;
                            }

                            #Read-Host "$count";
                        }                       
                    }   
        
        
                    if($MailboxesToExport.Count -gt 0)
                    {
                        $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

                        LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
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
    }
}
##CsvFile: Input csv should be Output of ExportMailboxWithPermissions script
function ExportDelegatesAndManagers3
{
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

    process 
    {   
        ##usage:
        ## ExportDelegatesAndManagers3 -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile 'delegates.csv'

        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $MailboxesToExport = @();

        if($Error.Count -eq 0 )
        {    
            if((test-path $CsvFile))
            {
                LogProgress "`nGetting all delegate mailboxes from CSV [$csvFile]";
                $csv = Import-Csv $CsvFile;
                $count =0;
                $total = 1;
                if($csv.Count -ne $null)
                {
                    $total = $csv.Count;
                }

                if($total -gt 0)
                {
                    $Delegates = @{};        
        
                    $count = 0;
                    foreach($row in $csv)
                    {
                        $user = $row.PrimarySmtpAddress.ToLower();

                        if([string]::IsNullOrEmpty($row.Delegates))
                        {
                            continue;
                        }            
            
                        $DelegatesArray = $row.Delegates.Split(";", [StringSplitOptions]::RemoveEmptyEntries);
            
                        foreach($d in $DelegatesArray)
                        {                
                            $count++;

                            LogProgress "`n[$count\$Total] Processing delegate [$d] for User [$User]";   
                    
                            $mailboxObj = $null

                            if($Delegates.ContainsKey($d.ToLower()))
                            {
                                #$Delegates+=$d.Tolower();                

                                $existingMailbox = $Delegates[$d.ToLower()];
                                
                                $mailboxObj = New-Object PSObject -Property @{
                                    UserEmail       = $User                            
                                    ManagerName       = $existingMailbox.ManagerName
                                    ManagerEmail = $existingMailbox.ManagerEmail                            
                                }
                            }
                            else
                            {
                                $mailbox= $null;
                                $mailbox = Get-Mailbox $d -ErrorAction:SilentlyContinue;    
                                if($mailbox)
                                {
                                    $mailboxObj = New-Object PSObject;
                                    $mailboxObj | Add-Member -MemberType NoteProperty -Name UserEmail -Value $User;                             
                
                                    LogInfo "Getting delegate [$($mailbox.Identity)] details";
                                    $delegateUser = Get-User $mailbox.Identity ;
                    
                                    $managerDn  = $delegateUser.Manager;
                                    $managerUser = $null;

                                    if(! [string]::IsNullOrEmpty($managerDn))
                                    {
                                        LogInfo "Getting manager [$managerDn] details";
                                        $managerUser = Get-Mailbox $managerDn;                        
                                        if($managerUser)
                                        {
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.DisplayName; 
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.PrimarySmtpAddress;                             
                                        }
                                        else
                                        {
                                            $managerUser = Get-User $managerDn;                        
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerName -Value $managerUser.DisplayName; 
                                            $mailboxObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value  $managerUser.WindowsEmailAddress;                             
                                        }                                
                                    }
                    
                                    $Delegates.Add($d.ToLower(), $mailboxObj);                        
                                }
                                else
                                {
                                    LogInfo "Mailbox could not be found for delegate '$d'";
                                }
                            }

                            if($mailboxObj -ne $null)
                            {                    
                                $MailboxesToExport+= $mailboxObj;
                            }

                            #Read-Host "$count";
                        }                       
                    }   
        
        
                    if($MailboxesToExport.Count -gt 0)
                    {
                        $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

                        LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
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
    }
}

function DeleteEmails
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $true)]
        [string]$Received,
        
        [Parameter(Mandatory = $true)]
        [string]$Sent
    )

    process 
    {          
        ##Usage:
        ## DeleteEmails  -User $User -Password $Password -ExchangeServer $ExchangeServer  -Mailbox user@domain.com -Sent "Wednesday, March 17, 2020 10:06:35 PM"  -Received "Wednesday, March 17, 2020 10:06:35 PM"

        ##Taking function name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();
        

        ConnectToExchange;

        $PFsToExport = @{};

        if($Error.Count -eq 0 )
        {
            LogProgress "`nVerifying mailbox [$Mailbox]";

	        $mailboxObject = Get-Mailbox $Mailbox -ErrorAction:SilentlyContinue;
    
	        $Error.Clear();

            $Mailboxes = @{};
            if($mailboxObject)
            {    
                LogSuccess "`nFound mailbox $($mailboxObject.DisplayName)";

                LogProgress "`nSearching mailbox [$Mailbox] for Received [$Received] and Sent [$Sent]";

                $mailboxObject | Search-mailbox -SearchQuery {(Received -lt $Received) -or (Sent -lt $Sent)} -DeleteContent -Force

                if($Error.Count -eq 0)
                {
                    LogSuccess "Command executed successfully.";
                }
                else
                {
                    LogError "`nError. $Error`n";
                }        
            }
            else
            {
                LogError "`nCould not find $mailbox. $Error`n";
            }
        }
    }
}

function ExportPFPermissions
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {           
        ##Usage:
        ## ExportPFPermissions -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13
        
        ##Taking function name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;
        
        $PFsToExport = @{};

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all public folders";

    	    #$PFs = Get-PublicFolder \ -recurse -ResultSize Unlimited -ErrorAction:SilentlyContinue;
            $PFs = Get-PublicFolder -GetChildren|   Get-PublicFolder -recurse -ResultSize Unlimited -ErrorAction:SilentlyContinue;
    
    	    $Error.Clear();

            $Mailboxes = @{};
            if($PFs.Count -gt 0)
            {    
                LogSuccess "`nFound $($PFs.Count) public folders.";        

                LogProgress "`nGetting folder permissions";    
       
                $Error.Clear();

                $Stats = $PFs  |%{Get-PublicFolderClientPermission $_.Identity | ?{$_.User -ne 'Default' -and $_.User -ne 'Anonymous'}}
            
                if($Error.Count -eq 0)
                {     
                    $count = 1;
                    if($Stats.Count -ne $null)
                    {
                        $count = $Stats.Count;
                    }

                    LogSuccess "`nFound ($Count) public folder permissions";            

                    $Stats | Select @{Label = "PublicFolder";Expression = { $_.Identity}},@{Label = "User";Expression = {  $User = $_.User;if($Mailboxes.ContainsKey($User))  { ($Mailboxes[$User]).PrimarySmtpAddress} else { $Mbx =Get-Mailbox $User -ErrorAction:SilentlyContinue;if($Mbx){$junk= $Mailboxes.Add($User, $Mbx); $Mbx.PrimarySmtpAddress;}}}}, @{Label = "Permission";Expression = { $_.AccessRights}} |Export-Csv $OutputFile -NoTypeInformation;                        
                }    
                else
                {
                    LogError "$Error";
                }
            }
            else
            {
                LogError "`nCould not get groups. $Error`n";
            }
        }     
    }
}

function ExportPFStats
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
    
        [Parameter(Mandatory = $true)]
        [string]$Password,
    
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {   
        ##usage:
        ##ExportPFStats -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13

        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $PFsToExport = @{};

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all public folders";

	        ##Getting all PFs excluding top level root / PF
            
            $PFs = Get-PublicFolder -GetChildren|   Get-PublicFolder -recurse -ResultSize Unlimited -ErrorAction:SilentlyContinue;
    
	        $Error.Clear();

            if($PFs.Count -gt 0)
            {    
                LogSuccess "`nFound $($PFs.Count) public folders.`n";

                foreach($pf in $PFs)
                {
                    $PFsToExport.Add($pf.EntryId, $pf);            
                }

                LogProgress "`nGetting folder stats";    
       
                $Error.Clear();

                $Stats = $PFs| Get-PublicFolderStatistics  #-ResultSize Unlimited 
        
                if($Error.Count -eq 0)
                {     
                    $count = 1;
                    if($Stats.Count -ne $null)
                    {
                        $count = $Stats.Count;
                    }

                    LogSuccess "`nFound ($Count) public folder stats`n";
                    #Name, Path, ChildrenPFCount, NumberOfItems,Size

                    if($Count -gt 0)
                    {         
                        $Stats | Select Name,@{Label = "Path";Expression = { ($PFsToExport[$_.EntryId]).Identity}} ,
                        @{Label = "FolderType";Expression = { if(($PFsToExport[$_.EntryId]).FolderType -eq $null){"IPF.Note"}else{($PFsToExport[$_.EntryId]).FolderType}}}, ItemCount,
                        @{Label = "Size (KB)";Expression = { $_.TotalItemSize.Value.ToKB()}},LastUserAccessTime |Export-Csv $OutputFile -NoTypeInformation;
                    }
                }    
                else
                {
                    LogError "$Error";
                }
            }
            else
            {
                LogError "`nCould not get groups. $Error`n";
            }
        }    
    }
}

function ExportMailboxWithPermissions
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $false)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $false)]
        [string]$CsvFile
    )

    process 
    {   
        ##usage:
        ## ExportMailboxWithPermissions -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -CsvFile '.\Input\mailboxes.csv'
        ## ExportMailboxWithPermissions -User sp3\administrator -Password ok -ExchangeServer 192.168.10.13 -EmailAddress test1@sp3.local
        
        ##Taking function name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name        

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $MailboxesToExport = @();

        if($Error.Count -eq 0 )
        {    
            if(-not [string]::IsNullOrEmpty($EmailAddress))
            {
                LogInfo "Working for single user '$EmailAddress'";
                $mailbox = Get-Mailbox $EmailAddress -ErrorAction:SilentlyContinue;    
                if($mailbox -ne $null)
                {
                    $mailboxObj = New-Object PSObject;
                    $mailboxObj | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias;
                    $mailboxObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName;
                    $mailboxObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress;                
                
                    $Delegates = "";
                    $permissions = Get-MailboxPermission $mailbox.PrimarySmtpAddress;
                    foreach($permission in $permissions)
                    {
                        if(! $permission.User.Contains(" ") -and $permission.AccessRights.Contains("FullAccess"))
                        {
                            Write-Host "Processing possible delegate $($permission.User) with permission $($permission.AccessRights)";
                            $delegate = Get-Mailbox $permission.User -ErrorAction:SilentlyContinue;    

                            if($delegate -ne $null)
                            {
                                $Delegates+= $delegate.PrimarySmtpAddress.Tostring()+";";
                            }
                        }
                    }

                    $Delegates= $Delegates.Trim(';');
                    $mailboxObj | Add-Member -MemberType NoteProperty -Name Delegates -Value $Delegates;

                    $MailboxesToExport+= $mailboxObj;   
                }
            }
            elseif((-not [string]::IsNullOrEmpty($CsvFile)) -and (test-path $CsvFile))
            {
                LogProgress "`nGetting all mailboxes from CSV [$CsvFile]";
                
                $csv = Import-Csv $CsvFile;
                $count =0;

                foreach($row in $csv)
                {
                    if([string]::IsNullOrEmpty($row.PrimarySmtpAddress))
                    {
                        continue;
                    }

                    $count++;
                    LogProgress "[$count] Getting mailbox info and permissions for user [$($row.PrimarySmtpAddress)]";

                    $mailbox = Get-Mailbox $row.PrimarySmtpAddress -ErrorAction:SilentlyContinue;    
                    if($mailbox -ne $null)
                    {
                        $mailboxObj = New-Object PSObject;
                        $mailboxObj | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias;
                        $mailboxObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName;
                        $mailboxObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress;                
                
                        $Delegates = "";
                        $permissions = Get-MailboxPermission $mailbox.PrimarySmtpAddress;

                        foreach($permission in $permissions)
                        {
                            if(! $permission.User.Contains(" ") -and $permission.AccessRights.Contains("FullAccess"))
                            {
                                Write-Host "Processing possible delegate $($permission.User) with permission $($permission.AccessRights)";
                                $delegate = Get-Mailbox $permission.User -ErrorAction:SilentlyContinue;    

                                if($delegate -ne $null)
                                {
                                    $Delegates+= $delegate.PrimarySmtpAddress.Tostring()+";";
                                }
                            }
                        }

                        $Delegates= $Delegates.Trim(';');
                            
                        $mailboxObj | Add-Member -MemberType NoteProperty -Name Delegates -Value $Delegates;

                        $MailboxesToExport+= $mailboxObj;   
                    }              
                }
            }
            else
            {
                LogError "Please input either -EmailAddress or -CsvFile `n";
            }

            if($MailboxesToExport.Count -gt 0)
            {
                $MailboxesToExport | Export-CSV $OutputFile -NoTypeInformation;    

                LogSuccess "`nExported $($MailboxesToExport.Count) records to [$OutputFile]`n";                 
            }
        }
        else
        {
            LogError "$Error`n";
        }   
    }
}


function ExportSharedMailboxLogins
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter (Mandatory = $false)]
        [string]$NumberOfDays =30
    )

    process 
    { 
        ##usage:
        Write-Host "new";
        ##To export all mailboxes who have not logged on since last 30 days [30 is the default period here in script]
        ## ExportSharedMailboxLogins -User AD\administrator -Password ok -ExchangeServer 192.168.10.10

        ##To export all mailboxes who have not logged on since last 60 days 
        ## ExportSharedMailboxLogins -User AD\administrator -Password ok -ExchangeServer 192.168.10.10 -NumberOfDays 60
  
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();
        

        ConnectToExchange;

        $MailboxesDic =@{};

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting shared mailbox list";

            $time = (get-date).AddDays(-$NumberOfDays);    

            $Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -ErrorAction:SilentlyContinue 
    
            if($Error.Count -eq 0 )
            {
                foreach($mbx in $Mailboxes)
                {
                    $MailboxesDic.Add($mbx.ExchangeGuid, $mbx.PrimarySmtpAddress);
                }

                LogProgress "`nGetting user logons";    
       
                $Error.Clear();

                $Stats = $Mailboxes| Get-mailboxStatistics | ? {$_.LastLogonTime -lt $time}|Select DisplayName, LastLogonTime,ItemCount, MailboxGuid ; 
        
                if($Error.Count -eq 0)
                {     
                    $count = 1;
                    if($Stats.Count -ne $null)
                    {
                        $count = $Stats.Count;
                    }

                    LogSuccess "`nFound ($Count) mailbox stats`n";
                    if($Count -gt 0)
                    {         
                        $Stats | Select DisplayName,@{Label = "EmailAddress";Expression = { $MailboxesDic[$_.MailboxGuid]}} ,LastLogonTime, ItemCount |Export-Csv $OutputFile -NoTypeInformation;
                    }
                }    
                else
                {
                    LogError "$Error";
                }
            }
            else
            {
                LogError "$Error";
            }
        }
        else
        {
            LogError "$Error";
        }    
    }
}

function ExportUserLogins
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter (Mandatory = $false)]
        [string]$NumberOfDays =30
    )

    process 
    {   
        ##usage:

        ##To export all mailboxes who have not logged on since last 30 days [30 is the default period here in script]
        ## ExportUserLogins -User AD\administrator -Password ok -ExchangeServer 192.168.10.10

        ##To export all mailboxes who have not logged on since last 60 days 
        ## ExportUserLogins -User AD\administrator -Password ok -ExchangeServer 192.168.10.10 -NumberOfDays 60

        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;
        
        $MailboxesDic =@{};

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting mailbox list";

            $time = (get-date).AddDays(-$NumberOfDays);

            $Mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction:SilentlyContinue 
    
            if($Error.Count -eq 0 )
            {
                foreach($mbx in $Mailboxes)
                {
                    $MailboxesDic.Add($mbx.ExchangeGuid, $mbx.PrimarySmtpAddress);
                }

                LogProgress "`nGetting user logons";
    
                $Stats = $Mailboxes| Get-mailboxStatistics | ? {$_.LastLogonTime -lt $time}|Select DisplayName, LastLogonTime,ItemCount, MailboxGuid ;
       
                $Error.Clear();

                LogSuccess "`nFound $($Stats.Count) mailbox stats`n";
                if($Stats.Count -gt 0)
                {         
                    $Stats | Select DisplayName,@{Label = "EmailAddress";Expression = { $MailboxesDic[$_.MailboxGuid]}} ,LastLogonTime, ItemCount |Export-Csv $OutputFile -NoTypeInformation;
                }    
            }
            else
            {
                LogError "$Error";
            }
        }
        else
        {
            LogError "$Error";
        }       
    }
}

function ExportDeletedDLs
{
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$User,
        
        [Parameter(Mandatory = $false)]
        [string]$Password,
        
        [Parameter(Mandatory = $false)]
        [string]$ADServer
    )

    process 
    {   
        ##usage:
        ## ExportDeletedDLs
        ## ExportDeletedDLs -User ad\administrator -Password ok -ADServer 192.168.10.10

        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        #ConnectToExchange;

        $scriptPath = GetScriptPath;
        $AllDLsOutputFolder = "$scriptPath\Output\DeletedDLs";       


        ##\\Output\DeletedDLs folder
        if(-not (Test-Path $AllDLsOutputFolder))
        {
            New-item $AllDLsOutputFolder -ItemType Directory | Out-Null 
        }


        $Error.Clear();

        LogInfo "`nImporting ActiveDirectory module";

        Import-Module ActiveDirectory

        $ADParameters =@{};

        $ADParameters.Add("IncludeDeletedObjects", $true);

        if(![string]::IsNullOrEmpty($User) -and ![string]::IsNullOrEmpty($Password))
        {
            $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force;
            $Credentials = New-Object System.Management.Automation.PSCredential $User, $SecurePassword ;

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

    }
}

function ExportAliasDLsAll
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {   
        ##usage:
        ## ExportAliasDLsAll -User AD\administrator -Password ok -ExchangeServer 192.168.10.10
        
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();
        

        ConnectToExchange;

        $DLsToExport = @();

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all groups";

            $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
            $Error.Clear();
            if($Error.Count -eq 0 )
            {
                if($DLs  -ne $null -and $DLs.Count -gt 0)
                {    
                    LogSuccess "`nFound $($DLs.Count) groups`n";

                    if($DLs.Count -gt 0)
                    {            
                        $DLs |Select Name, Alias,PrimarySmtpAddress| Export-CSV $OutputFile -NoTypeInformation;    

                        LogSuccess "`nExported $($DLs.Count) records to [$OutputFile]`n";         
                    }
                    else
                    {
                        LogSuccess "`nNo DL found`n";
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
    }
}

function ExportDLsAll
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {   
        ##usage:
        ## ExportDLsAll -User AD\administrator -Password ok -ExchangeServer 192.168.10.10
        
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $scriptPath = GetScriptPath;

        $AllDLsOutputFolder = "$scriptPath\Output\AllDLs";


        ##\\Output\AllDLs folder
        if(-not (Test-Path $AllDLsOutputFolder))
        {
            New-item $AllDLsOutputFolder -ItemType Directory | Out-Null 
        }


        $DLsToExport = @();

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all groups";

            $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
            $Error.Clear();

            if($DLs  -ne $null -and $DLs.Count -gt 0)
            {    
                LogSuccess "`nFound $($DLs.Count) groups`n";

                ##Iterating over $csv rows
                $count =0;
                
                foreach($DL  in $DLs)
                {
                    $count++;    

                    $Identity = $DL.PrimarySmtpAddress;
                    LogProgress "$count# Processing group [$($DL.DisplayName) ($Identity)]";  
            
                    $managerDisplayName = $null;
                    $managerEmail = $null;
                    $managerDisplayName = $null;
                    $managerEmail = $null;
             
                    if($DL.ManagedBy -ne $null)##here make sure that conditional operator is -ne
                    {
                        $manager = Get-User $DL.ManagedBy[0];

                        if($manager -ne $null)##here make sure that conditional operator is -ne
                        {
                            $managerDisplayName = $manager.DisplayName;
                            $managerEmail = $manager.WindowsEmailAddress;  
                        }
                    } 
            
                    ##Current DL has memebr(s), so export all members to a csv with group name in it                      
                    $members = Get-DistributionGroupMember $Identity | Select-Object Name,PrimarySMTPAddress,Manager
                    if($members.Count -ne 0)
                    {
                        $path = "$AllDLsOutputFolder\$Identity.csv";      
                        LogSuccess "Exporting DL. Member count: $($members.Count)";
                        $members | Export-Csv $path -Append -NoTypeInformation
                    }                        
            
                    $groupObj = New-Object PSObject;
                    $groupObj | Add-Member -MemberType NoteProperty -Name Name -Value $DL.Name;
                    $groupObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DL.DisplayName;
                    $groupObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $DL.PrimarySmtpAddress;
                    $groupObj | Add-Member -MemberType NoteProperty -Name ManagerDisplayName -Value $managerDisplayName;
                    $groupObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value $managerEmail;

                    $DLsToExport+= $groupObj;                           
                }

                if($DLsToExport.Count -gt 0)
                {            
                    $DLsToExport | Export-CSV $OutputFile -NoTypeInformation;    

                    LogSuccess "`nExported $($DLsToExport.Count) records to [$OutputFile]`n";         
                }
                else
                {
                    LogSuccess "`nNo DL found`n";
                }
            }
            else
            {
                LogError "`nCould not get groups. $Error`n";
            }
        } 
    }
}

function ExportDLs
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {   
        ##usage:
        ## ExportDLs -User AD\administrator -Password ok -ExchangeServer 192.168.10.10
        
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $scriptPath = GetScriptPath;

        $AllDLsOutputFolder = "$scriptPath\Output\AllDLs";

        ##\\Output\AllDLs folder
        if(-not (Test-Path $AllDLsOutputFolder))
        {
            New-item $AllDLsOutputFolder -ItemType Directory | Out-Null 
        }

        $DLsToExport = @();
        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all groups";

            $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
            $Error.Clear();

            if($DLs  -ne $null -and $DLs.Count -gt 0)
            {    
                LogSuccess "`nFound $($DLs.Count) groups`n";

                ##Iterating over $csv rows
                $count =0;
                foreach($DL  in $DLs)
                {
                    $count++;    

                    $Identity = $DL.PrimarySmtpAddress;
                    LogProgress "$count# Processing group [$($DL.DisplayName) ($Identity)]";  
            
                    $managerDisplayName = $null;
                    $managerEmail = $null;
             
                    if($DL.ManagedBy -ne $null)##here make sure that conditional operator is -ne
                    {
                        $manager = Get-User $DL.ManagedBy[0];

                        if($manager -ne $null)##here make sure that conditional operator is -ne
                        {
                            $managerDisplayName = $manager.DisplayName;
                            $managerEmail = $manager.WindowsEmailAddress;  
                        }
                    }
                    else
                    {
                        ##Current DL has no manager, so export all members to a csv with group name in it
                        $path = "$AllDLsOutputFolder\$Identity.csv";                

                        $members = Get-DistributionGroupMember $Identity | Select-Object Name,PrimarySMTPAddress,Manager
                        LogSuccess "Exporting DL-with-no-manager members. Member count: $($members.Count)";
                        $members | Export-Csv $path -Append -NoTypeInformation
                    }            
            
                    $groupObj = New-Object PSObject;
                    $groupObj | Add-Member -MemberType NoteProperty -Name Name -Value $DL.Name;
                    $groupObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DL.DisplayName;
                    $groupObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $DL.PrimarySmtpAddress;
                    $groupObj | Add-Member -MemberType NoteProperty -Name ManagerDisplayName -Value $managerDisplayName;
                    $groupObj | Add-Member -MemberType NoteProperty -Name ManagerEmail -Value $managerEmail;

                    $DLsToExport+= $groupObj;                           
                }

                if($DLsToExport.Count -gt 0)
                {            
                    $DLsToExport | Export-CSV $OutputFile -NoTypeInformation;    
    
                    LogSuccess "`nExported $($DLsToExport.Count) records to [$OutputFile]`n";         
                }
                else
                {
                    LogSuccess "`nNo DL found`n";
                }
            }
            else
            {
                LogError "`nCould not get groups. $Error`n";
            }
        }
    }
}

function RemoveDLWithNoManager
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $false)]
        [string]$CsvFile
    )

    process 
    {   
        ##usage:
        ## RemoveDLWithNoManager -User AD\administrator -Password ok -ExchangeServer 192.168.10.10 -CsvFile "C:\Scripts\File.csv";

        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

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
                            }
                            else
                            {
                                LogSuccess "Manager [$manager] is set.";
                            }                    
                        }

                        $Error.Clear();                        
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
    }
}

function ZeroMemberExportDL
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {   
        ##usage:
        ## ZeroMemberExportDL.ps1 -User AD\administrator -Password ok -ExchangeServer 192.168.10.10

        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();
        

        ConnectToExchange;
        
        $DLsToExport = @();

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all groups";

            $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
            $Error.Clear();

            if($DLs  -ne $null -and $DLs.Count -gt 0)
            {    
                LogSuccess "`nFound $($DLs.Count) groups`n";

                ##Iterating over $csv rows
                $count =0;
                foreach($DL  in $DLs)
                {
                    $count++;    

                    $Identity = $DL.PrimarySmtpAddress;
                    LogProgress "$count# Processing group [$($DL.DisplayName) ($Identity)]";  
                
                    $managerDisplayName = $null;
                    $managerEmail = $null;
             
                    $members = Get-DistributionGroupMember $Identity 
                    if($members.Count -eq 0)
                    {
                        LogSuccess "Exporting DL-with-no-member. Member count: $($members.Count)";
                        $groupObj = New-Object PSObject;
                        $groupObj | Add-Member -MemberType NoteProperty -Name Distro_Name -Value $DL.Alias;
                        $groupObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DL.DisplayName;
                        $groupObj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $DL.PrimarySmtpAddress;                

                        $DLsToExport+= $groupObj;    
                    }                                                          
                }

                if($DLsToExport.Count -gt 0)
                {            
                    $DLsToExport | Export-CSV $OutputFile -NoTypeInformation;    

                    LogSuccess "`nExported $($DLsToExport.Count) records to [$OutputFile]`n";         
                }
                else
                {
                    LogSuccess "`nNo DL found`n";
                }
            }
            else
            {
                LogError "`nCould not get groups. $Error`n";
            }
        }        
    }
}

function VerifyManagers
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )

    process 
    {   
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $DLsWithNoManagers = @();
        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting all groups";

            $DLs = Get-distributionGroup -ResultSize Unlimited -ErrorAction:SilentlyContinue;
            $Error.Clear();

            if($DLs  -ne $null -and $DLs.Count -gt 0)
            {    
                LogSuccess "`nFound $($DLs.Count) groups`n";

                ##Iterating over $csv rows
                $count =0;
                foreach($DL  in $DLs)
                {
                    $count++;    

                    $Identity = $DL.Alias;
                    LogProgress "$count# Processing group [$Identity ($($DL.PrimarySmtpAddress))]";               
                
                    if($DL.ManagedBy -eq $null)
                    {
                        LogInfo "Manager is not set for group [$Identity]";
                        $DLsWithNoManagers+= $DL;
                    }      
                }

                if($DLsWithNoManagers.Count -gt 0)
                {
                    $DLsWithNoManagers |Select Name, DisplayName, PrimarySmtpAddress, DistinguishedName| Export-CSV $OutputFile -NoTypeInformation;    
            
                    LogSuccess "`nExported $($DLsWithNoManagers.Count) records to [$OutputFile]`n";         
                }
                else
                {
                    LogSuccess "`nNo DL found withougt manager`n";
                }
            }
            else
            {
                LogError "`nCould not get groups. $Error`n";
            }
        }       
    }
}

function SearchEmailTrackingLogs
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $true)]
        [string]$CsvFile,
        
        [Parameter(Mandatory = $false)]
        [int]$NumberOfDays = 30
    )

    process 
    {
        $IPsOfInterestArray = @("fe80::59fe:381f:4503:3cb0","192.168.10.13")
        $HostNamesOfInterestArray = @("Ex2010sp3","Ex2010sp3.sp3.local")

        $IPsOfInterest = New-Object System.Collections.ArrayList(,$IPsOfInterestArray);
        $HostNamesOfInterest = New-Object System.Collections.ArrayList(,$HostNamesOfInterestArray);   
        
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();        

        ConnectToExchange;

        $PFsToExport = @{};

        if($Error.Count -eq 0 )
        {
            LogProgress "`nGetting Hub Transport server list";
            $htServers = get-exchangeserver |? {$_.serverrole -match "hubtransport"} |% {$_.name}

            LogSuccess "`nFound $($htServers.count) servers";
    
            $hubServerCount =1;
            if($htServers.Count -ne $null)
            {
                $hubServerCount =$htServers.Count ;
            }
            if($hubServerCount-gt 0)
            {
                $start = (Get-Date).AddDays(-$NumberOfDays);

                #Converting all hostnames to lower case, to use for matching
                #$HostNamesOfInterest= $HostNamesOfInterest.Tolower();

                $SearchedRecords = @();
    
                $count =0;
                $total = 0;

                foreach ($ht in $htServers)
                {
	                LogInfo "`nSearching email logs on server [$ht] with start time [$start]"

	                $SentEmailLogs=  get-messagetrackinglog -Server $ht -Start $start -resultsize unlimited -EventId "DELIVER"  | ? {$_.source -eq "STOREDRIVER" };

                    LogSuccess "`nFound $($SentEmailLogs.count) sent email records from server [$ht]";
            
                    $total += $SentEmailLogs.Count;

                    if($SentEmailLogs.Count -gt 0)
                    {
                        foreach($log in $SentEmailLogs)
                        {      
                            $ClientIp  = $log.ClientIp;
                            $ClientHostname  = $log.ClientHostname;
                            $Time =$log.Timestamp;
                            $Sender = $log.Sender;
                            $count++;

                            if( ($IPsOfInterest -Contains $ClientIp) -or ($HostNamesOfInterest -Contains $ClientHostname.ToLower()) )
                            {
                                LogProgress "[$count/$total]Processing email [ClientIp: $ClientIp] [ClientHostname: $ClientHostname] [Sender: $Sender] [Time: $Time]";
            
                                $emailObj = New-Object PSObject;
                                $emailObj | Add-Member -MemberType NoteProperty -Name ClientIp -Value $ClientIp;
                                $emailObj | Add-Member -MemberType NoteProperty -Name ClientHostname -Value $ClientHostname;
                                $emailObj | Add-Member -MemberType NoteProperty -Name Sender -Value $Sender
                                $emailObj | Add-Member -MemberType NoteProperty -Name Timestamp -Value $Time;
                                $emailObj | Add-Member -MemberType NoteProperty -Name MessageSubject -Value $log.MessageSubject;                        
                                $emailObj | Add-Member -MemberType NoteProperty -Name Recipients -Value $log.Recipients;  
                        
                                $SearchedRecords+=$emailObj;
                            }
                        }
                    }
                }

                if($SearchedRecords.Count -gt 0)
                {
                    $SearchedRecords | Export-CSV $OutputFile -NoTypeInformation;    

                    LogSuccess "`nExported $($SearchedRecords.Count) records to [$OutputFile]`n";                 
                }
            }
        }  
     }
}
function TemplateFuncion
{
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

    process 
    {   
        ##Taking script name from PS environment  
        $scriptName = $MyInvocation.MyCommand.Name;

        $OutputFile = InitializeLoggging -OutputFileName $scriptName

        $Error.Clear();
        

        ConnectToExchange;

        $MailboxesToExport = @();        
    }
}
#endregion Functions End

#Export-ModuleMember "GetScriptPath, RemovePSSession";