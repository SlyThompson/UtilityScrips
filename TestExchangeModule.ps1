
$User = "sp3\administrator";
$Password = "ok";
$ExchangeServer = "192.168.10.13";


#Import-Module "D:\Projects\Powershell\Ex\DG ManagedBy (Sylvester)\Scripts\ExchangeModule.psm1";
Import-Module ".\ExchangeModule.psm1";

#ExportPFPermissionsForSenders -User $User -Password $Password -ExchangeServer $ExchangeServer  -CsvFile '.\Input\ExportManagersForSenders-05-06-2020-02.csv'
#ExportManagersForSenders      -User $User -Password $Password -ExchangeServer $ExchangeServer  -CsvFile '.\Input\SearchEmailTrackingLogs-30-05-2020-02.csv'
#ExportDelegatesAndManagers3   -User $User -Password $Password -ExchangeServer $ExchangeServer  -CsvFile '.\Input\ExportMailboxWithPermissions-18-04-2020-04.csv'
#ExportDelegatesAndManagers2   -User $User -Password $Password -ExchangeServer $ExchangeServer  -CsvFile '.\Input\ExportMailboxWithPermissions-18-04-2020-04.csv' #ExportMailboxWithPermissions-18-04-2020-05
#ExportDelegatesAndManagers   -User $User -Password $Password -ExchangeServer $ExchangeServer  -CsvFile '.\Input\ExportMailboxWithPermissions-18-04-2020-04.csv' #ExportMailboxWithPermissions-18-04-2020-05

#DeleteEmails  -User $User -Password $Password -ExchangeServer $ExchangeServer  -Mailbox test1@sp3.local -Sent "Wednesday, March 17, 2020 10:06:35 PM"  -Received "Wednesday, March 17, 2020 10:06:35 PM"

#ExportPFPermissions   -User $User -Password $Password -ExchangeServer $ExchangeServer  
#ExportPFStats   -User $User -Password $Password -ExchangeServer $ExchangeServer  

#ExportMailboxWithPermissions -User $User -Password $Password -ExchangeServer $ExchangeServer -CsvFile '.\Input\mailboxes.csv'

#ExportSharedMailboxLogins -User $User -Password $Password -ExchangeServer $ExchangeServer -NumberOfDays 90
# ExportUserLogins -User $User -Password $Password -ExchangeServer $ExchangeServer -NumberOfDays 90

#ExportUserLogins -User $User -Password $Password -ExchangeServer $ExchangeServer -NumberOfDays 90

#ExportDeletedDLs -User $User -Password $Password -ADServer $ExchangeServer

#ExportAliasDLsAll -User $User -Password $Password -ExchangeServer $#ExchangeServer
#ExportDLsAll -User $User -Password $Password -ExchangeServer $ExchangeServer
#ExportDLs -User $User -Password $Password -ExchangeServer $ExchangeServer

#RemoveDLWithNoManager -User $User -Password $Password -ExchangeServer $ExchangeServer #-CsvFile "C:\Scripts\File.csv";
#ZeroMemberExportDL -User $User -Password $Password -ExchangeServer $ExchangeServer 

#ZeroMemberExportDL -User $User -Password $Password -ExchangeServer $ExchangeServer 

VerifyManagers -User $User -Password $Password -ExchangeServer $ExchangeServer 
Remove-Module "ExchangeModule";