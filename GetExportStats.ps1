
$requests = Get-MailboxExportRequest
$requestCount =1;
if($requests -eq $null)
{
	$requestCount=0;
}
elseif($requests.Count -ne $null)
{
	$requestCount= $requests.Count;
}

Write-Host "Found $requestCount mailbox export requests";
$Queued = $requests |?{$_.Status -eq 'Queued'};
$InProgress = $requests |?{$_.Status -eq 'InProgress'};
$Completed = $requests |?{$_.Status -eq 'Completed'};

$QueuedCount =1
if($Queued -eq $null)
{
	$QueuedCount=0;
}
elseif($Queued.Count -ne $null)
{
	$QueuedCount= $Queued.Count;
}

$InProgressCount=1;
if($InProgress -eq $null)
{
	$InProgressCount=0;
}
elseif($InProgress.Count -ne $null)
{
	$InProgressCount= $InProgress.Count;
}

$CompletedCount=1;
if($Completed -eq $null)
{
	$CompletedCount=0;
}
elseif($Completed.Count -ne $null)
{
	$CompletedCount= $Completed.Count;
}

Write-Host "Found $QueuedCount queued requets" -ForegroundColor Yellow
Write-Host "Found $InProgressCount in progress requets" -ForegroundColor White 
Write-Host "Found $CompletedCount completed requets`n" -ForegroundColor Cyan

if($QueuedCount -gt 0 -or $InProgressCount -gt 0)
{
	do
	{
		Sleep -Seconds 5;
		$requests = Get-MailboxExportRequest
		$requestCount =1;
		if($requests -eq $null)
		{
			$requestCount=0;
		}
		elseif($requests.Count -ne $null)
		{
			$requestCount= $requests.Count;
		}
		
        $time =Get-Date;
		Write-Host "`n[$time] Found $requestCount mailbox export requests";		
		
		$Completed = $requests |?{$_.Status -eq 'Completed'};
		$InProgress = $requests |?{$_.Status -eq 'InProgress'};
		$Queued = $requests |?{$_.Status -eq 'Queued'};			
		
		$QueuedCount =1
		if($Queued -eq $null)
		{
			$QueuedCount=0;
		}
		elseif($Queued.Count -ne $null)
		{
			$QueuedCount= $Queued.Count;
		}

		$InProgressCount=1;
		if($InProgress -eq $null)
		{
			$InProgressCount=0;
		}
		elseif($InProgress.Count -ne $null)
		{
			$InProgressCount= $InProgress.Count;
		}

		$CompletedCount=1;
		if($Completed -eq $null)
		{
			$CompletedCount=0;
		}
		elseif($Completed.Count -ne $null)
		{
			$CompletedCount= $Completed.Count;
		}

		Write-Host "Found $QueuedCount queued requets" -ForegroundColor Yellow
		Write-Host "Found $InProgressCount in progress requets" -ForegroundColor White 
		Write-Host "Found $CompletedCount completed requets`n" -ForegroundColor Cyan
	}
	while($Queued -or $InProgress);
}

