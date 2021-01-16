param(
    [parameter(Mandatory=$true,Position=0)]
    [String]$webUrl,
    [parameter(Mandatory=$true,Position=1)]
    [string]$calendarListName
)

# Install PnP PowerShell Module.
#Install-Module SharePointPnPPowerShellOnline

# Get Japan holidays list (CSV) from Cabinet Office in JAPAN
$holidayList = Invoke-WebRequest -uri "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
$holidayListUTF8 = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetString( [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($holidayList.Content))

Connect-PnPOnline -url $webUrl -UseWebLogin

#If no exsisting calendar list, creat new one
#New-PnPList -Title "<CalendarListName>" -Url "lists/jpnholidaycalendar" -Template Events

$currentYear = Get-Date -Format "yyyy"
foreach($line in $holidayListUTF8 -split "`r`n")
{
    if($line.StartsWith($currentYear))
    {
        [datetime]$date = $line.split(",")[0]
        $title = $line.split(",")[1]
        $startDate = $date.ToString("yyyy/MM/dd 00:00")
        $endDate = $date.ToString("yyyy/MM/dd 23:59")
        
        $result = Add-PnPListItem -list $calendarListName -Values @{ "Title"=$title; "EventDate"=$startDate; "EndDate"=$endDate; "fAllDayEvent"=$True}

        Write-Host $title, $date.ToString("yyyy/MM/dd")
    }
}
