$year = "2023"
$selectedMonth = "4"

$selectedMonth = [int]$selectedMonth
$selectedYear = [int]$year

$postDate = Get-Date "4/7/2023"

$intYear = [int]$postDate.Year

if ($intYear -eq $selectedYear) {
	Write-Host "true!"
}

if ([int]$postDate.Year -eq $selectedYear) {
    Write-Host "true!"
}



