$excelData = Import-Excel $budgetPath -WorksheetName "Jul" -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200

$testDate = [string](Get-Date $excelData[0].P1 -Format "MM/dd/yyyy")

$testDate.GetType()
Write-Host $testDate