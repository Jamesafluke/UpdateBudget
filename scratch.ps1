$excelData = Import-Excel $budgetPath -WorksheetName "Jul" -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200

foreach($item in $excelData){
    if($item.P1 -ne $null ){

        Write-Host "$item isn't null."
    }
}