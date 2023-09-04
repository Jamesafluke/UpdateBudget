
$budgetxlsxPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"

$abbMonthName = "Jul"


$excelData = Import-Excel $budgetxlsxPath -WorksheetName $abbMonthName -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200

foreach($item in $excelData){
    if($null -ne $item.P1 ){

        Write-Host "$item isn't null."
    }
}



