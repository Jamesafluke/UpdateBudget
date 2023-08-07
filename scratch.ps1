$budgetPath = "C:\PersonalMyCode\UpdateBudget\TestBudgetSimple.csv"
$csvPath = "C:\PersonalMyCode\UpdateBudget\rewards.csv"
$outputPath = "C:\PersonalMyCode\UpdateBudget\output"


$bankData = Import-Csv $csvPath
$oldBudgetData = Import-Csv -Path $budgetPath

$aprilBankData = @()

foreach ($item in $bankData){

    $existingExpense = $oldBudgetData | Where-Object { $_.Date -eq $item.'asdf date' -and $_.Amount -eq $item.Amount}
    $isDuplicate = 
    if ($item."asdf date"[0] -eq '8'){
        $newExpense = [PSCustomObject]@{
            Date = $item."asdf date"
            Item = $item.Description
            Method = ""
            Amount = $item.Debit
        }
        $aprilBankData += $newExpense
    }
}

Write-Host $aprilBankData.Date