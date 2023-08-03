# Define paths
# $budgetPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
$budgetPath = "C:\Users\james\OneDrive\Budget\TestBudget.xlsx"
$rewardsCSVPath = "C:\Users\james\Downloads\rewards.csv"

# Load ImportExcel module
Import-Module -Name ImportExcel

# Get current and previous months
$currentMonth = (Get-Date).Month
$previousMonth = if ($currentMonth -eq 1) { 12 } else { $currentMonth - 1 }

# Function to compare expenses and update spreadsheet
function Update-Budget($csvPath) {
    $csvData = Import-Csv $csvPath | Where-Object { $_.Credit -eq "" }

    $budgetData = Import-Excel -Path $budgetPath -WorksheetName (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($currentMonth) -StartRow 8

    $previousMonthBudgetData = Import-Excel -Path $budgetPath -WorksheetName (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($previousMonth) -StartRow 8

    $updatedBudgetData = @()

    foreach ($item in $csvData) {
        $existingExpense = $budgetData | Where-Object { $_.Date -eq $item.Date -and $_.Amount -eq $item.Amount }
        $isDuplicate = $existingExpense -ne $null

        $previousMonthExpense = $previousMonthBudgetData | Where-Object { $_.Date -eq $item.Date -and $_.Amount -eq $item.Amount }
        $isPreviousMonthDuplicate = $previousMonthExpense -ne $null

        if (-not ($isDuplicate -or $isPreviousMonthDuplicate)) {
            $newExpense = [PSCustomObject]@{
                Date   = $item.Date
                Item   = $item.Description
                Method = "Rewards"
                Amount = $item.Amount
            }
            $updatedBudgetData += $newExpense
        }
    }

    if ($updatedBudgetData.Count -gt 0) {
        Export-Excel -Path $budgetPath -WorkSheetName (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($currentMonth) -AutoSize -Append -TableName "Expenses" -InputObject $updatedBudgetData
        Write-Host "Updated expenses in $((Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($currentMonth))!"
    } else {
        Write-Host "No new expenses to update."
    }
}

# Update expenses
Update-Budget -csvPath $rewardsCSVPath
