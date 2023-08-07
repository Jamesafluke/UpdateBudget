# Define paths
# $budgetPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
# $budgetPath = "C:\Users\james\OneDrive\Budget\TestBudget.xlsx"
# $csvPath = "C:\Users\james\Downloads\rewards.csv"
# $outputPath = "C:\Users\james\OneDrive\Budget\output.csv"

# $budgetPath = "C:\PersonalMyCode\UpdateBudget\TestBudget.xlsx"
$budgetPath = "C:\PersonalMyCode\UpdateBudget\TestBudgetSimple.csv"
$csvPath = "C:\PersonalMyCode\UpdateBudget\rewards.csv"
$outputPath = "C:\PersonalMyCode\UpdateBudget\output"


# Load ImportExcel module
Import-Module -Name ImportExcel

# #Get month.
# $userInput = read-Host "Provide the number of the month"
# $monthNames = @("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
# $month = $monthNames[$userInput - 1]
# Write-Host $month
$month = "Sheet1"


# Function to compare expenses and update spreadsheet
function Update-Budget($csvPath) {
    $bankData = Import-Csv $csvPath
    $oldBudgetData = Import-Csv -Path $budgetPath    
    
    $updatedoldBudgetData = @()

    foreach ($item in $bankData) {
        # $existingExpense = $oldBudgetData | Where-Object { $_."Date" -eq $item.PostDate -and $_.Amount -eq $item.Amount }\
        # Write-Host "bankData date is $item.date"
        Write-Host "oldBudgetData date is " $_.date
        $existingExpense = $oldBudgetData | Where-Object { $_."Date" -eq $item.Date }
        $isDuplicate = $null -ne $existingExpense

        if (-not ($isDuplicate)) {
            $newExpense = [PSCustomObject]@{
                Date   = $item.Date
                Item   = $item.Description
                Method = ""
                Amount = $item.Amount
            }
            # Write-Host "Item date: $item.PostDate"
            $updatedBudgetData += $newExpense
        }
    }

    if ($updatedoldBudgetData.Count -gt 0) {
        # Export-Excel -Path $budgetPath -WorkSheetName (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($currentMonth) -AutoSize -Append -TableName "Expenses" -InputObject $updatedoldBudgetData
        $updatedoldBudgetData | Export-Excel -Path $outputPath
        Write-Host "Done."
    } else {
        Write-Host "No new expenses to update."
    }
}

# Update expenses
Update-Budget -csvPath $csvPath
