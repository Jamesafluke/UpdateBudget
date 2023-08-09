# Define paths
# $budgetPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
# $budgetPath = "C:\Users\james\OneDrive\Budget\TestBudget.xlsx"
# $csvPath = "C:\Users\james\Downloads\rewards.csv"
# $outputPath = "C:\Users\james\OneDrive\Budget\output.csv"

# $budgetPath = "C:\PersonalMyCode\UpdateBudget\TestBudget.xlsx"
$budgetPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$csvPath = "C:\PersonalMyCode\UpdateBudget\AccountHistory.csv"
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.xlsx"


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
    
    $newExpenses = @()
    $counter = 0

    foreach ($bankEntry in $bankData) {

        

        $existingExpense = $oldBudgetData | Where-Object { $_."Date" -eq $bankEntry."Post Date"  -and $_."Amount" -eq $bankEntry.Debit}
        
        if($null -ne $existingExpense){
            $isDuplicate = $true
        }

        if (-not $isDuplicate) {
            $newExpense = [PSCustomObject]@{
                Date   = $bankEntry."Post Date"
                Item   = $bankEntry.Description
                Method = ""
                Amount = $bankEntry.Debit
            }
            # Write-Host "Item date: $item.PostDate"
            $newExpenses += $newExpense
        }
        $counter ++
    }
    foreach ($item in $newExpenses){
        Write-Host $item
    }
    Write-Host "Number of new expenses is " $newExpenses.Count


    if ($newExpenses.Count -gt 0) {
        $newExpenses | Export-Excel -Path $outputPath
        Write-Host "Done."
    } else {
        Write-Host "No new expenses to update."
    }
}

# Update expenses
Update-Budget -csvPath $csvPath
