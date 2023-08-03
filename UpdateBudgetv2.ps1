# Define paths
# $budgetPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
$budgetPath = "C:\Users\james\OneDrive\Budget\TestBudget.xlsx"
$csvPath = "C:\Users\james\Downloads\rewards.csv"
$outputPath = "C:\Users\james\OneDrive\Budget\output.csv"

# Load ImportExcel module
Import-Module -Name ImportExcel

#Get month.
$userInput = read-Host "Provide the number of the month"
$monthNames = @("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
$month = $monthNames[$userInput - 1]
Write-Host $month


# Function to compare expenses and update spreadsheet
function Update-Budget($csvPath) {
    Write-Host "Import csv."
    $csvData = Import-Csv $csvPath | Where-Object { $_.Credit -eq "" }
    
    Write-Host "Import data from budget."
    $worksheet = Import-Excel -Path $budgetPath -WorksheetName $month
    # Filter data from columns S through W, starting from row 8 and below
    $dataToExport = $mayWorksheet | Select-Object -Skip 7 | Select-Object S,W
    
    $updatedBudgetData = @()

    foreach ($item in $csvData) {
        $existingExpense = $budgetData | Where-Object { $_.Date -eq $item.Date -and $_.Amount -eq $item.Amount }
        $isDuplicate = $existingExpense -ne $null

        if (-not ($isDuplicate)) {
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
Update-Budget -csvPath $csvPath
