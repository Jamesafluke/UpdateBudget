# Paths to input and output files
$oldBudgetDataPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$accountHistoryPaths = @(
    "C:\PersonalMyCode\UpdateBudget\AccountHistory.csv",
    "C:\PersonalMyCode\UpdateBudget\AccountHistory (1).csv"
)
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"

# Read the old budget data
$oldBudgetData = Import-Csv $oldBudgetDataPath

# Ask user to choose a month
# $selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
$selectedMonth = "4"
$year = "2023"

# Convert the selected month to an integer
$selectedMonth = [int]$selectedMonth

#Year
$selectedYear = [int]$year

# Iterate through account history files
function InputAccountHistory([String]$accountHistoryPath) {

    Write-Host "`n`nRunning for $accountHistoryPath"

    # Read the account history data
    $accountHistoryData = Import-Csv $accountHistoryPath

    
    # Filter account history data by selected month and specific condition
    $filteredAccountHistory = $accountHistoryData | Where-Object {
        $entry = $_
        $postDate = Get-Date $entry."Post Date"
        
        # Check if the year is 2023 and the month matches
        if ([int]$postDate.Year -ne $selectedYear) {
            return $false
        }   
        if ([int]$postDate.Month -ne $selectedMonth){
            return $false
        }
        return $true
    }
    return $filteredAccountHistory
}

function DeDup($thisMonthExpenses){

    $uniqueExpenses = @()

    foreach ($entry in $thisMonthExpenses) {
        $postDate = $entry."Post Date"
        $debit = [decimal]$entry."Debit"
        
        # Check if there's a matching entry in budget data
        $matchingBudgetEntry = $oldBudgetData | Where-Object { $_."Date" -eq $postDate -and $_."Amount" -eq $debit }
        
        # If no match found, consider it a non-duplicate
        if (-not $matchingBudgetEntry) {
            $newExpense = [PSCustomObject]@{
                Date = $entry."Post Date"
                Item = $entry.Description
                Method = ""
                Category = ""
                Amount = $entry.Debit
            }
            $uniqueExpenses += $newExpense
        }
    }        

    return $uniqueExpenses
}

function Export($uniqueExpenses){
    Write-Host "Export!"
    $uniqueExpenses | Export-Csv $outputPath -NoTypeInformation
}








$thisMonthExpenses = InputAccountHistory($accountHistoryPaths[0])
$thisMonthExpenses += InputAccountHistory($accountHistoryPaths[1])
# $thisMonthExpenses | Export-Csv "C:\PersonalMyCode\UpdateBudget\thisMonthExpenses.csv"

$uniqueExpenses = DeDup($thisMonthExpenses)
# $uniqueExpenses | Export-Csv "C:\PersonalMyCode\UpdateBudget\uniqueexpenses.csv"

Export($uniqueExpenses)




