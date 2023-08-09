# Paths to input and output files
$oldBudgetDataPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$accountHistoryPaths = @(
    "C:\Users\james\Downloads\AccountHistory.csv",
    "C:\Users\james\Downloads\AccountHistory (1).csv"
)
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"

# Read the old budget data
$oldBudgetData = Import-Csv $oldBudgetDataPath

# Ask user to choose a month
$selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
$year = "2023"

# Convert the selected month to an integer
$selectedMonth = [int]$selectedMonth

#Year
$selectedYear = [int]$year

# Iterate through account history files
foreach ($accountHistoryPath in $accountHistoryPaths) {
    # Read the account history data
    $accountHistoryData = Import-Csv $accountHistoryPath
    
    # Filter account history data by selected month and specific condition
    $filteredAccountHistory = $accountHistoryData | Where-Object {
        $entry = $_
        $postDate = Get-Date $entry."Post Date"

        # Check if the year is 2023 and the month matches
        if ($postDate.Year -ne $selectedYear -or $postDate.Month -ne $selectedMonth) {
            Write-Host "postDate.Year is $postDate.Year so returning false" 
            return $false
        }
        
        return $true
    }
    
    # Filter out duplicates and write to the output file
    $uniqueExpenses = $filteredAccountHistory | Where-Object {
        $entry = $_
        -not ($oldBudgetData | Where-Object { $_.date -eq $entry."Post Date" -and $_.amount -eq [decimal]$entry.Debit })
    }
    $uniqueExpenses | Export-Csv $outputPath -Append -NoTypeInformation
}
