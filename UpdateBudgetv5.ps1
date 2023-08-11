
# Stuff I need to add:
#Handle Credit and Debit
#Handle Method


# Paths to input and output files
$oldBudgetDataPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$accountHistoryPaths = @(
    "C:\PersonalMyCode\UpdateBudget\AccountHistory.csv",
    "C:\PersonalMyCode\UpdateBudget\AccountHistory (1).csv"
)
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"

$rewardsAccountNumber = "313235393200"
$checkingAccountNumber = "750501095729"

# Read the old budget data
$oldBudgetData = Import-Csv $oldBudgetDataPath

# Ask user to choose a month
$selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
# $selectedMonth = "4"
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

    $count = $filteredAccountHistory.Count
    Write-Host "$accountHistoryPath has $count items."
    return $filteredAccountHistory
}

function DeDup($thisMonthExpenses){

    $uniqueExpenses = @()

    foreach ($entry in $thisMonthExpenses) {

        

        $postDate = $entry."Post Date"
        $debit = [decimal]$entry."Debit"
        $credit = [decimal]$entry."Credit"
        
        # Check if there's a matching entry in budget data
        $matchingBudgetEntry = $oldBudgetData | Where-Object { $_."Date" -eq $postDate -and $_."Amount" -eq $debit }
        
        # If no match found, consider it a non-duplicate
        if (-not $matchingBudgetEntry) {
            
            #Determine method
            $method = $null
            if ($entry."Account Number" -eq $rewardsAccountNumber){
                $method = "Rewards"
            }elseif ($entry."Account Number" -eq $checkingAccountNumber) {
                $method = "Checking"
            }

            #Determine debit or credit.
            $amount = $null
            if ($debit -ne ""){
                $amount = $debit
            }else{
                $amount = $credit * -1
            }

            #Arbitrary exceptions.
            $description = ""
            $category = ""
            if ($entry.Description -eq "PennyMac") {
                $description = "Mortgage"
                $category = "Mortgage"
            }
            if ($entry.Description -eq "Walmart") {
                $category = "Groceries"
            }
            if ($entry.Description -eq "Payson City Debits") {
                $description = "Electricity"
                $category = "Electricity"
            }
            if ($entry.Description -eq "Wasatch Property") {
                $description = "HOA"
                $category = "HOA"
            }
            if ($entry.Description -eq "Maverik") {
                $description = "Gasoline"
                $category = "Gasoline"
            }
            if ($entry.Description -eq "American Funds") {
                $description = "This will become millions"
                $category = "Investment"
            }
            if ($entry.Description -eq "Allstate") {
                $description = "Insurance"
                $category = "Insurance"
            }
            if ($entry.Description -eq "Fast Gas Convenience Store") {
                $description = ""
                $category = ""
            }
            if ($entry.Description -eq "Dep Cloud Bee Direct Deposit") {
                $amount = ""
            }
            if ($entry.Description -eq "Dominion Energy") {
                $description = "Dominion Energy"
                $category = "Dominion"
            }
            if ($entry.Description -eq "Credit Card Payment") {
                $amount = ""
            }            
            if ($entry.Description -eq "Fluckiger") {
                $amount = ""
            }
            if ($entry.Description -eq "Costa Vida") {
                $description = ""
                $category = "Eating Out/Fun"
            }
            if ($entry.Description -eq "Xfinity Mobile") {
                $description = "Phones"
                $category = "Phones"
            }
            if ($entry.Description -eq "YouTube Premium") {
                $description = "YouTube Premium"
                $category = "YouTube Premium"
            }


            $newExpense = [PSCustomObject]@{
                Date = $entry."Post Date"
                Item = $entry.Description
                Description = $description
                Method = $method
                Category = $category
                Amount = $amount
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
$count = $thisMonthExpenses.Count
Write-Host "Importing a total of $count items."

$uniqueExpenses = DeDup($thisMonthExpenses)

$count = $oldBudgetDataPath.Count
Write-Host "Existing budget has $count items."

$count = $uniqueExpenses.Count
Write-Host "Grand total of $count items."


Export($uniqueExpenses)




