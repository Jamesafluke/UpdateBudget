Import-Module ImportExcel

# Paths to input and output files

$budgetPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
$oldBudgetDataPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$accountHistoryPaths = @(
    "C:\PersonalMyCode\UpdateBudget\AccountHistory.csv",
    "C:\PersonalMyCode\UpdateBudget\AccountHistory (1).csv"
)
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"

$rewardsAccountNumber = "313235393200"
$checkingAccountNumber = "750501095729"

# Ask user to choose a month
# $selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
$selectedMonth = "7"
$months=@("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
$monthName = $months[$selectedMonth -1]
Write-Host "Selected: $monthName"
$year = "2023"

# Convert the selected month to an integer
$selectedMonth = [int]$selectedMonth
$selectedYear = [int]$year


function InputExistingBudgetData(){
    $excelData = Import-Excel $budgetPath -WorksheetName "Jul" -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200

    #Remove blank items.
    $refinedData = ""
    foreach($item in $excelData){
        if($null -ne $item.P1){

            $nonBlankExpense = [PSCustomObject]@{
                Date = [string](Get-Date $item.P1 -Format "MM/dd/yyyy")
                Item = $item.P2
                Description = $item.P3
                Method = $item.P4
                Category = $item.P5
                Amount = [decimal]$item.P6
            }
        $refinedData += $nonBlankExpense
        }
    }
}

#Convert date to date
#Convert double to decimal


# Iterate through account history files
function InputAccountHistory() {
    $combinedCsv = ""
    $accountHistory0 = Import-Csv $accountHistoryPaths[0]
    $accountHistory1 = Import-Csv $accountHistoryPaths[1]
    $combinedCsv = $accountHistory0 + $accountHistory1
    # $combinedCsv = @($accountHistory0, $accountHistory1)

    Write-Host "Both csvs combined have $($combinedCsv.Count) items total."
    $filteredAccountHistory = ""
    
    # Filter account history data by selected month and specific condition
    $filteredAccountHistory = $combinedCsv | Where-Object {
        $entry = $_
        $postDate = Get-Date $entry."Post Date"
        
        # Check if the year is 2023 and the month matches
        if ([int]$postDate.Year -ne $selectedYear) {
            # Write-Host "Year doesn't match $selectedYear"
            return $false
        }   
        if ([int]$postDate.Month -ne $selectedMonth){
            # Write-Host "Month doesn't match $selectedMonth"
            return $false
        }
        return $true
    }
    Write-Host "Selected month has $($filteredAccountHistory.Count) items."
    return $filteredAccountHistory
}

function Deduplicate($thisMonthExpenses){

    #Remove the dollar sign and whitespace
    $oldBudgetData = Import-Csv $oldBudgetDataPath
    foreach ($entry in $oldBudgetData){
        $entry.Amount = $entry.Amount.Replace('$', '')
        $entry.Amount = $entry.Amount.Replace(' ', '')
    }

    Write-Host "Existing budget data has $($oldBudgetData.Count) items."
    $uniqueExpenses = @()

    foreach ($entry in $thisMonthExpenses) {
        $postDate = $entry."Post Date"
        $debit = [decimal]$entry."Debit"
        $credit = [decimal]$entry."Credit"
        
        
        # Check if there's a matching entry in old budget data
        $matchingBudgetEntry = $oldBudgetData | Where-Object { $_.Date -eq $postDate -and [decimal]$_.Amount -eq $debit}

        # If no match found, consider it a non-duplicate
        if (-not $matchingBudgetEntry) {
            
            #Determine method.
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
                $description = "Car Insurance"
                $category = "Car Insurance"
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
            if ($entry.Description -eq "Costco Gas") {
                $description = "Gasoline"
                $category = "Gasoline"
            }
            if ($entry.Description -eq "Comcast") {
                $description = "Internet"
                $category = "Internet"
            }
            if ($entry.Description -eq "Chevron") {
                $description = "Gasoline"
                $category = "Gasoline"
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


Write-Host "Starting!"

InputExistingBudgetData

$thisMonthExpenses = InputAccountHistory

if($thisMonthExpenses -ne $null){
    $uniqueExpenses = Deduplicate($thisMonthExpenses)
}

if($uniqueExpenses -ne $null){
    Write-Host "Exporting $($uniqueExpenses.Count) items."
    Export($uniqueExpenses)
}else{
    Write-Host "No expenses to add."
}