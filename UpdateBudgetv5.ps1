Import-Module ImportExcel
# Install-Module Recycle


$testMode = $false
# Paths to input and output files
$budgetxlsxPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
# $budgetxlsxPath = "C:\Users\jfluckiger\OneDrive\Budget\2023Budget.xlsx"
$budgetcsvPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"
$rewardsAccountNumber = "313235393200"
$checkingAccountNumber = "750501095729"
$pendingItems =""
$selectedYear = ""
$selectedMonth = ""
$abbMonthName = ""
$abbMonths=@("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
$fullMonths=@("January","February","March","April","May","June","July","August","September","October","November","December")
$accountHistoryPaths = @()
$oldBudgetData = ""

function AccountHistoryPaths{
    if($testMode){
        $accountHistoryPaths = @(
            "C:\PersonalMyCode\UpdateBudget\AccountHistory.csv",
            "C:\PersonalMyCode\UpdateBudget\AccountHistory (1).csv"
        )
    }else{
        $accountHistoryPaths = @(
            "C:\Users\james\Downloads\AccountHistory.csv",
            "C:\Users\james\Downloads\AccountHistory (1).csv"
        )
    }
}

function SelectMonthYear{
    if($testMode){
        $selectedMonth = "8"
        $selectedYear = "2023"
        $abbMonthName = $abbMonths[$selectedMonth -1]
        $fullMonthName = $fullMonths[$selectedMonth -1]
        Write-Host "Hard coded test month and year are " -NoNewline; Write-Host "$fullMonthName $selectedYear" -ForegroundColor Green
    }else{
        $userInput = Read-Host "Use current month? y/n"
        if($userInput -eq 'y'){
            $selectedMonth = Get-Date -Format "MM"
            $abbMonthName = $abbMonths[$selectedMonth -1]
            $fullMonthName = $fullMonths[$selectedMonth -1]
            Write-Host "Current month is " -NoNewline; Write-Host "$fullMonthName" -ForegroundColor Green
        }else{
            $selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
            # $selectedYear = Read-Host "Enter the year"
            $selectedYear = "2023"
            $abbMonthName = $abbMonths[$selectedMonth -1]
            $fullMonthName = $fullMonths[$selectedMonth -1]
            Write-Host "Selected month is " -NoNewline; Write-Host "$fullMonthName" -ForegroundColor Green
        }
    }
    # Convert the selected month to an integer
    $selectedMonth = [int]$selectedMonth
    $selectedYear = [int]$selectedYear
}

function ImportBudgetFromCsv(){
    Write-Host "Importing budget data from the local csv."
    $oldBudgetData = Import-Csv $budgetcsvPath
}

function ImportBudgetFromXlsx(){
    Write-Host "Importing budget data from 2023Budget.xlsx"
    try{
        $excelData = Import-Excel $budgetxlsxPath -WorksheetName $abbMonthName -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200
    }catch{
        Write-Host "Importing Excel data failed. Make sure it's closed."
    }

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

function ImportBankDataFromCsv() {
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
        if($entry.Status -eq "Pending"){
            Write-Host "Found a pending $entry"
            $pendingItems += $entry
        }else{

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
    }
    Write-Host "Selected month has $($filteredAccountHistory.Count) items."
    # Write-Host $filteredAccountHistory
    # foreach($item in $filteredAccountHistory){
    #     Write-Host $item
    # }
    return $filteredAccountHistory
}

function Deduplicate($thisMonthExpenses){
    #Remove the dollar sign and whitespace and change parenthesis to - sign.
    foreach ($entry in $oldBudgetData){
        $entry.Amount = $entry.Amount.Replace('$', '')
        $entry.Amount = $entry.Amount.Replace(' ', '')

        if($entry.Amount[0] -match "^\("){
            # $entry.Amount = $entry.Amount -replace "^\(", "-"
            $entry.Amount = "-" + $entry.Amount
            $entry.Amount = $entry.Amount.Remove(1,1)
            $entry.Amount = $entry.Amount.Substring(0, $entry.Amount.Length - 1)
        }
    }

    

    Write-Host "Existing budget data has $($oldBudgetData.Count) items."
    $uniqueExpenses = @()

    foreach ($entry in $thisMonthExpenses) {
        $postDate = $entry."Post Date"
        $debit = [decimal]$entry."Debit"
        $credit = [decimal]$entry."Credit"

        #Determine debit or credit.
        $amount = $null
        if ($debit -ne ""){
            $amount = $debit
        }else{
            $amount = $credit * -1
        }
        
        # Check if there's a matching entry in old budget data
        $matchingBudgetEntry = $oldBudgetData | Where-Object { $_.Date -eq $postDate -and [decimal]$_.Amount -eq $amount}

        # If no match found, consider it a non-duplicate
        if (-not $matchingBudgetEntry) {
            

            #Determine method.
            $method = $null
            if ($entry."Account Number" -eq $rewardsAccountNumber){
                $method = "Rewards"
            }elseif ($entry."Account Number" -eq $checkingAccountNumber) {
                $method = "Checking"
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

function ExportExpenses($uniqueExpenses){
    Write-Host "Export!"
    $uniqueExpenses | Export-Csv $outputPath -NoTypeInformation
}




#Test mode?
if($testMode){
    Write-Host "Test mode!" -ForegroundColor Yellow
}else{
    Write-Host "Starting!" -ForegroundColor Green
}

AccountHistoryPaths

Write-Host $accountHistoryPaths

SelectMonthYear

#Import existing data
if($testMode){
    $userInput = Read-Host "Import from local csv? y/n"
    # $userInput = "y"
    if($userInput){
        $thisMonthExpenses = ImportBankDataFromCsv
    }else{
        $thisMonthExpenses = ImportBudgetFromXlsx
    }
}else{
    $thisMonthExpenses = ImportBudgetFromXlsx
}


if($null -ne $thisMonthExpenses){
    $uniqueExpenses = Deduplicate($thisMonthExpenses)
}

if($null -ne $uniqueExpenses){
    Write-Host "Exporting $($uniqueExpenses.Count) items."
    ExportExpenses($uniqueExpenses)
}else{
    Write-Host "No expenses to add."
}

if(-not $testMode){
    $userInput = Read-Host "Delete AccountHistory files? y/n"
    if($userInput -eq 'y'){
        Write-Host "Deleting $accountHistoryPaths[0] and $accountHistoryPaths[1]"
        Remove-ItemSafely $accountHistoryPaths[0]
        Remove-ItemSafely $accountHistoryPaths[1]
    }
}

#Open output.csv
# Invoke-Item $outputPath