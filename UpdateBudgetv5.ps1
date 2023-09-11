Import-Module ImportExcel
# Install-Module Recycle

$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"
$backupPath = "C:\PersonalMyCode\UpdateBudget\BudgetBackup\"
$rewardsAccountNumber = "313235393200"
$checkingAccountNumber = "750501095729"
$pendingItems =""

$oldBudgetData = ""
$accountHistory = ""


function Main{

    $month = SelectMonth
    $year = SelectYear

    Write-Host "month is $month"
    Write-Host "year is $year"
    

    $accountHistory = ImportAccountHistory($year, $month)
    Write-Host $accountHistOry

    $existingBudget = ImportExistingBudget($month)
    Write-Host $existingBudget

    BackupBudget
    
    #Deduplicate
    $uniqueExpenses = Deduplicate
    
    Write-Host "Exporting $($uniqueExpenses.Count) items."
    ExportExpenses($uniqueExpenses)
    
    if(-not $testMode){
        # $userInput = Read-Host "Delete AccountHistory files? y/n"
        $userInput = 'n'
        if($userInput -eq 'y'){
            Write-Host "Deleting $accountHistoryPaths[0] and $accountHistoryPaths[1]"
            Remove-ItemSafely $accountHistoryPaths[0]
            Remove-ItemSafely $accountHistoryPaths[1]
        }
    }

#Open output.csv
# Invoke-Item $outputPath
}

function TestMode{
    #Test mode?
    $userInput = Read-Host "Turn on test mode? y/n"
    if($userInput -eq "y"){
        Write-Host "Test mode!" -ForegroundColor Yellow
        return $true
    }else{
        Write-Host "Test mode off"
        return $false
    }    
}

function SelectYear{
    return "2023"
}

function SelectMonth{ 
    param(
        
    )
    $fullMonths=@("January","February","March","April","May","June","July","August","September","October","November","December")

    $userInput = Read-Host "Use current month? y/n"
    if($userInput -eq 'y'){
        $selectedMonth = Get-Date -Format "MM"
    }else{
        $selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
    }
    
    $fullMonthName = $fullMonths[$selectedMonth -1]
    Write-Host "Selected month is " -NoNewline; Write-Host "$fullMonthName" -ForegroundColor Green
    # Convert the selected month to an integer
    return [int]$selectedMonth
}

function ImportAccountHistory{
    param(
        $year,
        $month
    )
    $accountHistoryPaths = @()
    if($env:computername -eq "PC_JFLUCKIGER"){
        $accountHistoryPaths = @(
            "C:\Users\jfluckiger\Downloads\AccountHistory.csv",
            "C:\Users\jfluckiger\Downloads\AccountHistory (1).csv"            
        )
    }else{
        $accountHistoryPaths = @(
            "C:\Users\james\Downloads\AccountHistory.csv",
            "C:\Users\james\Downloads\AccountHistory (1).csv"
        )
    }

    $accountHistory0 = Import-Csv $accountHistoryPaths[0]
    $accountHistory1 = Import-Csv $accountHistoryPaths[1]
    $combinedCsv = $accountHistory0 + $accountHistory1
    Write-Host "$($accountHistory0.Count) items in accountHistory0"
    Write-Host "$($accountHistory1.Count) items in accountHistory1"
    Write-Host "$($combinedCsv.Count) total items in account history."
    
    # Filter account history data by selected month and specific condition
    $filteredAccountHistory = $combinedCsv | Where-Object {
        $entry = $_
        if($entry.Status -eq "Pending"){
            Write-Host "Found a pending $entry"
            $pendingItems += $entry
        }else{

            $postDate = Get-Date $entry."Post Date"
            
            # Check if the year is 2023 and the month matches
            if ([int]$postDate.Year -ne $year) {
                # Write-Host "Year doesn't match $selectedYear"
                return $false
            }   
            if ([int]$postDate.Month -ne $month){
                # Write-Host "Month doesn't match $selectedMonth"
                return $false
            }
            return $true
        }
    }
    Write-Host "After trimming extraneous months and years there are $($filteredAccountHistory.Count) items."
    # Write-Host $filteredAccountHistory
    # foreach($item in $filteredAccountHistory){
    #     Write-Host $item
    # }
    return $filteredAccountHistory
}

function ImportExistingBudget{
    param(
        $laptop,
        $month #for $abbMonthName
    )

    #Determine csv or xlsx.   
    $userInput = Read-Host "c for local csv, x for 2023Budget.xlsx"
    if($userInput -eq "c"){
        $source = "csv"
    }elseif($userInput -eq "x"){
        $source = "xlsx"
    }

    if($source -eq "csv"){
        $budgetcsvPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
        Write-Host "Importing budget data from the local csv."
        return = Import-Csv $budgetcsvPath  

    }elseif($source -eq "xlsx"){
        #Determine $abbMonthName
        $abbMonths=@("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
        $abbMonthName = $abbMonths[$month -1]
        Write-Host $abbMonthName
        
        #Determine path.
        $xlsxPath = ""
        if($env:computername -eq "PC_JFLUCKIGER"){
            Write-Host "Laptop" -ForegroundColor Blue
            $xlsxPath = "C:\Users\jfluckiger\OneDrive\Budget\2023Budget.xlsx"
        }else{
            Write-Host "Desktop" -ForegroundColor Blue
            $xlsxPath = "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
        }
        Write-Host "Importing budget data from 2023Budget.xlsx"
        try{
            $rawXlsxData = Import-Excel $xlsxPath -WorksheetName $abbMonthName -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200
        }catch{
            Write-Host "Importing Excel data failed. Make sure it's closed."
            exit
        }

        #Remove blank items. Add to refined data.
        $refinedXlsxData = ""
        foreach($item in $rawXlsxData){
            if($null -ne $item.P1){
                $nonBlankExpense = [PSCustomObject]@{
                    Date = [string](Get-Date $item.P1 -Format "MM/dd/yyyy")
                    Item = $item.P2
                    Description = $item.P3
                    Method = $item.P4
                    Category = $item.P5
                    Amount = [decimal]$item.P6
                }
            $refinedXlsxData += $nonBlankExpense
            }
        }
    }
    return $refinedXlsxData
}






function ImportBudgetFromCsvOld(){ #Sets $script:oldBudgetData.
    Write-Host "Importing budget data from the local csv."
    $script:oldBudgetData = Import-Csv $budgetcsvPath
}

function ImportBudgetFromXlsxOld(){
    Write-Host "Importing budget data from 2023Budget.xlsx"
    try{
        # Write-Host $script:budgetXlsxPath
        # Write-Host $script:abbMonthName
        $rawXlsxData = Import-Excel $script:budgetXlsxPath -WorksheetName $script:abbMonthName -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200
    }catch{
        Write-Host "Importing Excel data failed. Make sure it's closed."
        exit
    }

    #Remove blank items. Add to refined data.
    $refinedXlsxData = ""
    foreach($item in $rawXlsxData){
        if($null -ne $item.P1){
            $nonBlankExpense = [PSCustomObject]@{
                Date = [string](Get-Date $item.P1 -Format "MM/dd/yyyy")
                Item = $item.P2
                Description = $item.P3
                Method = $item.P4
                Category = $item.P5
                Amount = [decimal]$item.P6
            }
        $refinedXlsxData += $nonBlankExpense
        }
    }
    # Write-Host $refinedXlsxData
    return $refinedXlsxData

}

function ImportAccountHistoryOld() {
    $combinedCsv = ""
    $accountHistory0 = Import-Csv $script:accountHistoryPaths[0]
    $accountHistory1 = Import-Csv $script:accountHistoryPaths[1]
    $combinedCsv = $accountHistory0 + $accountHistory1
    # $combinedCsv = @($accountHistory0, $accountHistory1)

    Write-Host "Both account history csvs combined have $($combinedCsv.Count) items total."
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
    Write-Host "After trimming extraneous months and years there are $($filteredAccountHistory.Count) items."
    # Write-Host $filteredAccountHistory
    # foreach($item in $filteredAccountHistory){
    #     Write-Host $item
    # }
    return $filteredAccountHistory
}

function BackupBudget(){
    $destination = $script:backupPath + (Get-DAte -Format "MM-dd-yyyy-hh-mm") 
    Copy-Item $script:budgetXlsxPath -Destination $destination 
}

function Deduplicate{
    #Remove the dollar sign and whitespace and change parenthesis to - sign.
    foreach ($entry in $script:oldBudgetData){
        $entry.Amount = $entry.Amount.Replace('$', '')
        $entry.Amount = $entry.Amount.Replace(' ', '')

        if($entry.Amount[0] -match "^\("){
            # $entry.Amount = $entry.Amount -replace "^\(", "-"
            $entry.Amount = "-" + $entry.Amount
            $entry.Amount = $entry.Amount.Remove(1,1)
            $entry.Amount = $entry.Amount.Substring(0, $entry.Amount.Length - 1)
        }
    }

    Write-Host "Existing budget data has $($script:oldBudgetData.Count) items."
    $uniqueExpenses = @()

    foreach ($entry in $script:accountHistory) {
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
        $matchingBudgetEntry = $script:oldBudgetData | Where-Object { $_.Date -eq $postDate -and [decimal]$_.Amount -eq $amount}

        # If no match found, consider it a non-duplicate
        if (-not $matchingBudgetEntry) {
            
            #Determine method.
            $method = $null
            if ($entry."Account Number" -eq $script:rewardsAccountNumber){
                $method = "Rewards"
            }elseif ($entry."Account Number" -eq $script:checkingAccountNumber) {
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

Main