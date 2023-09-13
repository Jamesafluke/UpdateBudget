Import-Module ImportExcel
# Install-Module Recycle

function Main{
    
    Write-Host "Welcome to Budginator!" -ForegroundColor Green

    StartAhk

    $outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"
    $month = SelectMonth
    $year = SelectYear

    $accountHistoryPaths = SetAccountHistoryPaths
    
    $accountHistory = ImportAccountHistory $year $month $accountHistoryPaths

    $existingBudget = ImportExistingBudget $month

    BackupBudget

    $verifiedExpenses = Deduplicate $accountHistory $existingBudget
    
    ExportExpenses $verifiedExpenses $outputPath

    DeleteAccountHistoryFiles $accountHistoryPaths
    
    OpenOutput $outputPath
}
function StartAhk{
    $userInput = Read-Host "Download account history? y/n"
    if($userInput -eq "y"){
        if($env:computername -eq "PC_JFLUCKIGER"){
            Invoke-Item "C:\PersonalMyCode\UpdateBudget\AHK\laptopDownloadUccu.ahk"
        }
        else{
            Invoke-Item "C:\PersonalMyCode\UpdateBudget\AHK\desktopDownloadUccu.ahk"
        }
    }
}

function SelectYear{
    $year = "2023"
    return [int]$year
}

function SelectMonth{ 
    param(
        
    )

    $userInput = Read-Host "Use current month? y/n"
    if($userInput -eq 'y'){
        $selectedMonth = Get-Date -Format "MM"
    }else{
        $selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
    }
    
    Write-Host "Selected month is " -NoNewline; Write-Host (GetFullMonthName $month) -ForegroundColor Green

    return [int]$selectedMonth
}

function SetAccountHistoryPaths{
    $accountHistoryPaths = @()
    if($env:computername -eq "PC_JFLUCKIGER"){
        Write-Host "Laptop detected."
        $accountHistoryPaths = @(
            "C:\Users\jfluckiger\Downloads\AccountHistory.csv",
            "C:\Users\jfluckiger\Downloads\AccountHistory (1).csv"            
        )
    }else{
        $accountHistoryPaths = @(
            Write-Host "Desktop detected."
            "C:\Users\james\Downloads\AccountHistory.csv",
            "C:\Users\james\Downloads\AccountHistory (1).csv"
        )
    }
    return $accountHistoryPaths
}

function GetXlsxPath{
    if($env:computername -eq "PC_JFLUCKIGER"){
        return "C:\Users\jfluckiger\OneDrive\Budget\2023Budget.xlsx"
    }else{
        return "C:\Users\james\OneDrive\Budget\2023Budget.xlsx"
    }
}

function GetFullMonthName{
    param(
        $month
    )
    $fullMonths=@("January","February","March","April","May","June","July","August","September","October","November","December")
    return $fullMonths[$month -1]
}

function ImportAccountHistory{
    param(
        $year,
        $month,
        $accountHistoryPaths
    )
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
            
            # Check if the year and month match
            if ([int]$postDate.Year -ne $year) {
                return $false
            }   
            if ([int]$postDate.Month -ne $month){
                return $false
            }
            return $true
        }
    }
    Write-Host "After trimming extraneous months and years there are " -NoNewLine; Write-Host $filteredAccountHistory.Count -NoNewLine -ForegroundColor Green; Write-Host "account history items in " -NoNewLine; Write-Host (GetFullMonthName $month) $year -NoNewLine -ForegroundColor Green; Write-Host "."
    return $filteredAccountHistory
}

function ImportExistingBudget{
    param(
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
        $budgetcsvPath = "C:\PersonalMyCode\UpdateBudget\existingBudgetData.csv"
        Write-Host "Importing budget data from the local csv."
        return Import-Csv $budgetcsvPath  

    }elseif($source -eq "xlsx"){
        $rawXlsxData = @()
        #Determine $abbMonthName
        $abbMonths=@("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
        $abbMonthName = $abbMonths[$month -1]
        
        #Determine path.
        $GetXlsxPath = (GetXlsxPath)
        Write-Host "Importing budget data from 2023Budget.xlsx"
        try{
            $rawXlsxData = Import-Excel $GetXlsxPath -WorksheetName $abbMonthName -NoHeader -ImportColumns @(19,20,21,22,23,24) -startrow 8 -endrow 200
        }catch{
            Write-Host "Importing Excel data failed. Make sure it's closed."
            exit
        }

        #Remove blank items. Add to refined data.
        $refinedXlsxData = @()
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

function BackupBudget(){
    $backupPath = "C:\PersonalMyCode\UpdateBudget\BudgetBackup\"
    $destination = $backupPath + (Get-Date -Format "MM-dd-yyyy-hh-mm")
    $path = GetXlsxPath
    Copy-Item $path -Destination $destination 
}

function Deduplicate{
    param(
        $accountHistory,
        $existingBudget
    )

    Write-Host "Existing budget data has " -NoNewLine; Write-Host $existingBudget.Count -NoNewLine -ForegroundColor Green; Write-Host " items."
    $verifiedExpenses = @()

    foreach ($entry in $accountHistory) {
        $postDate = $entry."Post Date"
        $debit = [decimal]$entry."Debit"
        $credit = [decimal]$entry."Credit"

        #Set $amount.
        $amount = $null
        if ($debit -ne ""){
            $amount = $debit
        }else{
            $amount = $credit * -1
        }
        
        # Check if there's a matching entry in existing budget data
        $discardableBudgetEntry = $existingBudget | Where-Object { $_.Date -eq $postDate -and [decimal]$_.Amount -eq $amount}

        $duplicateCount = 0
        # If no match found, consider it a non-duplicate
        if (-not $discardableBudgetEntry) {
            
            #Determine method.
            $method = $null
            if ($entry."Account Number" -eq "313235393200"){
                $method = "Rewards"
            }elseif ($entry."Account Number" -eq "750501095729") {
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
            $verifiedExpenses += $newExpense
        }else{
            #Is a duplicate.
            $duplicateCount ++
        }
    }        
    Write-Host $duplicateCount -NoNewLine -ForegroundColor Green; Write-Host " duplicates. There are " -NoNewLine; Write-Host $verifiedExpenses.count -NoNewLine -ForegroundColor Green; Write-Host " expenses ready to be exported."
    return $verifiedExpenses
}

function ExportExpenses{
    param(
        $verifiedExpenses,
        $outputPath
    )
    Write-Host "Exporting!"
    $verifiedExpenses | Export-Csv $outputPath -NoTypeInformation
}

function DeleteAccountHistoryFiles{
    param(
        $accountHistoryPaths
    )
    $userInput = Read-Host "Delete AccountHistory files? y/n"
    # $userInput = 'n'
    if($userInput -eq 'y'){
        Write-Host "Deleting $accountHistoryPaths[0] and $accountHistoryPaths[1]"
        Remove-ItemSafely $accountHistoryPaths[0]
        Remove-ItemSafely $accountHistoryPaths[1]
    }
}

function OpenOutput{
    param(
        $outputPath
    )
    $userInput = Read-Host "Open output.csv? y/n"
    if($userInput -eq "y"){
        Invoke-Item $outputPath
    }
}

Main