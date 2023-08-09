#This works.

# Paths to input and output files
$oldBudgetDataPath = "C:\PersonalMyCode\UpdateBudget\oldBudgetData.csv"
$bankDataPath = "C:\PersonalMyCode\UpdateBudget\bankData.csv"
$outputPath = "C:\PersonalMyCode\UpdateBudget\output.csv"

# Read the old budget data
$oldBudgetData = Import-Csv $oldBudgetDataPath

# Read the bank data
$bankData = Import-Csv $bankDataPath

# Ask user to choose a month
$selectedMonth = Read-Host "Enter a number between 1 and 12 for the desired month"
$selectedMonth = [int]$selectedMonth

# Filter bank data by selected month
$bankData = $bankData | Where-Object { $entry = $_; (Get-Date $entry."Post Date").Month -eq $selectedMonth }

# Filter out duplicates and write to the output file
$uniqueExpenses = $bankData | Where-Object { $entry = $_; -not ($oldBudgetData | Where-Object { $_.date -eq $entry."Post Date" -and $_.amount -eq [decimal]$entry.Debit }) }
$uniqueExpenses | Export-Csv $outputPath -NoTypeInformation4