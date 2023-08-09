function Compare-AccountHistory {
    param (
        [string]$accountHistoryPath,
        [string]$budgetDataPath
    )

    # Load CSV files into variables
    $accountHistory = Import-Csv $accountHistoryPath
    $budgetData = Import-Csv $budgetDataPath

    # Create an empty array to store non-duplicate entries
    $uniqueExpenses = @()

    # Loop through each entry in account history
    foreach ($entry in $accountHistory) {
        $postDate = $entry."Post Date"
        $debit = [decimal]$entry."Debit"
        
        # Check if there's a matching entry in budget data
        $matchingBudgetEntry = $budgetData | Where-Object { $_."Date" -eq $postDate -and $_."Amount" -eq $debit }
        
        # If no match found, consider it a non-duplicate
        if (-not $matchingBudgetEntry) {
            $uniqueExpenses += $entry
        }
    }

    # Return the non-duplicate entries
    return $uniqueExpenses
}

# Call the function and store the result in a variable
$filteredunique$uniqueExpenses = Compare-AccountHistory -accountHistoryPath "filteredAccountHistory.csv" -budgetDataPath "oldBudgetData.csv"
