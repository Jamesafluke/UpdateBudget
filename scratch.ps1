function Math{
    # $firstNumber = FirstNumber
    $result = (FirstNumber) + 3
    return $result
}

function FirstNumber{
    return 5
}

Write-Host (Math)