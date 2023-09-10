$dog = "Yahoo"

function CoolFunction(){
    
    # Write-Host $script:dog
    $script:dog = "Doink"
}

CoolFunction

Write-Host $dog
