


function Update-Progress([Int32]$percent, [String]$text){
    Write-Progress -Activity "Generating Report" `
        -Status "$($percent)% Complete - $($text)" -PercentComplete $percent  
}


Export-ModuleMember -Function 'Update-*'