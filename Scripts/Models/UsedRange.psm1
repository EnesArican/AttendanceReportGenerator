

$script:MaxUsedColumn = 1
$script:MaxUsedRow = 1


function Get-MaxUsedColumn(){
    return $script:MaxUsedColumn
}

function Set-MaxUsedColumn($value){
    $script:MaxUsedColumn = $value
}

function Get-MaxUsedRow(){
    return $script:MaxUsedRow
}

function Set-MaxUsedRow($value){
    $script:MaxUsedRow = $value
}


Export-ModuleMember -Function 'Set-*'
Export-ModuleMember -Function 'Get-*'



