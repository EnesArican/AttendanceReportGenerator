

$script:MaxUsedColumn = 1
$script:MaxUsedRow = 1
$script:CurrentMonth = ""

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

function Get-CurrentMonth(){
    return $script:CurrentMonth
}

function Set-CurrentMonth($value){
    $script:CurrentMonth = $value
}


Export-ModuleMember -Function 'Set-*'
Export-ModuleMember -Function 'Get-*'



