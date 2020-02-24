
Import-Module .\Scripts\Models\Variables.psm1
Import-Module .\Scripts\ProgressWriter.psm1

$script:DataHash = [ordered]@{}
$script:DatesList = New-Object System.Collections.Generic.List[System.Object]

function Get-Data($ws){
    Write-Host "Getting all data..." -NoNewline

    $nameString = 'Fatih'
    $range = $ws.Range("A1","A900")
    $recordSet = 0
    $previousDataRow = 0

    $nameSearch = $range.find($nameString,[Type]::Missing,[Type]::Missing,1)
    
    if ($null -ne $nameSearch) {
        $firstAddress = $nameSearch.Address()
       do {
            Add-DateToList -ws $ws -row $nameSearch.row -prevRow $previousDataRow

            $recordSet++
            $row = $nameSearch.row + 2

            # do {
            #     $lastName = $ws.cells.item($row,1).value()
            #     if ($lastName){ Add-AttendanceToHash -ws $ws -row $row -lastName $lastName -recordSet $recordSet }
            #     $row++
            # } while ($null -ne $lastName)
            Get-Attendance -ws $ws -row $row -recordSet $recordSet

            $previousDataRow = $row

            $absentNamesForDate = $script:DataHash.GetEnumerator() | ? { $_.Value.Count -lt $recordSet } 
            $absentNamesForDate | % { $_.Value.Add("emp") }
                        
            $nameSearch = $range.FindNext($nameSearch) 
        } while ( $null -ne $nameSearch -and $nameSearch.Address() -ne $firstAddress)
    }

    $script:DataHash = $script:DataHash.GetEnumerator() | sort-Object -Property name
    #$script:DataHash.GetEnumerator() | Out-String | Write-Host

    Write-Ok
}

function Add-DateToList($ws, $row, $prevDataRow){
    
    while ($row -ne $prevDataRow -and $rowValue -notmatch "Date") {
        $rowValue = $ws.cells.item($row,1).value() 
        $row--
    }     
    $script:DatesList.Add($rowValue)
}

function Get-Attendance($ws, $row, $recordSet){
    do {
        $lastName = $ws.cells.item($row,1).value()
        if ($lastName){ Add-AttendanceToHash -ws $ws -row $row -lastName $lastName -recordSet $recordSet }
        $row++
    } while ($null -ne $lastName)
}

function Add-AttendanceToHash($ws, $row, $lastName, $recordSet){
    $value = $ws.cells.item($row,3).value()
    $firstName =  $ws.cells.item($row,2).value()
    $key = $FirstName + ' ' + $LastName

    if($script:DataHash.Keys -contains $key){
        $script:DataHash[$key].Add($value)
    }else {
        $attendanceArr = New-Object System.Collections.Generic.List[System.Object]
        if($recordSet -ne 1){
            1..($recordSet-1) | % { $attendanceArr.Add("emp") }
        }
        $attendanceArr.Add($value)
        $script:DataHash.Add($key, $attendanceArr)
    }
}


function Set-Data($ws){
    Write-Host "Writing names and attendance records..." -NoNewline

    $row = 2
    foreach ($h in $script:DataHash.GetEnumerator()){
        $column = 1
        $ws.cells.Item($row, $column) = $h.Name
        $register = @($h.Value)
        [array]::Reverse($register)
        foreach ($v in $register){
            $column++
            $ws.cells.Item($row, $column) = $v
        }
        $row++
    }
    Set-MaxUsedRow -value $row

    Write-Ok    
}

function Set-Dates($ws){
    Write-Host "Writing dates..." -NoNewline

    $column = 2
    $datesArray = $script:DatesList | % { $_ }
    [array]::Reverse($datesArray)
    foreach ($date in $datesArray){
        $ws.cells.Item(1, $column) = $date
        $column++
    }
    Set-MaxUsedColumn -value ($column)

    Write-Ok
}


Export-ModuleMember -Function 'Get-Data'
Export-ModuleMember -Function 'Set-Data'
Export-ModuleMember -Function 'Set-Dates'