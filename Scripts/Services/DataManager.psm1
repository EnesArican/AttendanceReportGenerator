
Import-Module .\Scripts\Models\Variables.psm1
Import-Module .\Scripts\ProgressWriter.psm1

$script:DataHash = [ordered]@{}
$script:DatesList = New-Object System.Collections.Generic.List[System.Object]

function Get-Data($ws){
    Write-Host "Getting names and attendance records..." -NoNewline

    $nameString = 'Last Name'
    $range = $ws.Range("A1","A600")
    $recordSet = 0
    $nameSearch = $range.find($nameString)
    
    if ($null -ne $nameSearch) {
        $firstAddress = $nameSearch.Address()
       do {
            $recordSet++
            $row = $nameSearch.row + 1
            do {
                $lastName = $ws.cells.item($row,1).value()
                if ($lastName){ Add-AttendanceToHash -ws $ws -row $row -lastName $lastName -recordSet $recordSet }
                $row++
            } while ($null -ne $lastName)
            
            $absentNamesForDate = $script:DataHash.GetEnumerator() | ? { $_.Value.Count -lt $recordSet } 
            $absentNamesForDate | % { $_.Value.Add("emp") }
                        
            $nameSearch = $range.FindNext($nameSearch) 
        } while ( $null -ne $nameSearch -and $nameSearch.Address() -ne $firstAddress)
    }

    $script:DataHash = $script:DataHash.GetEnumerator() | sort-Object -Property name
    #$script:DataHash.GetEnumerator() | Out-String | Write-Host

    Write-Ok
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


function Get-Dates($ws){
    Write-Host "Getting dates..." -NoNewline

    $dateString = 'Date:*'
    $range = $ws.Range("A1","A600")
    
    $dateSearch = $range.find($dateString)
    if ($null -ne $dateSearch) {
        $FirstAddress = $dateSearch.Address()
       do { 
            $row = $dateSearch.row
            $date = $ws.cells.item($row,1).value()            
            $script:DatesList.Add($date)

    	    $dateSearch = $range.FindNext($dateSearch)
        
        } while ( $null -ne $dateSearch -and $dateSearch.Address() -ne $FirstAddress)
    }

    Write-Ok
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


Export-ModuleMember -Function 'Get-*'
Export-ModuleMember -Function 'Set-*'