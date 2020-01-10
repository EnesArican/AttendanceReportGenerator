


$script:DataHash = [ordered]@{}
$script:DatesArray = New-Object System.Collections.Generic.List[System.Object]

function Get-Data($ws){
    $nameString = 'Last Name'
    $range = $ws.Range("A1","A3000")
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
            
            $nameSearch = $range.FindNext($nameSearch) 
        } while ( $null -ne $nameSearch -and $nameSearch.Address() -ne $firstAddress)
    }

    $script:DataHash = $script:DataHash.GetEnumerator() | sort-Object -Property name
    #$script:DataHash.GetEnumerator() | Out-String | Write-Host
}

function Add-AttendanceToHash($ws, $row, $lastName, $recordSet){
    $value = $ws.cells.item($row,3).value()
    $firstName =  $ws.cells.item($row,2).value()
    $key = $FirstName + ' ' + $LastName

    if($script:DataHash.Keys -contains $key){
        $script:DataHash[$key].Add($value)
    }else {
        $attendanceArr = New-Object System.Collections.Generic.List[System.Object]
       
        #need to test this
        if($recordSet -ne 1){
            1..($recordSet-1) | % { $attendanceArr.Add("emp") }
        }
        
        $attendanceArr.Add($value)
        $script:DataHash.Add($key, $attendanceArr)
    }

}


function Get-Dates($ws){
    $dateString = 'Date:*'
    $range = $ws.Range("A1","A3000")
    
    $dateSearch = $range.find($dateString)
    if ($null -ne $dateSearch) {
        $FirstAddress = $dateSearch.Address()
       do { 
            $row = $dateSearch.row
            $date = $ws.cells.item($row,1).value()            
            $script:DatesArray.Add($date)

    	    $dateSearch = $range.FindNext($dateSearch)
        
        } while ( $null -ne $dateSearch -and $dateSearch.Address() -ne $FirstAddress)
    }
}


function Set-Data($ws){
    $row = 2
    foreach ($h in $script:DataHash.GetEnumerator){
        $column = 1
        $ws.cells.Item($row, $column) = $h.Name
        foreach ($v in $h.Value){
            $column++
            $worksheet.cells.Item($row, $column) = $v
        }
        $row++
    }
}

function Set-Data($ws, $nameArray, $attendanceHash){
    Set-IhvanNames -worksheet $ws -nameArray $nameArray
    Set-DatesAndRecords -worksheet $ws -attendanceHash $attendanceHash
}

function Set-DatesAndRecords($worksheet, $attendanceHash){
    $column = 2

    foreach ($h in $attendanceHash.GetEnumerator()) {
        $row = 1
        $worksheet.cells.Item($row, $column) = $h.Name
        $row++
        foreach ($v in $h.Value){
            $worksheet.cells.Item($row, $column) = $v
            $row++
        }
        $column++
    }
}

function Set-IhvanNames($worksheet, $nameArray){
    $row = 2
    foreach ($name in $nameArray){
        $worksheet.cells.Item($row, 1) = $name
        $row++
    }
    $global:MaxUsedRow = $row
}

Export-ModuleMember -Function 'Get-*'
Export-ModuleMember -Function 'Set-Data*'