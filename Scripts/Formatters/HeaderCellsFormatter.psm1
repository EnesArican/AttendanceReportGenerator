Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1

[Int32]$xlAlignCenter = -4108
[Int32]$xlAlignLeft = -4131

function Format-NumberingColumn([__ComObject]$Range){
    $Range.Font.Bold = $true
    $Range.Font.Size = 10
    $Range.ColumnWidth = 3.14
    $Range.Font.Name = "Times New Roman"
    $Range.HorizontalAlignment = $xlAlignLeft
    $Range.VerticalAlignment = $xlAlignCenter
    $Range.font.color = RGB 192 0 0 
    $Range.interior.color = RGB 228 223 236
}

function Format-Days([__ComObject]$Range){
    $Range.Font.Bold = $true
    $Range.Font.Size = 8
    $Range.Font.Name = "Times New Roman"
    $Range.HorizontalAlignment = $xlAlignCenter
    $Range.VerticalAlignment = $xlAlignCenter
    $Range.font.color = RGB 0 32 96
    $Range.interior.color = RGB 228 223 236
}
function Format-Dates([__ComObject]$Range){
    $Range.Font.Bold = $true
    $Range.Font.Size = 10
    $Range.Font.Name = "Times New Roman"
    $Range.HorizontalAlignment = $xlAlignCenter
    $Range.VerticalAlignment = $xlAlignCenter
    $Range.font.color = RGB 151 71 6
    $Range.interior.color = RGB 228 223 236
}

function Format-Title([__ComObject]$Range){
    $Range.Font.Bold = $true
    $Range.Font.Size = 16
    $Range.Font.Name = "Times New Roman"
    $Range.HorizontalAlignment = $xlAlignCenter
    $Range.VerticalAlignment = $xlAlignCenter
    $Range.font.color = RGB 192 0 0 
    $Range.interior.color = RGB 228 223 236
}


Export-ModuleMember -Function 'Format-*'
