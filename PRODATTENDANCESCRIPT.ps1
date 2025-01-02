Start-Process "C:\Program Files\Microsoft OneDrive\OneDrive.exe" "/sync"

Start-Sleep -Seconds 5

$bookpath = 'C:\Users\User\OneDrive\AttendanceAutomation-DO-NOT-MOVE.\PRODATTENDANCEBOOK.xlsx' #change to the prod book
# $dayofMonth = RowToUpdate
$dayofMonth = [int](Get-Date).Day + 5
# $month = BookToUpdate
$month = (Get-Date).ToString("MMMMMMMMMMMMMMMMMM")
$ $database = employeeid of users
$database = Import-Csv -Path 'C:\Users\User\OneDrive\PRODDATABASEIDS.csv' #move this to an online location
$report = Import-Csv -Path 'C:\Users\User\OneDrive\AttendanceAutomation-DO-NOT-MOVE.\reportdata.csv' #change this to the name of file PowerAutomate creates
$report.{Agent ID} | Set-Content -Path 'C:\Users\User\OneDrive\AttendanceAutomation-DO-NOT-MOVE\outdata.txt' #move this to an online location
$workbook = Open-ExcelPackage -Path $bookpath
$worksheet = $workbook.$month #change to workbook.$month
$rowCount = $worksheet.Dimension.End.Row
$rowNumb = $null
$newfile = Get-Content -Path 'C:\Users\User\OneDrive\AttendanceAutomation-DO-NOT-MOVE\outdata.txt' #change to the online location

foreach ($id in $database) {
    Write-Host "Current id in database is $id"
    $databaseids = $id.ID_List
    foreach($i in $databaseids) {
        Write-Host "Current DatabaseID is $i"
    if ($newfile -contains $i) {
    for ($row = 1; $row -le $rowCount; $row++) {
        if ($worksheet.Cells[$row, 5].Value -eq $i) { #change column number to a static column
            $rowNumb = $row
            Write-Host "$i has a match and is on row $rownumb"
            $worksheet.cells[$rowNumb, $dayofMonth].Value='Y'
            break
        } 
    }
    }
    else {
        for ($row = 1; $row -le $rowCount; $row++) {
            if ($worksheet.Cells[$row, 5].Value -eq $i) { #change column number to a static column
                $rowNumb = $row
                Write-Host "$i has no match but is on row $rownumb"
                if ($worksheet.Cells[$rownumb, $dayofMonth].Value -eq 'X' -or $null -eq $worksheet.Cells[$rowNumb, $dayofMonth].Value -or $worksheet.Cells[$rowNumb, $dayofMonth].Value -eq '') {
                    Write-Host "Employee is marked as Excused, skipping..."
                    break
                }
                else {
                    $worksheet.cells[$rowNumb, $dayofMonth].Value='N'
                }
                break
            }
    }
}
    }
    

}

Write-Host 'Saving Changes, Please wait...'
Start-Sleep -Seconds 15

################################################################################################
Close-ExcelPackage -ExcelPackage $workbook
################################################################################################

Write-Host 'Closing the Excel Book, Please wait...'
Start-Sleep -Seconds 5
Exit-PSSession