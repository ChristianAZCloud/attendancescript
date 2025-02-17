# Define an array of modules to check for
$modulesToCheck = @("PSWriteExcel", "ImportExcel")

# Function to check and install a module if it's not available
function Install-RequiredModules {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName is not installed. Installing now..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -Confirm:$false -Force -AllowClobber
            Write-Host "$ModuleName installed successfully!" -ForegroundColor Green
        } catch {
            Write-Host "Failed to install $ModuleName. Run PowerShell as Administrator and try again." -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "$ModuleName is already installed." -ForegroundColor Cyan
    }
}

# Loop through the array to ensure each module is installed and import it
foreach ($module in $modulesToCheck) {
    Install-RequiredModules -ModuleName $module
    Import-Module $module -ErrorAction Stop
}

function StartOneDrive {
    Write-Host 'Starting OneDrive...' -ForegroundColor Yellow

    Start-Process "C:\Program Files\Microsoft OneDrive\OneDrive.exe" "/sync"
}

function RunAttendance {
    # Path of the Excel Book where we will be making changes
    $bookpath = "$home\OneDrive\AttendanceAutomation-DO-NOT-MOVE\AttendanceBook.xlsx" 
    # Gets the current day and adds 5 to know which column to update
    $dayofMonth = [int](Get-Date).Day + 5
    # Gets the full month to know what sheet in the excel book to update
    $month = (Get-Date).ToString("MMMMMMMMMMMMMMMMMM")
    # Grabs a CSV file that contains all the employee ids of techs
    $database = Import-Csv -Path "$home\OneDrive\AttendanceAutomation-DO-NOT-MOVE\PRODDATABASEIDS.csv" 
    # Grabs the  attendance report that's emailed to me then stored in a sharepoint/onedrive location by Power Automate
    $report = Import-Csv -Path "$home\OneDrive\AttendanceAutomation-DO-NOT-MOVE.\reportdata.csv" #change this to the name of file PowerAutomate creates
    # Pulls Agent ID from the attendance report to determine who was logged in today {} is needed due to the space between Agent ID
    $report.{Agent ID} | Set-Content -Path "$home\OneDrive\AttendanceAutomation-DO-NOT-MOVE\outdata.txt"
    $workbook = Open-ExcelPackage -Path $bookpath
    #Defines the worksheet we'd like to update 
    $worksheet = $workbook.$month #change to workbook.$month
    $rowCount = $worksheet.Dimension.End.Row
    $rowNumb = $null
    # Grabs the Agent IDs of the agent who logged in today
    $newfile = Get-Content -Path "$home\OneDrive\AttendanceAutomation-DO-NOT-MOVE\outdata.txt" #change to the online location

    foreach ($id in $database) {
        # Remove ### from line 67 if needing to debug
        ### Write-Host "Current id in database is $id" -ForegroundColor Yellow
        $databaseids = $id.ID_List
        foreach($i in $databaseids) {
            # Remove ### from line 71 if needing to debug
            ### Write-Host "Current DatabaseID is $i" -ForegroundColor Yellow
         if ($newfile -contains $i) {
            for ($row = 1; $row -le $rowCount; $row++) {
                if ($worksheet.Cells[$row, 5].Value -eq $i) { # Searching each row on column 5 to locate the Agent ID
                    $rowNumb = $row
                    Write-Host "$i has a match and is on row $rownumb" -ForegroundColor Green
                    # Sets the new value of the Cell
                    $worksheet.cells[$rowNumb, $dayofMonth].Value='Y'
                    break
            }   
        }
        }
        else {
            # Loops through each row in the sheet 
            for ($row = 1; $row -le $rowCount; $row++) {
                # [$row,5] means current row and 5 is the column where we have stored employee IDs so the script knows that cell to update for each agent.
                if ($worksheet.Cells[$row, 5].Value -eq $i) { # Searching each row on column 5 to locate the Agent ID
                    $rowNumb = $row
                    Write-Host "$i has no match but is on row $rownumb" -ForegroundColor Red
                    # Skips Agents who do are marked as absent or don't work on this particular day
                    if ($worksheet.Cells[$rownumb, $dayofMonth].Value -eq 'X' -or $null -eq $worksheet.Cells[$rowNumb, $dayofMonth].Value -or $worksheet.Cells[$rowNumb, $dayofMonth].Value -eq '') {
                        Write-Host "Employee is marked as Excused, skipping..." -ForegroundColor Yellow
                        break
                    }
                    else {
                        # Sets the new value of the Cell
                        $worksheet.cells[$rowNumb, $dayofMonth].Value='N'
                    }
                    break
                }
        }
    }
        }
    

    }

    Write-Host 'Saving Changes, Please wait...' -ForegroundColor Yellow
    Start-Sleep -Seconds 15

    Write-Host 'Closing Excel Book' -ForegroundColor Yellow

    Start-Sleep -Seconds 5

################################################################################################
Close-ExcelPackage -ExcelPackage $workbook
################################################################################################

}

StartOneDrive
Start-Sleep -Seconds 15
RunAttendance

# Releasing data from memory
$workbook= $null
$bookpath= $null
$database= $null
$report= $null
$worksheet= $null
$newfile= $null

Exit-PSSession
