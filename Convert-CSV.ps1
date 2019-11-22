#requires -version 3
<#
.SYNOPSIS
  Convert CSV file from a password application to a custom format
.DESCRIPTION
  
  Script will convert a CSV based on conditions below.
  Output will be a CSV with converted values and a CSV with unknown values that couldn't be automatically converted

  Export CSV contains the following fields:
  Secret Name,Username,Password,URL,Notes,File,Folder

  Has to be converted to:
  Passportal ID (BLANK),Client Name,Credential,Username,Password,Description,Expires (Yes/No),Notes,URL,Folder(Optional)

  Due to the faulty nature of the export, the script has been made so it can concatenate CSV rows until a specified condition has been reached.
.INPUTS
  CSV file with data that needs to be converted
.OUTPUTS
  CSV file with converted data
  <Outputs if any, otherwise state None>
.NOTES
  Version:        1.0
  Author:         Bart Tacken
  Creation Date:  22-11-2019
  Purpose/Change: Initial script development
.PREREQUISITES 
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
  <Example explanation goes here>
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string    
Start-Transcript ('c:\windows\temp\' + $DateStr  + 'ConvertCSV.log') -Force # Start logging

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'
#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Enter customer and file details
$CustomerShortName = "<customer name>"
$NcentralCustomerName = "<customer name>"
$CredentialType = "Active Directory User"

# Reset some values to $Null
$ExportFile2 = $Null
$ExportFile = $Null
$ExportFile3 = $Null
$ExportFileNew = $Null
$PasswordArray = $Null

$ExportFile = Get-Content -Path "\TEST\$CustomerShortName.csv"
$NewCSVFile = "\TEST\$($CustomerShortName)_Converted_CSV.csv"
$ExportErrorCSVFile = "TEST\$($CustomerShortName)_ERROR_CSV.csv"

$PasswordArray = @() # Create an empty array
$NewLine = "`r`n" # Variable for adding a carriage return
#-----------------------------------------------------------[Execution]------------------------------------------------------------
$Collumstring = ($ExportFile | select -First 1) # Get all columns
$ExportFile2 = ($ExportFile | select -skip 1) # Export all rows to the file (excluding the header with column names)
[int]$Collumnumber = ($Collumstring.Split(",")).count # Extract column number
$FaultyRows = 0
$TotalRows = 0
$i = 0
$CompleteRowArray = @() # Array for adding up incomplete rows until the row is complete ("*\BEHEERKLANTEN\*")

# Loop through each entry in the export file 
ForEach ($Row in $ExportFile2) {
    $NewRowFlag = $Null # Flag for detecting if this is the first row
    $CompleteRowArray = $Null # Empty array for adding up incomplete rows until the row is complete ("*\BEHEERKLANTEN\*")
    [int]$CollumnRownumber = ($Row.split(",")).count # Extract column number for each row in the loop

    # Follow while loop if the current row does not end with "*\BEHEERKLANTEN\* or the CompleteRowArray is filled but does not have "*\BEHEERKLANTEN\* within it" 
    # This will create an completeRowArray array with the rows until it encounters a row with "*\BEHEERKLANTEN\
    While (($Row -notlike "*\BEHEERKLANTEN\*") -or (($CompleteRowArray -ne $Null) -and ($CompleteRowArray -notlike "*\BEHEERKLANTEN\*"))) {

        # ($Row -notlike "*\BEHEERKLANTEN\*") : run loop as long as there is no "*\BEHEERKLANTEN\*" in the row. This indicates that there row is not finished yet.
        # ($ExportFile[$i] -ne $null) for prevening a loop. This will detect the end of the input file.
        # ($CompleteRowArray -notlike "*\BEHEERKLANTEN\*"): Run loop as long as the CompleteRowArray does not contain \BEHEERKLANTEN. Break this loop when it does
        
        $Row = $Row -replace "`n|`r" # Remove the new line after each partial row entry. This way it can be concatenated to one row line.
        
        If ($NewRowFlag -eq $Null) { # If this isn't the first entry in the CompleteRowArray, make sure to add the current row and combine it with the next line.
            $CompleteRowArray += ($Row + ",") #+ $TempRow + "," } # Add current row to the array
            $CompleteRowArray += $ExportFile[($i + 1)] # Add the next row to the array
            $NewRowFlag = $True # Set flag to announce that this is not the first entry in the array anymore
        }
        Else {
            
            If ($ExportFileNew[$i + 1] -ne $null) {
                $CompleteRowArray += $ExportFile[($i)] + "," # If this isn't the first entry in the array, add the current line to the CompleteRowArray. 
            }
            Else {
                $CompleteRowArray += $ExportFile[($i)] 
            }
        }        
        $i++ # Add one to the counter
       
        # prevent adding new lines when the row already ended with "*\BEHEERKLANTEN\*"
        If ($CompleteRowArray -like "*\BEHEERKLANTEN\*") {
            
            #Write completeRowArray to the new content variable
            $ExportFileNew = $ExportFileNew + $CompleteRowArray + $NewLine
            $CompleteRowArray = $Null
        }

        # Exit out of while loop when the current(next) row is empty)
        Write-Host "Going through line [$i]..." -ForegroundColor Green
        If ($ExportFile[$i] -eq $null) {
            break # Break loop
        } 
        #Continue # Go to next item
        #$Row = $Null
    } #End While
    $i++

    # Prevent adding somethin that already is part of the array
    If ($ExportFileNew -notlike "*$Row*") {
        $ExportFileNew = $ExportFileNew + $Row + $NewLine
    }
}
$ExportFileNewArray = $ExportFileNew -split "`r`n" # Convert this string into array

ForEach ($Row in $ExportFileNewArray) {
    
    $Row = $Row.TrimEnd(',')
    [int]$CollumnRownumber = ($Row.split(",")).count

    If ($Collumnumber -ne $CollumnRownumber) {         # Row does not contain the right number of collums
        Write-Host ""
        write-host "This row has [$CollumnRownumber] columns and is therfore not equal to [$Collumnumber] columns:" -ForegroundColor Yellow
        write-host $Row -ForegroundColor White
        $FaultyRows++
        $Row | Out-File -FilePath $ExportErrorCSVFile -Append # Add current row to the export CSV file
        Continue
    }

    # Row has right number of collumns. 
    # Add to custom array
    $PasswordArray += New-Object -TypeName PSObject -Property @{ # Fill Array with custom objects
        'Passportal ID (BLANK)' = ""
        'Client Name' = $NcentralCustomerName
        'Credential' = $CredentialType
        'Username' = (($Row.split(","))[1])
        'Password' = (($Row.split(","))[2])
        'Description' = (($Row.split(","))[0])
        'URL' = (($Row.split(","))[3])
        'Expires (Yes/No)' = "No"
        'Notes' = (($Row.split(","))[4])
        'File' = ""
        'Folder(Optional)' = (($Row.split(","))[6])
    }
    $TotalRows++         
}

# Show all entries in Array
#$PasswordArray | select "Passportal ID (BLANK)","Client Name","Credential","Username","Password","Description","Expires (Yes/No)","Notes","URL","Folder(Optional)"
Write-Host "New converted file has [$TotalRows] rows." -ForegroundColor Green
Write-Host "New error file contains [$FaultyRows] rows" -ForegroundColor Magenta
$PasswordArray | Sort-Object Username | select "Passportal ID (BLANK)","Client Name","Credential","Username","Password","Description","Expires (Yes/No)","Notes","URL","Folder(Optional)" | Export-Csv $NewCSVFile -Delimiter ";" -Force -NoTypeInformation

Stop-Transcript
