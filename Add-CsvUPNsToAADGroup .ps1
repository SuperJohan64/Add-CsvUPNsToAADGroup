# Turns on Transcription to log all activity.
$TimeStamp = Get-Date -Format "yyyyMMddTHHmmssffff"
$LogFile = "$PSScriptRoot\Logs\$TimeStamp.txt"
Start-Transcript -Path $LogFile

# A small function to end the transcription and close the script
function EndTranscription {
    Stop-Transcript
    Invoke-Item $LogFile
    Disconnect-AzureAD -ErrorAction SilentlyContinue
}

# Connects to Azure AD and gets AAD Group details.
Try {Connect-AzureAD -ErrorAction Stop}
Catch {
    EndTranscription
    Exit
}

# Prompts the user for an AAD Group's Object ID then gets the Group's detials.
$AADGroupID = Read-Host "Enter an AAD Group's Object ID"
Try {$AADGroup = Get-AzureADGroup -ObjectId $AADGroupID -ErrorAction Stop}
Catch {
    EndTranscription
    Exit
}

# Displays AAD Group Detials and other instructions to the user.
Write-Host "AAD Group Details:"
Write-Host "`nObject ID:" $AADGroup.ObjectId
Write-Host "DisplayName:" $AADGroup.DisplayName
Write-Host "Description:" $AADGroup.Description
Write-Host "`nProvide a Csv file containing a column titled 'UserPrincipalName' containing a list of UserPrincipalName's from AD or AAD."
PAUSE
Write-Host ""

# Launches an open file dialog window from .Net to select the CSV File.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $PSScriptRoot
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.filename
$CsvFilePath = $OpenFileDialog.filename

# Imports the data from the CSV file.
Try {$CsvData = Import-Csv -Path $CsvFilePath -ErrorAction Stop}
Catch {
    EndTranscription
    Exit
}

# A loop that parses each user in the CSV file and adds them to the AAD Group.
foreach ($CsvUser in $CsvData) {
    $AadUser = Get-AzureADUser -ObjectId $CsvUser.UserPrincipalName
    Write-Host "Adding" $AadUser.DisplayName "(" $CsvUser.UserPrincipalName ")" "to" $AADGroup.DisplayName "in Azure AD."
    Add-AzureADGroupMember -ObjectId $AADGroupID -RefObjectId $AadUser.ObjectId
}

EndTranscription
