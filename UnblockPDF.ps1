# Unblock PDF files in a folder/subfolders that have changed recently and allows resume of progress. required for pdf Unblock of large archives.
# Chris Berry 31/01/2024
# saves last run details in separate file for reference:
# lastRanDate - indicates this full scan completed and now ok to run incrementally by modified date
# startDate - indicates the date to use to resume scan
# lastSubFolder - indicates the resume point for an interrupted scan
# 
$folderPath = Get-Location
$lastSubFolder = "" 
$newPath = "UnblockPDF.txt"
$startDate = Get-Date -Format 'yyyy-MM-dd'
$lastRanDate = ""
#Retreive defaults/progress

if (Test-Path -Path "$newPath") {
$values = Get-Content -Path "$newPath" | ForEach-Object {
    $var = $_.Split('=')
    New-Variable -Name $var[0] -Value $var[1] -Force    

}
} 


Write-Host "*  "
if($lastRanDate -ne "") {
Write-Host "* Sub-folder PDF/TIF Security Unblocker. BY FOLDER DATE. Run from Powershell console outside of Subfolders 1/2/3 etc. *"
} else {
Write-Host "* Sub-folder PDF/TIF Security Unblocker. FULL FOLDER SCAN. Run from Powershell console outside of Subfolders 1/2/3 etc. *"

}
Write-Host "*  "
Write-Host "* We are building the folder/file list. *"


if (Test-Path -Path $folderPath) {

if($lastRanDate -ne "") {
    
    # Date folder modified route.
    # Must ensure we use the sub-folder modified date NOT the archive folder mod date.
    $folders = Get-ChildItem -Path $folderPath -Directory | Where-Object { $_.FullName -ge $lastSubFolder }
    #Where-Object {$_.FullName -ge $lastSubFolder}
    Write-Host "* REGULAR SCAN MODE from " $startDate
    if ($lastSubFolder -ne "") {
    Write-Host "* Resuming from" $lastSubFolder
    }
    #MidPoint array built
    # Need extra search to dig out subfolders only
    #folder list   
    
        foreach ($folder in $folders) {
        #Write-Host $subfolder.FullName
        #NEW SUBFOLDER
        Write-Host "* We are now scanning/unblocking. * $folder"
        $subfolders = Get-ChildItem -Path $folder -recurse -Directory | Where-Object { $_.LastWriteTime  -gt $startDate } 

    #subfolder list   
        foreach ($subfolder in $subfolders) {
        Write-Host $subfolder.FullName
        $lastSubFolder =  $subfolder.FullName
           
            #Save folder progress in case we need to resume
            $hash = @{
            lastSubFolder = $folder.FullName #new
            lastRanDate = Get-Date
            startDate = $startDate
            }

            # Save folder progress
            $hash.GetEnumerator() | ForEach-Object {
            $_.Name + '=' + $_.Value
            } | Set-Content -Path $newPath 

            # File level
            $files = Get-ChildItem -Path $subfolder.FullName -recurse -include @("*.pdf") #| Sort-Object -Property Name 
              foreach ($file in $files) {
                Write-Host $file.FullName
                #Unblock
                unblock-file $file.FullName

              }
                
            }

        }


    } else {
    #All folders route.
    $subfolders = Get-ChildItem -Path $folderPath -recurse -Directory | Where-Object { $_.FullName -ge $lastSubFolder } #|  Sort-Object -Property Path #-Descending
    Where-Object {$_.FullName -ge $lastSubFolder}
    
    Write-Host "* Never completed here. FULL SCAN MODE*"
 
    #MidPoint array built
    Write-Host "* We are now scanning/unblocking. *"
        if ($lastSubFolder -ne "") {
        Write-Host "* Resuming from" $lastSubFolder
 
        }


    #folder list   
        foreach ($subfolder in $subfolders) {
        Write-Host $subfolder.FullName
        $lastSubFolder =  $subfolder.FullName
           
            #Save folder progress in case we need to resume
            $hash = @{
            lastSubFolder = $lastSubFolder
            startDate = $startDate
            }

            # Save folder progress
            $hash.GetEnumerator() | ForEach-Object {
            $_.Name + '=' + $_.Value
            } | Set-Content -Path $newPath 

            # File level
            $files = Get-ChildItem -Path $subfolder.FullName -recurse -include @("*.pdf") #| Sort-Object -Property Name 
              foreach ($file in $files) {
                Write-Host $file.FullName
                #Unblock
                unblock-file $file.FullName

              }
                
            }

}
}

#End - CLEAR the details for next full search
$lastSubFolder = ""
$endDate = Get-Date
#Use hash table for key-value pairs
$hash = @{
    lastSubFolder = $lastSubFolder
    lastRanDate = $endDate 
    startDate = Get-Date -Format 'yyyy-MM-dd'
    }

$hash.GetEnumerator() | ForEach-Object {
    $_.Name + '=' + $_.Value
} | Set-Content -Path $newPath 

Write-Host "*  "
Write-Host "* Unblock complete. Defaults saved to $newPath "