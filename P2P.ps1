
#logs
Start-Transcript -OutputDirectory "<Live Server Folder Location>\P2P\Logs"
#Start-Transcript -OutputDirectory "<Debugging on local Folder Location>\POD\P2P\Logs"

######################################################

#Details: Program to extract Proof of Delivery data and Proof of Delivery images from P2P API
#Developer Name: Velington Fernandes

# REMEMBER: Base Folder is Drive:POD\ . <Debugging on local Folder Location> replaced with Live Server location on production, if in PS ISE , then change the folder location accordingly
#######################################################

#######################################################
#Variables declaration for Destinations to save#
#######################################################
#P2P CN# list
#$CNLISTP2P="<Debugging on local Folder Location>\POD\P2P\"
$CNLISTP2P= <Live Server Folder Location>\P2P\Script\P2PDOLIST.csv

#ImagesP2P
#$ImagesDestinationP2P="<Debugging on local Folder Location>\POD\P2P\images"
$ImagesDestinationP2P="<Live Server Folder Location>"

#LogsP2P
#$LogsDestinationP2P="<Debugging on local Folder Location>\POD\P2P\Logs"
$LogsDestinationP2P="<Live Server Folder Location>\P2P\Logs"

#ArchivesP2P
#$ArchivesDestinationP2P="<Debugging on local Folder Location>\POD\P2P\archive"
$ArchivesDestinationP2P="<Live Server Folder Location>\P2P\archive"

#CSVP2P
#$CsvDestinationP2P="<Debugging on local Folder Location>\POD\P2P\csv"
$CsvDestinationP2P="<Live Server Folder Location>\P2P\csv"

#CSVP2P merged temp location
#$CsvDestinationP2P_M="<Debugging on local Folder Location>\POD\P2P\archive\merged"
$CsvDestinationP2P_M="<Live Server Folder Location>\P2P\archive\merged"

#CSVP2P individual files
#$CsvDestinationP2P_I="<Debugging on local Folder Location>\POD\P2P\archive\individualCNcsv"
$CsvDestinationP2P_I="<Live Server Folder Location>\P2P\archive\individualCNcsv"

#CSVP2P image csv files
#$CsvDestinationP2P_PIC="<Debugging on local Folder Location>\POD\P2P\archive\imagecsv"
$CsvDestinationP2P_PIC="<Live Server Folder Location>\P2P\archive\imagecsv"

#CSVP2P csv file containing list of CN# fetched
#$CsvDestinationP2P_MINUS="<Debugging on local Folder Location>\POD\P2P\archive\CNnofetched"
$CsvDestinationP2P_MINUS="<Live Server Folder Location>\P2P\archive\CNnofetched"

#CSVP2P csv file containing CN# sent from ABAP
#$CsvDestinationP2P_CNLISTABAP="<Debugging on local Folder Location>\POD\P2P\archive\inboundDOList"
$CsvDestinationP2P_CNLISTABAP="<Live Server Folder Location>\P2P\archive\inboundDOList"

#JSONP2P
#$JSONDestinationP2P="<Debugging on local Folder Location>\POD\P2P\json"
$JSONDestinationP2P="<Live Server Folder Location>\P2P\json"
#######################################################

#######################################################
#POD CN# list manipulation
#######################################################

#Set-Location <Debugging on local Folder Location>\POD\P2P\Script
Set-Location <Live Server Folder Location>\P2P\Script\

$importP2PList = $CNLISTP2P
[System.IO.File]::WriteAllText($importP2PList, [System.IO.File]::ReadAllText($importP2PList) -replace '[\r\n]+$')

#$files = Get-Content -path <Debugging on local Folder Location>\POD\P2P\Script\P2PDOLIST.csv
$files = Get-Content -path <Live Server Folder Location>\P2P\Script\P2PDOLIST.csv

#######################################################
#Extraction Program into JSON
#######################################################

#Set-Location <Debugging on local Folder Location>\POD\P2P\Script
Set-Location <Live Server Folder Location>\P2P\Script\

$systemdate = Get-Date
$systemtime=Get-Date -Format ss
Write-Host $systemdate
Write-Host $systemtime

ForEach ($file in $files)
{
$cnNovar= $file
Write-Host $cnNovar
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Accept-Language", "en-US")
$body = "{`"TrackingNo`": `"$cnNovar`",`"Category`": `"CNNo`"}"

Write-Host "Fetching Data from P2P for $cnNovar"

$response = Invoke-RestMethod 'https://cmswebapi.azurewebsites.net/api/external/tracking' -Method 'POST' -Headers $headers -Body $body
$response | ConvertTo-Json | Out-File "<Live Server Folder Location>\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).json"

Write-Host "Fetched Data from P2P $cnNovar"

$pathToJsonFile = "<Live Server Folder Location>\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).json"
$pathToOutputFile = "<Live Server Folder Location>\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv"
#$pathToJsonFile = "<Debugging on local Folder Location>\POD\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).json"
#$pathToOutputFile = "<Debugging on local Folder Location>\POD\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv"

#######################################################
#Extraction Program from JSON to CSV
#######################################################

Write-Host "Converting data for $cnNovar from JSON to CSV"
Write-Host "Now converting multi level JSON data into flat structure so that Header and detail can be shown in 1 row"

((Get-Content -Path $pathToJsonFile) | ConvertFrom-Json) | ForEach-Object {

$cnNo = $_.cnNo
$customerRef = $_.customerRef
$checkPointCode = $_.checkPointCode
$reasonCode = $_.reasonCode
$currentStation = $_.currentStation
$nextStation = $_.nextStation
$latestStatus = $_.latestStatus
$details = $_.details | ForEach-Object {
    [pscustomobject] @{   

       'cnNo' = $cnNo
       'customerRef' = $customerRef
       'checkPointCode' = $checkPointCode
       'reasonCode' = $reasoncode
       'currentStation' = $currentStation
       'nextStation' = $nextStation
       'latestStatus' = $latestStatus
       'transactionTime' = $_.transactionTime
       'currentStationDetails' = $_. currentStation
       'nextStationDetails' = $_.nextStation
       'checkPointCodeDetails' = $_.checkPointCode
       'reasonCodeDetails' = $_.reasonCode
       'isFinalStatus' = $_.isFinalStatus
       'currentStatus' = $_.currentStatus
       'receiverName' = $_.receiverName
       'podImagePath' = $_.podImagePath
       'remark' = $_.remark

      }
}
}

$details | Export-CSV -Delimiter "|" $pathToOutputFile -NoTypeInformation

Write-Host "Converted data for $cnNovar from JSON to CSV"
#######################################################
#Extraction Program from JSON to CSV completed
#######################################################

#######################################################
#Clean the csv file start
#######################################################
Write-Host "remove double quotes generated for CSV file and replace comma with empty and then change pipe to commas"

$importP2PdataFile = get-content -Path $pathToOutputFile
$importP2PdataFile | Select-Object -Skip 0 |foreach {%{ $_ -replace '"', ''}} | foreach {%{ $_ -replace ',', ''}}| Set-Content <Live Server Folder Location>\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv
(Get-Content -Path <Live Server Folder Location>\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv).Replace('|',',') | Set-Content -Path <Live Server Folder Location>\P2P\Script\P2P_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv

Write-Host "csv converter causes additional CRLF at the end of the P2P data file for $cnNovar, below code is used to remove the extra CRLF at the end of file"

$importP2PdataFile2 = $pathToOutputFile
[System.IO.File]::WriteAllText($importP2PdataFile2, [System.IO.File]::ReadAllText($importP2PdataFile2) -replace '[\r\n]+$')

#######################################################
#Cleaning the csv file ends
#######################################################

#######################################################
#Extract 1 column with URL and download image starts
#######################################################
Write-Host "Extract 1 column from above file which contains the image URLs for $cnNovar. This file will be without header and no double quotes"

$headers=@("podImagePath")
Import-Csv -path $pathToOutputFile | Select $headers | Where-Object { $_.PSObject.Properties.Value -ne '' } | Export-Csv  -Path "<Live Server Folder Location>\P2P\Script\P2P_image_list_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv" -NoTypeInformation
$importP2Pimagefile = get-content "<Live Server Folder Location>\P2P\Script\P2P_image_list_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv"
$importP2Pimagefile | Select-Object -Skip 1 |foreach {%{ $_ -replace '"', ''}} | Set-Content "<Live Server Folder Location>\P2P\Script\P2P_image_list_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv"

Write-Host  "csv converter causes additional CRLF at the end of the URLs file for $cnNovar, below code us used to remove the extra CRLF at the end of file"
$importP2Pimagefile = "<Live Server Folder Location>\P2P\Script\P2P_image_list_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv"
[System.IO.File]::WriteAllText($importP2Pimagefile, [System.IO.File]::ReadAllText($importP2Pimagefile) -replace '[\r\n]+$')

Write-Host "Download Images from the URL file using loop for $cnNovar"

$list = Get-Content "<Live Server Folder Location>\P2P\Script\P2P_image_list_${cnNovar}_$(Get-date -Format yyyyMMddhhmm).csv"
$clnt = New-Object System.Net.WebClient
foreach($url in $list)
    {
       #Get the filename
       $filename2 = [System.IO.Path]::GetFileName($url)
       #Create the output path
       $file2 = [System.IO.Path]::Combine($pwd.Path, $filename2)

       Write-Host -NoNewline "Getting ""$url""... "

       #Download the file using the WebClient
       $clnt.DownloadFile($url, $file2)
       Write-Host "done."
    }

#######################################################
#Extract 1 column with URL and download image ends
#######################################################

$systemdate2 = Get-Date

Write-Host "The current date time is $systemdate2"

$systemtime2 = Get-Date -Format ss

Write-Host "The current seconds is $systemtime2"

$Difference = 63-$systemtime2
if($Difference -gt 10){

Write-Host "The no. of seconds to start is more than 10 seconds, so pausing only for 3 secs and starting the next loop"

Start-Sleep -Seconds 3
}
else{Write-Host "The no. of seconds to pause until the loop starts again is $Difference"
Start-Sleep -Seconds $Difference
}
}
Write-Host $systemdate2
#######################################################
#Clean the folder of empty files starts
#######################################################
Write-Host "Delete empty files with 0kb data"
Get-ChildItem -Path <Live Server Folder Location>\P2P\Script\ |Where-Object {$_.Length -eq 0}| Remove-Item

#######################################################
#Clean the folder of empty files ends
#######################################################

#######################################################
#Moving generated files to relevant location starts
#######################################################

Write-Host "Move Images, json and imagelist csv file to the archive location first"

#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\P2P_image*.csv -Destination $CsvDestinationP2P_PIC -Force
#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\*.jpeg -Destination $ImagesDestinationP2P -Force
#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\*.jpg -Destination $ImagesDestinationP2P -Force
#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\*.json -Destination $jsonDestinationP2P -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\P2P_image*.csv -Destination $CsvDestinationP2P_PIC -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\*.jpeg -Destination $ImagesDestinationP2P -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\*.jpg -Destination $ImagesDestinationP2P -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\*.json -Destination $jsonDestinationP2P -Force

Write-Host "Images, json and imagelist csv file moved"

#######################################################
#Moving generated files to relevant location ends
#######################################################

#######################################################
#Merge individual csv files into 1 csv file starts
#######################################################
Write-Host "Import all csv files and merge them into 1 file. This will create a raw file with double quotes and extra line at the end"

Get-ChildItem <Live Server Folder Location>\P2P\Script\P2P_*.csv |  ForEach-Object { Import-Csv $_ } | Where-Object { $_.isFinalStatus -eq 'True' }|Select-Object * -Unique |  Export-Csv <Live Server Folder Location>\P2P\Script\P2P_$(Get-date -Format yyyyMMddhh).csv -NoTypeInformation

Write-Host "Then adding program to remove double quotes generated from final CSV file"

$importP2PdataFileFinal = get-content -Path <Live Server Folder Location>\P2P\Script\P2P_$(Get-date -Format yyyyMMddhh).csv
$importP2PdataFileFinal | Select-Object -Skip 0 |foreach {%{ $_ -replace '"', ''}} | Set-Content <Live Server Folder Location>\P2P\Script\P2P_$(Get-date -Format yyyyMMddhh).csv

Write-Host "and finally csv converter causes additional CRLF at the end of the P2P data file for merged CSV, next step to remove the extra CRLF at the end of file"

$importP2PdataFileFinal2 = "<Live Server Folder Location>\P2P\Script\P2P_$(Get-date -Format yyyyMMddhh).csv"
[System.IO.File]::WriteAllText($importP2PdataFileFinal2, [System.IO.File]::ReadAllText($importP2PdataFileFinal2) -replace '[\r\n]+$')

Write-Host "merged file is now cleaned and ready"

#######################################################
#Merge individual csv files into 1 csv file ends
#######################################################

#new csv file of only 1 column CNno
$headers2=@("cnNo")
Import-Csv -Path "<Live Server Folder Location>\P2P\Script\P2P_$(Get-date -Format yyyyMMddhh).csv" | Select-Object -Property cnNo -Unique | Export-Csv  -Path "<Live Server Folder Location>\P2P\Script\P2P_FETCHEDcnNO_$(Get-date -Format yyyyMMddhhmm).csv" -NoTypeInformation
$importP2PdataFileFinal3 = get-content -Path <Live Server Folder Location>\P2P\Script\P2P_FETCHEDcnNO_$(Get-date -Format yyyyMMddhhmm).csv
$importP2PdataFileFinal3 | Select-Object -Skip 0 |foreach {%{ $_ -replace '"', ''}} | Set-Content <Live Server Folder Location>\P2P\Script\P2P_FETCHEDcnNO_$(Get-date -Format yyyyMMddhhmm).csv
$importP2PdataFileFinal4 = "<Live Server Folder Location>\P2P\Script\P2P_FETCHEDcnNO_$(Get-date -Format yyyyMMddhhmm).csv"
[System.IO.File]::WriteAllText($importP2PdataFileFinal4, [System.IO.File]::ReadAllText($importP2PdataFileFinal4) -replace '[\r\n]+$')


Write-Host "merged file is now being moved to the temp location first"

#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\P2P_202*.csv -Destination $CsvDestinationP2P_M -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\P2P_202*.csv -Destination $CsvDestinationP2P_M -Force
Get-ChildItem -Path <Live Server Folder Location>\P2P\archive\merged\*.csv | Rename-Item -NewName "P2P_$(Get-date -Format yyyyMMddhhmm).csv"
#cnNo fetched records moved to POD\P2P\archive\CNnofetched folder for checking if the DO is fetched or not

Write-Host "Move the CN# list sent by P2P Server into CNnofetched folder within archive"

#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\P2P_FETCHEDcnNO_*.csv -Destination $CsvDestinationP2P_MINUS -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\P2P_FETCHEDcnNO_*.csv -Destination $CsvDestinationP2P_MINUS -Force

Write-Host "individual csv files are now moved the individual csv file location inside archive"

#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\P2P_*.csv -Destination $CsvDestinationP2P_I -Force
Move-Item -Path <Live Server Folder Location>\P2P\Script\P2P_*.csv -Destination $CsvDestinationP2P_I -Force

Write-Host "now merged file brought back from temp into the csv folder"

#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\archive\merged\P2P_202*.csv -Destination $CsvDestinationP2P -Force
Move-Item -Path <Live Server Folder Location>\P2P\archive\merged\P2P_202*.csv -Destination $CsvDestinationP2P -Force
#<Debugging on local Folder Location>\POD\P2P\archive\inboundDOList to move the daily P2P CN list into archive folder
#Write-Host "Move the DO List sent by ABAP into archive folder"
#Move-Item -Path <Debugging on local Folder Location>\POD\P2P\Script\P2PDOLIST*.csv -Destination $CsvDestinationP2P_CNLISTABAP -Force
#Move-Item -Path <Live Server Folder Location>\P2P\Script\P2PDOLIST*.csv -Destination $CsvDestinationP2P_CNLISTABAP -Force

Write-Host "merged file is moved and P2P program ends"

Stop-Transcript 

