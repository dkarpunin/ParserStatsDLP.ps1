$SourceFile = "db-conn-3.txt"
$CSVFile = "db-conn-3.csv"

$Stats = Get-Content -Path $PSScriptRoot\$SourceFile -Encoding UTF8

$Stats = $Stats -replace "\s" -replace "\|", ";"

for ($i=1; $i -le 5; $i++){

    $Stats = $Stats[1..($Stats.length-1)]

}

Remove-Item -Path $PSScriptRoot\$CSVFile

$Out = "dd;keylog_count;keylog_capacity;clipboard_count;clipboard_capacity;removable_count;removable_capacity;exchange_count;exchange_capacity;im_count;im_capacity;printjob_count;printjob_capacity;email_count;email_capacity;device_count;device_capacity;network_count;networ_capacityk;application_count;application_capacity;file_count;file_capacity;all_count;all_capacity"

Out-File -InputObject $Out -FilePath $PSScriptRoot\$CSVFile -Encoding utf8

function Capacity($Capacity) {

    if ($Capacity -match "bytes") {$Capacity = [int]($Capacity.Trim("bytes")); $Capacity}
    if ($Capacity -match "kB") {$Capacity = [int]($Capacity.Trim("kB")) * 1024; $Capacity} 
    if ($Capacity -match "MB") {$Capacity = [int]($Capacity.Trim("MB")) * [math]::Pow(1024,2); $Capacity}
    if ($Capacity -match "GB") {$Capacity = [int]($Capacity.Trim("GB")) * [math]::Pow(1024,3); $Capacity}

}

foreach ($Stat in $Stats) {

    $Var = $Stat.Split(";")

    $dd = $Var[0]
    $keylog_count = $Var[1].Split("/")[0]
    $keylog_capacity = Capacity ($Var[1].Split("/")[1])
    $clipboard_count = $Var[2].Split("/")[0]
    $clipboard_capacity =  Capacity ($Var[2].Split("/")[1])
    $removable_count = $Var[3].Split("/")[0]
    $removable_capacity =  Capacity ($Var[3].Split("/")[1])
    $exchange_count = $Var[4].Split("/")[0]
    $exchange_capacity =  Capacity ($Var[4].Split("/")[1])
    $im_count = $Var[5].Split("/")[0]
    $im_capacity =  Capacity ($Var[5].Split("/")[1])
    $printjob_count = $Var[6].Split("/")[0]
    $printjob_capacity =  Capacity ($Var[6].Split("/")[1])
    $email_count = $Var[7].Split("/")[0]
    $email_capacity =  Capacity ($Var[7].Split("/")[1])
    $device_count = $Var[8].Split("/")[0]
    $device_capacity =  Capacity ($Var[8].Split("/")[1])
    $network_count = $Var[9].Split("/")[0]
    $networ_capacityk =  Capacity ($Var[9].Split("/")[1])
    $application_count = $Var[10].Split("/")[0]
    $application_capacity =  Capacity ($Var[10].Split("/")[1])
    $file_count = $Var[11].Split("/")[0]
    $file_capacity =  Capacity ($Var[11].Split("/")[1])
    $all_count = $Var[12].Split("/")[0]
    $all_capacity =  Capacity ($Var[12].Split("/")[1])

    $Out = "$dd;$keylog_count;$keylog_capacity;$clipboard_count;$clipboard_capacity;$removable_count;$removable_capacity;$exchange_count;$exchange_capacity;$im_count;$im_capacity;$printjob_count;$printjob_capacity;$email_count;$email_capacity;$device_count;$device_capacity;$network_count;$networ_capacityk;$application_count;$application_capacity;$file_count;$file_capacity;$all_count;$all_capacity"

    Out-File -InputObject $Out -FilePath $PSScriptRoot\$CSVFile -Encoding utf8 -Append

}