$computers = Import-Csv -Path "computers.csv"
$source = "C:\FolderToMove"
$serviceName = "MyService"
$servicePath = "C:\DestinationFolder\MyService.exe"

foreach ($computer in $computers) {
    $destination = "\\$($computer.Name)\$servicePath"
    Move-Item $source $destination
    Invoke-Command -ComputerName $computer -ScriptBlock { param($servicePath) Invoke-Expression "sc.exe create $serviceName binPath=$servicePath start= auto" } -ArgumentList $servicePath
}


# Get-Service -Name ServiceName -ComputerName RemoteComputerName | Start-Service
