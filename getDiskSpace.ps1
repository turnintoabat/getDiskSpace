 <#
Script created by Brendan Sturges, reach out if you have any issues.
This script queries a file that the user chooses and queries for local disk names, free space, percent free & total drive size and exports to a txt file
#>
 
 
 Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).txt"
	return $nameWithExtension

}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName)

#open a file dialog window to save the output
$fileName = Save-File $fileName

#define "i" for progress bar
$i = 0
$ErrorActionPreference = 'Stop'

foreach ($server in $serverList){
	$server | out-file $fileName -append
	try{
		gwmi Win32_LogicalDisk -computername $server -Filter "DriveType=3" | select Name, FreeSpace,BlockSize,Size | % {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;$_.FreeSpace=($_.FreeSpace/1GB);$_.Size=($_.Size/1GB);$_} |
		Format-Table Name,@{n='Free Gb';e={'{0:N2}'-f $_.FreeSpace}}, @{n='Free %';e={'{0:N2}'-f $_.BlockSize}},@{n='Capacity Gb';e={'{0:N3}' -f $_.Size}} -AutoSize | out-file $fileName -Append
	}
	
	catch{
		if(Test-Connection -ComputerName $server -Count 2 -Quiet)
				{
				$ErrorMessage = $_.Exception.Message
				}
			else
				{
				$ErrorMessage = "Server is Offline`r`n"
				}
		$ErrorMessage | out-file $fileName -Append
	}
	$i++
	Write-Progress -activity "Checking server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)	
}
