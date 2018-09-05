
function createMyVM($vhdName,$memorySize,$vhdSize)
{
 $vhdPath = $vhdName + ".vhdx"
 New-VM –Name $vhdName -Generation 2 –MemoryStartupBytes $memorySize -Path "D:\Hyper-V\" -NewVHDPath $vhdPath -NewVHDSizeBytes $vhdSize –SwitchName Corpnet 
 
 #Configure VHDX  
 Write-Host -ForeColor Yellow "Copying Base Image for 2012 R2 with Patches. Be Patient. Very Patient"
 Remove-VMHardDiskDrive -VMName $vhdName -ControllerType SCSI -ControllerNumber 0 -ControllerLocation 0
 Remove-Item "D:\Hyper-V\$vhdName\Virtual Hard Disks\$vhdPath" -Force
 Copy-Item D:\Hyper-V\base2012image.vhdx "D:\Hyper-V\$vhdName\Virtual Hard Disks\$vhdPath"
 Add-VMHardDiskDrive -VMName $vhdName -ControllerType SCSI -ControllerNumber 0 -ControllerLocation 0 -Path "D:\Hyper-V\$vhdName\Virtual Hard Disks\$vhdPath"
 Start-VM $vhdName
}

function deleteMyVM($vmName)
{
 Stop-VM $vmName -TurnOff
 $vmPath = "D:\Hyper-V\" + $vmName + "\"
 Write-Host "Removing VM and Associated Files" -ForegroundColor Yellow
 Remove-VM -Name $vmName -Force -ErrorAction SilentlyContinue
 Remove-Item $vmPath -recurse -ErrorAction SilentlyContinue
}

#New-VmSwitch -Name Corpnet -SwitchType Private

<##>
createMyVM DC1 512MB 20GB
createMyVM WFE1 1GB 20GB
createMyVM APP1 2GB 20GB
createMyVM SQL1 4GB 30GB
createMyVM CLIENT1 2GB 80GB



<#
#Delete all the Farm Virtual Machines
deleteMyVM DC1
deleteMyVM WFE1
deleteMyVM APP1
deleteMyVM DC1
deleteMyVM SQL1
deleteMyVM CLIENT1
#>




