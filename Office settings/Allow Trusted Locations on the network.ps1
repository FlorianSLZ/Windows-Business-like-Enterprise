# Allow Trusted Locations on the network: Access
$RegistryPath = 'HKCU:\software\microsoft\office\16.0\access\security\trusted locations'
$Name         = 'allownetworklocations'
$Value        = '1'
$KeyType      = 'DWord'
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  
New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType $KeyType -Force 

# Allow Trusted Locations on the network: Word
$RegistryPath = 'HKCU:\software\microsoft\office\16.0\word\security\trusted locations'
$Name         = 'allownetworklocations'
$Value        = '1'
$KeyType      = 'DWord'
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  
New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType $KeyType -Force 

# Allow Trusted Locations on the network: Excel
$RegistryPath = 'HKCU:\software\microsoft\office\16.0\excel\security\trusted locations'
$Name         = 'allownetworklocations'
$Value        = '1'
$KeyType      = 'DWord'
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  
New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType $KeyType -Force 

# Allow Trusted Locations on the network: PowerPoint
$RegistryPath = 'HKCU:\software\microsoft\office\16.0\powerpoint\security\trusted locations'
$Name         = 'allownetworklocations'
$Value        = '1'
$KeyType      = 'DWord'
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  
New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType $KeyType -Force 

# Allow Trusted Locations on the network: Visio
$RegistryPath = 'HKCU:\software\microsoft\office\16.0\visio\security\trusted locations'
$Name         = 'allownetworklocations'
$Value        = '1'
$KeyType      = 'DWord'
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  
New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType $KeyType -Force 