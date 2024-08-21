
$regPath = "HKCU:\Software\Policies\Microsoft\office\16.0\excel\security\trusted locations"

Set-ItemProperty -Path $regPath -Name "allownetworklocations" -Value 1 -Type DWord

$locations = @(
    @{ Name = "Location10"; Path = "\\vault.local\Archive Data"; Description = "Z: Drive Archive Data" },
    @{ Name = "Location11"; Path = "\\vault.local\GIS Data"; Description = "G: Drive GIS Data" },
    @{ Name = "Location12"; Path = "\\vault.local\Project Data"; Description = "P: Drive Project Data" },
    @{ Name = "Location13"; Path = "\\vault.local\Kingdom Data"; Description = "K: Drive Kingdom Data" }
)

foreach ($location in $locations) {
    $locationName = $location.Name
    $locationPath = $location.Path
    $locationDescription = $location.Description

    $locationPathRegKey = "$regPath\$locationName"

    New-Item -Path $locationPathRegKey -Force | Out-Null
    Set-ItemProperty -Path $locationPathRegKey -Name "Path" -Value $locationPath -Type String
    Set-ItemProperty -Path $locationPathRegKey -Name "Description" -Value $locationDescription -Type String
    Set-ItemProperty -Path $locationPathRegKey -Name "AllowSubfolders" -Value 1 -Type DWord
}
