$WSB_FILE = "$PSScriptRoot\sandbox-config.wsb"
$wsbContent = Get-Content -Path $WSB_FILE -Raw
$updatedWsbContent = $wsbContent -replace '<local-path>', $PSScriptRoot
Set-Content -Path $WSB_FILE -Value $updatedWsbContent
Start-Process -FilePath "WindowsSandbox.exe" -ArgumentList $WSB_FILE