$TEMPLATE_WSB_FILE = "$PSScriptRoot\template.wsb" 
$templateWsbFile = $TEMPLATE_WSB_FILE -replace "src", "misc"
$wsb_file = $templateWsbFile -replace "template", "sandbox-config"
$wsbContent = Get-Content -Path $templateWsbFile -Raw
$localPath = "$PSScriptRoot" -replace "\\src", ""
$updatedWsbContent = $wsbContent -replace '<local-path>', "$localPath"
Set-Content -Path $wsb_file -Value $updatedWsbContent -Force
Start-Process -FilePath "WindowsSandbox.exe" -ArgumentList $wsb_file