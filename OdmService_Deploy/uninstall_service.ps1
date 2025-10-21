param(
    [string]$installPath = "C:\Program Files\OdmService"
)

# Vérifier les privilèges admin
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Demande d'élévation des privilèges..."
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" -installPath `"$installPath`"" -Verb RunAs
    exit
}

# Arrête le service
Write-Host "Arrêt du service..."
Stop-Service -Name "OdmService" -Force -ErrorAction SilentlyContinue

# Désinstalle le service
Write-Host "Désinstallation du service..."
Start-Process -FilePath "$installPath\OdmService.exe" -ArgumentList "remove" -Wait

# Supprime les fichiers
Write-Host "Nettoyage des fichiers..."
Remove-Item -Path $installPath -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n✅ Service désinstallé avec succès"