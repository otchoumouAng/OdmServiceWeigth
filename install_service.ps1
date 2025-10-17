param(
    [string]$installPath = "C:\Program Files\OdmService"
)

# Vérifier les privilèges admin
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`" -installPath `"$installPath`"" -Verb RunAs
    exit
}

# Crée le répertoire d'installation
New-Item -ItemType Directory -Force -Path $installPath

# Copie les fichiers
Copy-Item -Path ".\OdmService.exe" -Destination $installPath -Force

# Crée le répertoire de logs
$logDir = "$installPath\logs"
New-Item -ItemType Directory -Force -Path $logDir

# Installe le service
Start-Process -FilePath "$installPath\OdmService.exe" -ArgumentList "install" -Wait

# Configure le redémarrage automatique
sc.exe failure "OdmService" reset= 60 actions= restart/1000/restart/1000/restart/1000

# Démarrer le service
Start-Service -Name "OdmService"

# Configure les permissions NTFS
$acl = Get-Acl $installPath
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "Tout le monde", 
    "Modify", 
    "ContainerInherit,ObjectInherit", 
    "None", 
    "Allow"
)
$acl.AddAccessRule($rule)
Set-Acl -Path $installPath -AclObject $acl

# Affiche l'état
Get-Service -Name "OdmService"
Write-Host "Installation terminée. Logs: $logDir"
Write-Host "Redémarrage automatique configuré pour les échecs"