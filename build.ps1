param([string]$pythonVersion = "310")

# Crée un environnement virtuel
python -m venv .venv
.\.venv\Scripts\activate

# Installe les dépendances, y compris Flask et Flask-Cors
pip install pyinstaller pywin32 pyserial requests flask flask-cors

# Trouve le chemin des DLLs pywin32
$pywin32Path = (Get-Item .venv).FullName + "\Lib\site-packages\pywin32_system32"
$dllFiles = Get-ChildItem -Path $pywin32Path -Filter "*.dll" | Select-Object -ExpandProperty FullName

# Construit le tableau d'arguments
$pyInstallerArgs = @(
    "--onefile",
    "--name=OdmService",
    "--hidden-import=win32timezone",
    "--hidden-import=servicemanager",
    "--hidden-import=logging.handlers",
    "--hidden-import=flask",
    "--hidden-import=flask_cors",
    "--runtime-tmpdir=.",
    "--icon=NONE"
)

# Ajoute les DLLs au format correct
foreach ($dll in $dllFiles) {
    $pyInstallerArgs += "--add-data=$dll;."
}

# Ajoute le script principal
$pyInstallerArgs += "OdmService.py"

# Affiche la commande pour débogage
Write-Host "Exécution de PyInstaller avec les arguments:"
$pyInstallerArgs | ForEach-Object { Write-Host "  $_" }

# Compile l'exécutable
pyinstaller @pyInstallerArgs

# Vérifie la création de l'exécutable
$exePath = ".\dist\OdmService.exe"
if (Test-Path $exePath) {
    # Crée le package de déploiement
    $deployDir = "OdmService_Deploy"
    New-Item -ItemType Directory -Path $deployDir -Force
    Copy-Item -Path $exePath -Destination $deployDir
    Copy-Item -Path ".\install_service.ps1" -Destination $deployDir
    Copy-Item -Path ".\uninstall_service.ps1" -Destination $deployDir

    # Compresse le package
    Compress-Archive -Path "$deployDir\*" -DestinationPath "OdmService_Package.zip" -Force
    Write-Host "✅ Package créé: OdmService_Package.zip"
} else {
    Write-Host "❌ ERREUR: OdmService.exe non généré. Voir le log PyInstaller."
    exit 1
}
