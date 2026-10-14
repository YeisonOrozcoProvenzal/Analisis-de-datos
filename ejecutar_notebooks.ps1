# Ejecucion diaria de notebooks - Analisis de datos Provenzal
$BaseDir = "C:\Users\Asus\OneDrive - PROVENZAL S.A.S\Escritorio\Analisis-de-datos"
$LogFile = "$BaseDir\logs\ejecucion_$(Get-Date -Format 'yyyy-MM-dd').log"

# Crear carpeta de logs si no existe
New-Item -ItemType Directory -Force -Path "$BaseDir\logs" | Out-Null

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

$notebooks = @(
    "LimpiezaClientes\Limpieza_Clientes.ipynb",
    "Deduplicacion Clientes\TABLA_MASTER_ID.ipynb",
    "LimpiezaClientes\Limpieza_Contacto_Clientes.ipynb"
    
)

Log "=== Inicio de ejecucion diaria ==="

foreach ($nb in $notebooks) {
    $fullPath = "$BaseDir\$nb"
    Log "Ejecutando: $nb"

    $result = jupyter nbconvert --to notebook --execute --inplace --ExecutePreprocessor.timeout=3600 "$fullPath" 2>&1

    if ($LASTEXITCODE -eq 0) {
        Log "OK: $nb"
    } else {
        Log "ERROR en $nb`: $result"
        # Continua con el siguiente aunque falle uno
    }
}

Log "=== Fin de ejecucion diaria ==="
