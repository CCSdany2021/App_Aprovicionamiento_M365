# gestor_grupos_distribucion.ps1
# Script optimizado para agregar masivamente usuarios a grupos de seguridad/distribución
# Recibe un CSV y credenciales para evitar múltiples inicios de sesión.

param (
    [Parameter(Mandatory=$true)]
    [string]$RutaCsv,
    
    [Parameter(Mandatory=$true)]
    [string]$AdminUser,
    
    [Parameter(Mandatory=$true)]
    [string]$AdminPass,

    [Parameter(Mandatory=$true)]
    [string]$TenantId
)

# Configurar salida para capturar caracteres especiales
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "--- INICIO PROCESO DE SINCRONIZACION DE GRUPOS ---" -ForegroundColor Cyan
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"

# 1. Validar e Importar Módulo ExchangeOnlineManagement
try {
    # Intentar importar explícitamente primero para ver errores detallados
    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    Write-Host "[OK] Modulo ExchangeOnlineManagement importado correctamente."
} catch {
    Write-Host "[WARN] Fallo la importacion explicita: $_"
    Write-Host "Detalle Exception: $($_.Exception.Message)"
    # No salimos aqui, dejamos que el script intente continuar o instalar
}

if (-not (Get-Module -Name ExchangeOnlineManagement)) {
    Write-Host "Verificando instalacion..."
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "[ERROR] El módulo 'ExchangeOnlineManagement' no está instalado en este equipo." -ForegroundColor Red
        exit 1
    }
}

# 2. Conexión a Exchange Online (Intento robusto)
try {
    Write-Host "Conectando a Exchange Online..." -ForegroundColor Cyan
    
    # WORKAROUND: Construir SecureString con .NET puro para evitar 'ConvertTo-SecureString'
    # y los errores de carga del modulo Microsoft.PowerShell.Security
    $secPass = New-Object System.Security.SecureString
    $AdminPass.ToCharArray() | ForEach-Object { $secPass.AppendChar($_) }
    
    $cred = New-Object System.Management.Automation.PSCredential($AdminUser, $secPass)
    
    # Conectar (ShowProgress $false acelera proceso)
    Connect-ExchangeOnline -Credential $cred -ShowProgress $false -ErrorAction Stop
    
    Write-Host "[OK] Conexion Exitosa." -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] ERROR FATAL DE CONEXION: $_" -ForegroundColor Red
    exit 1
}

# 3. Procesamiento del CSV
if (-not (Test-Path $RutaCsv)) {
    Write-Host "[ERROR] No se encuentra el archivo CSV: $RutaCsv" -ForegroundColor Red
    exit 1
}

$datos = Import-Csv $RutaCsv
$total = $datos.Count
$actual = 0
$exitos = 0
$errores = 0

Write-Host "Comenzando procesamiento de $total registros..." -ForegroundColor Yellow

foreach ($fila in $datos) {
    $actual++
    $usuario = $fila.Member
    $grupo = $fila.Group
    
    if (-not $usuario -or -not $grupo) {
        Write-Host "[WARN] [$actual/$total] Fila invalida (Faltan datos). Saltando." -ForegroundColor Yellow
        continue
    }

    try {
        # Verificar si ya es miembro para evitar error de rojo innecesario
        # Nota: Get-DistributionGroupMember puede ser lento, para máxima velocidad intentamos agregar directo
        # y capturramos el error específico de "ya existe".
        
        Add-DistributionGroupMember -Identity $grupo -Member $usuario -ErrorAction Stop
        Write-Host "[OK] [$actual/$total] Agregado: $usuario -> $grupo" -ForegroundColor Green
        $exitos++
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -like "*already a member*") {
            Write-Host "[INFO] [$actual/$total] Ya existe: $usuario en $grupo" -ForegroundColor Gray
        }
        elseif ($msg -like "*Couldn't find object*") {
             Write-Host "[ERROR] [$actual/$total] No encontrado: Grupo '$grupo' o Usuario '$usuario' no existen." -ForegroundColor Red
             $errores++
        }
        else {
            Write-Host "[ERROR] [$actual/$total] Error: $usuario -> $grupo. Detalle: $msg" -ForegroundColor Red
            $errores++
        }
    }
}

Write-Host "--- RESUMEN FINAL ---" -ForegroundColor Cyan
Write-Host "Total Procesados: $total"
Write-Host "Exitosos: $exitos"
Write-Host "Errores: $errores"
Write-Host "--- FIN ---"
