<#
.SYNOPSIS
    Asigna el paquete de políticas Education_SecondaryStudent
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ArchivoEstudiantes,

    [Parameter(Mandatory = $false)]
    [string]$Dominio = "calasanzsuba.edu.co"
)

$ErrorActionPreference = 'Continue'
$PackageName = "Education_SecondaryStudent"
$MessagingPolicyName = "Estudiantes_Bloqueo_Chat"
$CarpetaLogs = "resultados\logs"

if (-not (Test-Path $CarpetaLogs)) {
    New-Item -ItemType Directory -Path $CarpetaLogs -Force | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArchivoLog = Join-Path $CarpetaLogs "politicas_estudiantes_$Timestamp.log"

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = "INFO")
    $Fecha = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Linea = "[$Fecha] [$Nivel] $Mensaje"
    if ($Nivel -eq "ERROR") { Write-Host $Linea -ForegroundColor Red }
    elseif ($Nivel -eq "SUCCESS") { Write-Host $Linea -ForegroundColor Green }
    elseif ($Nivel -eq "WARNING") { Write-Host $Linea -ForegroundColor Yellow }
    else { Write-Host $Linea }
    Add-Content -Path $ArchivoLog -Value $Linea
}

function Import-ArchivoEstudiantes {
    param([string]$Ruta)
    Write-Log "Cargando archivo: $Ruta"
    try {
        $Ext = [System.IO.Path]::GetExtension($Ruta)
        if ($Ext -like ".xls*") { $D = Import-Excel -Path $Ruta }
        elseif ($Ext -eq ".csv") { $D = Import-Csv -Path $Ruta -Encoding UTF8 }
        else { throw "Formato invalido" }
        Write-Log "Datos cargados: $($D.Count)" -Nivel "SUCCESS"
        return $D
    }
    catch {
        Write-Log "Error cargando archivo: $_" -Nivel "ERROR"
        throw
    }
}

function Get-UserPrincipalName {
    param($Fila)
    $Cols = @('UserPrincipalName', 'UPN', 'Email', 'Mail', 'CODIGO', 'Codigo')
    foreach ($C in $Cols) {
        if ($Fila.PSObject.Properties.Name -contains $C) {
            $V = $Fila.$C
            if (-not [string]::IsNullOrWhiteSpace($V)) {
                if ($V -like "*@*") { return $V.Trim() }
                else { return "$($V.Trim())@$Dominio" }
            }
        }
    }
    return $null
}

# MAIN
Write-Host "INICIANDO ASIGNACION POLITICAS" -ForegroundColor Cyan
Write-Log "Archivo: $ArchivoEstudiantes"

# Check Module
if ($ArchivoEstudiantes -like "*.xls*") {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        try { Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber } catch {}
    }
}

try {
    $Estudiantes = Import-ArchivoEstudiantes -Ruta $ArchivoEstudiantes
}
catch {
    exit 1
}

if ($Estudiantes.Count -eq 0) { exit 1 }

# Connect Teams
try {
    if (-not (Get-Module -Name MicrosoftTeams)) {
        Write-Log "Importing MicrosoftTeams module..."
        Import-Module MicrosoftTeams -ErrorAction Stop
    }

    # Verify session by running a lightweight command
    try {
        # Check if we can reach the tenant
        $null = Get-CsOnlineUser -Top 1 -ErrorAction Stop
        Write-Log "Sesión detectada y activa."
    }
    catch {
        Write-Log "No hay sesión activa o expiró. Conectando..."
        Connect-MicrosoftTeams
    }
}
catch {
    Write-Log "Error conectando a Teams: $_" -Nivel "ERROR"
    exit 1
}

# Process
$Total = $Estudiantes.Count
$Count = 0
$Exitosos = 0
$Errores = 0
$Ya = 0
$Results = @()

foreach ($Est in $Estudiantes) {
    $Count++
    $UPN = Get-UserPrincipalName -Fila $Est
    
    if (-not $UPN) {
        Write-Log "[$Count/$Total] Sin UPN" -Nivel "WARNING"
        continue
    }

    Write-Host "[$Count/$Total] Procesando: $UPN" -NoNewline
    
    # 1. Asignar Paquete (Package)
    try {
        Grant-CsUserPolicyPackage -Identity $UPN -PackageName $PackageName -ErrorAction Stop
        Write-Host " [Pack OK]" -NoNewline -ForegroundColor Green
    }
    catch {
        $PckMsg = $_.Exception.Message
        if ($PckMsg -like "*Status: OK*") {
            # Known false positive in MicrosoftTeams module
            Write-Host " [Pack Ignored]" -NoNewline -ForegroundColor Yellow
            Write-Log "[$Count/$Total] Package warning ignored (Status: OK bug): $PckMsg" -Nivel "WARNING"
        }
        elseif ($PckMsg -like "*already*" -or $PckMsg -like "*existe*") {
            Write-Host " [Pack Ya]" -NoNewline -ForegroundColor Yellow
        }
        else {
            Write-Host " [Pack Error]" -NoNewline -ForegroundColor Red
            Write-Log "[$Count/$Total] Package Error: $PckMsg" -Nivel "ERROR"
            $Errores++
        }
    }

    # 2. Asignar Bloqueo de Chat (Chat Policy) - CRÍTICO
    try {
        Grant-CsTeamsMessagingPolicy -Identity $UPN -PolicyName $MessagingPolicyName -ErrorAction Stop
        Write-Host " [Chat BLOQUEADO]" -ForegroundColor Green
        Write-Log "[$Count/$Total] Chat Policy Success: $UPN" -Nivel "SUCCESS"
        $Exitosos++
        $Results += [PSCustomObject]@{ UPN = $UPN; Estado = "OK" }
    }
    catch {
        $ChatMsg = $_.Exception.Message
        if ($ChatMsg -like "*Status: OK*") {
            Write-Host " [Chat Ignored/OK]" -ForegroundColor Green
            Write-Log "[$Count/$Total] Chat Policy Success (Status: OK bug handled): $ChatMsg" -Nivel "SUCCESS"
            $Exitosos++
            $Results += [PSCustomObject]@{ UPN = $UPN; Estado = "OK" }
        }
        elseif ($ChatMsg -like "*already*" -or $ChatMsg -like "*existe*") {
            Write-Host " [Chat Ya]" -ForegroundColor Green
            $Exitosos++ # Safety achieved
        }
        else {
            Write-Host " [Chat ERROR]" -ForegroundColor Red
            Write-Log "[$Count/$Total] Chat Policy Error: $ChatMsg" -Nivel "ERROR"
        }
    }
    
    Start-Sleep -Milliseconds 100
}

Write-Log "FIN. Exitosos: $Exitosos, Errores: $Errores"
