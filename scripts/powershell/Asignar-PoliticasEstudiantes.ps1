<#
.SYNOPSIS
    Asigna el paquete de políticas Education_SecondaryStudent a estudiantes

.DESCRIPTION
    Este script lee un archivo Excel o CSV con códigos de estudiantes y les asigna
    automáticamente el paquete de políticas 'Education_SecondaryStudent' de Microsoft Teams.

    Genera logs detallados y maneja errores de forma robusta.

.PARAMETER ArchivoEstudiantes
    Ruta al archivo Excel (.xlsx) o CSV (.csv) con la lista de estudiantes
    Columnas soportadas: CODIGO, UserPrincipalName, Email, UPN

.PARAMETER Dominio
    Dominio del correo institucional (ej: calasanzsuba.edu.co)
    Por defecto: calasanzsuba.edu.co

.EXAMPLE
    .\Asignar-PoliticasEstudiantes.ps1 -ArchivoEstudiantes "C:\archivos\estudiantes.xlsx"

.EXAMPLE
    .\Asignar-PoliticasEstudiantes.ps1 -ArchivoEstudiantes "estudiantes.csv" -Dominio "micolegio.edu.co"

.NOTES
    Autor: Aprovisionamiento Estudiantes
    Fecha: 2025
    Requiere: Módulo MicrosoftTeams instalado
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Ruta al archivo Excel o CSV con estudiantes")]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ArchivoEstudiantes,

    [Parameter(Mandatory=$false)]
    [string]$Dominio = "calasanzsuba.edu.co"
)

# =============================================================================
# CONFIGURACIÓN
# =============================================================================

$ErrorActionPreference = 'Continue'
$PackageName = "Education_SecondaryStudent"
$CarpetaLogs = "resultados\logs"

# Crear carpeta de logs si no existe
if (-not (Test-Path $CarpetaLogs)) {
    New-Item -ItemType Directory -Path $CarpetaLogs -Force | Out-Null
}

# Archivo de log con timestamp
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ArchivoLog = Join-Path $CarpetaLogs "politicas_estudiantes_$Timestamp.log"

# =============================================================================
# FUNCIONES
# =============================================================================

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = "INFO")

    $FechaHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LineaLog = "[$FechaHora] [$Nivel] $Mensaje"

    # Escribir en consola con color
    switch ($Nivel) {
        "ERROR"   { Write-Host $LineaLog -ForegroundColor Red }
        "SUCCESS" { Write-Host $LineaLog -ForegroundColor Green }
        "WARNING" { Write-Host $LineaLog -ForegroundColor Yellow }
        default   { Write-Host $LineaLog }
    }

    # Escribir en archivo
    Add-Content -Path $ArchivoLog -Value $LineaLog
}

function Import-ArchivoEstudiantes {
    param([string]$Ruta)

    Write-Log "Cargando archivo: $Ruta"

    try {
        $Extension = [System.IO.Path]::GetExtension($Ruta)

        if ($Extension -eq ".xlsx" -or $Extension -eq ".xls") {
            # Cargar Excel
            $Datos = Import-Excel -Path $Ruta
        }
        elseif ($Extension -eq ".csv") {
            # Cargar CSV
            $Datos = Import-Csv -Path $Ruta -Encoding UTF8
        }
        else {
            throw "Formato no soportado: $Extension. Use .xlsx o .csv"
        }

        Write-Log "✅ Archivo cargado: $($Datos.Count) registros" -Nivel "SUCCESS"
        return $Datos
    }
    catch {
        Write-Log "❌ Error cargando archivo: $_" -Nivel "ERROR"
        throw
    }
}

function Get-UserPrincipalName {
    param($Fila)

    # Intentar obtener UPN de diferentes columnas posibles
    $ColumnasPosibles = @('UserPrincipalName', 'UPN', 'Email', 'Mail', 'CODIGO', 'Codigo', 'codigo')

    foreach ($Columna in $ColumnasPosibles) {
        if ($Fila.PSObject.Properties.Name -contains $Columna) {
            $Valor = $Fila.$Columna

            if (-not [string]::IsNullOrWhiteSpace($Valor)) {
                # Si ya contiene @, es un UPN completo
                if ($Valor -like "*@*") {
                    return $Valor.Trim()
                }
                # Si es solo código, construir UPN
                else {
                    return "$($Valor.Trim())@$Dominio"
                }
            }
        }
    }

    return $null
}

# =============================================================================
# SCRIPT PRINCIPAL
# =============================================================================

Write-Host ""
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "  ASIGNACIÓN DE POLÍTICAS TEAMS - ESTUDIANTES" -ForegroundColor Cyan
Write-Host "  Paquete: $PackageName" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

Write-Log "Iniciando proceso de asignación de políticas para estudiantes"
Write-Log "Archivo de entrada: $ArchivoEstudiantes"
Write-Log "Dominio: $Dominio"

# Verificar módulo ImportExcel para archivos .xlsx
if ($ArchivoEstudiantes -like "*.xlsx" -or $ArchivoEstudiantes -like "*.xls") {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "⚠️  Módulo ImportExcel no instalado. Instalando..." -Nivel "WARNING"
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
            Write-Log "✅ Módulo ImportExcel instalado" -Nivel "SUCCESS"
        }
        catch {
            Write-Log "❌ Error instalando ImportExcel: $_" -Nivel "ERROR"
            Write-Host ""
            Write-Host "Por favor, instale el módulo manualmente:" -ForegroundColor Yellow
            Write-Host "Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
            exit 1
        }
    }
}

# 1. CARGAR ARCHIVO
# =============================================================================
try {
    $Estudiantes = Import-ArchivoEstudiantes -Ruta $ArchivoEstudiantes
}
catch {
    Write-Log "❌ No se pudo cargar el archivo. Abortando." -Nivel "ERROR"
    exit 1
}

if ($Estudiantes.Count -eq 0) {
    Write-Log "❌ El archivo está vacío" -Nivel "ERROR"
    exit 1
}

# 2. CONECTAR A MICROSOFT TEAMS
# =============================================================================
Write-Host ""
Write-Log "Verificando conexión a Microsoft Teams..."

try {
    # Verificar si hay sesión activa
    $Sesion = Get-CsOnlineSession -ErrorAction SilentlyContinue

    if (-not $Sesion) {
        Write-Log "🔐 No hay sesión activa. Conectando a Microsoft Teams..." -Nivel "WARNING"
        Connect-MicrosoftTeams
        Write-Log "✅ Conectado a Microsoft Teams" -Nivel "SUCCESS"
    }
    else {
        Write-Log "✅ Sesión activa de Microsoft Teams detectada" -Nivel "SUCCESS"
    }
}
catch {
    Write-Log "❌ Error conectando a Teams: $_" -Nivel "ERROR"
    Write-Host ""
    Write-Host "Asegúrese de tener instalado el módulo MicrosoftTeams:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name MicrosoftTeams -Force -AllowClobber" -ForegroundColor Yellow
    exit 1
}

# 3. PROCESAR ESTUDIANTES
# =============================================================================
Write-Host ""
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "  PROCESANDO ESTUDIANTES" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

$Contador = 0
$Exitosos = 0
$Errores = 0
$YaAsignados = 0
$Total = $Estudiantes.Count

$ResultadosDetallados = @()

foreach ($Estudiante in $Estudiantes) {
    $Contador++

    # Obtener UPN
    $UPN = Get-UserPrincipalName -Fila $Estudiante

    if ([string]::IsNullOrWhiteSpace($UPN)) {
        Write-Log "[$Contador/$Total] ⚠️  Fila sin UPN válido - Saltando" -Nivel "WARNING"
        $Errores++
        $ResultadosDetallados += [PSCustomObject]@{
            Numero = $Contador
            UPN = "N/A"
            Estado = "Sin UPN"
            Error = "No se pudo determinar el UserPrincipalName"
        }
        continue
    }

    Write-Host "[$Contador/$Total] Procesando: $UPN" -NoNewline

    try {
        # Asignar paquete de políticas
        Grant-CsUserPolicyPackage -Identity $UPN -PackageName $PackageName -ErrorAction Stop

        Write-Host " ✅" -ForegroundColor Green
        Write-Log "[$Contador/$Total] ✅ Política asignada: $UPN" -Nivel "SUCCESS"
        $Exitosos++

        $ResultadosDetallados += [PSCustomObject]@{
            Numero = $Contador
            UPN = $UPN
            Estado = "Exitoso"
            Error = ""
        }
    }
    catch {
        $MensajeError = $_.Exception.Message

        # Verificar si ya tiene la política asignada
        if ($MensajeError -like "*already*" -or $MensajeError -like "*existe*") {
            Write-Host " ⚠️  Ya asignado" -ForegroundColor Yellow
            Write-Log "[$Contador/$Total] ⚠️  Ya tiene la política: $UPN" -Nivel "WARNING"
            $YaAsignados++

            $ResultadosDetallados += [PSCustomObject]@{
                Numero = $Contador
                UPN = $UPN
                Estado = "Ya asignado"
                Error = ""
            }
        }
        else {
            Write-Host " ❌" -ForegroundColor Red
            Write-Log "[$Contador/$Total] ❌ Error en $UPN : $MensajeError" -Nivel "ERROR"
            $Errores++

            $ResultadosDetallados += [PSCustomObject]@{
                Numero = $Contador
                UPN = $UPN
                Estado = "Error"
                Error = $MensajeError
            }
        }
    }

    # Pequeña pausa para no saturar la API
    Start-Sleep -Milliseconds 200
}

# 4. RESUMEN FINAL
# =============================================================================
Write-Host ""
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "  RESUMEN DE ASIGNACIÓN" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""
Write-Host "Total procesados:    $Total" -ForegroundColor White
Write-Host "✅ Exitosos:         $Exitosos" -ForegroundColor Green
Write-Host "⚠️  Ya asignados:    $YaAsignados" -ForegroundColor Yellow
Write-Host "❌ Errores:          $Errores" -ForegroundColor Red
Write-Host ""
Write-Host "📝 Log guardado en: $ArchivoLog" -ForegroundColor Cyan
Write-Host ""

Write-Log "=========================================="
Write-Log "RESUMEN FINAL"
Write-Log "=========================================="
Write-Log "Total: $Total"
Write-Log "Exitosos: $Exitosos"
Write-Log "Ya asignados: $YaAsignados"
Write-Log "Errores: $Errores"
Write-Log "=========================================="

# Guardar resultados detallados en CSV
$ArchivoResultados = Join-Path $CarpetaLogs "resultados_estudiantes_$Timestamp.csv"
$ResultadosDetallados | Export-Csv -Path $ArchivoResultados -NoTypeInformation -Encoding UTF8
Write-Log "📊 Resultados detallados guardados en: $ArchivoResultados"

Write-Host "Proceso completado." -ForegroundColor Green
Write-Host ""
