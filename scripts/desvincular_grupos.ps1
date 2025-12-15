<#
.SYNOPSIS
    Desvincula estudiantes de grupos de distribución/seguridad (Mail-Enabled).
    Este script es necesario porque Graph API no soporta esta modificación para este tipo de grupos.

.DESCRIPTION
    Lee el archivo CSV/Excel de grupos y elimina todos sus miembros usando Exchange Online.

.PARAMETER Archivo
    Ruta al archivo CSV o Excel con la columna 'PrimarySmtpAddress'.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$Archivo
)

# Importar módulos necesarios
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Warning "Módulo ImportExcel no encontrado. Se intentará instalar..."
    Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
}

# Conectar a Exchange Online si no hay sesión
if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' })) {
    Write-Host "🔐 Conectando a Exchange Online... Por favor inicie sesión." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowProgress $true
}

# Leer archivo
try {
    Write-Host "📂 Leyendo archivo: $Archivo" -ForegroundColor Gray
    
    if ($Archivo -match '\.xlsx$') {
        $datos = Import-Excel -Path $Archivo
    }
    elseif ($Archivo -match '\.csv$') {
        $datos = Import-Csv -Path $Archivo
    }
    else {
        Write-Error "Formato no soportado. Use .csv o .xlsx"
        exit
    }
}
catch {
    Write-Error "Error leyendo archivo: $_"
    exit
}

$total = $datos.Count
$actual = 0

foreach ($fila in $datos) {
    $actual++
    $grupoEmail = $fila.PrimarySmtpAddress
    
    if (-not $grupoEmail) {
        $grupoEmail = $fila.Email
    }
    
    if (-not $grupoEmail) {
        Write-Warning "[$actual/$total] Fila sin correo electrónico. Saltando..."
        continue
    }

    Write-Host "`n[$actual/$total] 🔍 Procesando grupo: $grupoEmail" -ForegroundColor Cyan
    
    try {
        # Verificar si el grupo existe
        $grupo = Get-DistributionGroup -Identity $grupoEmail -ErrorAction SilentlyContinue
        
        if (-not $grupo) {
            Write-Warning "   ⚪ Grupo no encontrado en Exchange"
            continue
        }
        
        # Obtener miembros
        $miembros = Get-DistributionGroupMember -Identity $grupoEmail -ResultSize Unlimited
        $countMiembros = $miembros.Count
        
        if ($countMiembros -eq 0) {
            Write-Host "   ⚪ Grupo vacío, nada que eliminar." -ForegroundColor Gray
            continue
        }
        
        Write-Host "   👥 Eliminando $countMiembros miembros..." -ForegroundColor Yellow
        
        foreach ($miembro in $miembros) {
            # Eliminar miembro
            Remove-DistributionGroupMember -Identity $grupoEmail -Member $miembro.Identity -Confirm:$false -ErrorAction SilentlyContinue
        }
        
        Write-Host "   ✅ Grupo vaciado correctamente." -ForegroundColor Green
        
    }
    catch {
        Write-Error "   ❌ Error procesando grupo: $_"
    }
}

Write-Host "`n✨ Proceso finalizado." -ForegroundColor Green
