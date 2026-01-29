param(
    [Parameter(Mandatory = $true)]
    [string]$UPN
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " DIAGNÓSTICO DE POLÍTICAS TEAMS" -ForegroundColor Cyan
Write-Host " Usuario: $UPN" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

try {
    # Verificar conexión
    try {
        $session = Get-CsOnlineSession -ErrorAction SilentlyContinue
        if (-not $session) {
            Write-Host "Conectando a Microsoft Teams..." -ForegroundColor Yellow
            Connect-MicrosoftTeams
        }
    }
    catch {
        Write-Host "Error conectando a Teams. Asegúrate de tener el módulo instalado e internet." -ForegroundColor Red
        exit
    }

    Write-Host "Consultando información del usuario..." -ForegroundColor Yellow
    $user = Get-CsOnlineUser -Identity $UPN

    if (-not $user) {
        Write-Host "❌ Usuario no encontrado." -ForegroundColor Red
        exit
    }

    Write-Host "`n[1] POLÍTICA EFECTIVA (La que está funcionando)" -ForegroundColor Green
    Write-Host "------------------------------------------------"
    $policy = $user.TeamsMessagingPolicy
    if ([string]::IsNullOrWhiteSpace($policy)) {
        Write-Host "Política Actual: " -NoNewline
        Write-Host "Global (Org-wide default)" -ForegroundColor Yellow
        Write-Host "Nota: Si 'Global' permite chat, el usuario tendrá chat." -ForegroundColor Gray
    }
    else {
        Write-Host "Política Actual: " -NoNewline
        Write-Host "$policy" -ForegroundColor Cyan
        
        if ($policy -eq "Estudiantes_Bloqueo_Chat") {
            Write-Host "✅ La política correcta está asignada al usuario." -ForegroundColor Green
        }
        else {
            Write-Host "⚠️ El usuario tiene otra política asignada diferente a la esperada." -ForegroundColor Red
        }
    }

    Write-Host "`n[2] PAQUETE DE POLÍTICAS ASIGNADO" -ForegroundColor Green
    Write-Host "------------------------------------------------"
    # Verificar si tiene paquete asignado (requiere comando específico o ver propiedad)
    # Get-CsUserPolicyPackage no siempre está disponible o funciona diferente, usaremos info general si es posible
    # Nota: No hay propiedad directa simple en Get-CsOnlineUser para el paquete en todas las versiones, 
    # pero podemos inferirlo de las asignaciones si se usa PolicyPackage.
    
    $assignments = Get-CsUserPolicyAssignment -Identity $UPN
    if ($assignments) {
        Write-Host "Historial de Asignaciones (Directas y Grupo):"
        $assignments | Format-Table -Property PolicyType, PolicyName, AssignmentMethod -AutoSize
    }
    else {
        Write-Host "No se encontraron asignaciones directas explícitas." -ForegroundColor Gray
    }

    Write-Host "`n[3] RECOMENDACIONES" -ForegroundColor Green
    Write-Host "------------------------------------------------"
    Write-Host "1. Si en [1] ves 'Estudiantes_Bloqueo_Chat' pero el chat funciona:"
    Write-Host "   -> Verifica en Teams Admin Center que esa política tenga 'Chat = Off'."
    Write-Host "   -> Espera hasta 24-48h por propagación."
    Write-Host "   -> Prueba en Teams Web (Incógnito) para descartar caché."
    Write-Host "2. Si en [1] ves 'Global' u otra cosa:"
    Write-Host "   -> La asignación falló o fue sobrescrita. Vuelve a correr el script de asignación."

}
catch {
    $ErrorMsg = $_.Exception.Message
    Write-Host "Error inesperado: $ErrorMsg" -ForegroundColor Red
}
