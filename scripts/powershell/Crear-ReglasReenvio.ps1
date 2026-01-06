param(
    [Parameter(Mandatory=$true)]
    [string]$ArchivoEntrada,
    [string]$Dominio = "calasanzsuba.edu.co"
)

$ErrorActionPreference = "Continue"

# Importar datos
$Datos = Import-Csv -Path $ArchivoEntrada

Write-Host "Iniciando creación de reglas de reenvío..."
$Total = $Datos.Count
$Count = 0

foreach ($Fila in $Datos) {
    $Count++
    # Priorizar UPN ya normalizado desde Python (lowercase)
    $UPN = $Fila.UPN.Trim().ToLower()
    
    if ([string]::IsNullOrWhiteSpace($UPN)) {
        Write-Host "   ⚠️  Fila $Count sin UPN válido - Saltando"
        continue
    }

    try {
        # 1. Limpiar reglas anteriores similares para evitar duplicados
        Get-InboxRule -Mailbox $UPN | Where-Object { $_.Name -like "Reenvio de Comunicados CCS*" } | Remove-InboxRule -Confirm:$false -Force

        # 2. Crear regla para Mamá
        if (-not [string]::IsNullOrWhiteSpace($Fila.CORREOMAMA)) {
            New-InboxRule -Name "Reenvio de Comunicados CCS - Mama" `
                -Mailbox $UPN `
                -RedirectTo $Fila.CORREOMAMA `
                -MarkAsRead $true `
                -ExceptIfSubjectOrBodyContainsWords @('Microsoft Teams', 'Cancelled:', 'Cancelado:') `
                -StopProcessingRules $false
            Write-Host "   ✅ Regla Mamá creada -> $($Fila.CORREOMAMA)"
        }

        # 3. Crear regla para Papá
        if (-not [string]::IsNullOrWhiteSpace($Fila.CORREOPAPA)) {
            New-InboxRule -Name "Reenvio de Comunicados CCS - Papa" `
                -Mailbox $UPN `
                -RedirectTo $Fila.CORREOPAPA `
                -MarkAsRead $true `
                -ExceptIfSubjectOrBodyContainsWords @('Microsoft Teams', 'Cancelled:', 'Cancelado:') `
                -StopProcessingRules $false
            Write-Host "   ✅ Regla Papá creada -> $($Fila.CORREOPAPA)"
        }

        # 4. Crear regla para Acudiente
        if (-not [string]::IsNullOrWhiteSpace($Fila.CORREOACUDIENTE)) {
            New-InboxRule -Name "Reenvio de Comunicados CCS - Acudiente" `
                -Mailbox $UPN `
                -RedirectTo $Fila.CORREOACUDIENTE `
                -MarkAsRead $true `
                -ExceptIfSubjectOrBodyContainsWords @('Microsoft Teams', 'Cancelled:', 'Cancelado:') `
                -StopProcessingRules $false
            Write-Host "   ✅ Regla Acudiente creada -> $($Fila.CORREOACUDIENTE)"
        }
    }
    catch {
        Write-Host "   ❌ Error en buzón $UPN : $($_.Exception.Message)"
    }
}

Write-Host "Proceso completado exitosamente."
