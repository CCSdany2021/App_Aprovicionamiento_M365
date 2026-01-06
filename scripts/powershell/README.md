# Scripts PowerShell - Asignación de Políticas Teams

Scripts automatizados para asignar paquetes de políticas de Microsoft Teams a estudiantes y docentes.

## 📋 Contenido

- `Asignar-PoliticasEstudiantes.ps1` - Asigna el paquete **Education_SecondaryStudent** a estudiantes
- `Asignar-PoliticasDocentes.ps1` - Asigna el paquete **Education_Teacher** a docentes

---

## 🔧 Requisitos Previos

### 1. Módulo MicrosoftTeams

```powershell
# Instalar el módulo de Microsoft Teams
Install-Module -Name MicrosoftTeams -Force -AllowClobber

# Verificar instalación
Get-Module -ListAvailable MicrosoftTeams
```

### 2. Módulo ImportExcel (para archivos .xlsx)

```powershell
# Instalar el módulo ImportExcel
Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
```

**Nota:** Los scripts instalarán automáticamente ImportExcel si no está disponible.

### 3. Permisos

Necesitas una cuenta con permisos de **Teams Administrator** o **Global Administrator** en Microsoft 365.

---

## 📂 Formato de Archivos

Los scripts soportan archivos **Excel (.xlsx)** o **CSV (.csv)** con las siguientes columnas (cualquiera de estas):

### Para Estudiantes:
- `CODIGO` - Código del estudiante (ej: 40302001)
- `UserPrincipalName` - Email completo (ej: 40302001@calasanzsuba.edu.co)
- `Email` / `UPN` / `Mail` - Variaciones del email

### Para Docentes:
- `UserPrincipalName` - Email del docente
- `Email` / `UPN` / `Mail` / `Correo`
- `CODIGO` - Código del docente (si aplica)

### Ejemplo de archivo Excel:

| CODIGO    | UserPrincipalName              |
|-----------|--------------------------------|
| 40302001  | 40302001@calasanzsuba.edu.co  |
| 40302002  | 40302002@calasanzsuba.edu.co  |

O simplemente:

| CODIGO    |
|-----------|
| 40302001  |
| 40302002  |

**Los scripts construirán automáticamente el email si solo se proporciona el código.**

---

## 🚀 Uso

### Asignar Políticas a Estudiantes

```powershell
# Desde la carpeta del proyecto
cd C:\aprovisionamientoEstudiantes\scripts\powershell

# Ejecutar el script
.\Asignar-PoliticasEstudiantes.ps1 -ArchivoEstudiantes "..\..\archivos\estudiantes.xlsx"
```

#### Con dominio personalizado:

```powershell
.\Asignar-PoliticasEstudiantes.ps1 -ArchivoEstudiantes "estudiantes.csv" -Dominio "micolegio.edu.co"
```

---

### Asignar Políticas a Docentes

```powershell
# Desde la carpeta del proyecto
cd C:\aprovisionamientoEstudiantes\scripts\powershell

# Ejecutar el script
.\Asignar-PoliticasDocentes.ps1 -ArchivoDocentes "..\..\archivos\docentes.xlsx"
```

#### Con dominio personalizado:

```powershell
.\Asignar-PoliticasDocentes.ps1 -ArchivoDocentes "docentes.csv" -Dominio "micolegio.edu.co"
```

---

## 📊 Resultados

Los scripts generan:

### 1. Log de Ejecución
Ubicación: `resultados/logs/politicas_estudiantes_YYYYMMDD_HHMMSS.log`

Contiene:
- Fecha y hora de cada operación
- Usuarios procesados exitosamente
- Errores detallados
- Resumen final

### 2. Archivo de Resultados (CSV)
Ubicación: `resultados/logs/resultados_estudiantes_YYYYMMDD_HHMMSS.csv`

Columnas:
- `Numero` - Número de fila procesada
- `UPN` - UserPrincipalName del usuario
- `Estado` - Exitoso / Ya asignado / Error / Sin UPN
- `Error` - Mensaje de error (si aplica)

---

## 🎯 Ejemplo de Ejecución

```powershell
PS C:\aprovisionamientoEstudiantes\scripts\powershell> .\Asignar-PoliticasEstudiantes.ps1 -ArchivoEstudiantes "..\..\archivos\estudiantes.xlsx"

================================================================================
  ASIGNACIÓN DE POLÍTICAS TEAMS - ESTUDIANTES
  Paquete: Education_SecondaryStudent
================================================================================

[2025-01-15 14:30:00] [INFO] Iniciando proceso de asignación de políticas para estudiantes
[2025-01-15 14:30:01] [SUCCESS] ✅ Archivo cargado: 150 registros
[2025-01-15 14:30:05] [SUCCESS] ✅ Conectado a Microsoft Teams

================================================================================
  PROCESANDO ESTUDIANTES
================================================================================

[1/150] Procesando: 40302001@calasanzsuba.edu.co ✅
[2/150] Procesando: 40302002@calasanzsuba.edu.co ✅
[3/150] Procesando: 40302003@calasanzsuba.edu.co ⚠️  Ya asignado
...

================================================================================
  RESUMEN DE ASIGNACIÓN
================================================================================

Total procesados:    150
✅ Exitosos:         145
⚠️  Ya asignados:    3
❌ Errores:          2

📝 Log guardado en: resultados\logs\politicas_estudiantes_20250115_143000.log

Proceso completado.
```

---

## ⚠️ Solución de Problemas

### Error: "Execution of scripts is disabled"

```powershell
# Habilitar ejecución de scripts (como Administrador)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Error: "Connect-MicrosoftTeams : The term is not recognized"

```powershell
# Instalar el módulo de Teams
Install-Module -Name MicrosoftTeams -Force -AllowClobber
```

### Error: "Access Denied" al asignar políticas

- Verifica que tu cuenta tenga permisos de **Teams Administrator**
- Contacta al administrador global de Microsoft 365

### Usuarios no encontrados

- Verifica que los UPN sean correctos
- Asegúrate de que los usuarios existan en Azure AD
- Revisa que el dominio configurado sea el correcto

---

## 📝 Notas Importantes

1. **Los scripts NO modifican la app web** - Son herramientas independientes para ejecutar manualmente
2. **Autenticación interactiva** - Te pedirá login de Microsoft 365 la primera vez
3. **Rate limiting** - Los scripts incluyen pausas de 200ms entre usuarios para no saturar la API
4. **Logs automáticos** - Todos los procesos quedan registrados
5. **Idempotencia** - Puedes ejecutar el script múltiples veces, detecta usuarios ya asignados

---

## 🔒 Seguridad

- **NUNCA** compartas los logs que contienen información de usuarios
- Asegúrate de tener permisos adecuados antes de ejecutar
- Verifica el archivo de entrada antes de procesarlo

---

## 📞 Soporte

Para problemas o dudas:
1. Revisa los logs generados en `resultados/logs/`
2. Verifica que tienes los módulos instalados
3. Confirma que tienes permisos de administrador de Teams

---

## 📄 Licencia

Propiedad de Colegio Calasanz Suba - Uso interno únicamente
