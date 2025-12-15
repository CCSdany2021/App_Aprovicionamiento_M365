# Plantilla para Vaciar Equipos (Teams)

## Instrucciones de Uso

### Paso 1: Obtener el Inventario de Teams
1. Ve a la aplicación web
2. Haz clic en "Descargar Inventario de Teams"
3. Se descargará un archivo Excel con todos los equipos

### Paso 2: Preparar el Archivo
1. Abre el archivo de inventario descargado
2. Copia la columna **Id** de los equipos que quieres vaciar
3. Pega los IDs en un nuevo archivo CSV o Excel

### Paso 3: Formato del Archivo

#### Opción 1: Usar Team ID (RECOMENDADO)
```csv
TeamId,DisplayName
e1991d91-34fc-413f-9541-4e9d211b1908,Ejemplo Equipo 1
eb1887ba-4fed-4f74-bc55-a0a8fdd7c4f0,Ejemplo Equipo 2
```

**Ventajas:**
- ✅ Más confiable
- ✅ No hay problemas con caracteres especiales
- ✅ Funciona siempre

#### Opción 2: Usar Email (Compatible con versiones anteriores)
```csv
PrimarySmtpAddress
equipocurso1@calasanzsuba.edu.co
equipocurso2@calasanzsuba.edu.co
```

**Desventajas:**
- ⚠️ Puede fallar si el nombre tiene caracteres especiales
- ⚠️ Puede fallar si el equipo no tiene email configurado

### Columnas Aceptadas

El sistema buscará automáticamente una de estas columnas (en orden de prioridad):
1. `TeamId` o `Id` - **RECOMENDADO**
2. `GroupId`
3. `PrimarySmtpAddress`
4. `Email` o `Correo`

### Notas Importantes

> [!IMPORTANT]
> La columna `DisplayName` es opcional y solo sirve como referencia visual. El sistema usa únicamente el `TeamId` para identificar los equipos.

> [!WARNING]
> Este proceso eliminará TODOS los miembros y owners de los equipos especificados, excepto la cuenta `cap@calasanzsuba.edu.co`.

> [!TIP]
> Usa el inventario de Teams para obtener los IDs correctos. Los IDs son únicos y nunca cambian, mientras que los emails pueden cambiar.

## Ejemplo Completo

```csv
TeamId,DisplayName
e1991d91-34fc-413f-9541-4e9d211b1908,603 Dirección de grupo
eb1887ba-4fed-4f74-bc55-a0a8fdd7c4f0,701 Matemáticas
00399f0a-c320-4b57-bbec-afe0f26b3951,802 Biología
```

## Solución de Problemas

### "Equipo no encontrado"
- Verifica que el ID sea correcto (36 caracteres con guiones)
- Descarga un nuevo inventario para obtener IDs actualizados

### "Token expirado"
- El sistema ahora renueva automáticamente el token
- Si persiste, verifica las credenciales en el archivo `.env`

### "Error 401"
- El sistema reintentará automáticamente hasta 3 veces
- Verifica que las credenciales de la aplicación sean correctas
