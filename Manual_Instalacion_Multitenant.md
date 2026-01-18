# Guía de Instalación "Zero-Config" (Multi-tenant)

Esta guía explica cómo desplegar la aplicación en un nuevo colegio sin necesidad de editar archivos de código o `.env`.

## 1. Requisitos Iniciales
*   **Python 3.10+**.
*   **PowerShell 7.0+** (Módulos: `MicrosoftTeams`, `ImportExcel`).
*   **App Registration en Azure**: Necesitarás el Tenant ID, Client ID y Client Secret con permisos de `User.ReadWrite.All`, `Group.ReadWrite.All` y `Directory.ReadWrite.All`.

## 2. Instalación Rápida
1.  **Copiar la carpeta** del proyecto al servidor.
2.  **Preparar el entorno**:
    ```bash
    python -m venv venv
    .\venv\Scripts\activate
    pip install -r requirements.txt
    ```
3.  **Lanzar el sistema**:
    ```bash
    python app.py
    ```

## 3. Configuración Premium (Asistente Web)
1.  Abre tu navegador en `http://localhost:5000`.
2.  En la pantalla de Login, haz clic en **"Registrar Datos Nuevos"**.
3.  Sigue el **Asistente de 4 Pasos**:
    - **Paso 1 (Azure)**: Ingresa las llaves de tu tenant.
    - **Paso 2 (Colegio)**: Define el nombre del colegio y el año escolar (ej: 2026).
    - **Paso 3 (Seguridad)**: Crea tu usuario y contraseña de administrador. La contraseña se guardará de forma segura (hasheada).
    - **Paso 4 (Notificaciones)**: Configura el correo desde el que se enviarán las credenciales.
4.  Haz clic en **"Guardar Configuración"**.

## 4. Gestión del Ciclo Escolar
Cada enero o cierre de ciclo, simplemente entra al sistema, ve a **Configuración** y cambia el "Periodo Académico" al siguiente año. No necesitas reinstalar nada ni editar archivos.

> [!IMPORTANT]
> Los archivos en la carpeta `archivos/` son vitales. El archivo `.key` contiene la llave de cifrado de tus secretos de Azure. Mantenlo seguro.
