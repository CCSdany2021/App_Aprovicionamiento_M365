import pandas as pd
import os

def generar_plantillas():
    os.makedirs('plantillas', exist_ok=True)
    
    # 1. Plantilla Crear Estudiantes (SIN DOCUMENTO)
    cols_crear = ["CODIGO", "APELLIDOS", "NOMBRES", "GRADO", "CURSO", "CORREO_PERSONAL"]
    df_crear = pd.DataFrame(columns=cols_crear)
    # Add an example row
    df_crear.loc[0] = ["123456", "PEREZ LOPEZ", "JUAN CAMILO", "11", "1101", "juan.perez@gmail.com"]
    
    path_crear = 'plantillas/plantilla_crear_estudiantes.xlsx'
    df_crear.to_excel(path_crear, index=False)
    print(f"✅ Plantilla generada: {path_crear}")

    # 2. Plantilla Actualizar Estudiantes (SIN DOCUMENTO)
    cols_actualizar = ["CODIGO", "NOMBRES", "APELLIDOS", "CURSO"]
    df_actualizar = pd.DataFrame(columns=cols_actualizar)
    # Add an example row
    df_actualizar.loc[0] = ["123456", "JUAN CAMILO", "PEREZ LOPEZ", "1102"]
    
    path_actualizar = 'plantillas/plantilla_actualizar_estudiantes.xlsx'
    df_actualizar.to_excel(path_actualizar, index=False)
    print(f"✅ Plantilla generada (SIN DATOS SENSIBLES): {path_actualizar}")
    
    # 3. Plantilla Eliminar Estudiantes
    cols_eliminar = ["CODIGO"]
    df_eliminar = pd.DataFrame(columns=cols_eliminar)
    df_eliminar.loc[0] = ["123456"]
    
    path_eliminar = 'plantillas/plantilla_eliminar_estudiantes.xlsx'
    df_eliminar.to_excel(path_eliminar, index=False)
    print(f"✅ Plantilla generada: {path_eliminar}")

    # 4. Plantilla Manual Politicas (Opcional, pero util)
    cols_politicas = ["CODIGO"]
    df_politicas = pd.DataFrame(columns=cols_politicas)
    df_politicas.loc[0] = ["123456"]
    
    path_politicas = 'plantillas/plantilla_manual_politicas.xlsx'
    df_politicas.to_excel(path_politicas, index=False)
    print(f"✅ Plantilla generada: {path_politicas}")

if __name__ == "__main__":
    generar_plantillas()
