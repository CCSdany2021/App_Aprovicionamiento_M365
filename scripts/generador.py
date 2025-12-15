#!/usr/bin/env python3
"""
Generador de datos de prueba para estudiantes
Crea 200 registros para pruebas de creación y actualización
"""

import pandas as pd
import random
from datetime import datetime
import os

class GeneradorDatosPrueba:
    """Generador de datos de prueba para estudiantes"""
    
    def __init__(self):
        # Listas de nombres y apellidos colombianos
        self.nombres = [
            "Santiago", "Alejandra", "Miguel", "Paula", "Daniel", "Camila", "Sebastián", 
            "Isabella", "Juan", "Sofía", "Andrés", "Valeria", "Carlos", "Mariana", 
            "Diego", "Gabriela", "Luis", "Nicole", "David", "Andrea", "Felipe", 
            "Natalia", "Nicolás", "Juliana", "Alejandro", "María", "Gabriel", "Ana", 
            "Manuel", "Laura", "Ricardo", "Catalina", "Jorge", "Daniela", "Oscar", 
            "Fernanda", "Eduardo", "Carolina", "Antonio", "Valentina", "Francisco", 
            "Paola", "Rodrigo", "Lorena", "Esteban", "Melissa", "Mauricio", "Adriana",
            "Mateo", "Stephanie", "Kevin", "Tatiana", "Jhon", "Yesica", "Alexander",
            "Katherine", "Cristian", "Vanessa", "Jonathan", "Monica", "Freddy"
        ]
        
        self.apellidos = [
            "García", "Rodríguez", "Martínez", "López", "González", "Hernández", 
            "Pérez", "Sánchez", "Ramírez", "Torres", "Flores", "Rivera", "Gómez", 
            "Díaz", "Cruz", "Morales", "Ortiz", "Gutiérrez", "Jiménez", "Vargas",
            "Rojas", "Castro", "Ruiz", "Herrera", "Moreno", "Álvarez", "Romero",
            "Medina", "Aguilar", "Delgado", "Castillo", "Peña", "Reyes", "Vega",
            "León", "Ramos", "Guerrero", "Mendoza", "Espinoza", "Silva", "Campos",
            "Contreras", "Soto", "Figueroa", "Sandoval", "Navarro", "Cortés",
            "Muñoz", "Ríos", "Acosta", "Valencia", "Pineda", "Mosquera", "Cantor",
            "Ballesteros", "Quintero", "Mejía", "Cardona", "Henao", "Zapata"
        ]
        
        # Estructura de grados y cursos
        self.grados_cursos = {
            "Transicion": ["TR1", "TR2"],
            "Primero": ["101", "102"],
            "Segundo": ["201", "202", "203"],
            "Tercero": ["301", "302", "303"],
            "Cuarto": ["401", "402", "403"],
            "Quinto": ["501", "502", "503"],
            "Sexto": ["601", "602", "603"],
            "Septimo": ["701", "702", "703"],
            "Octavo": ["801", "802", "803"],
            "Noveno": ["901", "902", "903"],
            "Decimo": ["1001", "1002", "1003"],
            "Once": ["1101", "1102", "1103"]
        }
        
        # Mapeo de promoción de grados
        self.promocion_grados = {
            "Transicion": "Primero",
            "Primero": "Segundo", 
            "Segundo": "Tercero",
            "Tercero": "Cuarto",
            "Cuarto": "Quinto",
            "Quinto": "Sexto",
            "Sexto": "Septimo",
            "Septimo": "Octavo",
            "Octavo": "Noveno",
            "Noveno": "Decimo",
            "Decimo": "Once",
            "Once": "Graduado"
        }

    def generar_estudiantes_nuevos(self, cantidad: int = 200) -> pd.DataFrame:
        """Genera estudiantes nuevos para crear"""
        estudiantes = []
        
        for i in range(cantidad):
            # Generar código único
            codigo = 40302000 + i + 1
            
            # Generar documento único
            documento = 1223344556 + i
            
            # Seleccionar grado y curso aleatoriamente
            grado = random.choice(list(self.grados_cursos.keys()))
            curso = random.choice(self.grados_cursos[grado])
            
            # Generar nombres
            nombre = random.choice(self.nombres)
            segundo_nombre = random.choice(self.nombres) if random.random() > 0.6 else ""
            nombres_completos = f"{nombre} {segundo_nombre}".strip()
            
            # Generar apellidos
            primer_apellido = random.choice(self.apellidos)
            segundo_apellido = random.choice(self.apellidos)
            apellidos_completos = f"{primer_apellido} {segundo_apellido}"
            
            # Crear email
            email = f"{codigo}@calasanzsuba.edu.co"
            
            estudiante = {
                "CODIGO": codigo,
                "DOCUMENTO": documento,
                "GRADO": grado,
                "CURSO": curso,
                "APELLIDOS": apellidos_completos,
                "NOMBRES": nombres_completos,
                "USERPRINCIPALNAME": email
            }
            
            estudiantes.append(estudiante)
        
        return pd.DataFrame(estudiantes)

    def generar_estudiantes_actualizacion(self, df_nuevos: pd.DataFrame) -> pd.DataFrame:
        """Genera estudiantes para actualización (promovidos de grado)"""
        actualizados = []
        
        for _, estudiante in df_nuevos.iterrows():
            grado_actual = estudiante["GRADO"]
            
            # Promover al siguiente grado
            if grado_actual in self.promocion_grados:
                nuevo_grado = self.promocion_grados[grado_actual]
                
                # Si es graduado, omitir
                if nuevo_grado == "Graduado":
                    continue
                
                # Seleccionar nuevo curso
                nuevo_curso = random.choice(self.grados_cursos[nuevo_grado])
                
                estudiante_actualizado = {
                    "CODIGO": estudiante["CODIGO"],
                    "DOCUMENTO": estudiante["DOCUMENTO"],
                    "GRADO": nuevo_grado,
                    "CURSO": nuevo_curso,
                    "APELLIDOS": estudiante["APELLIDOS"],
                    "NOMBRES": estudiante["NOMBRES"]
                }
                
                actualizados.append(estudiante_actualizado)
        
        return pd.DataFrame(actualizados)

    def guardar_archivos(self, df_nuevos: pd.DataFrame, df_actualizacion: pd.DataFrame):
        """Guarda los archivos Excel en la carpeta archivos/"""
        
        # Crear carpeta si no existe
        os.makedirs("archivos", exist_ok=True)
        
        # Guardar archivo de estudiantes nuevos
        archivo_nuevos = "archivos/estudiantesNuevos_prueba.xlsx"
        with pd.ExcelWriter(archivo_nuevos, engine='openpyxl') as writer:
            df_nuevos.to_excel(writer, sheet_name='EstudiantesNuevos', index=False)
        
        # Guardar archivo de actualización
        archivo_actualizacion = "archivos/actualizacionEstudiantes_prueba.xlsx"
        with pd.ExcelWriter(archivo_actualizacion, engine='openpyxl') as writer:
            df_actualizacion.to_excel(writer, sheet_name='EstudiantesNuevos', index=False)
        
        print(f"✅ Archivos creados:")
        print(f"   - {archivo_nuevos} ({len(df_nuevos)} estudiantes)")
        print(f"   - {archivo_actualizacion} ({len(df_actualizacion)} estudiantes)")
        
        return archivo_nuevos, archivo_actualizacion

    def mostrar_resumen(self, df_nuevos: pd.DataFrame, df_actualizacion: pd.DataFrame):
        """Muestra resumen de los datos generados"""
        print("\n" + "="*60)
        print("📊 RESUMEN DE DATOS GENERADOS")
        print("="*60)
        print(f"📅 Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"🆕 Estudiantes nuevos: {len(df_nuevos)}")
        print(f"🔄 Estudiantes para actualizar: {len(df_actualizacion)}")
        
        print("\n📋 Distribución por grados (Nuevos):")
        distribucion = df_nuevos['GRADO'].value_counts().sort_index()
        for grado, cantidad in distribucion.items():
            print(f"   {grado}: {cantidad} estudiantes")
        
        print("\n📋 Distribución por grados (Actualización):")
        distribucion_act = df_actualizacion['GRADO'].value_counts().sort_index()
        for grado, cantidad in distribucion_act.items():
            print(f"   {grado}: {cantidad} estudiantes")
        
        print("\n🎯 Vista previa de estudiantes nuevos:")
        print(df_nuevos[['CODIGO', 'NOMBRES', 'APELLIDOS', 'GRADO', 'CURSO']].head())
        
        print("\n🎯 Vista previa de estudiantes actualizados:")
        print(df_actualizacion[['CODIGO', 'NOMBRES', 'APELLIDOS', 'GRADO', 'CURSO']].head())
        
        print("="*60)

def main():
    """Función principal para generar datos de prueba"""
    print("🎓 GENERADOR DE DATOS DE PRUEBA")
    print("📊 Generando 200 estudiantes para pruebas...")
    print("="*50)
    
    generador = GeneradorDatosPrueba()
    
    # Generar estudiantes nuevos
    print("🆕 Generando estudiantes nuevos...")
    df_nuevos = generador.generar_estudiantes_nuevos(200)
    
    # Generar estudiantes para actualización (promovidos)
    print("🔄 Generando estudiantes para actualización...")
    df_actualizacion = generador.generar_estudiantes_actualizacion(df_nuevos)
    
    # Guardar archivos
    print("💾 Guardando archivos...")
    generador.guardar_archivos(df_nuevos, df_actualizacion)
    
    # Mostrar resumen
    generador.mostrar_resumen(df_nuevos, df_actualizacion)
    
    print("\n✅ Datos de prueba generados exitosamente!")
    print("📁 Archivos guardados en la carpeta 'archivos/'")
    print("🚀 Ahora puedes probar tus scripts con estos datos")

if __name__ == "__main__":
    main()