import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as mpatches
from datetime import datetime
import numpy as np
import io
import tempfile
import os

# Configurar matplotlib para mostrar texto en español
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

def generar_reporte_inscritos(df):
    """
    Genera el reporte PDF a partir del DataFrame
    """
    # Limpiar datos: eliminar filas con valores nulos en columnas importantes
    df = df.dropna(subset=['Nombre y apellidos completos', 'Curso de interés', 'Hora de inicio'])

    # Convertir la columna de fecha a datetime
    df['Hora de inicio'] = pd.to_datetime(df['Hora de inicio'])

    # Calcular estadísticas principales
    personas_unicas = df['Nombre y apellidos completos'].nunique()
    fecha_minima = df['Hora de inicio'].min()
    fecha_maxima = df['Hora de inicio'].max()
    fecha_elaboracion = datetime.now().strftime("%d/%m/%Y")

    # Contar inscripciones por curso
    inscripciones_por_curso = df.groupby('Curso de interés')['Nombre y apellidos completos'].nunique().sort_values(ascending=False)

    # Crear archivo temporal para el PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        ruta_pdf = tmp_file.name

    with PdfPages(ruta_pdf) as pdf:
        # Página 1: Información general y gráfica
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 14))

        # Información general en la parte superior
        ax1.axis('off')
        ax1.text(0.5, 0.9, 'REPORTE DE INSCRIPCIONES A CURSOS',
                ha='center', va='center', fontsize=20, fontweight='bold',
                transform=ax1.transAxes)

        # Información del reporte
        info_text = f"""
        INFORMACIÓN GENERAL:

        • Total de personas inscritas: {personas_unicas}
        • Fecha de inicio de inscripciones: {fecha_minima.strftime('%d/%m/%Y %H:%M')}
        • Fecha de fin de inscripciones: {fecha_maxima.strftime('%d/%m/%Y %H:%M')}
        • Período de inscripciones: {(fecha_maxima - fecha_minima).days} días

        Fecha de elaboración: {fecha_elaboracion}
        Elaborado por: Samuel Barrera Meza
        """

        ax1.text(0.1, 0.6, info_text, ha='left', va='top', fontsize=12,
                transform=ax1.transAxes, bbox=dict(boxstyle="round,pad=0.3",
                facecolor="lightblue", alpha=0.7))

        # Gráfica de barras
        cursos = inscripciones_por_curso.index
        valores = inscripciones_por_curso.values

        # Crear colores para las barras
        colors = plt.cm.Set3(np.linspace(0, 1, len(cursos)))

        bars = ax2.bar(range(len(cursos)), valores, color=colors, alpha=0.8, edgecolor='black')

        # Configurar la gráfica
        ax2.set_xlabel('Cursos', fontsize=12, fontweight='bold')
        ax2.set_ylabel('Número de Inscritos', fontsize=12, fontweight='bold')
        ax2.set_title('Inscripciones por Curso', fontsize=14, fontweight='bold', pad=20)

        # Rotar etiquetas del eje x para mejor legibilidad
        ax2.set_xticks(range(len(cursos)))
        ax2.set_xticklabels(cursos, rotation=45, ha='right', fontsize=10)

        # Agregar valores en las barras
        for i, (bar, valor) in enumerate(zip(bars, valores)):
            ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                    str(valor), ha='center', va='bottom', fontweight='bold')

        # Ajustar el layout
        ax2.grid(axis='y', alpha=0.3)
        ax2.set_ylim(0, max(valores) * 1.1)

        plt.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Páginas adicionales: Tablas por curso
        for curso in inscripciones_por_curso.index:
            fig, ax = plt.subplots(figsize=(11, 8))
            ax.axis('off')

            # Filtrar datos por curso
            datos_curso = df[df['Curso de interés'] == curso][['Nombre y apellidos completos', 'Correo de contacto']].drop_duplicates()

            # Título de la tabla
            ax.text(0.5, 0.95, f'INSCRITOS EN: {curso}',
                   ha='center', va='top', fontsize=16, fontweight='bold',
                   transform=ax.transAxes)

            ax.text(0.5, 0.90, f'Total de inscritos: {len(datos_curso)}',
                   ha='center', va='top', fontsize=12,
                   transform=ax.transAxes)

            # Preparar datos para la tabla
            tabla_datos = []
            for idx, row in datos_curso.iterrows():
                tabla_datos.append([row['Nombre y apellidos completos'], row['Correo de contacto']])

            # Crear tabla
            if tabla_datos:
                tabla = ax.table(cellText=tabla_datos,
                               colLabels=['Nombre y Apellidos Completos', 'Correo de Contacto'],
                               cellLoc='left',
                               loc='center',
                               colWidths=[0.5, 0.5])

                # Configurar estilo de la tabla
                tabla.auto_set_font_size(False)
                tabla.set_fontsize(9)
                tabla.scale(1, 2)

                # Estilo del encabezado
                for i in range(len(tabla_datos[0])):
                    tabla[(0, i)].set_facecolor('#4CAF50')
                    tabla[(0, i)].set_text_props(weight='bold', color='white')

                # Alternar colores de filas
                for i in range(1, len(tabla_datos) + 1):
                    for j in range(len(tabla_datos[0])):
                        if i % 2 == 0:
                            tabla[(i, j)].set_facecolor('#f0f0f0')
                        else:
                            tabla[(i, j)].set_facecolor('white')

            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()

    return ruta_pdf, personas_unicas, fecha_minima, fecha_maxima, inscripciones_por_curso

def main():
    """
    Función principal de la aplicación Streamlit
    """
    st.set_page_config(
        page_title="Generador de Reportes de Inscripciones",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Generador de Reportes de Inscripciones")
    st.markdown("---")
    
    # Sidebar con información
    st.sidebar.header("ℹ️ Información")
    st.sidebar.info(
        "Esta aplicación genera reportes en PDF con las estadísticas "
        "de inscripciones a cursos a partir de un archivo Excel."
    )
    
    st.sidebar.markdown("### 📋 Columnas requeridas:")
    st.sidebar.markdown("""
    - **Nombre y apellidos completos**
    - **Curso de interés**
    - **Hora de inicio**
    - **Correo de contacto**
    """)
    
    # Área principal
    st.header("📂 Cargar Archivo Excel")
    
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel con los datos de inscripciones",
        type=['xlsx', 'xls'],
        help="El archivo debe contener las columnas: 'Nombre y apellidos completos', 'Curso de interés', 'Hora de inicio', 'Correo de contacto'"
    )
    
    if uploaded_file is not None:
        try:
            # Leer el archivo Excel
            df = pd.read_excel(uploaded_file)
            
            st.success("✅ Archivo cargado exitosamente")
            
            # Mostrar información del archivo
            st.subheader("📋 Información del Archivo")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total de registros", len(df))
            
            with col2:
                st.metric("Número de columnas", len(df.columns))
            
            with col3:
                registros_validos = len(df.dropna(subset=['Nombre y apellidos completos', 'Curso de interés', 'Hora de inicio']))
                st.metric("Registros válidos", registros_validos)
            
            # Mostrar columnas disponibles
            st.subheader("🔍 Columnas Disponibles")
            st.write(df.columns.tolist())
            
            # Verificar columnas requeridas
            columnas_requeridas = ['Nombre y apellidos completos', 'Curso de interés', 'Hora de inicio', 'Correo de contacto']
            columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
            
            if columnas_faltantes:
                st.error(f"❌ Faltan las siguientes columnas: {', '.join(columnas_faltantes)}")
                st.stop()
            
            # Mostrar vista previa de los datos
            st.subheader("👀 Vista Previa de los Datos")
            st.dataframe(df.head(10))
            
            # Botón para generar reporte
            if st.button("📊 Generar Reporte PDF", type="primary"):
                with st.spinner("Generando reporte..."):
                    try:
                        # Generar el reporte
                        ruta_pdf, personas_unicas, fecha_minima, fecha_maxima, inscripciones_por_curso = generar_reporte_inscritos(df)
                        
                        # Mostrar estadísticas
                        st.success("✅ Reporte generado exitosamente!")
                        
                        st.subheader("📈 Estadísticas Generales")
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.metric("Personas inscritas", personas_unicas)
                        
                        with col2:
                            st.metric("Fecha inicio", fecha_minima.strftime('%d/%m/%Y'))
                        
                        with col3:
                            st.metric("Fecha fin", fecha_maxima.strftime('%d/%m/%Y'))
                        
                        with col4:
                            st.metric("Período (días)", (fecha_maxima - fecha_minima).days)
                        
                        # Mostrar inscripciones por curso
                        st.subheader("📚 Inscripciones por Curso")
                        for curso, cantidad in inscripciones_por_curso.items():
                            st.write(f"• **{curso}**: {cantidad} personas")
                        
                        # Leer el archivo PDF generado
                        with open(ruta_pdf, "rb") as pdf_file:
                            pdf_bytes = pdf_file.read()
                        
                        # Botón de descarga
                        st.download_button(
                            label="📥 Descargar Reporte PDF",
                            data=pdf_bytes,
                            file_name=f"reporte_inscripciones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                            mime="application/pdf"
                        )
                        
                        # Limpiar archivo temporal
                        os.unlink(ruta_pdf)
                        
                    except Exception as e:
                        st.error(f"❌ Error al generar el reporte: {str(e)}")
                        
        except Exception as e:
            st.error(f"❌ Error al cargar el archivo: {str(e)}")
    else:
        st.info("👆 Por favor, carga un archivo Excel para comenzar")

if __name__ == "__main__":
    main()