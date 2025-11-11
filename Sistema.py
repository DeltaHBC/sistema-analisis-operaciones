import customtkinter as CTk
from tkinter import filedialog, simpledialog, scrolledtext, Listbox, Toplevel, Button, Frame
import tkinter as tk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image
from fpdf import FPDF
import os
import webbrowser
import numpy as np
from datetime import datetime, timedelta
import tempfile
CTk.set_appearance_mode("light")
CTk.set_default_color_theme("blue")
class SistemaAnalisisOperaciones:
    def __init__(self):
        self.categorias_materiales = {
            'metales': ['aluminio', 'cobre', 'bronce', 'fierro', 'hierro', 'metal', 'rin', 'rlc', 'acero', 'inox'],
            'plasticos': ['pet', 'plastico', 'polietileno', 'pvc', 'bote'],
            'papel_carton': ['carton', 'papel', 'caja'],
            'variados': ['boiler', 'archivo', 'radiador']
        }
    def analizar_operaciones_semanales(self, datos_semana):
        """Analiza las operaciones de compra de toda la semana"""
        try:
            kg_por_dia = datos_semana.groupby(datos_semana['FECHA'].dt.date)['KG'].sum()
            dia_max_compra = kg_por_dia.idxmax() if not kg_por_dia.empty else None
            max_kg_dia = kg_por_dia.max() if not kg_por_dia.empty else 0
            material_por_dia = {}
            for fecha in datos_semana['FECHA'].dt.date.unique():
                datos_dia = datos_semana[datos_semana['FECHA'].dt.date == fecha]
                if not datos_dia.empty:
                    material_max = datos_dia.groupby('PRODUCTO')['KG'].sum().idxmax()
                    material_por_dia[fecha] = material_max
            material_semana = datos_semana.groupby('PRODUCTO')['KG'].sum()
            material_mas_comprado = material_semana.idxmax() if not material_semana.empty else None
            kg_material_mas_comprado = material_semana.max() if not material_semana.empty else 0
            material_menos_comprado = material_semana.idxmin() if not material_semana.empty else None
            kg_material_menos_comprado = material_semana.min() if not material_semana.empty else 0
            dinero_por_dia = datos_semana.groupby(datos_semana['FECHA'].dt.date)['TOTAL COMPRA'].sum()
            dinero_total_semana = datos_semana['TOTAL COMPRA'].sum()
            venta_utilidad_material = {}
            for producto in datos_semana['PRODUCTO'].unique():
                datos_producto = datos_semana[datos_semana['PRODUCTO'] == producto]
                kg_total = datos_producto['KG'].sum()
                compra_total = datos_producto['TOTAL COMPRA'].sum()
                precio_venta = self._obtener_precio_venta(producto)
                if precio_venta > 0:
                    venta_aproximada = kg_total * precio_venta
                    utilidad_aproximada = venta_aproximada - compra_total
                else:
                    venta_aproximada = 0
                    utilidad_aproximada = 0
                
                venta_utilidad_material[producto] = {
                    'kg': kg_total,
                    'compra': compra_total,
                    'venta_aproximada': venta_aproximada,
                    'utilidad_aproximada': utilidad_aproximada
                }
            venta_total_aproximada = sum(item['venta_aproximada'] for item in venta_utilidad_material.values())
            utilidad_total_aproximada = sum(item['utilidad_aproximada'] for item in venta_utilidad_material.values()) 
            return {
                'kg_totales_semana': datos_semana['KG'].sum(),
                'kg_por_dia': kg_por_dia.to_dict(),
                'dia_max_compra': dia_max_compra,
                'max_kg_dia': max_kg_dia,
                'material_por_dia': material_por_dia,
                'material_mas_comprado': material_mas_comprado,
                'kg_material_mas_comprado': kg_material_mas_comprado,
                'material_menos_comprado': material_menos_comprado,
                'kg_material_menos_comprado': kg_material_menos_comprado,
                'dinero_por_dia': dinero_por_dia.to_dict(),
                'dinero_total_semana': dinero_total_semana,
                'venta_utilidad_material': venta_utilidad_material,
                'venta_total_aproximada': venta_total_aproximada,
                'utilidad_total_aproximada': utilidad_total_aproximada,
                'total_materiales': len(datos_semana['PRODUCTO'].unique()),
                'periodo_analizado': f"{datos_semana['FECHA'].min().strftime('%d/%m/%Y')} - {datos_semana['FECHA'].max().strftime('%d/%m/%Y')}"
            }
        except Exception as e:
            print(f"Error en análisis semanal: {e}")
            return {}
    def _obtener_precio_venta(self, producto):
        """Obtiene el precio de venta del producto"""
        try:
            precios_estimados = {
                'CARTON': 2.0, 'PET': 6.7, 'BOTE': 32.0, 'ALUMINIO': 27.0,
                'FIERRO': 4.5, 'CU 1': 184.5, 'CU 2': 177.5, 'BRONCE': 123.0,
                'ACERO INOX': 4.5, 'R/A': 30.0, 'ALUMINIO GRUESO (L)': 39.5,
                'ALUMINIO TRASTE (L)': 38.5, 'BRONCE (L)': 123.0, 'BRONCE (S)': 123.0,
                'RADIADOR LIN COBRE': 84.0, 'CAJA BOILER': 84.0, 'RIN ALUMINIO': 45.0,
                'ALUMINIO DELGADO': 30.5, 'ALUMINIO TUBO': 32.5, 'ALUMINIO CABLE (L)': 60.0,
                'ALUMINIO SPRAY (S)': 48.0, 'ALUMINIO PERFIL S/PINTURA (L)': 58.5,
                'ALUMINIO PERFIL C/PINTURA (L)': 48.5, 'CU ESTAÑADO': 180.5
            }
            producto_upper = str(producto).upper().strip()
            for key in precios_estimados:
                if key in producto_upper:
                    return precios_estimados[key]
            if 'CARTON' in producto_upper:
                return 2.0
            elif 'PET' in producto_upper:
                return 6.7
            elif 'BOTE' in producto_upper:
                return 32.0
            elif 'ALUMINIO' in producto_upper:
                return 27.0
            elif 'FIERRO' in producto_upper:
                return 4.5
            elif 'CU 1' in producto_upper or 'COBRE 1' in producto_upper:
                return 184.5
            elif 'CU 2' in producto_upper or 'COBRE 2' in producto_upper:
                return 177.5
            elif 'BRONCE' in producto_upper:
                return 123.0
            elif 'ACERO' in producto_upper or 'INOX' in producto_upper:
                return 4.5
            elif 'RADIADOR' in producto_upper:
                return 84.0
            else:
                return 15.0  
        except:
            return 15.0
global datos_operaciones, analizador_operaciones, analisis_actual
datos_operaciones = None
analizador_operaciones = SistemaAnalisisOperaciones()
analisis_actual = None
app = CTk.CTk(fg_color='#f8f9fa')
app.geometry("1200x800")
app.title("Sistema de Análisis de Operaciones Diarias")
app.resizable(True, True)
sidebar_frame = CTk.CTkFrame(master=app, fg_color="#2c3e50", width=250, corner_radius=0)
sidebar_frame.pack(side="left", fill="y")
main_frame = CTk.CTkFrame(master=app, fg_color="#f8f9fa")
main_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)
CTk.CTkLabel(sidebar_frame, text="DC-ACERO", 
             font=("Verdana", 20, "bold"), 
             text_color="white").pack(pady=30)
CTk.CTkLabel(sidebar_frame, text="Análisis de Operaciones", 
             font=("Verdana", 12), 
             text_color="#bdc3c7").pack(pady=5)
logo_placeholder = CTk.CTkFrame(sidebar_frame, width=120, height=120, fg_color="#34495e", corner_radius=15)
logo_placeholder.pack(pady=20)
CTk.CTkLabel(logo_placeholder, text="⚡", font=("Arial", 30), text_color="white").pack(expand=True)
def crear_boton_sidebar(texto, comando, color="#3498db"):
    boton = CTk.CTkButton(sidebar_frame, 
                         text=texto,
                         font=("Verdana", 12),
                         fg_color=color,
                         hover_color=color,
                         width=180,
                         height=35,
                         command=comando)
    boton.pack(pady=8)
    return boton
crear_boton_sidebar("📂 Cargar Excel", lambda: selectArchive(), "#3498db")
crear_boton_sidebar("🔄 Actualizar", lambda: mostrarAnalisisOperaciones(), "#2ecc71")
crear_boton_sidebar("📊 KG por Día", lambda: mostrarGraficoKGporDia(), "#9b59b6")
crear_boton_sidebar("📈 Top Materiales", lambda: mostrarGraficoTopMateriales(), "#e74c3c")
crear_boton_sidebar("💰 Dinero por Día", lambda: mostrarGraficoDineroDia(), "#f39c12")
crear_boton_sidebar("💵 Utilidades", lambda: mostrarGraficoUtilidades(), "#1abc9c")
crear_boton_sidebar("📄 Generar PDF", lambda: generarReporteCompletoPDF(), "#d35400")
crear_boton_sidebar("ℹ️ Ayuda", lambda: mostrarAyuda(), "#7f8c8d")
def selectArchive():
    global datos_operaciones, analisis_actual
    excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=(("Archivos de Excel", "*.xlsx;*.xlsm"), ("Todos los archivos", "*.*"))
    )
    if excel:
        try:
            archivo_excel = pd.ExcelFile(excel)
            print(f"Hojas disponibles: {archivo_excel.sheet_names}")
            hoja_operaciones = None
            nombres_posibles = ['OPERACIONES DIARIAS', 'OPERACIONES', 'OPERACIONES DIARIAS ', 'HOJA1', 'Sheet1', 'DATOS']
            for nombre in nombres_posibles:
                if nombre in archivo_excel.sheet_names:
                    hoja_operaciones = nombre
                    break
            if hoja_operaciones is None:
                hoja_operaciones = archivo_excel.sheet_names[0]
                print(f"Usando hoja: {hoja_operaciones}")
            datos_operaciones = pd.read_excel(excel, sheet_name=hoja_operaciones)
            print(f"Columnas cargadas: {datos_operaciones.columns.tolist()}")
            datos_operaciones.columns = datos_operaciones.columns.str.strip()
            mostrarInfoArchivoCargado()
            mostrarAnalisisOperaciones() 
        except Exception as e:
            mensaje_error = f"✗ Error: {str(e)}"
            print(mensaje_error)
            CTk.CTkLabel(main_frame, text=mensaje_error, 
                        text_color="red", font=("Verdana", 12)).pack(pady=5)
def mostrarInfoArchivoCargado():
    for widget in main_frame.winfo_children():
        if isinstance(widget, CTk.CTkLabel) and ("✓" in widget.cget("text") or "✗" in widget.cget("text")):
            widget.destroy()
    if datos_operaciones is not None:
        info_text = f"✓ Archivo cargado: {len(datos_operaciones)} registros, {len(datos_operaciones.columns)} columnas"
        CTk.CTkLabel(main_frame, text=info_text, 
                    text_color="green", font=("Verdana", 12)).pack(pady=5)
        print("Primeras 3 filas:")
        print(datos_operaciones.head(3))
def mostrarAnalisisOperaciones():
    global datos_operaciones, analisis_actual
    for widget in main_frame.winfo_children():
        widget.destroy()
    CTk.CTkLabel(main_frame, 
                text="ANÁLISIS DE OPERACIONES DIARIAS", 
                font=("Verdana", 22, "bold"),
                text_color="#2c3e50").pack(pady=15)
    if datos_operaciones is None:
        mostrarPantallaInicial()
        return
    try:
        datos_limpios = datos_operaciones.copy()
        print("Columnas disponibles:", datos_limpios.columns.tolist())
        columna_fecha = None
        columna_operacion = None
        columna_producto = None
        columna_kg = None
        columna_total = None
        for col in datos_limpios.columns:
            col_lower = str(col).lower()
            if 'fecha' in col_lower:
                columna_fecha = col
            elif 'operación' in col_lower or 'operacion' in col_lower:
                columna_operacion = col
            elif 'producto' in col_lower or 'material' in col_lower:
                columna_producto = col
            elif 'kg' in col_lower or 'kilo' in col_lower:
                columna_kg = col
            elif 'total' in col_lower and 'compra' in col_lower:
                columna_total = col
            elif 'total' in col_lower and columna_total is None:
                columna_total = col
        if columna_fecha is None and len(datos_limpios.columns) > 0:
            columna_fecha = datos_limpios.columns[0]
        if columna_operacion is None and len(datos_limpios.columns) > 1:
            columna_operacion = datos_limpios.columns[1]
        if columna_producto is None and len(datos_limpios.columns) > 2:
            columna_producto = datos_limpios.columns[2]
        if columna_kg is None and len(datos_limpios.columns) > 3:
            columna_kg = datos_limpios.columns[3]
        if columna_total is None and len(datos_limpios.columns) > 4:
            columna_total = datos_limpios.columns[4]
        print(f"Usando columnas - Fecha: {columna_fecha}, Operación: {columna_operacion}, Producto: {columna_producto}, KG: {columna_kg}, Total: {columna_total}")
        datos_limpios = datos_limpios.rename(columns={
            columna_fecha: 'FECHA',
            columna_operacion: 'OPERACIÓN',
            columna_producto: 'PRODUCTO', 
            columna_kg: 'KG',
            columna_total: 'TOTAL COMPRA'
        })
        datos_limpios = datos_limpios.dropna(subset=['FECHA', 'PRODUCTO', 'KG'])
        datos_limpios['FECHA'] = pd.to_datetime(datos_limpios['FECHA'], errors='coerce')
        datos_limpios = datos_limpios.dropna(subset=['FECHA'])
        datos_limpios['KG'] = pd.to_numeric(datos_limpios['KG'], errors='coerce')
        datos_limpios = datos_limpios.dropna(subset=['KG'])
        if 'TOTAL COMPRA' not in datos_limpios.columns or datos_limpios['TOTAL COMPRA'].isna().all():
            datos_limpios['TOTAL COMPRA'] = datos_limpios.apply(
                lambda row: row['KG'] * analizador_operaciones._obtener_precio_compra(row['PRODUCTO']), 
                axis=1
            )
        datos_compra = datos_limpios[datos_limpios['OPERACIÓN'] == 'COMPRA'].copy()
        if datos_compra.empty:
            datos_compra = datos_limpios.copy()
            CTk.CTkLabel(main_frame, text="⚠️ Usando todos los datos (no se encontraron solo COMPRAS)", 
                        text_color="orange", font=("Verdana", 11)).pack(pady=5) 
        if datos_compra.empty:
            CTk.CTkLabel(main_frame, text="No hay datos válidos para analizar", 
                        text_color="red", font=("Verdana", 12)).pack(pady=5)
            return
        fecha_max = datos_compra['FECHA'].max()
        inicio_semana = fecha_max - timedelta(days=30)  
        datos_semana = datos_compra[datos_compra['FECHA'] >= inicio_semana]
        if datos_semana.empty:
            datos_semana = datos_compra  
        analisis_actual = analizador_operaciones.analizar_operaciones_semanales(datos_semana)
        if not analisis_actual:
            CTk.CTkLabel(main_frame, text="Error en el análisis de datos", 
                        text_color="red", font=("Verdana", 12)).pack(pady=5)
            return
        mostrarResumenEjecutivo(analisis_actual, fecha_max)
        mostrarAnalisisDetallado(analisis_actual)
    except Exception as e:
        error_msg = f"Error en análisis: {str(e)}"
        print(error_msg)
        CTk.CTkLabel(main_frame, text=error_msg, 
                    text_color="red", font=("Verdana", 12)).pack(pady=5)
def _obtener_precio_compra(self, producto):
    """Obtiene el precio de compra del producto"""
    try:
        precios_compra = {
            'CARTON': 1.2, 'PET': 3.5, 'BOTE': 28.0, 'ALUMINIO': 18.0,
            'FIERRO': 3.3, 'CU 1': 156.0, 'CU 2': 150.0, 'BRONCE': 98.0,
            'ACERO INOX': 3.3, 'R/A': 24.0, 'ALUMINIO GRUESO (L)': 31.0,
            'ALUMINIO TRASTE (L)': 30.0, 'BRONCE (L)': 98.0, 'BRONCE (S)': 70.0,
            'RADIADOR LIN COBRE': 72.0, 'CAJA BOILER': 72.0, 'RIN ALUMINIO': 38.0,
            'ALUMINIO DELGADO': 24.0, 'ALUMINIO TUBO': 26.0, 'ALUMINIO CABLE (L)': 48.0,
            'ALUMINIO SPRAY (S)': 18.0, 'ALUMINIO PERFIL S/PINTURA (L)': 40.0,
            'ALUMINIO PERFIL C/PINTURA (L)': 33.0, 'CU ESTAÑADO': 153.0
        }
        producto_upper = str(producto).upper().strip()
        for key in precios_compra:
            if key in producto_upper:
                return precios_compra[key]
        if 'CARTON' in producto_upper:
            return 1.2
        elif 'PET' in producto_upper:
            return 3.5
        elif 'BOTE' in producto_upper:
            return 28.0
        elif 'ALUMINIO' in producto_upper:
            return 18.0
        elif 'FIERRO' in producto_upper:
            return 3.3
        elif 'CU 1' in producto_upper or 'COBRE 1' in producto_upper:
            return 156.0
        elif 'CU 2' in producto_upper or 'COBRE 2' in producto_upper:
            return 150.0
        elif 'BRONCE' in producto_upper:
            return 98.0
        elif 'ACERO' in producto_upper or 'INOX' in producto_upper:
            return 3.3
        elif 'RADIADOR' in producto_upper:
            return 72.0
        else:
            return 8.0 
    except:
        return 8.0
SistemaAnalisisOperaciones._obtener_precio_compra = _obtener_precio_compra
def mostrarPantallaInicial():
    CTk.CTkLabel(main_frame, 
                text="BIENVENIDO", 
                font=("Verdana", 28, "bold"),
                text_color="#2c3e50").pack(pady=30)
    CTk.CTkLabel(main_frame, 
                text="Sistema de Análisis de Operaciones Diarias", 
                font=("Verdana", 16),
                text_color="#7f8c8d").pack(pady=5)
    CTk.CTkButton(main_frame, 
                 text="📂 CARGAR ARCHIVO EXCEL", 
                 font=("Verdana", 18, "bold"), 
                 fg_color="#3498db",
                 hover_color="#2980b9",
                 width=300, height=50,
                 command=selectArchive).pack(pady=30)
    info_frame = CTk.CTkFrame(main_frame, fg_color="#e8f4fd", corner_radius=10)
    info_frame.pack(pady=20, padx=50, fill="x")
    CTk.CTkLabel(info_frame, 
                text="📋 EL SISTEMA ANALIZA:", 
                font=("Verdana", 14, "bold"),
                text_color="#2c3e50").pack(pady=10)
    analisis_items = [
        "• KG comprados a la semana", "• KG comprados por día", 
        "• Día de máxima compra y material", "• Material más comprado",
        "• Material menos comprado", "• Dinero invertido diario y total",
        "• Venta aproximada por material", "• Utilidad aproximada",
        "• Gráficas en ventanas independientes", "• Reporte PDF completo"
    ]
    for item in analisis_items:
        CTk.CTkLabel(info_frame, 
                    text=item, 
                    font=("Verdana", 12),
                    text_color="#34495e").pack(anchor="w", pady=2, padx=20)
def mostrarResumenEjecutivo(analisis, fecha_max):
    resumen_frame = CTk.CTkFrame(main_frame, fg_color="#e8f4fd", corner_radius=15)
    resumen_frame.pack(pady=10, padx=20, fill="x")
    CTk.CTkLabel(resumen_frame, 
                text="📊 RESUMEN EJECUTIVO", 
                font=("Verdana", 16, "bold"),
                text_color="#2c3e50").pack(pady=10)
    CTk.CTkLabel(resumen_frame, 
                text=f"📅 Período analizado: {analisis.get('periodo_analizado', 'N/A')}",
                font=("Verdana", 11, "bold"),
                text_color="#7f8c8d").pack(pady=5)
    metricas_frame = CTk.CTkFrame(resumen_frame, fg_color="transparent")
    metricas_frame.pack(pady=10, padx=20, fill="x")
    fila1 = CTk.CTkFrame(metricas_frame, fg_color="transparent")
    fila1.pack(fill="x", pady=5)
    CTk.CTkLabel(fila1, 
                text=f"📦 KG Totales: {analisis['kg_totales_semana']:,.1f} kg", 
                font=("Verdana", 12, "bold"),
                text_color="#27ae60").pack(side="left", padx=20)
    CTk.CTkLabel(fila1, 
                text=f"💰 Inversión Total: ${analisis['dinero_total_semana']:,.2f}", 
                font=("Verdana", 12, "bold"),
                text_color="#e74c3c").pack(side="left", padx=20)
    fila2 = CTk.CTkFrame(metricas_frame, fg_color="transparent")
    fila2.pack(fill="x", pady=5)
    material_mas = analisis['material_mas_comprado'] or "N/A"
    CTk.CTkLabel(fila2, 
                text=f"🏆 Material Top: {material_mas}", 
                font=("Verdana", 12, "bold"),
                text_color="#9b59b6").pack(side="left", padx=20)
    dia_max = analisis['dia_max_compra'] or "N/A"
    CTk.CTkLabel(fila2, 
                text=f"📈 Día Máximo: {dia_max}", 
                font=("Verdana", 12, "bold"),
                text_color="#f39c12").pack(side="left", padx=20)
    fila3 = CTk.CTkFrame(metricas_frame, fg_color="transparent")
    fila3.pack(fill="x", pady=5)
    CTk.CTkLabel(fila3, 
                text=f"💵 Venta Aprox: ${analisis['venta_total_aproximada']:,.2f}", 
                font=("Verdana", 12, "bold"),
                text_color="#27ae60").pack(side="left", padx=20)
    CTk.CTkLabel(fila3, 
                text=f"📊 Utilidad Aprox: ${analisis['utilidad_total_aproximada']:,.2f}", 
                font=("Verdana", 12, "bold"),
                text_color="#2ecc71").pack(side="left", padx=20)
    fila4 = CTk.CTkFrame(metricas_frame, fg_color="transparent")
    fila4.pack(fill="x", pady=5)
    CTk.CTkLabel(fila4, 
                text=f"📋 Total Materiales: {analisis.get('total_materiales', 0)}", 
                font=("Verdana", 12, "bold"),
                text_color="#3498db").pack(side="left", padx=20)
    material_menos = analisis['material_menos_comprado'] or "N/A"
    CTk.CTkLabel(fila4, 
                text=f"📉 Material Menos: {material_menos}", 
                font=("Verdana", 12, "bold"),
                text_color="#e67e22").pack(side="left", padx=20)
def mostrarGraficoKGporDia():
    if analisis_actual is None:
        CTk.CTkLabel(main_frame, text="Primero carga y analiza un archivo", 
                    text_color="orange", font=("Verdana", 12)).pack(pady=5)
        return
    ventana_grafico = CTk.CTkToplevel(app)
    ventana_grafico.title("KG Comprados por Día")
    ventana_grafico.geometry("900x600")
    ventana_grafico.fg_color = "#ffffff"
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('#f8f9fa')
    if analisis_actual['kg_por_dia']:
        dias = [str(d) for d in analisis_actual['kg_por_dia'].keys()]
        kgs = list(analisis_actual['kg_por_dia'].values())
        bars = ax.bar(dias, kgs, color='skyblue', alpha=0.7, edgecolor='navy')
        ax.set_title('KG COMPRADOS POR DÍA', fontsize=16, fontweight='bold', pad=20)
        ax.set_ylabel('KILOGRAMOS (KG)', fontweight='bold', fontsize=12)
        ax.set_xlabel('FECHAS', fontweight='bold', fontsize=12)
        ax.tick_params(axis='x', rotation=45)
        ax.grid(axis='y', alpha=0.3)
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + max(kgs)*0.01,
                   f'{height:,.0f} kg', ha='center', va='bottom', fontweight='bold', fontsize=9)
    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=ventana_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)
def mostrarGraficoTopMateriales():
    if analisis_actual is None:
        CTk.CTkLabel(main_frame, text="Primero carga y analiza un archivo", 
                    text_color="orange", font=("Verdana", 12)).pack(pady=5)
        return 
    ventana_grafico = CTk.CTkToplevel(app)
    ventana_grafico.title("Top Materiales por KG")
    ventana_grafico.geometry("1000x700")
    ventana_grafico.fg_color = "#ffffff"
    fig, ax = plt.subplots(figsize=(12, 8))
    fig.patch.set_facecolor('#f8f9fa')
    if analisis_actual['venta_utilidad_material']:
        materiales_kg = {k: v['kg'] for k, v in analisis_actual['venta_utilidad_material'].items()}
        top_materiales = dict(sorted(materiales_kg.items(), key=lambda x: x[1], reverse=True)[:15])        
        if top_materiales:
            y_pos = np.arange(len(top_materiales))
            bars = ax.barh(y_pos, list(top_materiales.values()), color='lightgreen', alpha=0.7, edgecolor='green')
            ax.set_yticks(y_pos)
            ax.set_yticklabels(list(top_materiales.keys()), fontsize=10)
            ax.set_xlabel('KILOGRAMOS (KG)', fontweight='bold', fontsize=12)
            ax.set_title('TOP 15 MATERIALES POR KG COMPRADOS', fontsize=16, fontweight='bold', pad=20)
            ax.grid(axis='x', alpha=0.3)
            for i, bar in enumerate(bars):
                width = bar.get_width()
                ax.text(width + max(top_materiales.values())*0.01, bar.get_y() + bar.get_height()/2,
                       f'{width:,.1f} kg', ha='left', va='center', fontweight='bold', fontsize=9)
    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=ventana_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)
def mostrarGraficoDineroDia():
    if analisis_actual is None:
        CTk.CTkLabel(main_frame, text="Primero carga y analiza un archivo", 
                    text_color="orange", font=("Verdana", 12)).pack(pady=5)
        return
    ventana_grafico = CTk.CTkToplevel(app)
    ventana_grafico.title("Dinero Invertido por Día")
    ventana_grafico.geometry("900x600")
    ventana_grafico.fg_color = "#ffffff"
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('#f8f9fa')
    if analisis_actual['dinero_por_dia']:
        dias = [str(d) for d in analisis_actual['dinero_por_dia'].keys()]
        dinero = list(analisis_actual['dinero_por_dia'].values())
        bars = ax.bar(dias, dinero, color='gold', alpha=0.7, edgecolor='darkorange')
        ax.set_title('DINERO INVERTIDO POR DÍA', fontsize=16, fontweight='bold', pad=20)
        ax.set_ylabel('PESOS ($)', fontweight='bold', fontsize=12)
        ax.set_xlabel('FECHAS', fontweight='bold', fontsize=12)
        ax.tick_params(axis='x', rotation=45)
        ax.grid(axis='y', alpha=0.3)
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + max(dinero)*0.01,
                   f'${height:,.0f}', ha='center', va='bottom', fontweight='bold', fontsize=9)
    plt.tight_layout()  
    canvas = FigureCanvasTkAgg(fig, master=ventana_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)
def mostrarGraficoUtilidades():
    if analisis_actual is None:
        CTk.CTkLabel(main_frame, text="Primero carga y analiza un archivo", 
                    text_color="orange", font=("Verdana", 12)).pack(pady=5)
        return
    ventana_grafico = CTk.CTkToplevel(app)
    ventana_grafico.title("Utilidades por Material")
    ventana_grafico.geometry("1000x700")
    ventana_grafico.fg_color = "#ffffff"
    fig, ax = plt.subplots(figsize=(12, 8))
    fig.patch.set_facecolor('#f8f9fa')
    if analisis_actual['venta_utilidad_material']:
        utilidades = {mat: data['utilidad_aproximada'] for mat, data in analisis_actual['venta_utilidad_material'].items()}
        top_utilidades = dict(sorted(utilidades.items(), key=lambda x: x[1], reverse=True)[:15])
        if top_utilidades:
            y_pos = np.arange(len(top_utilidades))
            colors = ['#2ecc71' if val >= 0 else '#e74c3c' for val in top_utilidades.values()]
            bars = ax.barh(y_pos, list(top_utilidades.values()), color=colors, alpha=0.7, edgecolor=colors)
            ax.set_yticks(y_pos)
            ax.set_yticklabels(list(top_utilidades.keys()), fontsize=10)
            ax.set_xlabel('UTILIDAD ($)', fontweight='bold', fontsize=12)
            ax.set_title('TOP 15 MATERIALES POR UTILIDAD APROXIMADA', fontsize=16, fontweight='bold', pad=20)
            ax.grid(axis='x', alpha=0.3)
            ax.axvline(x=0, color='black', linestyle='-', alpha=0.5)
            for i, bar in enumerate(bars):
                width = bar.get_width()
                ax.text(width + (max(list(top_utilidades.values())) * 0.01 if width >= 0 else width - (max(list(top_utilidades.values())) * 0.01)), 
                       bar.get_y() + bar.get_height()/2,
                       f'${width:,.0f}', ha='left' if width >= 0 else 'right', va='center', 
                       fontweight='bold', fontsize=9, color='white' if abs(width) < max(list(top_utilidades.values())) * 0.1 else 'black')
    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=ventana_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)
def mostrarAnalisisDetallado(analisis):
    frame_scroll = CTk.CTkScrollableFrame(main_frame, fg_color="#f8f9fa", height=300)
    frame_scroll.pack(pady=10, padx=20, fill="x")
    CTk.CTkLabel(frame_scroll, 
                text="📈 DETALLE COMPLETO POR MATERIAL", 
                font=("Verdana", 14, "bold"),
                text_color="#2c3e50").pack(pady=10)
    materiales_ordenados = sorted(analisis['venta_utilidad_material'].items(), 
                                 key=lambda x: x[1]['kg'], reverse=True)
    for producto, datos in materiales_ordenados:
        material_frame = CTk.CTkFrame(frame_scroll, fg_color="white", corner_radius=8)
        material_frame.pack(pady=3, padx=10, fill="x")  
        header_frame = CTk.CTkFrame(material_frame, fg_color="#3498db", corner_radius=6)
        header_frame.pack(pady=2, padx=8, fill="x")
        CTk.CTkLabel(header_frame, 
                    text=f"📦 {producto}", 
                    font=("Verdana", 11, "bold"),
                    text_color="white").pack(pady=3)
        datos_frame = CTk.CTkFrame(material_frame, fg_color="#f8f9fa", corner_radius=6)
        datos_frame.pack(pady=2, padx=8, fill="x")
        metricas_frame = CTk.CTkFrame(datos_frame, fg_color="transparent")
        metricas_frame.pack(fill="x", pady=2)
        CTk.CTkLabel(metricas_frame, 
                    text=f"⚖️ {datos['kg']:,.1f} kg", 
                    font=("Verdana", 9),
                    text_color="#2c3e50").pack(side="left", padx=8)
        CTk.CTkLabel(metricas_frame, 
                    text=f"💰 Compra: ${datos['compra']:,.0f}", 
                    font=("Verdana", 9),
                    text_color="#e74c3c").pack(side="left", padx=8)
        CTk.CTkLabel(metricas_frame, 
                    text=f"📊 Venta: ${datos['venta_aproximada']:,.0f}", 
                    font=("Verdana", 9),
                    text_color="#27ae60").pack(side="left", padx=8)
        utilidad_color = "#2ecc71" if datos['utilidad_aproximada'] >= 0 else "#e74c3c"
        CTk.CTkLabel(metricas_frame, 
                    text=f"💵 Utilidad: ${datos['utilidad_aproximada']:,.0f}", 
                    font=("Verdana", 9, "bold"),
                    text_color=utilidad_color).pack(side="left", padx=8)
def generarGraficosParaPDF():
    """Genera y guarda las gráficas para incluirlas en el PDF"""
    graficos = {}
    try:
        if analisis_actual['kg_por_dia']:
            fig1, ax1 = plt.subplots(figsize=(10, 6))
            dias = [str(d) for d in analisis_actual['kg_por_dia'].keys()]
            kgs = list(analisis_actual['kg_por_dia'].values())
            bars = ax1.bar(dias, kgs, color='skyblue', alpha=0.7, edgecolor='navy')
            ax1.set_title('KG COMPRADOS POR DÍA', fontsize=14, fontweight='bold', pad=15)
            ax1.set_ylabel('KILOGRAMOS (KG)', fontweight='bold', fontsize=10)
            ax1.set_xlabel('FECHAS', fontweight='bold', fontsize=10)
            ax1.tick_params(axis='x', rotation=45)
            ax1.grid(axis='y', alpha=0.3)
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + max(kgs)*0.01,
                       f'{height:,.0f} kg', ha='center', va='bottom', fontweight='bold', fontsize=8)
            plt.tight_layout()
            temp_file1 = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            plt.savefig(temp_file1.name, dpi=150, bbox_inches='tight')
            graficos['kg_por_dia'] = temp_file1.name
            plt.close(fig1)
        if analisis_actual['venta_utilidad_material']:
            fig2, ax2 = plt.subplots(figsize=(10, 8))
            materiales_kg = {k: v['kg'] for k, v in analisis_actual['venta_utilidad_material'].items()}
            top_materiales = dict(sorted(materiales_kg.items(), key=lambda x: x[1], reverse=True)[:10])
            if top_materiales:
                y_pos = np.arange(len(top_materiales))
                bars = ax2.barh(y_pos, list(top_materiales.values()), color='lightgreen', alpha=0.7, edgecolor='green')
                ax2.set_yticks(y_pos)
                ax2.set_yticklabels(list(top_materiales.keys()), fontsize=9)
                ax2.set_xlabel('KILOGRAMOS (KG)', fontweight='bold', fontsize=10)
                ax2.set_title('TOP 10 MATERIALES POR KG COMPRADOS', fontsize=14, fontweight='bold', pad=15)
                ax2.grid(axis='x', alpha=0.3)
                for i, bar in enumerate(bars):
                    width = bar.get_width()
                    ax2.text(width + max(top_materiales.values())*0.01, bar.get_y() + bar.get_height()/2,
                           f'{width:,.1f} kg', ha='left', va='center', fontweight='bold', fontsize=8)
            plt.tight_layout()
            temp_file2 = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            plt.savefig(temp_file2.name, dpi=150, bbox_inches='tight')
            graficos['top_materiales'] = temp_file2.name
            plt.close(fig2)
        if analisis_actual['dinero_por_dia']:
            fig3, ax3 = plt.subplots(figsize=(10, 6))
            dias = [str(d) for d in analisis_actual['dinero_por_dia'].keys()]
            dinero = list(analisis_actual['dinero_por_dia'].values())
            bars = ax3.bar(dias, dinero, color='gold', alpha=0.7, edgecolor='darkorange')
            ax3.set_title('DINERO INVERTIDO POR DÍA', fontsize=14, fontweight='bold', pad=15)
            ax3.set_ylabel('DÓLARES ($)', fontweight='bold', fontsize=10)
            ax3.set_xlabel('FECHAS', fontweight='bold', fontsize=10)
            ax3.tick_params(axis='x', rotation=45)
            ax3.grid(axis='y', alpha=0.3)    
            for bar in bars:
                height = bar.get_height()
                ax3.text(bar.get_x() + bar.get_width()/2., height + max(dinero)*0.01,
                       f'${height:,.0f}', ha='center', va='bottom', fontweight='bold', fontsize=8)
            plt.tight_layout()
            temp_file3 = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            plt.savefig(temp_file3.name, dpi=150, bbox_inches='tight')
            graficos['dinero_por_dia'] = temp_file3.name
            plt.close(fig3)
        if analisis_actual['venta_utilidad_material']:
            fig4, ax4 = plt.subplots(figsize=(10, 8))
            utilidades = {mat: data['utilidad_aproximada'] for mat, data in analisis_actual['venta_utilidad_material'].items()}
            top_utilidades = dict(sorted(utilidades.items(), key=lambda x: x[1], reverse=True)[:10])  
            if top_utilidades:
                y_pos = np.arange(len(top_utilidades))
                colors = ['#2ecc71' if val >= 0 else '#e74c3c' for val in top_utilidades.values()]
                bars = ax4.barh(y_pos, list(top_utilidades.values()), color=colors, alpha=0.7, edgecolor=colors)
                ax4.set_yticks(y_pos)
                ax4.set_yticklabels(list(top_utilidades.keys()), fontsize=9)
                ax4.set_xlabel('UTILIDAD ($)', fontweight='bold', fontsize=10)
                ax4.set_title('TOP 10 MATERIALES POR UTILIDAD APROXIMADA', fontsize=14, fontweight='bold', pad=15)
                ax4.grid(axis='x', alpha=0.3)
                ax4.axvline(x=0, color='black', linestyle='-', alpha=0.5)
                for i, bar in enumerate(bars):
                    width = bar.get_width()
                    text_offset = max([abs(val) for val in top_utilidades.values()]) * 0.02
                    x_position = width + text_offset if width >= 0 else width - text_offset
                    ha_position = 'left' if width >= 0 else 'right'   
                    ax4.text(x_position, 
                           bar.get_y() + bar.get_height()/2,
                           f'${width:,.0f}', 
                           ha=ha_position, va='center', 
                           fontweight='bold', fontsize=8, color='black')
            plt.tight_layout()
            temp_file4 = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            plt.savefig(temp_file4.name, dpi=150, bbox_inches='tight')
            graficos['utilidades'] = temp_file4.name
            plt.close(fig4)
    except Exception as e:
        print(f"Error generando gráficos para PDF: {e}")
    return graficos
def generarReporteCompletoPDF():
    if analisis_actual is None:
        CTk.CTkLabel(main_frame, text="Primero carga y analiza un archivo Excel", 
                    text_color="orange", font=("Verdana", 12)).pack(pady=5)
        return
    try:
        graficos = generarGraficosParaPDF()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 20)
        pdf.cell(0, 15, "REPORTE COMPLETO", 0, 1, 'C')
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 1)
        pdf.cell(0, 8, f"Período analizado: {analisis_actual.get('periodo_analizado', 'N/A')}", 0, 1)
        pdf.cell(0, 8, f"Total materiales analizados: {analisis_actual.get('total_materiales', 0)}", 0, 1)
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "RESUMEN EJECUTIVO", 0, 1)
        pdf.set_font("Arial", '', 12)
        pdf.cell(0, 8, f"KG totales comprados: {analisis_actual['kg_totales_semana']:,.1f} kg", 0, 1)
        pdf.cell(0, 8, f"Inversión total en compras: ${analisis_actual['dinero_total_semana']:,.2f}", 0, 1)
        pdf.cell(0, 8, f"Material más comprado: {analisis_actual['material_mas_comprado']} ({analisis_actual['kg_material_mas_comprado']:,.1f} kg)", 0, 1)
        pdf.cell(0, 8, f"Material menos comprado: {analisis_actual['material_menos_comprado']} ({analisis_actual['kg_material_menos_comprado']:,.1f} kg)", 0, 1)
        pdf.cell(0, 8, f"Día de máxima compra: {analisis_actual['dia_max_compra']} ({analisis_actual['max_kg_dia']:,.1f} kg)", 0, 1)
        pdf.cell(0, 8, f"Venta total aproximada: ${analisis_actual['venta_total_aproximada']:,.2f}", 0, 1)
        pdf.cell(0, 8, f"Utilidad total aproximada: ${analisis_actual['utilidad_total_aproximada']:,.2f}", 0, 1)
        pdf.ln(15)
        if 'kg_por_dia' in graficos:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "KG COMPRADOS POR DÍA", 0, 1)
            pdf.image(graficos['kg_por_dia'], x=10, y=30, w=190)
            pdf.ln(100)
        if 'top_materiales' in graficos:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "TOP 10 MATERIALES POR KG COMPRADOS", 0, 1)
            pdf.image(graficos['top_materiales'], x=10, y=30, w=190)
            pdf.ln(120)
        if 'dinero_por_dia' in graficos:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "DINERO INVERTIDO POR DÍA", 0, 1)
            pdf.image(graficos['dinero_por_dia'], x=10, y=30, w=190)
            pdf.ln(100)
        if 'utilidades' in graficos:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "TOP 10 MATERIALES POR UTILIDAD", 0, 1)
            pdf.image(graficos['utilidades'], x=10, y=30, w=190)
            pdf.ln(120)
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "DETALLE COMPLETO POR MATERIAL", 0, 1)
        pdf.ln(5)
        materiales_ordenados = sorted(analisis_actual['venta_utilidad_material'].items(), 
                                     key=lambda x: x[1]['kg'], reverse=True)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(60, 8, "MATERIAL", 1)
        pdf.cell(25, 8, "KG", 1)
        pdf.cell(35, 8, "COMPRA", 1)
        pdf.cell(35, 8, "VENTA APROX", 1)
        pdf.cell(35, 8, "UTILIDAD APROX", 1)
        pdf.ln(8)   
        pdf.set_font("Arial", '', 9)
        for producto, datos in materiales_ordenados:
            nombre = producto[:25] + "..." if len(producto) > 25 else producto      
            pdf.cell(60, 6, nombre, 1)
            pdf.cell(25, 6, f"{datos['kg']:,.1f}", 1)
            pdf.cell(35, 6, f"${datos['compra']:,.0f}", 1)
            pdf.cell(35, 6, f"${datos['venta_aproximada']:,.0f}", 1)
            utilidad_text = f"${datos['utilidad_aproximada']:,.0f}"
            if datos['utilidad_aproximada'] < 0:
                utilidad_text = f"-${abs(datos['utilidad_aproximada']):,.0f}"
            pdf.cell(35, 6, utilidad_text, 1)
            pdf.ln(6)
        output_file = f"Reporte_ACERO_XD_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        pdf.output(output_file)
        for temp_file in graficos.values():
            try:
                os.unlink(temp_file)
            except:
                pass
        CTk.CTkLabel(main_frame, 
                    text=f"✓ PDF generado exitosamente: {output_file}", 
                    text_color="green", font=("Verdana", 12)).pack(pady=5)
        try:
            webbrowser.open(output_file)
        except:
            pass
    except Exception as e:
        error_msg = f"✗ Error al generar PDF: {str(e)}"
        print(error_msg)
        CTk.CTkLabel(main_frame, text=error_msg, 
                    text_color="red", font=("Verdana", 12)).pack(pady=5)
def mostrarAyuda():
    ventana_ayuda = CTk.CTkToplevel(app)
    ventana_ayuda.title("Ayuda")
    ventana_ayuda.geometry("500x400")
    ventana_ayuda.fg_color = "#ffffff"
    CTk.CTkLabel(ventana_ayuda, 
                text="❓ AYUDA - CÓMO USAR EL SISTEMA", 
                font=("Verdana", 16, "bold"),
                text_color="#2c3e50").pack(pady=20)
    instrucciones = [
        "1. 📂 Cargar Excel: Selecciona tu archivo Excel (.xlsx o .xlsm)",
        "2. 🔄 El sistema detecta automáticamente las columnas",
        "3. 📊 Se analizan KG, dinero y utilidades automáticamente",
        "4. 📈 Cada gráfico se abre en ventana independiente:",
        "   - KG por Día", "   - Top Materiales", 
        "   - Dinero por Día", "   - Utilidades",
        "5. 📄 Generar PDF: Reporte completo con gráficas y datos",
        "",
        "📝 El sistema analiza todo lo solicitado:",
        "  • KG comprados por día y total",
        "  • Material más y menos comprado", 
        "  • Dinero invertido diario y total",
        "  • Ventas y utilidades aproximadas por material",
        "  • Día de máxima compra y material correspondiente",
        "",
        "⚠️ Compatible con archivos .xlsx y .xlsm"
    ]
    for instruccion in instrucciones:
        CTk.CTkLabel(ventana_ayuda, 
                    text=instruccion, 
                    font=("Verdana", 10),
                    text_color="#34495e",
                    justify="left").pack(anchor="w", pady=1, padx=20)
mostrarPantallaInicial()
if __name__ == "__main__":
    app.mainloop()