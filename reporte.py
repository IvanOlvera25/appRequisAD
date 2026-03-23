#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================
REPORTE ANUAL COMPLETO - AD17 SOLUTIONS
Sistema de Gestión de Costos
Análisis Integral de Fin de Año
Versión Mejorada - Diseño Profesional + Exportación de Gráficas
=====================================================
"""

import sqlite3
import pymysql
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.gridspec import GridSpec
import seaborn as sns
from datetime import datetime, timedelta
from calendar import month_name
import warnings
import json
from io import BytesIO
import xlsxwriter
from collections import Counter
import textwrap
import os
from pathlib import Path

warnings.filterwarnings('ignore')

# ============================================================
# CONFIGURACIÓN DE ESTILO PROFESIONAL
# ============================================================

# Configuración base de matplotlib
plt.rcParams.update({
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'grid.linestyle': '-',
    'grid.linewidth': 0.5,
    'font.family': 'sans-serif',
    'font.size': 10,
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'axes.labelsize': 11,
    'axes.labelweight': 'normal',
    'xtick.labelsize': 9,
    'ytick.labelsize': 9,
    'legend.fontsize': 9,
    'figure.titlesize': 16,
    'axes.spines.top': False,
    'axes.spines.right': False,
})

# Paleta de colores corporativa AD17
COLORS = {
    'primary': '#FF9800',
    'primary_dark': '#F57C00',
    'primary_light': '#FFB74D',
    'secondary': '#1976D2',
    'secondary_dark': '#1565C0',
    'success': '#43A047',
    'danger': '#E53935',
    'warning': '#FDD835',
    'info': '#29B6F6',
    'dark': '#37474F',
    'light': '#FAFAFA',
    'gray': '#9E9E9E',
    'white': '#FFFFFF'
}

# Paleta de colores para gráficas (más profesional)
CHART_COLORS = [
    '#FF9800', '#1976D2', '#43A047', '#E53935', '#9C27B0',
    '#00BCD4', '#FF5722', '#3F51B5', '#4CAF50', '#FFC107',
    '#795548', '#607D8B', '#E91E63', '#009688', '#CDDC39'
]


def truncate_text(text, max_length=25):
    """Trunca texto largo con elipsis"""
    if pd.isna(text):
        return ''
    text = str(text)
    if len(text) > max_length:
        return text[:max_length-3] + '...'
    return text


def wrap_labels(labels, max_width=15):
    """Envuelve etiquetas largas en múltiples líneas"""
    wrapped = []
    for label in labels:
        if pd.isna(label):
            wrapped.append('')
        else:
            wrapped.append('\n'.join(textwrap.wrap(str(label), max_width)))
    return wrapped


def format_currency(value, abbreviated=True):
    """Formatea valores monetarios"""
    if pd.isna(value) or value == 0:
        return '$0'
    if abbreviated:
        if abs(value) >= 1_000_000:
            return f'${value/1_000_000:.1f}M'
        elif abs(value) >= 1_000:
            return f'${value/1_000:.0f}K'
    return f'${value:,.0f}'


class AD17FinancialAnalyzer:
    """
    Analizador completo del sistema financiero AD17
    Versión mejorada con diseño profesional y exportación de gráficas
    """

    def __init__(self, year=None):
        self.year = year or datetime.now().year
        self.sqlite_db = 'database.db'
        self.mysql_config = {
            'host': 'ad17solutions.dscloud.me',
            'port': 3307,
            'user': 'IvanUriel',
            'password': 'iuOp20!!25',
            'charset': 'utf8mb4'
        }

        self.df_solicitudes = None
        self.df_creditos = None
        self.df_pagos_credito = None
        self.df_pagos_recurrentes = None
        self.df_proveedores = None

        self.metricas = {}

        # Configuración para exportación de imágenes
        self.export_images = True
        self.images_folder = None
        self.image_format = 'png'  # o 'svg' para vectorial
        self.image_dpi = 300  # Alta resolución
        self.contador_graficas = 0

        print(f"🚀 Iniciando análisis para el año {self.year}...")

    def crear_carpeta_imagenes(self):
        """Crea la carpeta para almacenar las imágenes"""
        folder_name = f'Graficas_Reporte_{self.year}'
        self.images_folder = Path(folder_name)

        # Crear carpeta si no existe
        self.images_folder.mkdir(exist_ok=True)

        print(f"📁 Carpeta de imágenes creada: {self.images_folder}")
        return self.images_folder

    def guardar_figura(self, fig, pdf, nombre_grafica):
        """
        Guarda la figura tanto en el PDF como en archivo individual

        Args:
            fig: Figura de matplotlib
            pdf: Objeto PdfPages
            nombre_grafica: Nombre descriptivo para el archivo de imagen
        """
        # Guardar en PDF
        pdf.savefig(fig, bbox_inches='tight', facecolor='white')

        # Guardar imagen individual si está habilitado
        if self.export_images and self.images_folder:
            self.contador_graficas += 1

            # Crear nombre de archivo seguro
            nombre_archivo = f"{self.contador_graficas:02d}_{nombre_grafica}"
            nombre_archivo = "".join(c for c in nombre_archivo if c.isalnum() or c in (' ', '-', '_')).rstrip()
            nombre_archivo = nombre_archivo.replace(' ', '_')

            # Guardar en el formato especificado
            if self.image_format == 'png':
                filepath = self.images_folder / f"{nombre_archivo}.png"
                fig.savefig(filepath, dpi=self.image_dpi, bbox_inches='tight',
                           facecolor='white', edgecolor='none')
            elif self.image_format == 'svg':
                filepath = self.images_folder / f"{nombre_archivo}.svg"
                fig.savefig(filepath, format='svg', bbox_inches='tight',
                           facecolor='white', edgecolor='none')
            else:  # Ambos formatos
                filepath_png = self.images_folder / f"{nombre_archivo}.png"
                filepath_svg = self.images_folder / f"{nombre_archivo}.svg"
                fig.savefig(filepath_png, dpi=self.image_dpi, bbox_inches='tight',
                           facecolor='white', edgecolor='none')
                fig.savefig(filepath_svg, format='svg', bbox_inches='tight',
                           facecolor='white', edgecolor='none')

            print(f"  💾 Guardada: {nombre_archivo}")

        # Cerrar figura
        plt.close(fig)

    def conectar_sqlite(self):
        """Conecta a la base de datos SQLite local"""
        try:
            conn = sqlite3.connect(self.sqlite_db, timeout=20)
            conn.row_factory = sqlite3.Row
            print("✅ Conexión SQLite establecida")
            return conn
        except Exception as e:
            print(f"❌ Error conectando a SQLite: {e}")
            return None

    def conectar_mysql(self, database):
        """Conecta a MySQL remoto"""
        try:
            conn = pymysql.connect(
                **self.mysql_config,
                database=database,
                cursorclass=pymysql.cursors.DictCursor
            )
            print(f"✅ Conexión MySQL ({database}) establecida")
            return conn
        except Exception as e:
            print(f"❌ Error conectando a MySQL ({database}): {e}")
            return None

    def extraer_solicitudes(self):
        """Extrae y procesa todas las solicitudes de pago"""
        print("\n📊 Extrayendo solicitudes de pago...")

        conn = self.conectar_sqlite()
        if not conn:
            return

        try:
            query = """
                SELECT
                    id, fp, nombre, destinatario, correo, departamento,
                    tipo_solicitud, tipo_pago, descripcion, datos_deposito,
                    banco, clabe, referencia, monto, estado, fecha, fecha_limite,
                    archivo_adjunto, archivo_factura, archivo_recibo, archivo_orden_compra,
                    anticipo, porcentaje_anticipo, monto_anticipo, monto_restante, tipo_anticipo,
                    tiene_comision, porcentaje_comision, monto_sin_comision, monto_comision,
                    es_programada, fecha_aprobado, fecha_liquidado, fecha_ultimo_cambio,
                    historial_estados
                FROM solicitudes
            """

            self.df_solicitudes = pd.read_sql_query(query, conn)

            # Conversiones de tipos
            self.df_solicitudes['monto'] = pd.to_numeric(self.df_solicitudes['monto'], errors='coerce')
            self.df_solicitudes['monto_comision'] = pd.to_numeric(self.df_solicitudes['monto_comision'], errors='coerce')
            self.df_solicitudes['porcentaje_comision'] = pd.to_numeric(self.df_solicitudes['porcentaje_comision'], errors='coerce')
            self.df_solicitudes['fecha'] = pd.to_datetime(self.df_solicitudes['fecha'], errors='coerce')
            self.df_solicitudes['fecha_limite'] = pd.to_datetime(self.df_solicitudes['fecha_limite'], errors='coerce')

            # Extraer año, mes, trimestre
            self.df_solicitudes['año'] = self.df_solicitudes['fecha'].dt.year
            self.df_solicitudes['mes'] = self.df_solicitudes['fecha'].dt.month
            self.df_solicitudes['mes_nombre'] = self.df_solicitudes['fecha'].dt.month_name()
            self.df_solicitudes['trimestre'] = self.df_solicitudes['fecha'].dt.quarter
            self.df_solicitudes['dia_semana'] = self.df_solicitudes['fecha'].dt.day_name()

            # Filtrar por año
            self.df_solicitudes = self.df_solicitudes[
                self.df_solicitudes['año'] == self.year
            ]

            print(f"✅ {len(self.df_solicitudes)} solicitudes extraídas para {self.year}")

        except Exception as e:
            print(f"❌ Error extrayendo solicitudes: {e}")
        finally:
            conn.close()

    def extraer_creditos(self):
        """Extrae información de créditos"""
        print("\n💳 Extrayendo créditos...")

        conn = self.conectar_sqlite()
        if not conn:
            return

        try:
            query_creditos = """
                SELECT
                    id, nombre, entidad, descripcion, monto_total, tasa_interes,
                    fecha_inicio, fecha_final, plazo_meses, estado, fecha_registro,
                    numero_cuenta, tipo_credito, pago_mensual, contacto, notas
                FROM creditos
            """
            self.df_creditos = pd.read_sql_query(query_creditos, conn)

            query_pagos = """
                SELECT
                    id, credito_id, monto, fecha, referencia,
                    descripcion, comprobante, tipo_pago
                FROM pagos_credito
            """
            self.df_pagos_credito = pd.read_sql_query(query_pagos, conn)

            self.df_creditos['monto_total'] = pd.to_numeric(self.df_creditos['monto_total'], errors='coerce')
            self.df_creditos['fecha_inicio'] = pd.to_datetime(self.df_creditos['fecha_inicio'], errors='coerce')
            self.df_creditos['fecha_final'] = pd.to_datetime(self.df_creditos['fecha_final'], errors='coerce')

            self.df_pagos_credito['monto'] = pd.to_numeric(self.df_pagos_credito['monto'], errors='coerce')
            self.df_pagos_credito['fecha'] = pd.to_datetime(self.df_pagos_credito['fecha'], errors='coerce')
            self.df_pagos_credito['año'] = self.df_pagos_credito['fecha'].dt.year

            self.df_pagos_credito = self.df_pagos_credito[
                self.df_pagos_credito['año'] == self.year
            ]

            print(f"✅ {len(self.df_creditos)} créditos y {len(self.df_pagos_credito)} pagos extraídos")

        except Exception as e:
            print(f"❌ Error extrayendo créditos: {e}")
        finally:
            conn.close()

    def extraer_pagos_recurrentes(self):
        """Extrae pagos recurrentes"""
        print("\n🔄 Extrayendo pagos recurrentes...")

        conn = self.conectar_sqlite()
        if not conn:
            return

        try:
            query = """
                SELECT
                    id, nombre, proveedor, descripcion, monto, metodo_pago,
                    banco, clabe, periodicidad, fecha_proximo_pago,
                    dias_recordatorio, correos, activo, fecha_creacion
                FROM pagos_recurrentes
            """
            self.df_pagos_recurrentes = pd.read_sql_query(query, conn)

            self.df_pagos_recurrentes['monto'] = pd.to_numeric(
                self.df_pagos_recurrentes['monto'], errors='coerce'
            )

            print(f"✅ {len(self.df_pagos_recurrentes)} pagos recurrentes extraídos")

        except Exception as e:
            print(f"❌ Error extrayendo pagos recurrentes: {e}")
        finally:
            conn.close()

    def extraer_proveedores(self):
        """Extrae información de proveedores"""
        print("\n👥 Extrayendo proveedores...")

        conn = self.conectar_mysql('AD17_Proveedores')
        if not conn:
            return

        try:
            query = """
                SELECT
                    i.id,
                    d.nombre,
                    d.rfc,
                    COUNT(DISTINCT p.regID) as num_metodos_pago,
                    c.email,
                    c.telefono
                FROM ID i
                LEFT JOIN (
                    SELECT * FROM Datos
                    WHERE regID IN (SELECT MAX(regID) FROM Datos GROUP BY provID)
                ) d ON d.provID = i.id
                LEFT JOIN (
                    SELECT * FROM Contactos
                    WHERE regID IN (SELECT MAX(regID) FROM Contactos GROUP BY provID)
                ) c ON c.provID = i.id
                LEFT JOIN MetodosDePago p ON p.provID = i.id
                GROUP BY i.id
            """

            with conn.cursor() as cursor:
                cursor.execute(query)
                results = cursor.fetchall()
                self.df_proveedores = pd.DataFrame(results)

            print(f"✅ {len(self.df_proveedores)} proveedores extraídos")

        except Exception as e:
            print(f"❌ Error extrayendo proveedores: {e}")
        finally:
            if conn:
                conn.close()

    def calcular_metricas_generales(self):
        """Calcula métricas clave del año"""
        print("\n📈 Calculando métricas generales...")

        try:
            # === SOLICITUDES ===
            total_solicitudes = len(self.df_solicitudes)
            monto_total_solicitudes = self.df_solicitudes['monto'].sum()

            solicitudes_por_estado = self.df_solicitudes.groupby('estado').agg({
                'id': 'count',
                'monto': 'sum'
            }).rename(columns={'id': 'cantidad', 'monto': 'monto_total'})

            monto_promedio = self.df_solicitudes['monto'].mean()
            monto_mediana = self.df_solicitudes['monto'].median()
            monto_max = self.df_solicitudes['monto'].max()
            monto_min = self.df_solicitudes['monto'].min()

            # Anticipos
            solicitudes_con_anticipo = self.df_solicitudes[
                self.df_solicitudes['anticipo'].str.lower() == 'si'
            ]
            total_anticipos = len(solicitudes_con_anticipo)
            monto_total_anticipos = pd.to_numeric(
                solicitudes_con_anticipo['monto_anticipo'], errors='coerce'
            ).sum()

            # === COMISIONES BBVA SIN FACTURA ===
            solicitudes_bbva_sin_factura = self.df_solicitudes[
                (self.df_solicitudes['tiene_comision'] == 1) &
                (self.df_solicitudes['banco'].str.upper().str.contains('BBVA', na=False))
            ]

            if len(solicitudes_bbva_sin_factura) == 0:
                solicitudes_con_comision = self.df_solicitudes[
                    self.df_solicitudes['tiene_comision'] == 1
                ]
            else:
                solicitudes_con_comision = solicitudes_bbva_sin_factura

            total_comisiones_bbva = solicitudes_con_comision['monto_comision'].sum()
            cantidad_comisiones_bbva = len(solicitudes_con_comision)

            comisiones_por_mes = solicitudes_con_comision.groupby('mes').agg({
                'monto_comision': 'sum',
                'id': 'count',
                'monto_sin_comision': 'sum'
            }).rename(columns={'id': 'cantidad'})

            # === CRÉDITOS ===
            total_creditos = len(self.df_creditos)
            monto_total_creditos = self.df_creditos['monto_total'].sum()
            total_pagado_creditos = self.df_pagos_credito['monto'].sum()

            creditos_activos = len(self.df_creditos[self.df_creditos['estado'] == 'Activo'])
            creditos_liquidados = len(self.df_creditos[self.df_creditos['estado'] == 'Liquidado'])

            # === PAGOS RECURRENTES ===
            total_recurrentes = len(self.df_pagos_recurrentes)
            recurrentes_activos = len(
                self.df_pagos_recurrentes[self.df_pagos_recurrentes['activo'] == 1]
            )
            monto_mensual_recurrentes = self.df_pagos_recurrentes[
                (self.df_pagos_recurrentes['activo'] == 1) &
                (self.df_pagos_recurrentes['periodicidad'] == 'mensual')
            ]['monto'].sum()

            # === PROVEEDORES ===
            total_proveedores = len(self.df_proveedores)

            # Guardar métricas
            self.metricas = {
                'solicitudes': {
                    'total': total_solicitudes,
                    'monto_total': monto_total_solicitudes,
                    'monto_promedio': monto_promedio,
                    'monto_mediana': monto_mediana,
                    'monto_max': monto_max,
                    'monto_min': monto_min,
                    'por_estado': solicitudes_por_estado.to_dict(),
                    'con_anticipo': total_anticipos,
                    'monto_anticipos': monto_total_anticipos
                },
                'comisiones_bbva': {
                    'total': total_comisiones_bbva,
                    'cantidad': cantidad_comisiones_bbva,
                    'por_mes': comisiones_por_mes.to_dict() if not comisiones_por_mes.empty else {},
                    'df_detalle': solicitudes_con_comision
                },
                'creditos': {
                    'total': total_creditos,
                    'activos': creditos_activos,
                    'liquidados': creditos_liquidados,
                    'monto_total': monto_total_creditos,
                    'total_pagado': total_pagado_creditos
                },
                'recurrentes': {
                    'total': total_recurrentes,
                    'activos': recurrentes_activos,
                    'monto_mensual': monto_mensual_recurrentes
                },
                'proveedores': {
                    'total': total_proveedores
                }
            }

            print("✅ Métricas calculadas exitosamente")

        except Exception as e:
            print(f"❌ Error calculando métricas: {e}")
            import traceback
            traceback.print_exc()

    # ============================================================
    # GENERACIÓN DE GRÁFICAS - DISEÑO PROFESIONAL
    # ============================================================

    def generar_portada(self, pdf):
        """Genera portada profesional sin superposición de elementos"""
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        # Header con color corporativo
        header = mpatches.FancyBboxPatch(
            (0, 0.82), 1, 0.18,
            boxstyle="square,pad=0",
            facecolor=COLORS['primary'],
            edgecolor='none',
            transform=ax.transAxes
        )
        ax.add_patch(header)

        # Logo/Título principal sobre el header
        ax.text(0.5, 0.91, 'AD17 SOLUTIONS',
               ha='center', va='center', fontsize=42,
               fontweight='bold', color='white',
               transform=ax.transAxes)

        ax.text(0.5, 0.84, 'Sistema de Gestión de Costos',
               ha='center', va='center', fontsize=16,
               color='white', style='italic',
               transform=ax.transAxes)

        # Título del reporte (área central)
        ax.text(0.5, 0.65, f'REPORTE ANUAL {self.year}',
               ha='center', va='center', fontsize=36,
               fontweight='bold', color=COLORS['primary_dark'],
               transform=ax.transAxes)

        ax.text(0.5, 0.57, 'Análisis Financiero Completo',
               ha='center', va='center', fontsize=18,
               color=COLORS['dark'],
               transform=ax.transAxes)

        # Línea decorativa
        ax.plot([0.25, 0.75], [0.52, 0.52],
               color=COLORS['primary'], linewidth=3,
               transform=ax.transAxes)

        # Cuadro de métricas
        metrics_box = mpatches.FancyBboxPatch(
            (0.15, 0.15), 0.7, 0.32,
            boxstyle="round,pad=0.02,rounding_size=0.02",
            facecolor=COLORS['light'],
            edgecolor=COLORS['primary'],
            linewidth=2,
            transform=ax.transAxes
        )
        ax.add_patch(metrics_box)

        # Métricas dentro del cuadro
        metrics = [
            ('Total de Solicitudes', f"{self.metricas['solicitudes']['total']:,}"),
            ('Monto Total Procesado', f"${self.metricas['solicitudes']['monto_total']:,.0f}"),
            ('Total de Proveedores', f"{self.metricas['proveedores']['total']}"),
            ('Créditos Gestionados', f"{self.metricas['creditos']['total']}")
        ]

        y_start = 0.42
        for i, (label, value) in enumerate(metrics):
            y_pos = y_start - (i * 0.065)
            ax.text(0.5, y_pos, f'{label}: {value}',
                   ha='center', va='center', fontsize=13,
                   color=COLORS['dark'], fontweight='medium',
                   transform=ax.transAxes)

        # Fecha de generación
        ax.text(0.5, 0.05, f'Generado el {datetime.now().strftime("%d de %B de %Y")}',
               ha='center', va='center', fontsize=11,
               color=COLORS['gray'],
               transform=ax.transAxes)

        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)

        plt.tight_layout(pad=0)
        self.guardar_figura(fig, pdf, "Portada")

    def generar_resumen_ejecutivo(self, pdf):
        """Página de resumen ejecutivo con KPIs"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')

        fig.suptitle(f'Resumen Ejecutivo - {self.year}',
                    fontsize=20, fontweight='bold',
                    color=COLORS['dark'], y=0.96)

        gs = GridSpec(3, 3, figure=fig, hspace=0.35, wspace=0.25,
                     left=0.08, right=0.92, top=0.88, bottom=0.08)

        # KPIs
        kpis = [
            ('Total Solicitudes', f"{self.metricas['solicitudes']['total']:,}", COLORS['primary']),
            ('Monto Procesado', format_currency(self.metricas['solicitudes']['monto_total']), COLORS['secondary']),
            ('Monto Promedio', format_currency(self.metricas['solicitudes']['monto_promedio']), COLORS['success']),
            ('Créditos Activos', f"{self.metricas['creditos']['activos']}", COLORS['info']),
            ('Total Proveedores', f"{self.metricas['proveedores']['total']}", COLORS['primary_dark']),
            ('Pagos Recurrentes', f"{self.metricas['recurrentes']['activos']}", COLORS['warning']),
            ('Comisiones BBVA', format_currency(self.metricas['comisiones_bbva']['total']), COLORS['danger']),
            ('Solicitudes c/Anticipo', f"{self.metricas['solicitudes']['con_anticipo']}", COLORS['success']),
            ('Monto en Anticipos', format_currency(self.metricas['solicitudes']['monto_anticipos']), COLORS['info'])
        ]

        for idx, (label, valor, color) in enumerate(kpis):
            row = idx // 3
            col = idx % 3
            ax = fig.add_subplot(gs[row, col])
            ax.axis('off')

            # Tarjeta con sombra
            shadow = mpatches.FancyBboxPatch(
                (0.03, 0.03), 0.94, 0.94,
                boxstyle="round,pad=0.02,rounding_size=0.05",
                facecolor='#E0E0E0',
                edgecolor='none'
            )
            ax.add_patch(shadow)

            card = mpatches.FancyBboxPatch(
                (0, 0.06), 0.94, 0.94,
                boxstyle="round,pad=0.02,rounding_size=0.05",
                facecolor='white',
                edgecolor=color,
                linewidth=3
            )
            ax.add_patch(card)

            # Barra de color superior
            color_bar = mpatches.FancyBboxPatch(
                (0, 0.85), 0.94, 0.15,
                boxstyle="round,pad=0,rounding_size=0.05",
                facecolor=color,
                edgecolor='none'
            )
            ax.add_patch(color_bar)

            # Valor y etiqueta
            ax.text(0.47, 0.5, str(valor), ha='center', va='center',
                   fontsize=26, fontweight='bold', color=color)
            ax.text(0.47, 0.2, label, ha='center', va='center',
                   fontsize=11, color=COLORS['dark'])

            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)

        self.guardar_figura(fig, pdf, "Resumen_Ejecutivo")

    def generar_analisis_estados(self, pdf):
        """Análisis de estados con diseño mejorado"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis por Estados - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        gs = GridSpec(2, 2, figure=fig, hspace=0.3, wspace=0.25,
                     left=0.08, right=0.92, top=0.88, bottom=0.1)

        # Datos
        estados_count = self.df_solicitudes['estado'].value_counts()
        estados_monto = self.df_solicitudes.groupby('estado')['monto'].sum()

        # Colores por estado
        estado_colors = {
            'Liquidado': COLORS['success'],
            'Aprobado': COLORS['info'],
            'Pendiente': COLORS['warning'],
            'Rechazado': COLORS['danger'],
            'Cancelado': COLORS['gray']
        }
        colors_list = [estado_colors.get(e, COLORS['primary']) for e in estados_count.index]

        # 1. Donut chart - Cantidad
        ax1 = fig.add_subplot(gs[0, 0])
        wedges, texts, autotexts = ax1.pie(
            estados_count.values,
            labels=None,
            autopct='%1.1f%%',
            colors=colors_list,
            startangle=90,
            pctdistance=0.75,
            wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
        )

        for autotext in autotexts:
            autotext.set_fontsize(10)
            autotext.set_fontweight('bold')

        ax1.legend(wedges, estados_count.index,
                  title="Estado", loc="center left",
                  bbox_to_anchor=(1, 0, 0.5, 1),
                  fontsize=9)
        ax1.set_title('Distribución por Cantidad', fontweight='bold', pad=15)

        # 2. Donut chart - Monto
        ax2 = fig.add_subplot(gs[0, 1])
        colors_monto = [estado_colors.get(e, COLORS['primary']) for e in estados_monto.index]

        wedges2, texts2, autotexts2 = ax2.pie(
            estados_monto.values,
            labels=None,
            autopct='%1.1f%%',
            colors=colors_monto,
            startangle=90,
            pctdistance=0.75,
            wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
        )

        for autotext in autotexts2:
            autotext.set_fontsize(10)
            autotext.set_fontweight('bold')

        ax2.legend(wedges2, estados_monto.index,
                  title="Estado", loc="center left",
                  bbox_to_anchor=(1, 0, 0.5, 1),
                  fontsize=9)
        ax2.set_title('Distribución por Monto', fontweight='bold', pad=15)

        # 3. Barras horizontales - Monto por estado
        ax3 = fig.add_subplot(gs[1, 0])
        estados_monto_sorted = estados_monto.sort_values()
        colors_sorted = [estado_colors.get(e, COLORS['primary']) for e in estados_monto_sorted.index]

        bars = ax3.barh(range(len(estados_monto_sorted)),
                       estados_monto_sorted.values,
                       color=colors_sorted,
                       edgecolor='white',
                       linewidth=1)

        ax3.set_yticks(range(len(estados_monto_sorted)))
        ax3.set_yticklabels(estados_monto_sorted.index)
        ax3.set_xlabel('Monto Total')
        ax3.set_title('Monto Total por Estado', fontweight='bold')
        ax3.spines['left'].set_visible(False)
        ax3.tick_params(axis='y', length=0)

        for i, (bar, v) in enumerate(zip(bars, estados_monto_sorted.values)):
            ax3.text(v + estados_monto_sorted.max() * 0.02, i,
                    format_currency(v), va='center', fontsize=9, fontweight='medium')

        ax3.set_xlim(0, estados_monto_sorted.max() * 1.25)

        # 4. Tabla resumen
        ax4 = fig.add_subplot(gs[1, 1])
        ax4.axis('off')

        tabla_data = []
        for estado in estados_count.index:
            cantidad = estados_count[estado]
            monto = estados_monto.get(estado, 0)
            promedio = monto / cantidad if cantidad > 0 else 0
            tabla_data.append([
                estado,
                f'{cantidad:,}',
                format_currency(monto, False),
                format_currency(promedio, False)
            ])

        tabla = ax4.table(
            cellText=tabla_data,
            colLabels=['Estado', 'Cantidad', 'Monto Total', 'Promedio'],
            cellLoc='center',
            loc='center',
            colWidths=[0.25, 0.15, 0.30, 0.30]
        )
        tabla.auto_set_font_size(False)
        tabla.set_fontsize(10)
        tabla.scale(1, 2.2)

        for i in range(len(tabla_data) + 1):
            for j in range(4):
                cell = tabla[(i, j)]
                if i == 0:
                    cell.set_facecolor(COLORS['primary'])
                    cell.set_text_props(weight='bold', color='white')
                else:
                    cell.set_facecolor('#F5F5F5' if i % 2 == 0 else 'white')
                cell.set_edgecolor('#E0E0E0')

        ax4.set_title('Resumen por Estado', fontweight='bold', pad=20)

        self.guardar_figura(fig, pdf, "Analisis_Estados")

    def generar_analisis_temporal(self, pdf):
        """Análisis temporal con mejor diseño"""
        fig = plt.figure(figsize=(16, 12))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis Temporal - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.97)

        gs = GridSpec(3, 2, figure=fig, hspace=0.35, wspace=0.25,
                     left=0.08, right=0.92, top=0.90, bottom=0.08)

        meses_labels = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
                       'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']

        # 1. Evolución mensual de montos
        ax1 = fig.add_subplot(gs[0, 0])
        monthly_monto = self.df_solicitudes.groupby('mes')['monto'].sum()
        monthly_monto = monthly_monto.reindex(range(1, 13), fill_value=0)

        ax1.fill_between(range(1, 13), monthly_monto.values,
                        alpha=0.3, color=COLORS['primary'])
        ax1.plot(range(1, 13), monthly_monto.values,
                marker='o', linewidth=2.5, markersize=7,
                color=COLORS['primary'], markerfacecolor='white',
                markeredgewidth=2)

        ax1.set_xlabel('Mes')
        ax1.set_ylabel('Monto')
        ax1.set_title('Evolución Mensual de Montos', fontweight='bold')
        ax1.set_xticks(range(1, 13))
        ax1.set_xticklabels(meses_labels, fontsize=9)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 2. Solicitudes por mes
        ax2 = fig.add_subplot(gs[0, 1])
        monthly_count = self.df_solicitudes.groupby('mes').size()
        monthly_count = monthly_count.reindex(range(1, 13), fill_value=0)

        bars = ax2.bar(range(1, 13), monthly_count.values,
                      color=COLORS['secondary'], alpha=0.8,
                      edgecolor='white', linewidth=1)

        ax2.set_xlabel('Mes')
        ax2.set_ylabel('Cantidad')
        ax2.set_title('Solicitudes por Mes', fontweight='bold')
        ax2.set_xticks(range(1, 13))
        ax2.set_xticklabels(meses_labels, fontsize=9)

        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        f'{int(height)}', ha='center', va='bottom', fontsize=8)

        # 3. Distribución por trimestre
        ax3 = fig.add_subplot(gs[1, 0])
        quarterly = self.df_solicitudes.groupby('trimestre')['monto'].sum()
        quarterly = quarterly.reindex(range(1, 5), fill_value=0)

        colors_q = [COLORS['primary'], COLORS['secondary'],
                   COLORS['success'], COLORS['info']]

        bars_q = ax3.bar([f'Q{i}' for i in range(1, 5)], quarterly.values,
                        color=colors_q, alpha=0.85,
                        edgecolor='white', linewidth=2)

        ax3.set_ylabel('Monto')
        ax3.set_title('Distribución por Trimestre', fontweight='bold')
        ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        for bar in bars_q:
            height = bar.get_height()
            if height > 0:
                ax3.text(bar.get_x() + bar.get_width()/2., height,
                        format_currency(height), ha='center', va='bottom',
                        fontsize=9, fontweight='bold')

        # 4. Por día de la semana
        ax4 = fig.add_subplot(gs[1, 1])
        weekday_data = self.df_solicitudes.groupby('dia_semana')['monto'].sum()
        weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        weekday_labels_es = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
        weekday_data = weekday_data.reindex(weekday_order, fill_value=0)

        ax4.barh(range(len(weekday_labels_es)), weekday_data.values,
                color=COLORS['warning'], alpha=0.85,
                edgecolor='white', linewidth=1)

        ax4.set_yticks(range(len(weekday_labels_es)))
        ax4.set_yticklabels(weekday_labels_es)
        ax4.set_xlabel('Monto')
        ax4.set_title('Monto por Día de la Semana', fontweight='bold')
        ax4.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))
        ax4.spines['left'].set_visible(False)
        ax4.tick_params(axis='y', length=0)

        # 5. Heatmap simplificado
        ax5 = fig.add_subplot(gs[2, 0])
        heatmap_data = self.df_solicitudes.pivot_table(
            values='monto',
            index='tipo_solicitud',
            columns='mes',
            aggfunc='sum',
            fill_value=0
        )

        if not heatmap_data.empty and len(heatmap_data) > 0:
            heatmap_data.index = [truncate_text(str(x), 20) for x in heatmap_data.index]

            sns.heatmap(heatmap_data, ax=ax5, cmap='YlOrRd',
                       cbar_kws={'label': 'Monto', 'shrink': 0.8},
                       linewidths=0.5, linecolor='white',
                       xticklabels=meses_labels[:len(heatmap_data.columns)])
            ax5.set_xlabel('Mes')
            ax5.set_ylabel('')
            ax5.set_title('Monto por Tipo de Solicitud y Mes', fontweight='bold')
            ax5.tick_params(axis='y', labelsize=8)
        else:
            ax5.text(0.5, 0.5, 'Sin datos suficientes', ha='center', va='center')
            ax5.axis('off')

        # 6. Acumulado anual
        ax6 = fig.add_subplot(gs[2, 1])
        df_sorted = self.df_solicitudes.sort_values('fecha').copy()
        df_sorted['monto_acumulado'] = df_sorted['monto'].cumsum()

        ax6.fill_between(range(len(df_sorted)), df_sorted['monto_acumulado'].values,
                        alpha=0.3, color=COLORS['success'])
        ax6.plot(range(len(df_sorted)), df_sorted['monto_acumulado'].values,
                color=COLORS['success'], linewidth=2)

        ax6.set_xlabel('Solicitudes (orden cronológico)')
        ax6.set_ylabel('Monto Acumulado')
        ax6.set_title('Acumulado Anual', fontweight='bold')
        ax6.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        total = df_sorted['monto_acumulado'].iloc[-1] if len(df_sorted) > 0 else 0
        ax6.text(0.95, 0.95, f'Total: {format_currency(total)}',
                transform=ax6.transAxes, ha='right', va='top',
                fontsize=11, fontweight='bold',
                bbox=dict(boxstyle='round', facecolor='white',
                         edgecolor=COLORS['success'], alpha=0.9))

        self.guardar_figura(fig, pdf, "Analisis_Temporal")

    def generar_analisis_proveedores(self, pdf):
        """Análisis de proveedores con nombres truncados y mejor diseño"""
        fig = plt.figure(figsize=(16, 12))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis de Proveedores - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.97)

        gs = GridSpec(2, 2, figure=fig, hspace=0.35, wspace=0.3,
                     left=0.1, right=0.92, top=0.90, bottom=0.08)

        # Datos
        prov_data = self.df_solicitudes.groupby('destinatario').agg({
            'monto': 'sum',
            'id': 'count'
        }).sort_values('monto', ascending=False)

        # 1. Top 12 por monto
        ax1 = fig.add_subplot(gs[0, 0])
        top_12 = prov_data.head(12)
        nombres_truncados = [truncate_text(str(n), 22) for n in top_12.index]

        bars = ax1.barh(range(len(top_12)), top_12['monto'].values,
                       color=COLORS['primary'], alpha=0.85,
                       edgecolor='white', linewidth=1)

        ax1.set_yticks(range(len(top_12)))
        ax1.set_yticklabels(nombres_truncados, fontsize=9)
        ax1.set_xlabel('Monto Total')
        ax1.set_title('Top 12 Proveedores por Monto', fontweight='bold')
        ax1.spines['left'].set_visible(False)
        ax1.tick_params(axis='y', length=0)
        ax1.invert_yaxis()

        max_val = top_12['monto'].max()
        for i, v in enumerate(top_12['monto'].values):
            ax1.text(v + max_val * 0.02, i, format_currency(v),
                    va='center', fontsize=8, fontweight='medium')

        ax1.set_xlim(0, max_val * 1.25)

        # 2. Curva de Pareto
        ax2 = fig.add_subplot(gs[0, 1])
        prov_data['monto_pct'] = (prov_data['monto'] / prov_data['monto'].sum()) * 100
        prov_data_sorted = prov_data.sort_values('monto_pct', ascending=False)
        prov_data_sorted['acumulado'] = prov_data_sorted['monto_pct'].cumsum()

        x_range = range(1, len(prov_data_sorted) + 1)

        ax2.fill_between(x_range, prov_data_sorted['acumulado'].values,
                        alpha=0.3, color=COLORS['danger'])
        ax2.plot(x_range, prov_data_sorted['acumulado'].values,
                color=COLORS['danger'], linewidth=2.5, marker='')
        ax2.axhline(y=80, color=COLORS['gray'], linestyle='--',
                   alpha=0.7, linewidth=1.5, label='80% del monto')

        ax2.set_xlabel('Proveedores (ordenados por monto)')
        ax2.set_ylabel('% Acumulado del Monto')
        ax2.set_title('Curva de Concentración (Pareto)', fontweight='bold')
        ax2.legend(loc='lower right')
        ax2.set_ylim(0, 105)

        provs_80 = (prov_data_sorted['acumulado'] <= 80).sum()
        ax2.annotate(f'{provs_80} proveedores\ngeneran 80%',
                    xy=(provs_80, 80), xytext=(provs_80 + len(prov_data_sorted)*0.1, 60),
                    fontsize=9, ha='left',
                    arrowprops=dict(arrowstyle='->', color=COLORS['dark']))

        # 3. Top 12 por frecuencia
        ax3 = fig.add_subplot(gs[1, 0])
        top_freq = prov_data.sort_values('id', ascending=False).head(12)
        nombres_freq = [truncate_text(str(n), 22) for n in top_freq.index]

        bars_freq = ax3.barh(range(len(top_freq)), top_freq['id'].values,
                            color=COLORS['secondary'], alpha=0.85,
                            edgecolor='white', linewidth=1)

        ax3.set_yticks(range(len(top_freq)))
        ax3.set_yticklabels(nombres_freq, fontsize=9)
        ax3.set_xlabel('Número de Solicitudes')
        ax3.set_title('Top 12 Proveedores por Frecuencia', fontweight='bold')
        ax3.spines['left'].set_visible(False)
        ax3.tick_params(axis='y', length=0)
        ax3.invert_yaxis()

        for i, v in enumerate(top_freq['id'].values):
            ax3.text(v + top_freq['id'].max() * 0.03, i, str(v),
                    va='center', fontsize=9, fontweight='medium')

        # 4. Ticket promedio
        ax4 = fig.add_subplot(gs[1, 1])
        prov_data['ticket_promedio'] = prov_data['monto'] / prov_data['id']

        prov_min_freq = prov_data[prov_data['id'] >= 3]
        top_ticket = prov_min_freq.nlargest(10, 'ticket_promedio')
        nombres_ticket = [truncate_text(str(n), 22) for n in top_ticket.index]

        bars_ticket = ax4.barh(range(len(top_ticket)), top_ticket['ticket_promedio'].values,
                              color=COLORS['success'], alpha=0.85,
                              edgecolor='white', linewidth=1)

        ax4.set_yticks(range(len(top_ticket)))
        ax4.set_yticklabels(nombres_ticket, fontsize=9)
        ax4.set_xlabel('Ticket Promedio')
        ax4.set_title('Top 10 Proveedores por Ticket Promedio\n(mín. 3 solicitudes)', fontweight='bold')
        ax4.spines['left'].set_visible(False)
        ax4.tick_params(axis='y', length=0)
        ax4.invert_yaxis()
        ax4.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        self.guardar_figura(fig, pdf, "Analisis_Proveedores")

    def generar_analisis_departamentos(self, pdf):
        """Análisis de departamentos con diseño mejorado"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis por Departamentos - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        gs = GridSpec(2, 2, figure=fig, hspace=0.35, wspace=0.3,
                     left=0.1, right=0.92, top=0.88, bottom=0.1)

        # Datos
        dept_data = self.df_solicitudes.groupby('departamento').agg({
            'monto': 'sum',
            'id': 'count'
        }).sort_values('monto', ascending=False)
        dept_data['promedio'] = dept_data['monto'] / dept_data['id']

        # 1. Top 10 por monto
        ax1 = fig.add_subplot(gs[0, 0])
        top_10 = dept_data.head(10)
        nombres_dept = [truncate_text(str(n), 20) for n in top_10.index]

        bars = ax1.barh(range(len(top_10)), top_10['monto'].values,
                       color=COLORS['primary'], alpha=0.85,
                       edgecolor='white', linewidth=1)

        ax1.set_yticks(range(len(top_10)))
        ax1.set_yticklabels(nombres_dept, fontsize=9)
        ax1.set_xlabel('Monto Total')
        ax1.set_title('Top 10 Departamentos por Monto', fontweight='bold')
        ax1.spines['left'].set_visible(False)
        ax1.tick_params(axis='y', length=0)
        ax1.invert_yaxis()
        ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 2. Top 10 por cantidad
        ax2 = fig.add_subplot(gs[0, 1])
        top_count = dept_data.sort_values('id', ascending=False).head(10)
        nombres_count = [truncate_text(str(n), 20) for n in top_count.index]

        bars2 = ax2.barh(range(len(top_count)), top_count['id'].values,
                        color=COLORS['secondary'], alpha=0.85,
                        edgecolor='white', linewidth=1)

        ax2.set_yticks(range(len(top_count)))
        ax2.set_yticklabels(nombres_count, fontsize=9)
        ax2.set_xlabel('Cantidad de Solicitudes')
        ax2.set_title('Top 10 Departamentos por Cantidad', fontweight='bold')
        ax2.spines['left'].set_visible(False)
        ax2.tick_params(axis='y', length=0)
        ax2.invert_yaxis()

        for i, v in enumerate(top_count['id'].values):
            ax2.text(v + top_count['id'].max() * 0.02, i, str(v),
                    va='center', fontsize=9)

        # 3. Top 10 por promedio
        ax3 = fig.add_subplot(gs[1, 0])
        top_avg = dept_data.nlargest(10, 'promedio')
        nombres_avg = [truncate_text(str(n), 20) for n in top_avg.index]

        bars3 = ax3.barh(range(len(top_avg)), top_avg['promedio'].values,
                        color=COLORS['success'], alpha=0.85,
                        edgecolor='white', linewidth=1)

        ax3.set_yticks(range(len(top_avg)))
        ax3.set_yticklabels(nombres_avg, fontsize=9)
        ax3.set_xlabel('Monto Promedio por Solicitud')
        ax3.set_title('Top 10 Departamentos por Promedio', fontweight='bold')
        ax3.spines['left'].set_visible(False)
        ax3.tick_params(axis='y', length=0)
        ax3.invert_yaxis()
        ax3.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 4. Distribución
        ax4 = fig.add_subplot(gs[1, 1])
        top_5 = dept_data.head(5)
        otros = dept_data.iloc[5:]['monto'].sum() if len(dept_data) > 5 else 0

        sizes = list(top_5['monto'].values) + ([otros] if otros > 0 else [])
        labels = [truncate_text(str(n), 15) for n in top_5.index] + (['Otros'] if otros > 0 else [])
        colors_tree = CHART_COLORS[:len(sizes)]

        wedges, texts = ax4.pie(sizes, labels=None, colors=colors_tree,
                               startangle=90,
                               wedgeprops=dict(width=0.6, edgecolor='white', linewidth=2))

        total = sum(sizes)
        legend_labels = [f'{l} ({s/total*100:.1f}%)' for l, s in zip(labels, sizes)]
        ax4.legend(wedges, legend_labels, title="Departamento",
                  loc="center left", bbox_to_anchor=(1, 0.5),
                  fontsize=9)

        ax4.set_title('Distribución del Gasto por Departamento', fontweight='bold')

        self.guardar_figura(fig, pdf, "Analisis_Departamentos")

    def generar_analisis_comisiones_bbva(self, pdf):
        """Análisis de comisiones BBVA por pagos sin factura"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis de Comisiones BBVA (Sin Factura) - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        gs = GridSpec(2, 2, figure=fig, hspace=0.35, wspace=0.3,
                     left=0.1, right=0.92, top=0.88, bottom=0.1)

        df_comisiones = self.metricas['comisiones_bbva']['df_detalle'].copy()

        if len(df_comisiones) == 0:
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, 'No hay datos de comisiones BBVA\npor pagos sin factura en este período',
                   ha='center', va='center', fontsize=14,
                   transform=ax.transAxes)
            ax.axis('off')
            self.guardar_figura(fig, pdf, "Analisis_Comisiones_BBVA")
            return

        meses_labels = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
                       'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']

        # 1. Evolución mensual de comisiones
        ax1 = fig.add_subplot(gs[0, 0])
        com_por_mes = df_comisiones.groupby('mes')['monto_comision'].sum()
        com_por_mes = com_por_mes.reindex(range(1, 13), fill_value=0)

        ax1.fill_between(range(1, 13), com_por_mes.values,
                        alpha=0.3, color=COLORS['danger'])
        ax1.plot(range(1, 13), com_por_mes.values,
                marker='o', linewidth=2.5, markersize=7,
                color=COLORS['danger'], markerfacecolor='white',
                markeredgewidth=2)

        ax1.set_xlabel('Mes')
        ax1.set_ylabel('Comisiones Pagadas')
        ax1.set_title('Comisiones BBVA por Mes', fontweight='bold')
        ax1.set_xticks(range(1, 13))
        ax1.set_xticklabels(meses_labels, fontsize=9)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        total_com = df_comisiones['monto_comision'].sum()
        ax1.text(0.95, 0.95, f'Total: {format_currency(total_com)}',
                transform=ax1.transAxes, ha='right', va='top',
                fontsize=11, fontweight='bold',
                bbox=dict(boxstyle='round', facecolor='white',
                         edgecolor=COLORS['danger'], alpha=0.9))

        # 2. Cantidad de operaciones con comisión por mes
        ax2 = fig.add_subplot(gs[0, 1])
        cant_por_mes = df_comisiones.groupby('mes').size()
        cant_por_mes = cant_por_mes.reindex(range(1, 13), fill_value=0)

        bars = ax2.bar(range(1, 13), cant_por_mes.values,
                      color=COLORS['warning'], alpha=0.85,
                      edgecolor='white', linewidth=1)

        ax2.set_xlabel('Mes')
        ax2.set_ylabel('Cantidad de Operaciones')
        ax2.set_title('Operaciones con Comisión por Mes', fontweight='bold')
        ax2.set_xticks(range(1, 13))
        ax2.set_xticklabels(meses_labels, fontsize=9)

        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        f'{int(height)}', ha='center', va='bottom', fontsize=9)

        # 3. Distribución de porcentajes de comisión
        ax3 = fig.add_subplot(gs[1, 0])

        if 'porcentaje_comision' in df_comisiones.columns:
            porcentajes = df_comisiones['porcentaje_comision'].dropna()
            if len(porcentajes) > 0:
                pct_counts = porcentajes.value_counts().sort_index()

                ax3.bar(pct_counts.index.astype(str), pct_counts.values,
                       color=COLORS['info'], alpha=0.85,
                       edgecolor='white', linewidth=1)

                ax3.set_xlabel('Porcentaje de Comisión (%)')
                ax3.set_ylabel('Número de Operaciones')
                ax3.set_title('Distribución por Porcentaje de Comisión', fontweight='bold')

                for i, (pct, count) in enumerate(zip(pct_counts.index, pct_counts.values)):
                    ax3.text(i, count, str(count), ha='center', va='bottom', fontsize=9)
            else:
                ax3.text(0.5, 0.5, 'Sin datos de porcentaje', ha='center', va='center')
                ax3.axis('off')
        else:
            ax3.text(0.5, 0.5, 'Sin datos de porcentaje', ha='center', va='center')
            ax3.axis('off')

        # 4. Top proveedores con más comisiones
        ax4 = fig.add_subplot(gs[1, 1])

        com_por_prov = df_comisiones.groupby('destinatario')['monto_comision'].sum()
        top_com_prov = com_por_prov.nlargest(10)
        nombres_prov = [truncate_text(str(n), 20) for n in top_com_prov.index]

        bars_prov = ax4.barh(range(len(top_com_prov)), top_com_prov.values,
                            color=COLORS['danger'], alpha=0.85,
                            edgecolor='white', linewidth=1)

        ax4.set_yticks(range(len(top_com_prov)))
        ax4.set_yticklabels(nombres_prov, fontsize=9)
        ax4.set_xlabel('Comisiones Pagadas')
        ax4.set_title('Top 10 Proveedores por Comisiones', fontweight='bold')
        ax4.spines['left'].set_visible(False)
        ax4.tick_params(axis='y', length=0)
        ax4.invert_yaxis()
        ax4.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        self.guardar_figura(fig, pdf, "Analisis_Comisiones_BBVA")

    def generar_analisis_creditos(self, pdf):
        """Análisis de créditos con diseño mejorado"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis de Créditos - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        gs = GridSpec(2, 3, figure=fig, hspace=0.35, wspace=0.3,
                     left=0.08, right=0.92, top=0.88, bottom=0.1)

        if len(self.df_creditos) == 0:
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, 'No hay datos de créditos disponibles',
                   ha='center', va='center', fontsize=14)
            ax.axis('off')
            self.guardar_figura(fig, pdf, "Analisis_Creditos")
            return

        # 1. Distribución por estado
        ax1 = fig.add_subplot(gs[0, 0])
        estados = self.df_creditos['estado'].value_counts()

        colors_estado = [COLORS['success'] if e == 'Activo' else COLORS['info']
                        for e in estados.index]

        wedges, texts, autotexts = ax1.pie(
            estados.values, labels=estados.index, autopct='%1.1f%%',
            colors=colors_estado, startangle=90,
            wedgeprops=dict(edgecolor='white', linewidth=2)
        )
        ax1.set_title('Distribución por Estado', fontweight='bold')

        # 2. Distribución por tipo
        ax2 = fig.add_subplot(gs[0, 1])
        tipos = self.df_creditos['tipo_credito'].value_counts()

        ax2.bar(range(len(tipos)), tipos.values,
               color=COLORS['primary'], alpha=0.85,
               edgecolor='white', linewidth=1)
        ax2.set_xticks(range(len(tipos)))
        ax2.set_xticklabels([truncate_text(str(t), 12) for t in tipos.index],
                          rotation=45, ha='right', fontsize=9)
        ax2.set_ylabel('Cantidad')
        ax2.set_title('Créditos por Tipo', fontweight='bold')

        # 3. Monto por entidad
        ax3 = fig.add_subplot(gs[0, 2])
        entidad_monto = self.df_creditos.groupby('entidad')['monto_total'].sum().sort_values()
        nombres_entidad = [truncate_text(str(e), 15) for e in entidad_monto.index]

        ax3.barh(range(len(entidad_monto)), entidad_monto.values,
                color=COLORS['secondary'], alpha=0.85,
                edgecolor='white', linewidth=1)
        ax3.set_yticks(range(len(entidad_monto)))
        ax3.set_yticklabels(nombres_entidad, fontsize=9)
        ax3.set_xlabel('Monto Total')
        ax3.set_title('Monto por Entidad', fontweight='bold')
        ax3.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 4. Evolución de pagos
        ax4 = fig.add_subplot(gs[1, 0])
        if not self.df_pagos_credito.empty:
            pagos_mes = self.df_pagos_credito.groupby(
                self.df_pagos_credito['fecha'].dt.month
            )['monto'].sum()
            pagos_mes = pagos_mes.reindex(range(1, 13), fill_value=0)

            meses_labels = ['E', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']

            ax4.fill_between(range(1, 13), pagos_mes.values,
                            alpha=0.3, color=COLORS['success'])
            ax4.plot(range(1, 13), pagos_mes.values,
                    marker='o', color=COLORS['success'], linewidth=2)
            ax4.set_xticks(range(1, 13))
            ax4.set_xticklabels(meses_labels)
            ax4.set_xlabel('Mes')
            ax4.set_ylabel('Monto Pagado')
            ax4.set_title('Pagos de Créditos por Mes', fontweight='bold')
            ax4.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))
        else:
            ax4.text(0.5, 0.5, 'Sin pagos registrados', ha='center', va='center')
            ax4.axis('off')

        # 5. Distribución de tasas
        ax5 = fig.add_subplot(gs[1, 1])
        tasas = self.df_creditos['tasa_interes'].dropna()
        if len(tasas) > 0:
            ax5.hist(tasas, bins=10, color=COLORS['warning'], alpha=0.85,
                    edgecolor='white', linewidth=1)
            ax5.axvline(x=tasas.mean(), color=COLORS['danger'], linestyle='--',
                       linewidth=2, label=f'Media: {tasas.mean():.1f}%')
            ax5.set_xlabel('Tasa de Interés (%)')
            ax5.set_ylabel('Frecuencia')
            ax5.set_title('Distribución de Tasas', fontweight='bold')
            ax5.legend()
        else:
            ax5.text(0.5, 0.5, 'Sin datos de tasas', ha='center', va='center')
            ax5.axis('off')

        # 6. Progreso de pagos
        ax6 = fig.add_subplot(gs[1, 2])
        creditos_activos = self.df_creditos[self.df_creditos['estado'] == 'Activo']

        if not creditos_activos.empty and not self.df_pagos_credito.empty:
            progress_data = []
            for _, credito in creditos_activos.iterrows():
                pagado = self.df_pagos_credito[
                    self.df_pagos_credito['credito_id'] == credito['id']
                ]['monto'].sum()
                porcentaje = (pagado / credito['monto_total']) * 100 if credito['monto_total'] > 0 else 0
                progress_data.append({
                    'nombre': truncate_text(str(credito['nombre']), 18),
                    'porcentaje': min(porcentaje, 100)
                })

            if progress_data:
                df_progress = pd.DataFrame(progress_data).sort_values('porcentaje', ascending=False)

                bars = ax6.barh(range(len(df_progress)), df_progress['porcentaje'].values,
                               color=COLORS['info'], alpha=0.85,
                               edgecolor='white', linewidth=1)
                ax6.set_yticks(range(len(df_progress)))
                ax6.set_yticklabels(df_progress['nombre'], fontsize=9)
                ax6.set_xlabel('% Pagado')
                ax6.set_title('Progreso de Pago', fontweight='bold')
                ax6.axvline(x=50, color=COLORS['warning'], linestyle='--', alpha=0.7)
                ax6.set_xlim(0, 105)
            else:
                ax6.text(0.5, 0.5, 'Sin datos de progreso', ha='center', va='center')
                ax6.axis('off')
        else:
            ax6.text(0.5, 0.5, 'Sin créditos activos', ha='center', va='center')
            ax6.axis('off')

        self.guardar_figura(fig, pdf, "Analisis_Creditos")

    def generar_analisis_anticipos(self, pdf):
        """Análisis de anticipos con diseño mejorado"""
        anticipos = self.df_solicitudes[
            self.df_solicitudes['anticipo'].str.lower() == 'si'
        ].copy()

        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis de Anticipos - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        if anticipos.empty:
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, 'No hay datos de anticipos en este período',
                   ha='center', va='center', fontsize=14)
            ax.axis('off')
            self.guardar_figura(fig, pdf, "Analisis_Anticipos")
            return

        gs = GridSpec(2, 2, figure=fig, hspace=0.35, wspace=0.3,
                     left=0.1, right=0.92, top=0.88, bottom=0.1)

        anticipos['monto_anticipo'] = pd.to_numeric(anticipos['monto_anticipo'], errors='coerce')
        anticipos['porcentaje_anticipo'] = pd.to_numeric(anticipos['porcentaje_anticipo'], errors='coerce')

        meses_labels = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
                       'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']

        # 1. Evolución mensual
        ax1 = fig.add_subplot(gs[0, 0])
        ant_mes = anticipos.groupby('mes')['monto_anticipo'].sum()
        ant_mes = ant_mes.reindex(range(1, 13), fill_value=0)

        ax1.fill_between(range(1, 13), ant_mes.values,
                        alpha=0.3, color=COLORS['success'])
        ax1.plot(range(1, 13), ant_mes.values,
                marker='o', linewidth=2.5, markersize=7,
                color=COLORS['success'], markerfacecolor='white',
                markeredgewidth=2)

        ax1.set_xlabel('Mes')
        ax1.set_ylabel('Monto en Anticipos')
        ax1.set_title('Evolución Mensual de Anticipos', fontweight='bold')
        ax1.set_xticks(range(1, 13))
        ax1.set_xticklabels(meses_labels, fontsize=9)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 2. Distribución de porcentajes
        ax2 = fig.add_subplot(gs[0, 1])
        porcentajes = anticipos['porcentaje_anticipo'].dropna()

        if len(porcentajes) > 0:
            ax2.hist(porcentajes, bins=15, color=COLORS['primary'], alpha=0.85,
                    edgecolor='white', linewidth=1)
            ax2.axvline(x=porcentajes.mean(), color=COLORS['danger'], linestyle='--',
                       linewidth=2, label=f'Media: {porcentajes.mean():.1f}%')
            ax2.set_xlabel('Porcentaje de Anticipo')
            ax2.set_ylabel('Frecuencia')
            ax2.set_title('Distribución de Porcentajes', fontweight='bold')
            ax2.legend()
        else:
            ax2.text(0.5, 0.5, 'Sin datos', ha='center', va='center')
            ax2.axis('off')

        # 3. Por estado
        ax3 = fig.add_subplot(gs[1, 0])
        ant_estado = anticipos.groupby('estado')['monto_anticipo'].sum().sort_values()

        estado_colors = {
            'Liquidado': COLORS['success'],
            'Aprobado': COLORS['info'],
            'Pendiente': COLORS['warning'],
            'Rechazado': COLORS['danger']
        }
        colors_est = [estado_colors.get(e, COLORS['gray']) for e in ant_estado.index]

        ax3.barh(range(len(ant_estado)), ant_estado.values,
                color=colors_est, alpha=0.85,
                edgecolor='white', linewidth=1)
        ax3.set_yticks(range(len(ant_estado)))
        ax3.set_yticklabels(ant_estado.index)
        ax3.set_xlabel('Monto en Anticipos')
        ax3.set_title('Anticipos por Estado', fontweight='bold')
        ax3.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 4. Top departamentos
        ax4 = fig.add_subplot(gs[1, 1])
        ant_dept = anticipos.groupby('departamento')['monto_anticipo'].sum().nlargest(10)
        nombres_dept = [truncate_text(str(n), 18) for n in ant_dept.index]

        ax4.barh(range(len(ant_dept)), ant_dept.values,
                color=COLORS['secondary'], alpha=0.85,
                edgecolor='white', linewidth=1)
        ax4.set_yticks(range(len(ant_dept)))
        ax4.set_yticklabels(nombres_dept, fontsize=9)
        ax4.set_xlabel('Monto en Anticipos')
        ax4.set_title('Top 10 Departamentos con Anticipos', fontweight='bold')
        ax4.spines['left'].set_visible(False)
        ax4.tick_params(axis='y', length=0)
        ax4.invert_yaxis()
        ax4.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        self.guardar_figura(fig, pdf, "Analisis_Anticipos")

    def generar_analisis_recurrentes(self, pdf):
        """Análisis de pagos recurrentes"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Análisis de Pagos Recurrentes - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        if self.df_pagos_recurrentes.empty:
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, 'No hay pagos recurrentes registrados',
                   ha='center', va='center', fontsize=14)
            ax.axis('off')
            self.guardar_figura(fig, pdf, "Analisis_Pagos_Recurrentes")
            return

        gs = GridSpec(2, 2, figure=fig, hspace=0.35, wspace=0.3,
                     left=0.1, right=0.92, top=0.88, bottom=0.1)

        # 1. Por periodicidad
        ax1 = fig.add_subplot(gs[0, 0])
        period = self.df_pagos_recurrentes['periodicidad'].value_counts()

        wedges, texts, autotexts = ax1.pie(
            period.values, labels=period.index, autopct='%1.1f%%',
            colors=CHART_COLORS[:len(period)], startangle=90,
            wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
        )
        ax1.set_title('Distribución por Periodicidad', fontweight='bold')

        # 2. Monto por periodicidad
        ax2 = fig.add_subplot(gs[0, 1])
        period_monto = self.df_pagos_recurrentes.groupby('periodicidad')['monto'].sum()

        ax2.bar(range(len(period_monto)), period_monto.values,
               color=COLORS['primary'], alpha=0.85,
               edgecolor='white', linewidth=1)
        ax2.set_xticks(range(len(period_monto)))
        ax2.set_xticklabels(period_monto.index, fontsize=10)
        ax2.set_ylabel('Monto Total')
        ax2.set_title('Monto por Periodicidad', fontweight='bold')
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 3. Top proveedores recurrentes
        ax3 = fig.add_subplot(gs[1, 0])
        top_prov = self.df_pagos_recurrentes.groupby('proveedor')['monto'].sum().nlargest(10)
        nombres_prov = [truncate_text(str(n), 18) for n in top_prov.index]

        ax3.barh(range(len(top_prov)), top_prov.values,
                color=COLORS['secondary'], alpha=0.85,
                edgecolor='white', linewidth=1)
        ax3.set_yticks(range(len(top_prov)))
        ax3.set_yticklabels(nombres_prov, fontsize=9)
        ax3.set_xlabel('Monto')
        ax3.set_title('Top 10 Proveedores Recurrentes', fontweight='bold')
        ax3.spines['left'].set_visible(False)
        ax3.tick_params(axis='y', length=0)
        ax3.invert_yaxis()
        ax3.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 4. Estado activo/inactivo
        ax4 = fig.add_subplot(gs[1, 1])
        activos = self.df_pagos_recurrentes['activo'].value_counts()
        labels_activo = ['Activo' if x == 1 else 'Inactivo' for x in activos.index]
        colors_activo = [COLORS['success'] if x == 1 else COLORS['gray'] for x in activos.index]

        wedges2, texts2, autotexts2 = ax4.pie(
            activos.values, labels=labels_activo, autopct='%1.1f%%',
            colors=colors_activo, startangle=90,
            wedgeprops=dict(edgecolor='white', linewidth=2)
        )
        ax4.set_title('Estado de Pagos Recurrentes', fontweight='bold')

        self.guardar_figura(fig, pdf, "Analisis_Pagos_Recurrentes")

    def generar_proyecciones(self, pdf):
        """Proyecciones y KPIs finales"""
        fig = plt.figure(figsize=(16, 10))
        fig.patch.set_facecolor('white')
        fig.suptitle(f'Proyecciones e Indicadores Clave - {self.year}',
                    fontsize=18, fontweight='bold', color=COLORS['dark'], y=0.96)

        gs = GridSpec(2, 2, figure=fig, hspace=0.35, wspace=0.25,
                     left=0.08, right=0.92, top=0.88, bottom=0.08)

        meses_labels = ['E', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']

        # 1. Proyección mensual
        ax1 = fig.add_subplot(gs[0, 0])
        monthly = self.df_solicitudes.groupby('mes')['monto'].sum()
        monthly = monthly.reindex(range(1, 13), fill_value=0)
        avg_monthly = monthly[monthly > 0].mean() if len(monthly[monthly > 0]) > 0 else 0

        ax1.plot(range(1, 13), monthly.values, marker='o', linewidth=2.5,
                color=COLORS['primary'], label='Real', markersize=6)

        ax1.axhline(y=avg_monthly, color=COLORS['gray'], linestyle='--',
                   alpha=0.7, linewidth=1.5, label=f'Promedio: {format_currency(avg_monthly)}')

        ax1.set_xlabel('Mes')
        ax1.set_ylabel('Monto')
        ax1.set_title('Evolución y Promedio Mensual', fontweight='bold')
        ax1.set_xticks(range(1, 13))
        ax1.set_xticklabels(meses_labels)
        ax1.legend(loc='upper left')
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        # 2. Comparativa trimestral
        ax2 = fig.add_subplot(gs[0, 1])
        quarterly = self.df_solicitudes.groupby('trimestre')['monto'].sum()
        quarterly = quarterly.reindex(range(1, 5), fill_value=0)

        colors_q = [COLORS['info'], COLORS['success'], COLORS['warning'], COLORS['danger']]
        bars = ax2.bar([f'Q{i}' for i in range(1, 5)], quarterly.values,
                      color=colors_q, alpha=0.85,
                      edgecolor='white', linewidth=2)

        ax2.plot([f'Q{i}' for i in range(1, 5)], quarterly.values,
                marker='D', color=COLORS['dark'], linewidth=2, markersize=8)

        ax2.set_ylabel('Monto')
        ax2.set_title('Evolución Trimestral', fontweight='bold')
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        format_currency(height), ha='center', va='bottom',
                        fontsize=9, fontweight='bold')

        # 3. Acumulado con meta
        ax3 = fig.add_subplot(gs[1, 0])
        monthly_cumsum = monthly.cumsum()

        ax3.fill_between(range(1, 13), monthly_cumsum.values,
                        alpha=0.3, color=COLORS['success'])
        ax3.plot(range(1, 13), monthly_cumsum.values,
                color=COLORS['success'], linewidth=2.5, marker='o')

        ax3.set_xlabel('Mes')
        ax3.set_ylabel('Acumulado')
        ax3.set_title('Acumulado Anual', fontweight='bold')
        ax3.set_xticks(range(1, 13))
        ax3.set_xticklabels(meses_labels)
        ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_currency(x)))

        total_anual = monthly_cumsum.iloc[-1] if len(monthly_cumsum) > 0 else 0
        ax3.text(0.95, 0.05, f'Total: {format_currency(total_anual)}',
                transform=ax3.transAxes, ha='right', va='bottom',
                fontsize=12, fontweight='bold',
                bbox=dict(boxstyle='round', facecolor='white',
                         edgecolor=COLORS['success'], alpha=0.9))

        # 4. Panel de KPIs
        ax4 = fig.add_subplot(gs[1, 1])
        ax4.axis('off')

        # Calcular KPIs
        solicitudes_mes = len(self.df_solicitudes) / 12
        ticket_prom = self.df_solicitudes['monto'].mean()
        proveedores_unicos = self.df_solicitudes['destinatario'].nunique()
        departamentos = self.df_solicitudes['departamento'].nunique()

        if len(quarterly) >= 2 and quarterly.iloc[0] > 0:
            crecimiento = ((quarterly.iloc[-1] / quarterly.iloc[0]) - 1) * 100
        else:
            crecimiento = 0

        kpis = [
            ('📊', 'Solicitudes/Mes', f'{solicitudes_mes:.0f}'),
            ('💰', 'Ticket Promedio', format_currency(ticket_prom)),
            ('📈', 'Var. Trimestral', f'{crecimiento:+.1f}%'),
            ('👥', 'Proveedores Únicos', f'{proveedores_unicos}'),
            ('🏢', 'Departamentos', f'{departamentos}'),
            ('💳', 'Créditos Activos', f"{self.metricas['creditos']['activos']}"),
            ('🔄', 'Pagos Recurrentes', f"{self.metricas['recurrentes']['activos']}"),
            ('💸', 'Comisiones BBVA', format_currency(self.metricas['comisiones_bbva']['total']))
        ]

        y_pos = 0.95
        for emoji, label, value in kpis:
            ax4.text(0.1, y_pos, f'{emoji}  {label}:', fontsize=12,
                    transform=ax4.transAxes, va='top', fontweight='medium')
            ax4.text(0.9, y_pos, value, fontsize=12,
                    transform=ax4.transAxes, va='top', ha='right',
                    fontweight='bold', color=COLORS['primary'])
            y_pos -= 0.11

        ax4.set_title('Indicadores Clave de Desempeño', fontweight='bold', pad=20)

        rect = mpatches.FancyBboxPatch(
            (0.02, 0.02), 0.96, 0.96,
            boxstyle="round,pad=0.02,rounding_size=0.02",
            facecolor='none',
            edgecolor=COLORS['primary'],
            linewidth=2,
            transform=ax4.transAxes
        )
        ax4.add_patch(rect)

        self.guardar_figura(fig, pdf, "Proyecciones_KPIs")

    # ============================================================
    # GENERACIÓN DE REPORTES
    # ============================================================

    def generar_reporte_pdf(self, filename=None):
        """Genera el reporte completo en PDF"""
        if filename is None:
            filename = f'Reporte_Anual_{self.year}_AD17_Solutions.pdf'

        print(f"\n🎨 Generando reporte PDF: {filename}")

        # Crear carpeta para imágenes
        if self.export_images:
            self.crear_carpeta_imagenes()

        with PdfPages(filename) as pdf:
            # 1. Portada
            self.generar_portada(pdf)

            # 2. Resumen ejecutivo
            self.generar_resumen_ejecutivo(pdf)

            # 3. Análisis de estados
            self.generar_analisis_estados(pdf)

            # 4. Análisis temporal
            self.generar_analisis_temporal(pdf)

            # 5. Análisis de proveedores
            self.generar_analisis_proveedores(pdf)

            # 6. Análisis de departamentos
            self.generar_analisis_departamentos(pdf)

            # 7. Análisis de comisiones BBVA
            self.generar_analisis_comisiones_bbva(pdf)

            # 8. Análisis de créditos
            self.generar_analisis_creditos(pdf)

            # 9. Análisis de anticipos
            self.generar_analisis_anticipos(pdf)

            # 10. Pagos recurrentes
            self.generar_analisis_recurrentes(pdf)

            # 11. Proyecciones y KPIs
            self.generar_proyecciones(pdf)

            # Metadata
            d = pdf.infodict()
            d['Title'] = f'Reporte Anual {self.year} - AD17 Solutions'
            d['Author'] = 'Sistema de Gestión de Costos'
            d['Subject'] = 'Análisis Financiero Completo'
            d['Keywords'] = 'Finanzas, Análisis, Costos, AD17'
            d['CreationDate'] = datetime.now()

        print(f"✅ Reporte PDF generado: {filename}")

        if self.export_images:
            print(f"✅ {self.contador_graficas} gráficas exportadas a: {self.images_folder}")

        return filename

    def generar_reporte_excel(self):
        """Genera reporte complementario en Excel"""
        filename = f'Reporte_Anual_{self.year}_AD17_Datos.xlsx'

        print(f"\n📊 Generando reporte Excel: {filename}")

        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Formatos
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#FF9800',
                'font_color': 'white',
                'border': 1
            })

            # Hoja 1: Resumen
            df_resumen = pd.DataFrame({
                'Métrica': [
                    'Total Solicitudes',
                    'Monto Total',
                    'Monto Promedio',
                    'Total Créditos',
                    'Total Proveedores',
                    'Pagos Recurrentes Activos',
                    'Comisiones BBVA Pagadas'
                ],
                'Valor': [
                    self.metricas['solicitudes']['total'],
                    self.metricas['solicitudes']['monto_total'],
                    self.metricas['solicitudes']['monto_promedio'],
                    self.metricas['creditos']['total'],
                    self.metricas['proveedores']['total'],
                    self.metricas['recurrentes']['activos'],
                    self.metricas['comisiones_bbva']['total']
                ]
            })
            df_resumen.to_excel(writer, sheet_name='Resumen', index=False)

            # Hoja 2: Por Estado
            df_estados = pd.DataFrame(self.metricas['solicitudes']['por_estado']).T
            df_estados.to_excel(writer, sheet_name='Por Estado')

            # Hoja 3: Top Proveedores
            top_prov = self.df_solicitudes.groupby('destinatario').agg({
                'monto': 'sum',
                'id': 'count'
            }).sort_values('monto', ascending=False).head(20)
            top_prov.columns = ['Monto Total', 'Cantidad']
            top_prov.to_excel(writer, sheet_name='Top Proveedores')

            # Hoja 4: Comisiones BBVA
            df_com = self.metricas['comisiones_bbva']['df_detalle']
            if len(df_com) > 0:
                df_com_export = df_com[['fecha', 'destinatario', 'monto',
                                       'monto_comision', 'porcentaje_comision',
                                       'departamento']].copy()
                df_com_export.to_excel(writer, sheet_name='Comisiones BBVA', index=False)

            # Hoja 5: Por Departamento
            dept_summary = self.df_solicitudes.groupby('departamento').agg({
                'monto': 'sum',
                'id': 'count'
            }).sort_values('monto', ascending=False)
            dept_summary.columns = ['Monto Total', 'Cantidad']
            dept_summary.to_excel(writer, sheet_name='Por Departamento')

            # Hoja 6: Datos completos
            self.df_solicitudes.to_excel(writer, sheet_name='Datos Completos', index=False)

        print(f"✅ Excel generado: {filename}")
        return filename

    def generar_reporte_completo(self):
        """Ejecuta el análisis completo"""
        print("\n" + "="*60)
        print("🚀 INICIANDO ANÁLISIS ANUAL - AD17 SOLUTIONS")
        print("="*60)

        # Extraer datos
        self.extraer_solicitudes()
        self.extraer_creditos()
        self.extraer_pagos_recurrentes()
        self.extraer_proveedores()

        # Calcular métricas
        self.calcular_metricas_generales()

        # Generar reportes
        pdf_file = self.generar_reporte_pdf()
        excel_file = self.generar_reporte_excel()

        print("\n" + "="*60)
        print("✅ ANÁLISIS COMPLETADO")
        print("="*60)
        print(f"📄 Reporte PDF: {pdf_file}")
        print(f"📊 Reporte Excel: {excel_file}")

        if self.export_images:
            print(f"🖼️  Gráficas exportadas: {self.images_folder}/")
            print(f"    Total de imágenes: {self.contador_graficas}")

        return pdf_file, excel_file


def main():
    """Función principal"""
    import sys

    year = int(sys.argv[1]) if len(sys.argv) > 1 else datetime.now().year

    analyzer = AD17FinancialAnalyzer(year=year)
    pdf_file, excel_file = analyzer.generar_reporte_completo()

    print(f"\n✅ Archivos generados:")
    print(f"   - {pdf_file}")
    print(f"   - {excel_file}")
    if analyzer.export_images:
        print(f"   - {analyzer.images_folder}/ ({analyzer.contador_graficas} imágenes)")


if __name__ == "__main__":
    main()