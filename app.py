import csv
import os
import smtplib
from email.message import EmailMessage
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, send_from_directory, jsonify
from datetime import datetime, timedelta, date

import io
import pandas as pd
import json
import sqlite3
import mimetypes
from werkzeug.utils import secure_filename
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import mysql.connector

import pymysql  # Nueva importación
import threading
import time
from apscheduler.schedulers.background import BackgroundScheduler  # pip install apscheduler
import atexit

# === BLUEPRINTS EXTERNOS ===
# Importa el blueprint de Comisiones (no ejecuta nada hasta que se llame una ruta)
from comisiones import comisiones_bp



# === CONFIGURACIÓN DE BASE DE DATOS REMOTA ===
def get_remote_db_connection():
    """
    Conecta a la base de datos MySQL remota
    """
    try:
        connection = pymysql.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Costos",  # Cambiado de "AD17_Costos.Pagos" a "AD17_Costos"
            charset='utf8mb4',
            cursorclass=pymysql.cursors.DictCursor
        )
        return connection
    except Exception as e:
        print(f"Error conectando a base de datos remota: {e}")
        return None

def verify_and_fix_remote_tables():
    """
    Verifica y corrige la estructura de las tablas remotas
    """
    connection = get_remote_db_connection()
    if not connection:
        print("No se pudo conectar a la base de datos remota")
        return False

    try:
        with connection.cursor() as cursor:
            # Verificar si la tabla Pagos existe y su estructura
            cursor.execute("SHOW TABLES LIKE 'Pagos'")
            table_exists = cursor.fetchone()

            if table_exists:
                # La tabla existe, verificar columnas
                cursor.execute("DESCRIBE Pagos")
                columns = [row['Field'] for row in cursor.fetchall()]
                print(f"Columnas existentes en tabla Pagos: {columns}")

                # Columnas requeridas
                required_columns = {
                    'fp': "VARCHAR(50) NOT NULL",
                    'nombre': "VARCHAR(255) NOT NULL",
                    'destinatario': "VARCHAR(255) NOT NULL DEFAULT ''",
                    'correo': "TEXT NOT NULL",
                    'departamento': "VARCHAR(255) NOT NULL",
                    'tipo_solicitud': "VARCHAR(100) NOT NULL",
                    'tipo_pago': "VARCHAR(100) NOT NULL",
                    'descripcion': "TEXT NOT NULL",
                    'datos_deposito': "TEXT NOT NULL",
                    'banco': "VARCHAR(255) NOT NULL",
                    'clabe': "VARCHAR(255) NOT NULL",
                    'monto': "DECIMAL(15,2) NOT NULL",
                    'estado': "VARCHAR(50) NOT NULL",
                    'fecha': "DATETIME NOT NULL",
                    'fecha_limite': "DATE NOT NULL",
                    'archivo_adjunto': "VARCHAR(255) NOT NULL DEFAULT ''",
                    'anticipo': "VARCHAR(10) NOT NULL DEFAULT 'No'",
                    'porcentaje_anticipo': "DECIMAL(5,2) NOT NULL DEFAULT 0.0",
                    'monto_restante': "DECIMAL(15,2) NOT NULL DEFAULT 0.0",
                    'fecha_sincronizacion': "DATETIME DEFAULT CURRENT_TIMESTAMP",
                    'es_programada': "TINYINT(1) NOT NULL DEFAULT 0",
                    'fecha_aprobado': "DATETIME DEFAULT NULL",
                    'fecha_liquidado': "DATETIME DEFAULT NULL",
                    'fecha_ultimo_cambio': "DATETIME DEFAULT NULL",
                    'historial_estados': "TEXT DEFAULT '[]'",
                    'categoria_administrativa': "VARCHAR(255) NOT NULL DEFAULT ''"
                }

                # Agregar columnas faltantes
                for column, definition in required_columns.items():
                    if column not in columns:
                        print(f"Agregando columna faltante: {column}")
                        cursor.execute(f"ALTER TABLE Pagos ADD COLUMN {column} {definition}")

                # Verificar si necesitamos agregar índices
                cursor.execute("SHOW INDEX FROM Pagos")
                existing_indexes = [row['Key_name'] for row in cursor.fetchall()]

                indexes_to_create = {
                    'idx_fp': 'fp',
                    'idx_fecha': 'fecha',
                    'idx_estado': 'estado',
                    'idx_destinatario': 'destinatario'
                }

                for index_name, column in indexes_to_create.items():
                    if index_name not in existing_indexes:
                        print(f"Creando índice: {index_name}")
                        cursor.execute(f"CREATE INDEX {index_name} ON Pagos ({column})")

                # Ampliar columna clabe si es VARCHAR(20)
                try:
                    cursor.execute("SELECT CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Pagos' AND COLUMN_NAME='clabe' AND TABLE_SCHEMA=DATABASE()")
                    clabe_info = cursor.fetchone()
                    if clabe_info and clabe_info.get('CHARACTER_MAXIMUM_LENGTH', 0) < 255:
                        print("Ampliando columna clabe de VARCHAR(20) a VARCHAR(255)...")
                        cursor.execute("ALTER TABLE Pagos MODIFY COLUMN clabe VARCHAR(255) NOT NULL")
                        print("Columna clabe ampliada exitosamente")
                except Exception as e:
                    print(f"Error verificando/ampliando columna clabe: {e}")

            else:
                # La tabla no existe, crearla completamente
                print("Creando tabla Pagos desde cero...")
                cursor.execute("""
                    CREATE TABLE Pagos (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        fp VARCHAR(50) NOT NULL,
                        nombre VARCHAR(255) NOT NULL,
                        destinatario VARCHAR(255) NOT NULL DEFAULT '',
                        correo TEXT NOT NULL,
                        departamento VARCHAR(255) NOT NULL,
                        tipo_solicitud VARCHAR(100) NOT NULL,
                        tipo_pago VARCHAR(100) NOT NULL,
                        descripcion TEXT NOT NULL,
                        datos_deposito TEXT NOT NULL,
                        banco VARCHAR(255) NOT NULL,
                        clabe VARCHAR(255) NOT NULL,
                        monto DECIMAL(15,2) NOT NULL,
                        estado VARCHAR(50) NOT NULL,
                        fecha DATETIME NOT NULL,
                        fecha_limite DATE NOT NULL,
                        archivo_adjunto VARCHAR(255) NOT NULL DEFAULT '',
                        anticipo VARCHAR(10) NOT NULL DEFAULT 'No',
                        porcentaje_anticipo DECIMAL(5,2) NOT NULL DEFAULT 0.0,
                        monto_restante DECIMAL(15,2) NOT NULL DEFAULT 0.0,
                        fecha_sincronizacion DATETIME DEFAULT CURRENT_TIMESTAMP,
                        INDEX idx_fp (fp),
                        INDEX idx_fecha (fecha),
                        INDEX idx_estado (estado),
                        INDEX idx_destinatario (destinatario)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)

            # Verificar tabla Creditos
            cursor.execute("SHOW TABLES LIKE 'Creditos'")
            creditos_exists = cursor.fetchone()

            if creditos_exists:
                # Verificar si tiene la columna nombre_proveedor
                cursor.execute("DESCRIBE Creditos")
                creditos_columns = [row['Field'] for row in cursor.fetchall()]

                if 'nombre_proveedor' not in creditos_columns:
                    print("Agregando columna nombre_proveedor a tabla Creditos")
                    cursor.execute("ALTER TABLE Creditos ADD COLUMN nombre_proveedor TEXT DEFAULT ''")
            else:
                # Crear tabla Creditos
                print("Creando tabla Creditos desde cero...")
                cursor.execute("""
                    CREATE TABLE Creditos (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        nombre VARCHAR(255) NOT NULL,
                        entidad VARCHAR(255) NOT NULL,
                        descripcion TEXT,
                        monto_total DECIMAL(15,2) NOT NULL,
                        tasa_interes DECIMAL(5,2) NOT NULL,
                        fecha_inicio DATE NOT NULL,
                        fecha_final DATE NOT NULL,
                        plazo_meses INT NOT NULL,
                        estado VARCHAR(50) NOT NULL DEFAULT 'Activo',
                        fecha_registro DATETIME NOT NULL,
                        numero_cuenta VARCHAR(50),
                        tipo_credito VARCHAR(100) NOT NULL,
                        pago_mensual DECIMAL(15,2) NOT NULL,
                        contacto VARCHAR(255),
                        notas TEXT,
                        nombre_proveedor TEXT DEFAULT '',
                        fecha_sincronizacion DATETIME DEFAULT CURRENT_TIMESTAMP,
                        INDEX idx_nombre (nombre),
                        INDEX idx_entidad (entidad),
                        INDEX idx_estado (estado)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)

            # Verificar tabla PagosCredito
            cursor.execute("SHOW TABLES LIKE 'PagosCredito'")
            pagos_credito_exists = cursor.fetchone()

            if not pagos_credito_exists:
                print("Creando tabla PagosCredito desde cero...")
                cursor.execute("""
                    CREATE TABLE PagosCredito (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        credito_id INT NOT NULL,
                        monto DECIMAL(15,2) NOT NULL,
                        fecha DATE NOT NULL,
                        referencia VARCHAR(255),
                        descripcion TEXT,
                        comprobante VARCHAR(255),
                        tipo_pago VARCHAR(100) NOT NULL,
                        fecha_sincronizacion DATETIME DEFAULT CURRENT_TIMESTAMP,
                        INDEX idx_credito_id (credito_id),
                        INDEX idx_fecha (fecha)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)

        connection.commit()
        print("Verificación y corrección de tablas completada exitosamente")
        return True
    except Exception as e:
        print(f"Error verificando/corrigiendo tablas remotas: {e}")
        connection.rollback()
        return False
    finally:
        connection.close()

def init_remote_tables():
    """
    Inicializa las tablas en la base de datos remota si no existen
    """
    return verify_and_fix_remote_tables()

def sync_solicitudes_to_remote():
    """
    Sincroniza todas las solicitudes de SQLite a MySQL
    """
    print("Iniciando sincronización de solicitudes...")

    # Conectar a bases de datos
    local_conn = get_db_connection()
    remote_conn = get_remote_db_connection()

    if not remote_conn:
        print("Error: No se pudo conectar a la base de datos remota")
        local_conn.close()
        return False

    try:
        # Verificar que la tabla tenga la estructura correcta
        with remote_conn.cursor() as cursor:
            cursor.execute("DESCRIBE Pagos")
            columns_info = cursor.fetchall()
            column_names = [col['Field'] for col in columns_info]
            print(f"Columnas disponibles en Pagos: {column_names}")

            # Verificar que existan las columnas críticas
            required_cols = ['fp', 'nombre', 'destinatario', 'correo', 'monto', 'estado']
            missing_cols = [col for col in required_cols if col not in column_names]

            if missing_cols:
                print(f"Error: Faltan columnas críticas: {missing_cols}")
                return False

        # Obtener todos los registros de SQLite
        local_solicitudes = local_conn.execute("SELECT * FROM solicitudes").fetchall()
        print(f"Solicitudes a sincronizar: {len(local_solicitudes)}")

        with remote_conn.cursor() as cursor:
            # Obtener FPs existentes en la base remota
            cursor.execute("SELECT fp FROM Pagos")
            fps_existentes = {row['fp'] for row in cursor.fetchall()}
            print(f"FPs existentes en remoto: {len(fps_existentes)}")

            # Insertar o actualizar registros
            nuevos = 0
            actualizados = 0
            errores = 0

            for solicitud in local_solicitudes:
                try:
                    data = dict(solicitud)

                    if data['fp'] in fps_existentes:
                        # Actualizar registro existente
                        cursor.execute("""
                            UPDATE Pagos SET
                                nombre = %s, destinatario = %s, correo = %s, departamento = %s,
                                tipo_solicitud = %s, tipo_pago = %s, descripcion = %s,
                                datos_deposito = %s, banco = %s, clabe = %s, monto = %s,
                                estado = %s, fecha = %s, fecha_limite = %s,
                                archivo_adjunto = %s, anticipo = %s, porcentaje_anticipo = %s,
                                monto_restante = %s, categoria_administrativa = %s, fecha_sincronizacion = NOW()
                            WHERE fp = %s
                        """, (
                            data['nombre'], data.get('destinatario', ''), data['correo'], data['departamento'],
                            data['tipo_solicitud'], data['tipo_pago'], data['descripcion'],
                            data['datos_deposito'], data['banco'], data['clabe'], data['monto'],
                            data['estado'], data['fecha'], data['fecha_limite'],
                            data.get('archivo_adjunto', ''), data.get('anticipo', 'No'),
                            data.get('porcentaje_anticipo', 0.0), data.get('monto_restante', 0.0),
                            data.get('categoria_administrativa', ''), data['fp']
                        ))
                        actualizados += 1
                    else:
                        # Insertar nuevo registro
                        cursor.execute("""
                            INSERT INTO Pagos (
                                fp, nombre, destinatario, correo, departamento, tipo_solicitud,
                                tipo_pago, descripcion, datos_deposito, banco, clabe, monto,
                                estado, fecha, fecha_limite, archivo_adjunto, anticipo,
                                porcentaje_anticipo, monto_restante, categoria_administrativa, fecha_sincronizacion
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NOW())
                        """, (
                            data['fp'], data['nombre'], data.get('destinatario', ''), data['correo'],
                            data['departamento'], data['tipo_solicitud'], data['tipo_pago'],
                            data['descripcion'], data['datos_deposito'], data['banco'],
                            data['clabe'], data['monto'], data['estado'], data['fecha'],
                            data['fecha_limite'], data.get('archivo_adjunto', ''), data.get('anticipo', 'No'),
                            data.get('porcentaje_anticipo', 0.0), data.get('monto_restante', 0.0),
                            data.get('categoria_administrativa', '')
                        ))
                        nuevos += 1

                except Exception as e:
                    print(f"Error procesando solicitud FP {data.get('fp', 'unknown')}: {e}")
                    errores += 1
                    continue

        remote_conn.commit()
        print(f"Sincronización completada: {nuevos} nuevos, {actualizados} actualizados, {errores} errores")
        return errores == 0  # Retorna True solo si no hubo errores

    except Exception as e:
        print(f"Error durante la sincronización de solicitudes: {e}")
        remote_conn.rollback()
        return False
    finally:
        local_conn.close()
        remote_conn.close()
def sync_creditos_to_remote():
    """
    Sincroniza créditos, pagos de créditos e historial de montos a la base remota
    """
    print("Iniciando sincronización de créditos...")

    local_conn = get_db_connection()
    remote_conn = get_remote_db_connection()

    if not remote_conn:
        print("Error: No se pudo conectar a la base de datos remota")
        local_conn.close()
        return False

    try:
        with remote_conn.cursor() as cursor:
            # ============ SINCRONIZAR CRÉDITOS ============
            local_creditos = local_conn.execute("SELECT * FROM creditos").fetchall()

            # Obtener IDs existentes en la base remota
            cursor.execute("SELECT id FROM Creditos")
            ids_existentes = {row['id'] for row in cursor.fetchall()}

            creditos_nuevos = 0
            creditos_actualizados = 0

            for credito in local_creditos:
                data = dict(credito)

                if data['id'] in ids_existentes:
                    # Actualizar registro existente
                    cursor.execute("""
                        UPDATE Creditos SET
                            nombre = %s, entidad = %s, descripcion = %s, monto_total = %s,
                            tasa_interes = %s, fecha_inicio = %s, fecha_final = %s,
                            plazo_meses = %s, estado = %s, fecha_registro = %s,
                            numero_cuenta = %s, tipo_credito = %s, pago_mensual = %s,
                            contacto = %s, notas = %s, nombre_proveedor = %s, fecha_sincronizacion = NOW()
                        WHERE id = %s
                    """, (
                        data['nombre'], data['entidad'], data['descripcion'], data['monto_total'],
                        data['tasa_interes'], data['fecha_inicio'], data['fecha_final'],
                        data['plazo_meses'], data['estado'], data['fecha_registro'],
                        data['numero_cuenta'], data['tipo_credito'], data['pago_mensual'],
                        data['contacto'], data['notas'], data.get('nombre_proveedor', ''), data['id']
                    ))
                    creditos_actualizados += 1
                else:
                    # Insertar nuevo registro
                    cursor.execute("""
                        INSERT INTO Creditos (
                            id, nombre, entidad, descripcion, monto_total, tasa_interes,
                            fecha_inicio, fecha_final, plazo_meses, estado, fecha_registro,
                            numero_cuenta, tipo_credito, pago_mensual, contacto, notas,
                            nombre_proveedor, fecha_sincronizacion
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NOW())
                    """, (
                        data['id'], data['nombre'], data['entidad'], data['descripcion'],
                        data['monto_total'], data['tasa_interes'], data['fecha_inicio'],
                        data['fecha_final'], data['plazo_meses'], data['estado'],
                        data['fecha_registro'], data['numero_cuenta'], data['tipo_credito'],
                        data['pago_mensual'], data['contacto'], data['notas'],
                        data.get('nombre_proveedor', '')
                    ))
                    creditos_nuevos += 1

            print(f"Créditos: {creditos_nuevos} nuevos, {creditos_actualizados} actualizados")

            # ============ SINCRONIZAR PAGOS DE CRÉDITOS ============
            local_pagos = local_conn.execute("SELECT * FROM pagos_credito").fetchall()

            # Obtener IDs existentes de pagos
            cursor.execute("SELECT id FROM PagosCredito")
            ids_pagos_existentes = {row['id'] for row in cursor.fetchall()}

            pagos_nuevos = 0
            pagos_actualizados = 0

            for pago in local_pagos:
                data = dict(pago)

                if data['id'] in ids_pagos_existentes:
                    # Actualizar pago existente
                    cursor.execute("""
                        UPDATE PagosCredito SET
                            credito_id = %s, monto = %s, fecha = %s, referencia = %s,
                            descripcion = %s, comprobante = %s, tipo_pago = %s,
                            fecha_sincronizacion = NOW()
                        WHERE id = %s
                    """, (
                        data['credito_id'], data['monto'], data['fecha'], data['referencia'],
                        data['descripcion'], data['comprobante'], data['tipo_pago'], data['id']
                    ))
                    pagos_actualizados += 1
                else:
                    # Insertar nuevo pago
                    cursor.execute("""
                        INSERT INTO PagosCredito (
                            id, credito_id, monto, fecha, referencia, descripcion,
                            comprobante, tipo_pago, fecha_sincronizacion
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
                    """, (
                        data['id'], data['credito_id'], data['monto'], data['fecha'],
                        data['referencia'], data['descripcion'], data['comprobante'],
                        data['tipo_pago']
                    ))
                    pagos_nuevos += 1

            print(f"Pagos: {pagos_nuevos} nuevos, {pagos_actualizados} actualizados")

            # ============ SINCRONIZAR HISTORIAL DE MONTO ============
            local_historial = local_conn.execute("SELECT * FROM historial_monto_credito").fetchall()

            # Obtener IDs existentes del historial
            cursor.execute("SELECT id FROM HistorialMontoCredito")
            ids_historial_existentes = {row['id'] for row in cursor.fetchall()}

            historial_nuevos = 0
            historial_actualizados = 0

            for historial in local_historial:
                data = dict(historial)

                if data['id'] in ids_historial_existentes:
                    # Actualizar registro existente
                    cursor.execute("""
                        UPDATE HistorialMontoCredito SET
                            credito_id = %s, monto_anterior = %s, monto_nuevo = %s,
                            fecha_cambio = %s, motivo = %s, usuario = %s,
                            fecha_sincronizacion = NOW()
                        WHERE id = %s
                    """, (
                        data['credito_id'], data['monto_anterior'], data['monto_nuevo'],
                        data['fecha_cambio'], data['motivo'], data['usuario'], data['id']
                    ))
                    historial_actualizados += 1
                else:
                    # Insertar nuevo registro
                    cursor.execute("""
                        INSERT INTO HistorialMontoCredito (
                            id, credito_id, monto_anterior, monto_nuevo,
                            fecha_cambio, motivo, usuario, fecha_sincronizacion
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
                    """, (
                        data['id'], data['credito_id'], data['monto_anterior'],
                        data['monto_nuevo'], data['fecha_cambio'], data['motivo'],
                        data['usuario']
                    ))
                    historial_nuevos += 1

            print(f"Historial de montos: {historial_nuevos} nuevos, {historial_actualizados} actualizados")

        remote_conn.commit()
        print("✅ Sincronización de créditos completada exitosamente")
        return True

    except Exception as e:
        print(f"❌ Error durante la sincronización de créditos: {e}")
        import traceback
        traceback.print_exc()
        remote_conn.rollback()
        return False
    finally:
        local_conn.close()
        remote_conn.close()

def sync_all_data():
    """
    Sincroniza todos los datos (solicitudes y créditos)
    """
    print(f"=== SINCRONIZACIÓN INICIADA: {datetime.now()} ===")

    # Inicializar tablas si es necesario
    if not init_remote_tables():
        print("Error: No se pudieron inicializar las tablas remotas")
        return

    # Sincronizar solicitudes
    if sync_solicitudes_to_remote():
        print("✓ Solicitudes sincronizadas")
    else:
        print("✗ Error sincronizando solicitudes")

    # Sincronizar créditos
    if sync_creditos_to_remote():
        print("✓ Créditos sincronizados")
    else:
        print("✗ Error sincronizando créditos")

    print(f"=== SINCRONIZACIÓN COMPLETADA: {datetime.now()} ===")


# ===== OPTIMIZACIONES DE RENDIMIENTO: índices, PRAGMAs y consultas paginadas =====
import math

PAGE_SIZE_DEFAULT = 15

def ensure_sqlite_indexes():
    """
    Crea índices en columnas usadas en filtros/orden y acelera ORDER BY fecha.
    Seguro de ejecutar múltiples veces (IF NOT EXISTS).
    """
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("PRAGMA foreign_keys = ON")
        # Índices clave para búsquedas del panel
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_fecha ON solicitudes(fecha)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_estado ON solicitudes(estado)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_fp ON solicitudes(fp)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_monto ON solicitudes(monto)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_destinatario ON solicitudes(destinatario)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_departamento ON solicitudes(departamento)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sol_tipo_solicitud ON solicitudes(tipo_solicitud)")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[perf] Error creando índices: {e}")

def apply_sqlite_pragmas_once():
    """
    Ajusta PRAGMAs para mejor rendimiento de lectura. Ejecuta 1 sola vez por proceso.
    """
    try:
        if getattr(app, "_sqlite_pragmas_applied", False):
            return
    except Exception:
        # Si aún no existe app (por orden de carga), no hacemos nada; se aplicará luego.
        return

    try:
        conn = get_db_connection()
        cur = conn.cursor()
        # Modo WAL = mejores lecturas concurrentes
        cur.execute("PRAGMA journal_mode=WAL")
        # Sincronización NORMAL = buen balance seguridad/rendimiento
        cur.execute("PRAGMA synchronous=NORMAL")
        # Más caché en memoria (valor negativo = KB; -20000 ≈ ~20MB)
        cur.execute("PRAGMA cache_size=-20000")
        # Tablas temporales en memoria
        cur.execute("PRAGMA temp_store=MEMORY")
        conn.commit()
        conn.close()
        app._sqlite_pragmas_applied = True
        print("[perf] PRAGMAs aplicados")
    except Exception as e:
        print(f"[perf] Error aplicando PRAGMAs: {e}")

def _build_filters(estado_filtro: str, busqueda: str, criterio: str):
    """
    Devuelve (where_sql, params) para filtros de panel.
    """
    where = []
    params = []

    if estado_filtro:
        where.append("LOWER(estado) = LOWER(?)")
        params.append(estado_filtro)

    if busqueda:
        q = f"%{busqueda}%"
        if criterio == "fp":
            where.append("fp LIKE ?")
            params.append(q)
        elif criterio == "monto":
            # Comparar como texto para subcadena
            where.append("CAST(monto AS TEXT) LIKE ?")
            params.append(q)
        elif criterio == "nombre":
            where.append("LOWER(nombre) LIKE LOWER(?)")
            params.append(q)
        elif criterio == "destinatario":
            where.append("LOWER(destinatario) LIKE LOWER(?)")
            params.append(q)
        else:
            # Búsqueda amplia
            where.append("""(
                LOWER(destinatario) LIKE LOWER(?)
                OR LOWER(correo) LIKE LOWER(?)
                OR LOWER(departamento) LIKE LOWER(?)
                OR LOWER(tipo_solicitud) LIKE LOWER(?)
                OR LOWER(descripcion) LIKE LOWER(?)
                OR LOWER(fp) LIKE LOWER(?)
            )""")
            params.extend([q, q, q, q, q, q])

    where_sql = ("WHERE " + " AND ".join(where)) if where else ""
    return where_sql, params

def query_solicitudes_paginated(page=1, page_size=PAGE_SIZE_DEFAULT, estado_filtro="", busqueda="", criterio="todos"):
    """
    Consulta paginada de solicitudes con filtros, búsqueda y contadores por estado.
    Incluye soporte para los 3 nuevos tipos de archivos.
    """
    conn = get_db_connection()

    # --- Condiciones WHERE ---
    where_clauses = []
    params = []

    if estado_filtro:
        where_clauses.append("estado = ?")
        params.append(estado_filtro)

    if busqueda:
        criterio = criterio.lower()
        if criterio == "fp":
            where_clauses.append("fp LIKE ?")
            params.append(f"%{busqueda}%")
        elif criterio == "monto":
            where_clauses.append("CAST(monto AS TEXT) LIKE ?")
            params.append(f"%{busqueda}%")
        elif criterio == "nombre":
            where_clauses.append("nombre LIKE ?")
            params.append(f"%{busqueda}%")
        elif criterio == "destinatario":
            where_clauses.append("destinatario LIKE ?")
            params.append(f"%{busqueda}%")
        else:  # "todos"
            where_clauses.append("""
                (fp LIKE ? OR
                 nombre LIKE ? OR
                 destinatario LIKE ? OR
                 tipo_solicitud LIKE ? OR
                 CAST(monto AS TEXT) LIKE ?)
            """)
            params.extend([f"%{busqueda}%"] * 5)

    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)

    # --- Contadores por estado ---
    counts_query = f"""
        SELECT
            LOWER(estado) as estado,
            COUNT(*) as total
        FROM solicitudes
        {where_sql}
        GROUP BY LOWER(estado)
    """
    rows_counts = conn.execute(counts_query, params).fetchall()
    counts = {row["estado"]: row["total"] for row in rows_counts}

    # Total de registros
    total_query = f"SELECT COUNT(*) as total FROM solicitudes {where_sql}"
    total_count = conn.execute(total_query, params).fetchone()["total"]

    # --- Paginación ---
    total_pages = max(1, (total_count + page_size - 1) // page_size)
    page = max(1, min(page, total_pages))
    offset = (page - 1) * page_size

    # ===== ACTUALIZADO: Query principal incluye los nuevos campos de archivos =====
    items_query = f"""
        SELECT
            id, fp, nombre, destinatario, correo, departamento,
            tipo_solicitud, tipo_pago, monto, estado, fecha, fecha_limite,
            archivo_adjunto, archivo_factura, archivo_recibo, archivo_orden_compra,
            banco, clabe, referencia, descripcion, datos_deposito,
            anticipo, porcentaje_anticipo, monto_anticipo, monto_restante, tipo_anticipo,
            tiene_comision, porcentaje_comision, monto_sin_comision, monto_comision,
            historial_estados, fecha_aprobado, fecha_liquidado, fecha_ultimo_cambio
        FROM solicitudes
        {where_sql}
        ORDER BY fecha DESC
        LIMIT ? OFFSET ?
    """
    # ===== FIN ACTUALIZADO =====

    items = conn.execute(items_query, params + [page_size, offset]).fetchall()
    conn.close()

    return {
        "items": items,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "total_count": total_count,
        "counts": counts
    }


# === PROGRAMADOR DE TAREAS ===
def start_scheduler():
    """
    Inicia el programador de tareas para sincronización automática
    """
    if getattr(app, "_scheduler_started", False):
        print("Scheduler ya estaba iniciado")
        return

    scheduler = BackgroundScheduler()

    # Programar sincronización diaria a las 02:00 AM
    scheduler.add_job(
        func=sync_all_data,
        trigger="cron",
        hour=2,
        minute=0,
        id='sync_daily',
        replace_existing=True

    )

    # Programar sincronización cada 6 horas como respaldo
    scheduler.add_job(
        func=sync_all_data,
        trigger="cron",
        hour="*/6",
        id='sync_backup',
        replace_existing=True

    )
        # Recordatorios de pagos recurrentes: 8:00, 13:00 y 18:00
    scheduler.add_job(
        func=check_recurring_payment_reminders,
        trigger="cron",
        hour="13,18",
        minute=0,
        id="recurring_payment_reminders",
        replace_existing=True

    )


    scheduler.start()
    app._scheduler_started = True

    print("Programador de sincronización iniciado")

    # Ejecutar sincronización inicial al iniciar la aplicación
    threading.Thread(target=sync_all_data, daemon=True).start()

    # Asegurar que el scheduler se cierre apropiadamente
    atexit.register(lambda: scheduler.shutdown())

# === APP & CONFIG ===
app = Flask(__name__)
app.secret_key = "clave-secreta-para-sesiones"  # En producción, usa variables de entorno
app.permanent_session_lifetime = timedelta(days=30)

# Config por defecto para correos y carpeta de comprobantes (no rompe si ya lo configuras en otro lado)
app.config.setdefault("EMAIL_USER", "ad17solutionsbot@gmail.com")
app.config.setdefault("EMAIL_PASS", "misvtfhrnwbmiptb")
app.config.setdefault("COMM_UPLOAD_FOLDER", os.path.join(app.root_path, "uploads", "comisiones"))
os.makedirs(app.config["COMM_UPLOAD_FOLDER"], exist_ok=True)

# Registro del Blueprint de Comisiones (sin url_prefix para mantener /comisiones)
app.register_blueprint(comisiones_bp, url_prefix="/comisiones")

# --- Inicialización segura (se ejecuta 1 sola vez por worker) ---
# --- Inicialización perezosa (una sola vez por proceso) ---
INIT_DONE = False
INIT_LOCK = threading.Lock()

@app.before_request
def _init_once():
    global INIT_DONE
    if INIT_DONE:
        return
    with INIT_LOCK:
        if INIT_DONE:
            return
        # Inyección perezosa de dependencias para el blueprint de Comisiones
        # (estas funciones existen más abajo o en tu app completa)
        try:
            if "GET_DB_CONNECTION" not in app.config and "get_db_connection" in globals():
                app.config["GET_DB_CONNECTION"] = get_db_connection
            if "READ_EMPLOYEES" not in app.config and "read_employees" in globals():
                app.config["READ_EMPLOYEES"] = read_employees
        except Exception as e:
            print(f"[init_once] Aviso inyección dependencias comisiones: {e}")

        try:
            ensure_recurring_tables()
        except Exception as e:
            print(f"[init_once] Error creando/verificando pagos_recurrentes: {e}")
        try:
            start_scheduler()
        except Exception as e:
            print(f"[init_once] Error iniciando scheduler: {e}")
        INIT_DONE = True


# ---- Hook de arranque de rendimiento (1 vez por worker) ----
@app.before_request
def _perf_bootstrap_once():
    # Evita trabajo repetido
    if getattr(app, "_perf_bootstrapped", False):
        return
    try:
        ensure_sqlite_indexes()
        apply_sqlite_pragmas_once()
        app._perf_bootstrapped = True
        print("[perf] Bootstrap de rendimiento listo")
    except Exception as e:
        print(f"[perf] Error en bootstrap: {e}")

DATABASE = 'database.db'
import json
@app.template_filter('fromjson')
def fromjson_filter(s):
    try:
        return json.loads(s)
    except Exception:
        return {}
def get_db_connection():
    conn = sqlite3.connect(DATABASE, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn




def init_historial_monto_table():
    """
    Inicializa la tabla historial_monto_credito si no existe
    """
    try:
        conn = get_db_connection()
        conn.execute("""
            CREATE TABLE IF NOT EXISTS historial_monto_credito (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                credito_id INTEGER NOT NULL,
                monto_anterior REAL NOT NULL,
                monto_nuevo REAL NOT NULL,
                fecha_cambio DATETIME NOT NULL,
                motivo TEXT,
                usuario TEXT,
                FOREIGN KEY (credito_id) REFERENCES creditos(id) ON DELETE CASCADE
            )
        """)
        conn.commit()
        conn.close()
        print("✅ Tabla historial_monto_credito verificada/creada")
    except Exception as e:
        print(f"Error inicializando tabla historial_monto_credito: {e}")

# Llamar a esta función cuando inicia la app
init_historial_monto_table()

def migrate_db():
    """
    Agrega columnas nuevas si no existen
    """
    conn = get_db_connection()
    cursor = conn.execute("PRAGMA table_info(solicitudes)")
    columns = [row["name"] for row in cursor.fetchall()]
    cursor.close()

    # Columnas existentes
    if "destinatario" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN destinatario TEXT NOT NULL DEFAULT ''")
        print("Columna 'destinatario' agregada a la tabla 'solicitudes'.")
    if "archivo_adjunto" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN archivo_adjunto TEXT NOT NULL DEFAULT ''")
        print("Columna 'archivo_adjunto' agregada a la tabla 'solicitudes'.")

    # NUEVAS COLUMNAS PARA 3 ARCHIVOS
    if "archivo_factura" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN archivo_factura TEXT NOT NULL DEFAULT ''")
        print("Columna 'archivo_factura' agregada a la tabla 'solicitudes'.")
    if "archivo_recibo" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN archivo_recibo TEXT NOT NULL DEFAULT ''")
        print("Columna 'archivo_recibo' agregada a la tabla 'solicitudes'.")
    if "archivo_orden_compra" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN archivo_orden_compra TEXT NOT NULL DEFAULT ''")
        print("Columna 'archivo_orden_compra' agregada a la tabla 'solicitudes'.")

    if "anticipo" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN anticipo TEXT NOT NULL DEFAULT 'No'")
        print("Columna 'anticipo' agregada a la tabla 'solicitudes'.")
    if "porcentaje_anticipo" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN porcentaje_anticipo REAL NOT NULL DEFAULT 0.0")
        print("Columna 'porcentaje_anticipo' agregada a la tabla 'solicitudes'.")
    if "monto_restante" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN monto_restante REAL NOT NULL DEFAULT 0.0")
        print("Columna 'monto_restante' agregada a la tabla 'solicitudes'.")
    if "es_programada" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN es_programada INTEGER NOT NULL DEFAULT 0")
        print("Columna 'es_programada' agregada a la tabla 'solicitudes'.")
    if "fecha_aprobado" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN fecha_aprobado TEXT")
        print("Columna 'fecha_aprobado' agregada a la tabla 'solicitudes'.")
    if "fecha_liquidado" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN fecha_liquidado TEXT")
        print("Columna 'fecha_liquidado' agregada a la tabla 'solicitudes'.")
    if "fecha_ultimo_cambio" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN fecha_ultimo_cambio TEXT")
        print("Columna 'fecha_ultimo_cambio' agregada a la tabla 'solicitudes'.")
    if "historial_estados" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN historial_estados TEXT DEFAULT '[]'")
        print("Columna 'historial_estados' agregada a la tabla 'solicitudes'.")
    if "tiene_comision" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN tiene_comision INTEGER NOT NULL DEFAULT 0")
        print("Columna 'tiene_comision' agregada a la tabla 'solicitudes'.")
    if "porcentaje_comision" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN porcentaje_comision REAL NOT NULL DEFAULT 0.0")
        print("Columna 'porcentaje_comision' agregada a la tabla 'solicitudes'.")
    if "monto_comision" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN monto_comision REAL NOT NULL DEFAULT 0.0")
        print("Columna 'monto_comision' agregada a la tabla 'solicitudes'.")
    if "monto_sin_comision" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN monto_sin_comision REAL NOT NULL DEFAULT 0.0")
        print("Columna 'monto_sin_comision' agregada a la tabla 'solicitudes'.")
    if "tipo_anticipo" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN tipo_anticipo TEXT NOT NULL DEFAULT 'porcentaje'")
        print("Columna 'tipo_anticipo' agregada a la tabla 'solicitudes'.")
    if "monto_anticipo" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN monto_anticipo REAL NOT NULL DEFAULT 0.0")
        print("Columna 'monto_anticipo' agregada a la tabla 'solicitudes'.")
    if "referencia" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN referencia TEXT NOT NULL DEFAULT ''")
        print("Columna 'referencia' agregada a la tabla 'solicitudes'.")

    # NUEVA COLUMNA PARA CATEGORÍA ADMINISTRATIVA
    if "categoria_administrativa" not in columns:
        conn.execute("ALTER TABLE solicitudes ADD COLUMN categoria_administrativa TEXT NOT NULL DEFAULT ''")
        print("Columna 'categoria_administrativa' agregada a la tabla 'solicitudes'.")

        # Migrar datos existentes: extraer categoría del campo descripcion
        # Buscar registros con tipo_solicitud="Administrativos" que tengan "Categoría: " en descripcion
        solicitudes = conn.execute("""
            SELECT id, descripcion FROM solicitudes
            WHERE tipo_solicitud = 'Administrativos' AND descripcion LIKE 'Categoría: %'
        """).fetchall()

        for sol in solicitudes:
            descripcion = sol['descripcion']
            # Extraer la categoría (primera línea después de "Categoría: ")
            if descripcion.startswith("Categoría: "):
                lines = descripcion.split('\n', 1)
                categoria = lines[0].replace("Categoría: ", "").strip()
                nueva_descripcion = lines[1].strip() if len(lines) > 1 else ""

                # Actualizar el registro
                conn.execute("""
                    UPDATE solicitudes
                    SET categoria_administrativa = ?, descripcion = ?
                    WHERE id = ?
                """, (categoria, nueva_descripcion, sol['id']))

        print(f"Migrados {len(solicitudes)} registros con categoría administrativa.")

    conn.commit()
    conn.close()

def init_db():
    conn = get_db_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS solicitudes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fp TEXT NOT NULL,
            nombre TEXT NOT NULL,
            destinatario TEXT NOT NULL,
            correo TEXT NOT NULL,
            departamento TEXT NOT NULL,
            tipo_solicitud TEXT NOT NULL,
            tipo_pago TEXT NOT NULL,
            descripcion TEXT NOT NULL,
            datos_deposito TEXT NOT NULL,
            banco TEXT NOT NULL,
            clabe TEXT NOT NULL,
            monto REAL NOT NULL,
            estado TEXT NOT NULL,
            fecha TEXT NOT NULL,
            fecha_limite TEXT NOT NULL,
            archivo_adjunto TEXT NOT NULL DEFAULT '',
            archivo_factura TEXT NOT NULL DEFAULT '',
            archivo_recibo TEXT NOT NULL DEFAULT '',
            archivo_orden_compra TEXT NOT NULL DEFAULT '',
            anticipo TEXT NOT NULL DEFAULT 'No',
            porcentaje_anticipo REAL NOT NULL DEFAULT 0.0,
            monto_restante REAL NOT NULL DEFAULT 0.0,
            tiene_comision INTEGER NOT NULL DEFAULT 0,
            porcentaje_comision REAL NOT NULL DEFAULT 0.0,
            monto_comision REAL NOT NULL DEFAULT 0.0,
            monto_sin_comision REAL NOT NULL DEFAULT 0.0,
            tipo_anticipo TEXT NOT NULL DEFAULT 'porcentaje',
            monto_anticipo REAL NOT NULL DEFAULT 0.0,
            categoria_administrativa TEXT NOT NULL DEFAULT ''
        )
    """)

    # Crear tabla de créditos
    conn.execute("""
        CREATE TABLE IF NOT EXISTS creditos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            entidad TEXT NOT NULL,
            descripcion TEXT,
            monto_total REAL NOT NULL,
            tasa_interes REAL NOT NULL,
            fecha_inicio TEXT NOT NULL,
            fecha_final TEXT NOT NULL,
            plazo_meses INTEGER NOT NULL,
            estado TEXT NOT NULL DEFAULT 'Activo',
            fecha_registro TEXT NOT NULL,
            numero_cuenta TEXT,
            tipo_credito TEXT NOT NULL,
            pago_mensual REAL NOT NULL,
            contacto TEXT,
            notas TEXT
        )
    """)

    # Crear tabla de pagos de créditos
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pagos_credito (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            credito_id INTEGER NOT NULL,
            monto REAL NOT NULL,
            fecha TEXT NOT NULL,
            referencia TEXT,
            descripcion TEXT,
            comprobante TEXT,
            tipo_pago TEXT NOT NULL,
            FOREIGN KEY (credito_id) REFERENCES creditos(id)
        )
    """)

    conn.commit()
    conn.close()

# Inicializa la base de datos y ejecuta migraciones
init_db()
migrate_db()


def ensure_recurring_tables():
    """
    Crea/actualiza la tabla pagos_recurrentes para que permita periodicidad 'bimestral'.
    Si la tabla existe pero su DDL no contiene 'bimestral', se reconstruye de forma segura.
    """
    conn = get_db_connection()
    cur = conn.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='pagos_recurrentes'")
    row = cur.fetchone()

    ddl_nueva = """
        CREATE TABLE IF NOT EXISTS pagos_recurrentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            proveedor TEXT,
            descripcion TEXT,
            monto REAL NOT NULL,
            metodo_pago TEXT,
            banco TEXT,
            clabe TEXT,
            periodicidad TEXT NOT NULL CHECK (periodicidad IN ('mensual','semanal','anual','quincenal','bimestral')),
            fecha_proximo_pago TEXT NOT NULL,
            dias_recordatorio INTEGER NOT NULL DEFAULT 2,
            correos TEXT NOT NULL DEFAULT '',
            activo INTEGER NOT NULL DEFAULT 1,
            fecha_creacion TEXT NOT NULL,
            fecha_ultimo_recordatorio TEXT,
            ultimo_recordatorio_para_fecha TEXT
        )
    """

    if not row:
        conn.executescript(ddl_nueva)
        conn.commit()
        conn.close()
        return

    ddl_actual = row[0] if isinstance(row, tuple) else row["sql"]
    if "bimestral" not in (ddl_actual or ""):
        # reconstrucción
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pagos_recurrentes_tmp (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                proveedor TEXT,
                descripcion TEXT,
                monto REAL NOT NULL,
                metodo_pago TEXT,
                banco TEXT,
                clabe TEXT,
                periodicidad TEXT NOT NULL CHECK (periodicidad IN ('mensual','semanal','anual','quincenal','bimestral')),
                fecha_proximo_pago TEXT NOT NULL,
                dias_recordatorio INTEGER NOT NULL DEFAULT 2,
                correos TEXT NOT NULL DEFAULT '',
                activo INTEGER NOT NULL DEFAULT 1,
                fecha_creacion TEXT NOT NULL,
                fecha_ultimo_recordatorio TEXT,
                ultimo_recordatorio_para_fecha TEXT
            )
        """)
        # Copia de datos 1:1 (todas las columnas existentes)
        conn.execute("""
            INSERT INTO pagos_recurrentes_tmp
            (id,nombre,proveedor,descripcion,monto,metodo_pago,banco,clabe,periodicidad,
             fecha_proximo_pago,dias_recordatorio,correos,activo,fecha_creacion,
             fecha_ultimo_recordatorio,ultimo_recordatorio_para_fecha)
            SELECT
             id,nombre,proveedor,descripcion,monto,metodo_pago,banco,clabe,periodicidad,
             fecha_proximo_pago,dias_recordatorio,correos,activo,fecha_creacion,
             fecha_ultimo_recordatorio,ultimo_recordatorio_para_fecha
            FROM pagos_recurrentes
        """)
        conn.execute("DROP TABLE pagos_recurrentes")
        conn.execute("ALTER TABLE pagos_recurrentes_tmp RENAME TO pagos_recurrentes")
        conn.commit()

    conn.close()
def _add_months(d: date, months: int) -> date:
    """Suma meses cuidando fin de mes."""
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    # días por mes (manejo simple de bisiesto)
    dim = [31, 29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    day = min(d.day, dim)
    return date(y, m, day)

def compute_next_due_date(fecha_proximo_pago_str: str, periodicidad: str) -> str:
    """Calcula la siguiente fecha de pago según periodicidad."""
    d = datetime.strptime(fecha_proximo_pago_str, "%Y-%m-%d").date()
    if periodicidad == "mensual":
        nd = _add_months(d, 1)
    elif periodicidad == "bimestral":
        nd = _add_months(d, 2)
    elif periodicidad == "semanal":
        nd = d + timedelta(days=7)
    elif periodicidad == "quincenal":
        nd = d + timedelta(days=14)
    else:  # anual
        try:
            nd = date(d.year + 1, d.month, d.day)
        except ValueError:  # 29 feb
            nd = date(d.year + 1, d.month, 28)
    return nd.strftime("%Y-%m-%d")

def send_recurring_payment_reminder(row: dict) -> bool:
    """Email de recordatorio de pago recurrente"""

    recipients = get_recipients(row.get('correos', ''))
    if not recipients:
        return False

    monto = float(row.get('monto', 0) or 0)

    highlight_html = f"""
    <div class="highlight-box" style="background: linear-gradient(135deg, #FFF3E0 0%, #FFE0B2 100%); border-left-color: #F57C00;">
        <h2 style="color: #E65100;">🔔 Recordatorio de Pago Recurrente</h2>
        <div class="critical-info">
            <div class="info-item" style="border-color: #F57C00;">
                <div class="info-label">👤 Proveedor/Destinatario</div>
                <div class="info-value" style="color: #E65100;">{row.get('proveedor', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #F57C00;">
                <div class="info-label">📝 Concepto</div>
                <div class="info-value" style="color: #E65100;">{row.get('nombre', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #F57C00;">
                <div class="info-label">🔄 Periodicidad</div>
                <div class="info-value" style="color: #E65100;">{row.get('periodicidad', 'N/A').capitalize()}</div>
            </div>
            <div class="info-item" style="border-color: #F57C00;">
                <div class="info-label">💰 Monto</div>
                <div class="info-value monto" style="color: #E65100;">${monto:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>🔔 Recordatorio de Pago Recurrente</h3>
        <p>Hola,</p>
        <p>Este es un recordatorio automático de un pago recurrente próximo a vencer:</p>

        <table class="details-table">
            <tr>
                <td>📅 Fecha Límite:</td>
                <td><strong style="color: #E65100; font-size: 18px;">{row.get('fecha_proximo_pago', 'N/A')}</strong></td>
            </tr>
            <tr>
                <td>📝 Descripción:</td>
                <td>{row.get('descripcion', 'N/A')}</td>
            </tr>
        </table>

        <h3>💳 Información de Pago</h3>
        <table class="details-table">
            <tr>
                <td>💳 Método:</td>
                <td>{row.get('metodo_pago', 'N/A')}</td>
            </tr>
            <tr>
                <td>🏦 Banco:</td>
                <td>{row.get('banco', 'N/A')}</td>
            </tr>
            <tr>
                <td>🔢 CLABE:</td>
                <td><code style="background: #f5f5f5; padding: 4px 8px; border-radius: 4px;">{row.get('clabe', 'N/A')}</code></td>
            </tr>
        </table>

        <div class="alert-box" style="background-color: #FFF3E0; border-left-color: #FF9800;">
            <p style="color: #E65100;">
                <strong>⏰ Este recordatorio se envía {row.get('dias_recordatorio', 2)} día(s) antes de la fecha límite.</strong>
            </p>
        </div>
    """

    html_content = get_email_html_template(
        title="Recordatorio de Pago Recurrente",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"🔔 Recordatorio: {row.get('nombre', '(sin nombre)')} vence el {row.get('fecha_proximo_pago', '')}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error enviando recordatorio de recurrente: {e}")
        return False


def check_recurring_payment_reminders():
    """
    Recorre pagos activos y envía recordatorios cuando falten <= dias_recordatorio días.
    También avanza fecha_proximo_pago si ya pasó.
    """
    try:
        conn = get_db_connection()
        today = datetime.now().date()
        rows = conn.execute("SELECT * FROM pagos_recurrentes WHERE activo = 1").fetchall()
        for r in rows:
            r = dict(r)
            due_date = datetime.strptime(r['fecha_proximo_pago'], "%Y-%m-%d").date()
            dias = int(r.get('dias_recordatorio') or 2)

            # 1) Enviar recordatorio si estamos en ventana y no se ha enviado para esta fecha
            if 0 <= (due_date - today).days <= dias:
                if (r.get('ultimo_recordatorio_para_fecha') or "") != r['fecha_proximo_pago']:
                    if send_recurring_payment_reminder(r):
                        conn.execute(
                            "UPDATE pagos_recurrentes SET fecha_ultimo_recordatorio=?, ultimo_recordatorio_para_fecha=? WHERE id=?",
                            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), r['fecha_proximo_pago'], r['id'])
                        )
                        conn.commit()

            # 2) Si la fecha ya pasó, avanzar a la siguiente
            if today > due_date:
                nueva = compute_next_due_date(r['fecha_proximo_pago'], r['periodicidad'])
                conn.execute(
                    "UPDATE pagos_recurrentes SET fecha_proximo_pago=?, ultimo_recordatorio_para_fecha=NULL WHERE id=?",
                    (nueva, r['id'])
                )
                conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error en check_recurring_payment_reminders: {e}")

# ---------- Rutas ----------
@app.route("/pagos_recurrentes", methods=["GET"])
def pagos_recurrentes():
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a Pagos Recurrentes.", "error")
        return redirect(url_for("admin_dashboard"))
    conn = get_db_connection()
    rows = conn.execute(
        "SELECT * FROM pagos_recurrentes ORDER BY activo DESC, fecha_proximo_pago ASC, nombre ASC"
    ).fetchall()
    conn.close()
    pagos = [dict(r) for r in rows]
    return render_template("pagos_recurrentes.html", pagos=pagos)

@app.route("/pagos_recurrentes/nuevo", methods=["POST"])
def crear_pago_recurrente():
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))

    f = request.form
    nombre = (f.get("nombre") or "").strip()
    proveedor = (f.get("proveedor") or "").strip()
    descripcion = f.get("descripcion") or ""
    monto_str = (f.get("monto") or "0").replace(",", ".")
    metodo_pago = f.get("metodo_pago") or ""
    banco = f.get("banco") or ""
    clabe = f.get("clabe") or ""
    periodicidad = f.get("periodicidad") or ""
    fecha_proximo_pago = f.get("fecha_proximo_pago") or ""
    dias_recordatorio = f.get("dias_recordatorio") or "2"
    correos = f.get("correos") or ""

    # Validaciones rápidas
    try:
        monto = float(monto_str)
        dias_rec = int(dias_recordatorio)
        datetime.strptime(fecha_proximo_pago, "%Y-%m-%d")
    except Exception:
        flash("Revisa monto, días de recordatorio y formato de fecha (YYYY-MM-DD).", "error")
        return redirect(url_for("pagos_recurrentes"))

    if periodicidad not in ("mensual","bimestral", "semanal", "anual", "quincenal"):
        flash("Periodicidad inválida.", "error")
        return redirect(url_for("pagos_recurrentes"))
    if not nombre:
        flash("El campo 'Concepto / Nombre' es obligatorio.", "error")
        return redirect(url_for("pagos_recurrentes"))

    conn = get_db_connection()
    conn.execute("""
        INSERT INTO pagos_recurrentes
        (nombre, proveedor, descripcion, monto, metodo_pago, banco, clabe, periodicidad,
         fecha_proximo_pago, dias_recordatorio, correos, activo, fecha_creacion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
    """, (
        nombre, proveedor, descripcion, monto, metodo_pago, banco, clabe, periodicidad,
        fecha_proximo_pago, dias_rec, correos, datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))
    conn.commit()
    conn.close()
    flash("Pago recurrente creado.", "success")
    return redirect(url_for("pagos_recurrentes"))

@app.route("/pagos_recurrentes/<int:pid>.json", methods=["GET"])
def pago_recurrente_json(pid):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        return jsonify({"error": "No autorizado"}), 403
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM pagos_recurrentes WHERE id=?", (pid,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"error": "No encontrado"}), 404
    return jsonify(dict(row))


@app.route("/pagos_recurrentes/<int:pid>/editar", methods=["POST"])
def editar_pago_recurrente(pid):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))

    f = request.form
    nombre = (f.get("nombre") or "").strip()
    proveedor = (f.get("proveedor") or "").strip()
    descripcion = f.get("descripcion") or ""
    monto_str = (f.get("monto") or "0").replace(",", ".")
    metodo_pago = f.get("metodo_pago") or ""
    banco = f.get("banco") or ""
    clabe = f.get("clabe") or ""
    periodicidad = f.get("periodicidad") or ""
    fecha_proximo_pago = f.get("fecha_proximo_pago") or ""
    dias_recordatorio = f.get("dias_recordatorio") or "2"
    correos = f.get("correos") or ""

    try:
        monto = float(monto_str)
        dias_rec = int(dias_recordatorio)
        datetime.strptime(fecha_proximo_pago, "%Y-%m-%d")
    except Exception:
        flash("Revisa monto, días de recordatorio y formato de fecha (YYYY-MM-DD).", "error")
        return redirect(url_for("pagos_recurrentes"))

    if periodicidad not in ("mensual", "bimestral", "semanal", "anual", "quincenal"):
        flash("Periodicidad inválida.", "error")
        return redirect(url_for("pagos_recurrentes"))
    if not nombre:
        flash("El campo 'Concepto / Nombre' es obligatorio.", "error")
        return redirect(url_for("pagos_recurrentes"))

    conn = get_db_connection()
    conn.execute("""
        UPDATE pagos_recurrentes SET
            nombre=?, proveedor=?, descripcion=?, monto=?, metodo_pago=?, banco=?, clabe=?,
            periodicidad=?, fecha_proximo_pago=?, dias_recordatorio=?, correos=?
        WHERE id=?
    """, (nombre, proveedor, descripcion, monto, metodo_pago, banco, clabe,
          periodicidad, fecha_proximo_pago, dias_rec, correos, pid))
    conn.commit()
    conn.close()
    flash("Pago recurrente actualizado.", "success")
    return redirect(url_for("pagos_recurrentes"))


@app.route("/pagos_recurrentes/<int:pid>/toggle", methods=["POST"])
def toggle_pago_recurrente(pid):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))
    conn = get_db_connection()
    row = conn.execute("SELECT activo FROM pagos_recurrentes WHERE id = ?", (pid,)).fetchone()
    if not row:
        conn.close()
        flash("Registro no encontrado.", "error")
        return redirect(url_for("pagos_recurrentes"))
    nuevo = 0 if row["activo"] else 1
    conn.execute("UPDATE pagos_recurrentes SET activo=? WHERE id=?", (nuevo, pid))
    conn.commit()
    conn.close()
    flash("Estado actualizado.", "success")
    return redirect(url_for("pagos_recurrentes"))

@app.route("/pagos_recurrentes/<int:pid>/eliminar", methods=["POST"])
def eliminar_pago_recurrente(pid):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))
    conn = get_db_connection()
    conn.execute("DELETE FROM pagos_recurrentes WHERE id = ?", (pid,))
    conn.commit()
    conn.close()
    flash("Pago recurrente eliminado.", "success")
    return redirect(url_for("pagos_recurrentes"))

@app.route("/pagos_recurrentes/<int:pid>/avanzar", methods=["POST"])
def avanzar_fecha_pago_recurrente(pid):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))
    conn = get_db_connection()
    row = conn.execute("SELECT fecha_proximo_pago, periodicidad FROM pagos_recurrentes WHERE id = ?", (pid,)).fetchone()
    if row:
        nueva = compute_next_due_date(row["fecha_proximo_pago"], row["periodicidad"])
        conn.execute(
            "UPDATE pagos_recurrentes SET fecha_proximo_pago=?, ultimo_recordatorio_para_fecha=NULL WHERE id=?",
            (nueva, pid)
        )
        conn.commit()
    conn.close()
    flash("Fecha próxima actualizada.", "success")
    return redirect(url_for("pagos_recurrentes"))

@app.route("/admin/run_recurring_reminders_now")
def run_recurring_reminders_now():
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))
    try:
        check_recurring_payment_reminders()
        flash("Proceso de recordatorios ejecutado.", "success")
    except Exception as e:
        flash(f"Error ejecutando recordatorios: {e}", "error")
    return redirect(url_for("pagos_recurrentes"))



# Carpeta para almacenar archivos adjuntos
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Variables globales y credenciales
ADMIN_USER = "dagarcia"
ADMIN_PASS = "daGt20!!23"
COORDINADOR_USER = "Gadelarosa"
COORDINADOR_PASS = "05360"
capital_total = 100000.0

def read_employees():
    employees = []
    csv_file_path = os.path.join("data", "empleados.csv")
    if os.path.exists(csv_file_path):
        with open(csv_file_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                employees.append({
                    "nombre": row["Nompropio"],
                    "departamento": row["Departamento"]
                })

    return employees

# Función auxiliar para procesar múltiples correos
def get_recipients(correos):
    return [correo.strip() for correo in correos.split(",") if correo.strip()]

def send_notification_email(solicitud):
    """Email para notificar nueva solicitud al administrador"""

    # Sección destacada con info crítica
    highlight_html = f"""
    <div class="highlight-box">
        <h2>📋 Información Principal</h2>
        <div class="critical-info">
            <div class="info-item">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value">{solicitud.get('tipo_solicitud', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value">{solicitud.get('tipo_pago', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto">${solicitud.get('monto', 0):,.2f}</div>
            </div>
        </div>
    </div>
    """

    # Información adicional de comisión si aplica
    comision_html = ""
    if solicitud.get('tiene_comision', 0) == 1:
        comision_html = f"""
        <div class="alert-box">
            <p><strong>⚠️ Comisión BBVA Sin Factura: {solicitud.get('porcentaje_comision', 0)}%</strong></p>
            <p>Monto sin comisión: <strong>${solicitud.get('monto_sin_comision', 0):,.2f}</strong></p>
            <p>Monto de comisión: <strong>${solicitud.get('monto_comision', 0):,.2f}</strong></p>
        </div>
        """

    # Información de anticipo si aplica
    anticipo_html = ""
    if solicitud.get('anticipo', 'No') == 'Si':
        anticipo_html = f"""
        <div class="alert-box">
            <p><strong>💵 Solicitud con Anticipo</strong></p>
            <p>Tipo: {solicitud.get('tipo_anticipo', 'porcentaje')}</p>
            <p>Porcentaje: <strong>{solicitud.get('porcentaje_anticipo', 0)}%</strong></p>
            <p>Monto anticipo: <strong>${solicitud.get('monto_anticipo', 0):,.2f}</strong></p>
            <p>Monto restante: <strong>${solicitud.get('monto_restante', 0):,.2f}</strong></p>
        </div>
        """

    # Contenido del correo
    content_html = f"""
        <h3>🆕 Nueva Solicitud de Pago</h3>
        <p>Se ha recibido una nueva solicitud de pago que requiere tu atención:</p>

        {comision_html}
        {anticipo_html}

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud.get('fp', 'N/A')}</strong></td>
            </tr>
            <tr>
                <td>👤 Solicitante:</td>
                <td>{solicitud.get('nombre', 'N/A')}</td>
            </tr>
            <tr>
                <td>📧 Correo(s):</td>
                <td>{solicitud.get('correo', 'N/A')}</td>
            </tr>
            <tr>
                <td>🏢 Departamento:</td>
                <td>{solicitud.get('departamento', 'N/A')}</td>
            </tr>
            <tr>
                <td>📅 Fecha Límite:</td>
                <td><strong style="color: #E65100;">{solicitud.get('fecha_limite', 'N/A')}</strong></td>
            </tr>
            <tr>
                <td>🏦 Banco:</td>
                <td>{solicitud.get('banco', 'N/A')}</td>
            </tr>
            <tr>
                <td>🔢 CLABE:</td>
                <td><code style="background: #f5f5f5; padding: 4px 8px; border-radius: 4px;">{solicitud.get('clabe', 'N/A')}</code></td>
            </tr>
            <tr>
                <td>📝 Descripción:</td>
                <td>{solicitud.get('descripcion', 'N/A')}</td>
            </tr>
            <tr>
                <td>📌 Datos de Depósito:</td>
                <td>{solicitud.get('datos_deposito', 'N/A')}</td>
            </tr>
            <tr>
                <td>📊 Estado:</td>
                <td><span class="status-badge status-pendiente">{solicitud.get('estado', 'Pendiente')}</span></td>
            </tr>
            <tr>
                <td>🕒 Fecha de Solicitud:</td>
                <td>{solicitud.get('fecha', 'N/A')}</td>
            </tr>
        </table>
    """

    html_content = get_email_html_template(
        title="Nueva Solicitud de Pago",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"🆕 Nueva Solicitud de Pago - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    msg['To'] = "dagarcia@ad17solutions.com"
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar notificación: {e}")
        return False

def send_approval_email(solicitud):
    """Email de aprobación de solicitud"""

    highlight_html = f"""
    <div class="highlight-box">
        <h2>✅ Solicitud Aprobada</h2>
        <div class="critical-info">
            <div class="info-item">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value">{solicitud['tipo_solicitud']}</div>
            </div>
            <div class="info-item">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value">{solicitud['tipo_pago']}</div>
            </div>
            <div class="info-item">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto">${solicitud['monto']:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>✅ ¡Tu solicitud ha sido aprobada!</h3>
        <p>Hola <strong>{solicitud['nombre']}</strong>,</p>
        <p>Nos complace informarte que tu solicitud de pago ha sido <strong style="color: #2E7D32;">APROBADA</strong>.</p>

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud['fp']}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha Límite de Pago:</td>
                <td><strong style="color: #E65100;">{solicitud['fecha_limite']}</strong></td>
            </tr>
            <tr>
                <td>🕒 Fecha de Solicitud:</td>
                <td>{solicitud['fecha']}</td>
            </tr>
        </table>

        <h3>🏦 Datos de la Cuenta Destino</h3>
        <table class="details-table">
            <tr>
                <td>🏦 Banco:</td>
                <td><strong>{solicitud['banco']}</strong></td>
            </tr>
            <tr>
                <td>🔢 CLABE:</td>
                <td><code style="background: #f5f5f5; padding: 4px 8px; border-radius: 4px; font-size: 14px;">{solicitud['clabe']}</code></td>
            </tr>
        </table>

        <p style="margin-top: 25px; color: #666;">
            El pago se realizará según los tiempos establecidos. Recibirás una notificación cuando el pago sea liquidado.
        </p>
    """

    html_content = get_email_html_template(
        title="Solicitud Aprobada",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"✅ Tu solicitud de pago ha sido aprobada - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    recipients = get_recipients(solicitud['correo'])
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar correo de aprobación: {e}")
        return False


def send_approval_anticipo_email(solicitud):
    """Email de aprobación con anticipo"""

    highlight_html = f"""
    <div class="highlight-box">
        <h2>✅ Solicitud Aprobada con Anticipo</h2>
        <div class="critical-info">
            <div class="info-item">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value">{solicitud.get('tipo_solicitud', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value">{solicitud.get('tipo_pago', 'N/A')}</div>
            </div>
            <div class="info-item">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto">${solicitud['monto']:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>✅ ¡Tu solicitud con anticipo ha sido aprobada!</h3>
        <p>Hola <strong>{solicitud['nombre']}</strong>,</p>
        <p>Tu solicitud de pago con anticipo ha sido <strong style="color: #2E7D32;">APROBADA</strong>.</p>

        <div class="alert-box">
            <p><strong>💵 Detalles del Anticipo</strong></p>
            <p>Porcentaje de anticipo: <strong>{solicitud['porcentaje_anticipo']}%</strong></p>
            <p>Monto del anticipo: <strong>${solicitud.get('monto_anticipo', solicitud['monto'] * solicitud['porcentaje_anticipo'] / 100):,.2f}</strong></p>
            <p>Monto restante: <strong style="color: #E65100;">${solicitud['monto_restante']:,.2f}</strong></p>
        </div>

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud['fp']}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha Límite:</td>
                <td><strong style="color: #E65100;">{solicitud['fecha_limite']}</strong></td>
            </tr>
            <tr>
                <td>🕒 Fecha de Solicitud:</td>
                <td>{solicitud['fecha']}</td>
            </tr>
        </table>

        <h3>🏦 Datos de la Cuenta Destino</h3>
        <table class="details-table">
            <tr>
                <td>🏦 Banco:</td>
                <td><strong>{solicitud['banco']}</strong></td>
            </tr>
            <tr>
                <td>🔢 CLABE:</td>
                <td><code style="background: #f5f5f5; padding: 4px 8px; border-radius: 4px; font-size: 14px;">{solicitud['clabe']}</code></td>
            </tr>
        </table>

        <p style="margin-top: 25px; color: #666;">
            Se realizará el pago del anticipo según los tiempos establecidos. El monto restante se pagará posteriormente.
        </p>
    """

    html_content = get_email_html_template(
        title="Solicitud Aprobada con Anticipo",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"✅ Tu solicitud con anticipo ha sido aprobada - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    recipients = get_recipients(solicitud['correo'])
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar correo de aprobación con anticipo: {e}")
        return False

def send_declined_email(solicitud):
    """Email de solicitud declinada"""

    highlight_html = f"""
    <div class="highlight-box" style="background: linear-gradient(135deg, #FFEBEE 0%, #FFCDD2 100%); border-left-color: #C62828;">
        <h2 style="color: #C62828;">❌ Solicitud Declinada</h2>
        <div class="critical-info">
            <div class="info-item" style="border-color: #C62828;">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value" style="color: #C62828;">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #C62828;">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value" style="color: #C62828;">{solicitud['tipo_solicitud']}</div>
            </div>
            <div class="info-item" style="border-color: #C62828;">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value" style="color: #C62828;">{solicitud['tipo_pago']}</div>
            </div>
            <div class="info-item" style="border-color: #C62828;">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto" style="color: #C62828;">${solicitud['monto']:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>❌ Tu solicitud ha sido declinada</h3>
        <p>Hola <strong>{solicitud['nombre']}</strong>,</p>
        <p>Lamentamos informarte que tu solicitud de pago ha sido <strong style="color: #C62828;">DECLINADA</strong>.</p>

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud['fp']}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha Límite:</td>
                <td>{solicitud['fecha_limite']}</td>
            </tr>
            <tr>
                <td>🕒 Fecha de Solicitud:</td>
                <td>{solicitud['fecha']}</td>
            </tr>
        </table>

        <div class="alert-box" style="background-color: #FFEBEE; border-left-color: #C62828;">
            <p style="color: #C62828;">
                <strong>📞 ¿Tienes dudas?</strong><br>
                Por favor contacta con el departamento administrativo para obtener más información sobre el motivo de la declinación.
            </p>
        </div>
    """

    html_content = get_email_html_template(
        title="Solicitud Declinada",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"❌ Tu solicitud de pago ha sido declinada - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    recipients = get_recipients(solicitud['correo'])
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar correo de declinación: {e}")
        return False


def send_liquidacion_total_email(solicitud, attachment_file=None):
    """Email de liquidación total"""

    # Calcular montos
    total = solicitud.get("monto", 0.0)
    porcentaje = float(solicitud.get("porcentaje_anticipo", 0.0))
    anticipo_amount = total * (porcentaje / 100)

    highlight_html = f"""
    <div class="highlight-box" style="background: linear-gradient(135deg, #E3F2FD 0%, #BBDEFB 100%); border-left-color: #1565C0;">
        <h2 style="color: #1565C0;">💰 Liquidación Total Completada</h2>
        <div class="critical-info">
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('tipo_solicitud', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('tipo_pago', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto" style="color: #1565C0;">${total:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>💰 ¡Liquidación Total Completada!</h3>
        <p>Hola <strong>{solicitud['nombre']}</strong>,</p>
        <p>Te informamos que se ha realizado la <strong style="color: #1565C0;">LIQUIDACIÓN TOTAL</strong> de tu solicitud de pago.</p>

        <div class="alert-box" style="background-color: #E3F2FD; border-left-color: #1565C0;">
            <p style="color: #1565C0;"><strong>📊 Desglose del Pago</strong></p>
            <p>Anticipo: <strong>{porcentaje}%</strong> (${anticipo_amount:,.2f})</p>
            <p>Último pago (Liquidación): <strong>${solicitud.get('monto_restante', 0.0):,.2f}</strong></p>
            <p style="font-size: 18px; margin-top: 10px;">
                <strong>Total pagado: ${total:,.2f}</strong>
            </p>
        </div>

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud['fp']}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha de Liquidación:</td>
                <td><strong>{datetime.now().strftime('%d/%m/%Y %H:%M')}</strong></td>
            </tr>
        </table>

        <p style="margin-top: 25px; color: #666;">
            {'📎 El comprobante del pago final está adjunto en este correo.' if attachment_file else ''}
        </p>

        <p style="color: #2E7D32; font-weight: 600;">
            ✅ Con esta liquidación se ha completado el pago total de tu solicitud.
        </p>
    """

    html_content = get_email_html_template(
        title="Liquidación Total",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"💰 Liquidación Total Completada - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    recipients = get_recipients(solicitud['correo'])
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    if attachment_file:
        try:
            file_data = attachment_file.read()
            mime_type, _ = mimetypes.guess_type(attachment_file.filename)
            if not mime_type:
                mime_type = "application/octet-stream"
            maintype, subtype = mime_type.split("/", 1)
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=attachment_file.filename)
        except Exception as e:
            print("Error al adjuntar el archivo:", e)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar correo de liquidación total: {e}")
        return False


def send_liquidado_email(solicitud, attachment_file=None):
    """Email de pago liquidado"""

    highlight_html = f"""
    <div class="highlight-box" style="background: linear-gradient(135deg, #E3F2FD 0%, #BBDEFB 100%); border-left-color: #1565C0;">
        <h2 style="color: #1565C0;">💸 Pago Liquidado</h2>
        <div class="critical-info">
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value" style="color: #1565C0;">{solicitud['tipo_solicitud']}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value" style="color: #1565C0;">{solicitud['tipo_pago']}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto" style="color: #1565C0;">${solicitud['monto']:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>💸 ¡Tu pago ha sido liquidado!</h3>
        <p>Hola <strong>{solicitud['nombre']}</strong>,</p>
        <p>Te informamos que tu solicitud de pago ha sido <strong style="color: #1565C0;">LIQUIDADA</strong> exitosamente.</p>

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud['fp']}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha de Liquidación:</td>
                <td><strong>{datetime.now().strftime('%d/%m/%Y %H:%M')}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha Límite (Original):</td>
                <td>{solicitud['fecha_limite']}</td>
            </tr>
            <tr>
                <td>🕒 Fecha de Solicitud:</td>
                <td>{solicitud['fecha']}</td>
            </tr>
        </table>

        <p style="margin-top: 25px; color: #666;">
            {'📎 El comprobante del pago está adjunto en este correo.' if attachment_file else ''}
        </p>

        <p style="color: #2E7D32; font-weight: 600;">
            ✅ El pago ha sido procesado y completado exitosamente.
        </p>
    """

    html_content = get_email_html_template(
        title="Pago Liquidado",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"💸 Tu solicitud de pago ha sido liquidada - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    recipients = get_recipients(solicitud['correo'])
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    if attachment_file:
        try:
            file_data = attachment_file.read()
            mime_type, _ = mimetypes.guess_type(attachment_file.filename)
            if not mime_type:
                mime_type = "application/octet-stream"
            maintype, subtype = mime_type.split("/", 1)
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=attachment_file.filename)
        except Exception as e:
            print("Error al adjuntar el archivo:", e)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar correo de liquidación: {e}")
        return False


def send_liquidado_anticipo_email(solicitud, attachment_file=None):
    """Email de pago liquidado con anticipo"""

    highlight_html = f"""
    <div class="highlight-box" style="background: linear-gradient(135deg, #E3F2FD 0%, #BBDEFB 100%); border-left-color: #1565C0;">
        <h2 style="color: #1565C0;">💸 Anticipo Liquidado</h2>
        <div class="critical-info">
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">👤 Destinatario</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('destinatario', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">📝 Tipo de Solicitud</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('tipo_solicitud', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">💳 Tipo de Pago</div>
                <div class="info-value" style="color: #1565C0;">{solicitud.get('tipo_pago', 'N/A')}</div>
            </div>
            <div class="info-item" style="border-color: #1565C0;">
                <div class="info-label">💰 Monto Total</div>
                <div class="info-value monto" style="color: #1565C0;">${solicitud['monto']:,.2f}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>💸 ¡El anticipo ha sido liquidado!</h3>
        <p>Hola <strong>{solicitud['nombre']}</strong>,</p>
        <p>Te informamos que el <strong style="color: #1565C0;">ANTICIPO</strong> de tu solicitud ha sido liquidado.</p>

        <div class="alert-box" style="background-color: #FFF3E0; border-left-color: #FF9800;">
            <p style="color: #E65100;"><strong>💵 Detalles del Pago</strong></p>
            <p>Anticipo: <strong>{solicitud['porcentaje_anticipo']}%</strong></p>
            <p>Monto del anticipo liquidado: <strong>${solicitud.get('monto_anticipo', solicitud['monto'] * solicitud['porcentaje_anticipo'] / 100):,.2f}</strong></p>
            <p>Monto restante pendiente: <strong style="color: #E65100;">${solicitud['monto_restante']:,.2f}</strong></p>
        </div>

        <table class="details-table">
            <tr>
                <td>📋 FP:</td>
                <td><strong>{solicitud['fp']}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha de Liquidación:</td>
                <td><strong>{datetime.now().strftime('%d/%m/%Y %H:%M')}</strong></td>
            </tr>
            <tr>
                <td>📅 Fecha Límite:</td>
                <td>{solicitud['fecha_limite']}</td>
            </tr>
            <tr>
                <td>🕒 Fecha de Solicitud:</td>
                <td>{solicitud['fecha']}</td>
            </tr>
        </table>

        <p style="margin-top: 25px; color: #666;">
            {'📎 El comprobante del anticipo está adjunto en este correo.' if attachment_file else ''}
        </p>

        <p style="color: #FF9800; font-weight: 600;">
            ⚠️ Recuerda que aún queda pendiente el pago del monto restante.
        </p>
    """

    html_content = get_email_html_template(
        title="Anticipo Liquidado",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"💸 El anticipo de tu solicitud ha sido liquidado - FP {solicitud['fp']}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    recipients = get_recipients(solicitud['correo'])
    msg['To'] = ", ".join(recipients)
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    if attachment_file:
        try:
            file_data = attachment_file.read()
            mime_type, _ = mimetypes.guess_type(attachment_file.filename)
            if not mime_type:
                mime_type = "application/octet-stream"
            maintype, subtype = mime_type.split("/", 1)
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=attachment_file.filename)
        except Exception as e:
            print("Error al adjuntar el archivo:", e)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login('ad17solutionsbot@gmail.com', 'misvtfhrnwbmiptb')
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error al enviar correo de liquidación con anticipo: {e}")
        return False


@app.route("/")
def index():
    return redirect(url_for("admin_dashboard"))

@app.route("/solicitar_pago", methods=["GET", "POST"])
def solicitar_pago():
    import re
    # helper robusto para floats (quita símbolos, espacios y miles)
    def _to_float(s):
        if s is None:
            return 0.0
        s = str(s)
        # deja sólo dígitos, punto y coma, luego elimina miles y usa punto como decimal
        s = re.sub(r"[^\d,.\-]", "", s)
        # si viene con miles tipo 1,234.56, quita comas
        s = s.replace(",", "")
        try:
            return float(s)
        except ValueError:
            return 0.0

    employees = read_employees()
    conceptos_indirectos = []
    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Costos",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor(dictionary=True)
        cursor.execute("""SELECT regID as id, concepto
                          FROM AD17_Costos.Conceptos_Indirectos
                          WHERE habilitado = 1
                          ORDER BY concepto ASC""")
        conceptos_indirectos = cursor.fetchall()
        cursor.close()
        remote_conn.close()
    except Exception as e:
        print("Error al obtener conceptos indirectos:", e)
        conceptos_indirectos = []

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor(dictionary=True)
        query = """
            SELECT
                i.id AS id,
                d.nombre AS nombre,
                d.rfc AS rfc,
                d.direccion AS direccion,
                d.referencia AS referencia,
                c.regID AS contID,
                c.contacto AS contacto,
                c.telefono AS telefono,
                c.email AS email,
                p.regID AS metID,
                m.forma AS metodo,
                p.banco AS banco,
                p.beneficiario AS beneficiario,
                p.clabe AS clabe
            FROM AD17_Proveedores.ID AS i
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.Datos
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.Datos GROUP BY provID
                )
            ) AS d ON d.provID LIKE i.id
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.Contactos
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.Contactos GROUP BY provID
                )
            ) AS c ON c.provID LIKE i.id
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.MetodosDePago
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.MetodosDePago GROUP BY provID
                )
            ) AS p ON p.provID LIKE i.id
            LEFT JOIN AD17_Proveedores.Metodos AS m on m.regID LIKE p.metodo
            ORDER BY d.nombre ASC;
        """
        cursor.execute(query)
        proveedores = cursor.fetchall()
        cursor.close()
        remote_conn.close()
    except Exception as e:
        print("Error al obtener proveedores:", e)
        proveedores = []

    if request.method == "POST":
        # -------- campos base --------
        fp = request.form.get("fp")
        selected_nombre = request.form.get("selected_nombre")
        if selected_nombre == "otro":
            nombre = request.form.get("nombre_otro")
            departamento = request.form.get("departamento_otro")
        else:
            nombre = selected_nombre
            departamento = request.form.get("departamento_hidden")

        tipo_solicitud = request.form.get("tipo_solicitud")
        categoria_administrativa = request.form.get("categoria_administrativos", "")

        tipo_pago = request.form.get("tipo_pago")
        descripcion = request.form.get("descripcion", "")
        # Ya no concatenamos la categoría al campo descripción
        # La categoría se guarda en su propia columna: categoria_administrativa

        datos_deposito = request.form.get("datos_deposito", "")

        # -------- montos y comisión --------
        monto_sin_comision = _to_float(request.form.get("monto"))
        tiene_comision = 1 if tipo_pago == "BBVA Sin factura" else 0
        porcentaje_comision = 6.0 if tiene_comision else 0.0
        if tiene_comision:
            # total bruto necesario para que neto sea monto_sin_comision
            monto = round(monto_sin_comision / 0.94, 2)
            monto_comision = round(monto - monto_sin_comision, 2)
        else:
            monto = round(monto_sin_comision, 2)
            monto_comision = 0.0

        # -------- correos --------
        correos = request.form.getlist("correo")
        correo = ", ".join([c.strip() for c in correos if c.strip()])

        # -------- datos bancarios y fechas --------
        banco = request.form.get("banco")
        clabe = request.form.get("clabe")
        referencia = request.form.get("referencia", "")

        fecha_limite = request.form.get("fecha_limite")
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        destinatario = request.form.get("destinatario")

        # -------- ARCHIVOS ADJUNTOS (3 archivos + legacy) --------
        archivo_adjunto = ""
        archivo_factura = ""
        archivo_recibo = ""
        archivo_orden_compra = ""

        # Archivo adjunto principal (legacy)
        adjunto_file = request.files.get("adjunto_file")
        if adjunto_file and adjunto_file.filename != "":
            filename = secure_filename(adjunto_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            adjunto_file.save(filepath)
            archivo_adjunto = filename

        # Factura
        factura_file = request.files.get("factura_file")
        if factura_file and factura_file.filename != "":
            filename = secure_filename(factura_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            factura_file.save(filepath)
            archivo_factura = filename

        # Recibo
        recibo_file = request.files.get("recibo_file")
        if recibo_file and recibo_file.filename != "":
            filename = secure_filename(recibo_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            recibo_file.save(filepath)
            archivo_recibo = filename

        # Orden de Compra
        orden_file = request.files.get("orden_compra_file")
        if orden_file and orden_file.filename != "":
            filename = secure_filename(orden_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            orden_file.save(filepath)
            archivo_orden_compra = filename

        # -------- programado --------
        es_programada = request.form.get("es_programada")
        estado = "Programado" if es_programada else "Pendiente"
        es_programada_val = 1 if es_programada else 0

        # -------- ANTICIPO (robusto) --------
        anticipo_val = "Si" if request.form.get("anticipo") else "No"
        tipo_anticipo = request.form.get("tipo_anticipo", "porcentaje")

        porcentaje_anticipo = _to_float(request.form.get("porcentaje_anticipo"))
        monto_anticipo = _to_float(request.form.get("monto_anticipo"))

        if anticipo_val == "Si":
            if tipo_anticipo == "porcentaje" and porcentaje_anticipo > 0:
                monto_anticipo = round(monto * (porcentaje_anticipo / 100.0), 2)
            elif tipo_anticipo == "cantidad" and monto_anticipo > 0:
                porcentaje_anticipo = round(((monto_anticipo / monto) * 100.0), 2) if monto > 0 else 0.0
            else:
                # entrada inválida -> 0
                porcentaje_anticipo = 0.0
                monto_anticipo = 0.0
        else:
            porcentaje_anticipo = 0.0
            monto_anticipo = 0.0
            tipo_anticipo = "porcentaje"

        monto_restante = round(max(monto - monto_anticipo, 0.0), 2)

        # -------- VIÁTICOS (recalcula total y anticipo con el nuevo total) --------
        if tipo_solicitud == "Viáticos":
            persona_nombres = request.form.getlist("persona_nombre[]")
            persona_montos = request.form.getlist("persona_monto[]")
            persona_clabes = request.form.getlist("persona_clabe[]")
            detalle_personas = []
            total_calc = 0.0

            for i in range(len(persona_nombres)):
                pnombre = (persona_nombres[i] or "").strip()
                pmonto = _to_float(persona_montos[i]) if i < len(persona_montos) else 0.0
                p_clabe = (persona_clabes[i] or "").strip() if i < len(persona_clabes) else ""
                detalle_personas.append({"nombre": pnombre, "monto": pmonto, "clabe": p_clabe})
                total_calc += pmonto

            monto_sin_comision = round(total_calc, 2)
            if tiene_comision:
                monto = round(monto_sin_comision / 0.94, 2)
                monto_comision = round(monto - monto_sin_comision, 2)
            else:
                monto = monto_sin_comision
                monto_comision = 0.0

            if anticipo_val == "Si":
                if tipo_anticipo == "porcentaje":
                    monto_anticipo = round(monto * (porcentaje_anticipo / 100.0), 2)
                else:
                    if monto_anticipo > monto:
                        monto_anticipo = monto
                        porcentaje_anticipo = 100.0
                monto_restante = round(monto - monto_anticipo, 2)

            descripcion += "\nDetalle de Viáticos: " + json.dumps(detalle_personas, ensure_ascii=False)

        # -------- INSERT --------
        conn = get_db_connection()
        conn.execute("""
            INSERT INTO solicitudes
            (fp, nombre, destinatario, correo, departamento, tipo_solicitud, tipo_pago, descripcion,
             datos_deposito, banco, clabe, referencia, monto, estado, fecha, fecha_limite, archivo_adjunto,
             archivo_factura, archivo_recibo, archivo_orden_compra,
             anticipo, porcentaje_anticipo, monto_restante, es_programada, tiene_comision,
             porcentaje_comision, monto_comision, monto_sin_comision, tipo_anticipo, monto_anticipo,
             categoria_administrativa)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            fp, nombre, destinatario, correo, departamento, tipo_solicitud, tipo_pago, descripcion,
            datos_deposito, banco, clabe, referencia, monto, estado, fecha, fecha_limite, archivo_adjunto,
            archivo_factura, archivo_recibo, archivo_orden_compra,
            anticipo_val, porcentaje_anticipo, monto_restante, es_programada_val, tiene_comision,
            porcentaje_comision, monto_comision, monto_sin_comision, tipo_anticipo, monto_anticipo,
            categoria_administrativa
        ))
        conn.commit()
        conn.close()

        # -------- notificación --------
        solicitud = {
            "fp": fp,
            "nombre": nombre,
            "destinatario": destinatario,
            "correo": correo,
            "departamento": departamento,
            "tipo_solicitud": tipo_solicitud,
            "tipo_pago": tipo_pago,
            "descripcion": descripcion,
            "datos_deposito": datos_deposito,
            "banco": banco,
            "clabe": clabe,
            "monto": monto,
            "monto_sin_comision": monto_sin_comision,
            "tiene_comision": tiene_comision,
            "porcentaje_comision": porcentaje_comision,
            "monto_comision": monto_comision,
            "estado": estado,
            "fecha": fecha,
            "fecha_limite": fecha_limite,
            "archivo_adjunto": archivo_adjunto,
            "anticipo": anticipo_val,
            "porcentaje_anticipo": porcentaje_anticipo,
            "monto_anticipo": monto_anticipo,
            "monto_restante": monto_restante,
            "tipo_anticipo": tipo_anticipo
        }
        send_notification_email(solicitud)

        # sync remoto en segundo plano
        try:
            threading.Thread(target=sync_solicitudes_to_remote, daemon=True).start()
        except Exception as e:
            print(f"Error al iniciar sincronización: {e}")

        flash("¡Solicitud enviada exitosamente!", "success")
        return redirect(url_for("solicitar_pago"))

    return render_template("solicitar_pago.html",
                           employees=employees,
                           proveedores=proveedores,
                           conceptos_indirectos=conceptos_indirectos)

# === ENDPOINTS PARA SINCRONIZACIÓN MANUAL ===
@app.route("/admin/sync_now")
def sync_now():
    """
    Endpoint para sincronización manual (solo administradores)
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para realizar esta acción.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        # Ejecutar sincronización en hilo separado para no bloquear la respuesta
        threading.Thread(target=sync_all_data, daemon=True).start()
        flash("Sincronización iniciada exitosamente. Revisa los logs para más detalles.", "success")
    except Exception as e:
        flash(f"Error al iniciar sincronización: {e}", "error")

    return redirect(url_for("admin_dashboard"))

@app.route("/admin/test_remote_connection")
def test_remote_connection():
    """
    Endpoint para probar la conexión remota
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para realizar esta acción.", "error")
        return redirect(url_for("admin_dashboard"))

    connection = get_remote_db_connection()
    if connection:
        try:
            with connection.cursor() as cursor:
                cursor.execute("SELECT 1")
                result = cursor.fetchone()
            connection.close()
            flash("✓ Conexión a base de datos remota exitosa", "success")
        except Exception as e:
            flash(f"Error en consulta: {e}", "error")
    else:
        flash("✗ Error al conectar con la base de datos remota", "error")

    return redirect(url_for("admin_dashboard"))

@app.route("/admin/verify_remote_structure")
def verify_remote_structure():
    """
    Endpoint para verificar y corregir la estructura de la base remota
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para realizar esta acción.", "error")
        return redirect(url_for("admin_dashboard"))

    if verify_and_fix_remote_tables():
        flash("✓ Estructura de base de datos verificada y corregida", "success")
    else:
        flash("✗ Error al verificar estructura de base de datos", "error")

    return redirect(url_for("admin_dashboard"))


@app.route("/admin_login", methods=["GET", "POST"])
def admin_login():
    if session.get("admin_logged_in"):
        return redirect(url_for("admin_dashboard"))
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        remember = request.form.get("remember_me")
        session.permanent = True if remember else False
        if ((username == ADMIN_USER and password == ADMIN_PASS) or
            (username == "BrandonViNu" and password == "Bvn2016!") or
            (username == "Rubengarcia" and password == "Dany1712") or
            (username == "DiegoGarciaToledano" and password == "daGt20!!25") or
            (username == "Gadelarosa" and password == "05360") or
            (username == "mrivero" and password == "B230163z")):
            session["admin_logged_in"] = True
            session["role"] = "admin"
            flash("Has iniciado sesión correctamente (Administrador).", "success")
            return redirect(url_for("admin_dashboard"))
        coordinators = {

            "ddelarosa": "082291",
            "gildardo": "gilad17",
            "vmejia": "FOME1005",
            "DafneDeLaRosa": "dDLRz20!!25"
        }
        if username in coordinators and password == coordinators[username]:
            session["admin_logged_in"] = True
            session["role"] = "coordinador"
            flash("Has iniciado sesión correctamente (Coordinador).", "success")
            return redirect(url_for("admin_dashboard"))
        flash("Credenciales incorrectas.", "error")
        return render_template("admin_login.html")
    return render_template("admin_login.html")

@app.route("/admin_logout")
def admin_logout():
    session.pop("admin_logged_in", None)
    session.pop("role", None)
    flash("Has cerrado sesión.", "success")
    return redirect(url_for("admin_login"))

@app.route("/actualizar_flujo", methods=["POST"])
def actualizar_flujo():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))
    if session.get("role") != "admin":
        flash("No tienes permiso para actualizar el flujo de capital.", "error")
        return redirect(url_for("admin_dashboard"))
    global capital_total
    new_value = request.form.get("capital_total")
    try:
        capital_total = float(new_value)
        flash("Flujo de capital actualizado.", "success")
    except ValueError:
        flash("Valor no válido.", "error")
    return redirect(url_for("admin_dashboard"))
@app.route("/admin_dashboard", methods=["GET", "POST"], endpoint="admin_dashboard")
def admin_dashboard():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    estado_filtro = (request.args.get("estado") or "").strip()
    busqueda = (request.args.get("busqueda") or "").strip()
    criterio_busqueda = (request.args.get("criterio_busqueda") or "todos").strip().lower()
    page = int(request.args.get("p", 1))
    page_size = int(request.args.get("page_size", PAGE_SIZE_DEFAULT))

    data = query_solicitudes_paginated(
        page=page,
        page_size=page_size,
        estado_filtro=estado_filtro,
        busqueda=busqueda,
        criterio=criterio_busqueda
    )

    def _empty(v):
        return v is None or (isinstance(v, str) and v.strip() == "")

    def _f(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    def _i(v):
        try:
            return int(v)
        except (TypeError, ValueError):
            return 0

    items = []
    missing_ids = []

    for r in data["items"]:
        row = dict(r)

        it = {
            "id": row["id"],
            "fp": row.get("fp", ""),
            "nombre": row.get("nombre", ""),
            "destinatario": row.get("destinatario", ""),
            "correo": row.get("correo", ""),
            "departamento": row.get("departamento", ""),
            "tipo_solicitud": row.get("tipo_solicitud", ""),
            "tipo_pago": row.get("tipo_pago", ""),
            "monto": _f(row.get("monto")),
            "estado": row.get("estado", ""),
            "fecha": row.get("fecha", ""),
            "fecha_limite": row.get("fecha_limite", ""),
            "archivo_adjunto": row.get("archivo_adjunto", ""),

            # ===== NUEVO: Archivos adicionales =====
            "archivo_factura": row.get("archivo_factura", ""),
            "archivo_recibo": row.get("archivo_recibo", ""),
            "archivo_orden_compra": row.get("archivo_orden_compra", ""),
            # ===== FIN NUEVO =====

            # detalles
            "banco": row.get("banco") or "",
            "clabe": row.get("clabe") or "",
            "referencia": row.get("referencia") or "",
            "descripcion": row.get("descripcion") or "",
            "datos_deposito": row.get("datos_deposito") or "",

            # --- CAMPOS DE ANTICIPO (CLAVE PARA EL BADGE) ---
            "anticipo": row.get("anticipo"),  # "Si"/"No" o None
            "porcentaje_anticipo": row.get("porcentaje_anticipo"),
            "monto_anticipo": row.get("monto_anticipo"),
            "monto_restante": row.get("monto_restante"),
            "tipo_anticipo": row.get("tipo_anticipo"),

            # comisión (opcional pero útil)
            "tiene_comision": _i(row.get("tiene_comision")),
            "porcentaje_comision": _f(row.get("porcentaje_comision")),
            "monto_sin_comision": _f(row.get("monto_sin_comision")),
            "monto_comision": _f(row.get("monto_comision")),

            # ===== NUEVO: Campos de historial =====
            "historial_estados": row.get("historial_estados", "[]"),
            "fecha_aprobado": row.get("fecha_aprobado", ""),
            "fecha_liquidado": row.get("fecha_liquidado", ""),
            "fecha_ultimo_cambio": row.get("fecha_ultimo_cambio", ""),
            # ===== FIN NUEVO =====
        }
        items.append(it)

        # si algo clave falta, lo traeremos directo de la tabla
        if (
            _empty(it["banco"]) or _empty(it["clabe"]) or it["referencia"] is None
            or _empty(it["descripcion"]) or _empty(it["datos_deposito"])
            or it["anticipo"] is None
            or it["porcentaje_anticipo"] is None
            or it["monto_anticipo"] is None
            or it["monto_restante"] is None
            or it["tipo_anticipo"] is None
            # ===== NUEVO: Verificar archivos adicionales =====
            or _empty(it["archivo_factura"])
            or _empty(it["archivo_recibo"])
            or _empty(it["archivo_orden_compra"])
            # ===== FIN NUEVO =====
        ):
            missing_ids.append(it["id"])

    if missing_ids:
        try:
            conn = get_db_connection()
            qmarks = ",".join("?" for _ in missing_ids)
            # ===== ACTUALIZADO: Query incluye los nuevos campos de archivos =====
            rows = conn.execute(
                f"""
                SELECT id, banco, clabe, referencia, descripcion, datos_deposito,
                       anticipo, porcentaje_anticipo, monto_anticipo, monto_restante, tipo_anticipo,
                       tiene_comision, porcentaje_comision, monto_sin_comision, monto_comision, monto,
                       archivo_factura, archivo_recibo, archivo_orden_compra,
                       historial_estados, fecha_aprobado, fecha_liquidado, fecha_ultimo_cambio
                FROM solicitudes WHERE id IN ({qmarks})
                """,
                missing_ids
            ).fetchall()
            # ===== FIN ACTUALIZADO =====
            conn.close()
            extra = {row["id"]: dict(row) for row in rows}

            for it in items:
                ex = extra.get(it["id"])
                if not ex:
                    continue
                # bancarios / descripción
                if _empty(it["banco"]): it["banco"] = ex.get("banco") or ""
                if _empty(it["clabe"]): it["clabe"] = ex.get("clabe") or ""
                if it["referencia"] in (None, ""): it["referencia"] = ex.get("referencia") or ""
                if _empty(it["descripcion"]): it["descripcion"] = ex.get("descripcion") or ""
                if _empty(it["datos_deposito"]): it["datos_deposito"] = ex.get("datos_deposito") or ""

                # ===== NUEVO: Archivos adicionales =====
                if _empty(it["archivo_factura"]): it["archivo_factura"] = ex.get("archivo_factura") or ""
                if _empty(it["archivo_recibo"]): it["archivo_recibo"] = ex.get("archivo_recibo") or ""
                if _empty(it["archivo_orden_compra"]): it["archivo_orden_compra"] = ex.get("archivo_orden_compra") or ""
                # ===== FIN NUEVO =====

                # ===== NUEVO: Campos de historial =====
                if _empty(it.get("historial_estados")): it["historial_estados"] = ex.get("historial_estados") or "[]"
                if _empty(it.get("fecha_aprobado")): it["fecha_aprobado"] = ex.get("fecha_aprobado") or ""
                if _empty(it.get("fecha_liquidado")): it["fecha_liquidado"] = ex.get("fecha_liquidado") or ""
                if _empty(it.get("fecha_ultimo_cambio")): it["fecha_ultimo_cambio"] = ex.get("fecha_ultimo_cambio") or ""
                # ===== FIN NUEVO =====

                # anticipo
                if it["anticipo"] is None: it["anticipo"] = ex.get("anticipo") or "No"
                if it["tipo_anticipo"] is None: it["tipo_anticipo"] = ex.get("tipo_anticipo") or "porcentaje"
                if it["porcentaje_anticipo"] is None: it["porcentaje_anticipo"] = ex.get("porcentaje_anticipo")
                if it["monto_anticipo"] is None: it["monto_anticipo"] = ex.get("monto_anticipo")
                if it["monto_restante"] is None: it["monto_restante"] = ex.get("monto_restante")
                # numéricos a float
                for k in ("porcentaje_anticipo","monto_anticipo","monto_restante"):
                    if it[k] is not None: it[k] = _f(it[k])

                # comisión
                if it["tiene_comision"] in (None, ""): it["tiene_comision"] = _i(ex.get("tiene_comision"))
                if not it["porcentaje_comision"]: it["porcentaje_comision"] = _f(ex.get("porcentaje_comision"))
                if not it["monto_sin_comision"]: it["monto_sin_comision"] = _f(ex.get("monto_sin_comision"))
                if not it["monto_comision"]: it["monto_comision"] = _f(ex.get("monto_comision"))
                if not it["monto"]: it["monto"] = _f(ex.get("monto"))
        except Exception as e:
            print(f"Error al obtener datos faltantes: {e}")
            # asegurar llaves presentes
            for it in items:
                it["banco"] = it.get("banco") or ""
                it["clabe"] = it.get("clabe") or ""
                it["referencia"] = it.get("referencia") or ""
                it["descripcion"] = it.get("descripcion") or ""
                it["datos_deposito"] = it.get("datos_deposito") or ""
                # ===== NUEVO: Asegurar archivos adicionales =====
                it["archivo_factura"] = it.get("archivo_factura") or ""
                it["archivo_recibo"] = it.get("archivo_recibo") or ""
                it["archivo_orden_compra"] = it.get("archivo_orden_compra") or ""
                # ===== FIN NUEVO =====
                # ===== NUEVO: Asegurar campos de historial =====
                it["historial_estados"] = it.get("historial_estados") or "[]"
                it["fecha_aprobado"] = it.get("fecha_aprobado") or ""
                it["fecha_liquidado"] = it.get("fecha_liquidado") or ""
                it["fecha_ultimo_cambio"] = it.get("fecha_ultimo_cambio") or ""
                # ===== FIN NUEVO =====
                it["anticipo"] = it.get("anticipo") or "No"
                it["tipo_anticipo"] = it.get("tipo_anticipo") or "porcentaje"
                for k in ("porcentaje_anticipo","monto_anticipo","monto_restante",
                          "porcentaje_comision","monto_sin_comision","monto_comision","monto"):
                    it[k] = _f(it.get(k))

    return render_template(
        "admin_dashboard.html",
        solicitudes=items,
        estado_filtro=estado_filtro,
        busqueda=busqueda,
        criterio_busqueda=criterio_busqueda,
        capital_total=capital_total,
        page=data["page"],
        page_size=data["page_size"],
        total_pages=data["total_pages"],
        total_count=data["total_count"],
        counts=data["counts"]
    )

@app.route("/admin/solicitudes.json")
def admin_solicitudes_json():
    if not session.get("admin_logged_in"):
        return jsonify({"error": "No autorizado"}), 403

    estado_filtro = (request.args.get("estado") or "").strip()
    busqueda = (request.args.get("busqueda") or "").strip()
    criterio = (request.args.get("criterio") or "todos").strip().lower()

    try:
        page = int(request.args.get("p", 1) or 1)
    except ValueError:
        page = 1

    page_size_default = int(globals().get("PAGE_SIZE_DEFAULT", 15))
    hard_max = 100
    try:
        page_size = int(request.args.get("page_size", page_size_default))
    except ValueError:
        page_size = page_size_default
    page = max(1, page)
    page_size = max(1, min(page_size, hard_max))

    data = query_solicitudes_paginated(
        page=page,
        page_size=page_size,
        estado_filtro=estado_filtro,
        busqueda=busqueda,
        criterio=criterio
    )

    def _empty(v):
        return v is None or (isinstance(v, str) and v.strip() == "")

    items = []
    missing_ids = []

    for r in data["items"]:
        row = dict(r)

        item = {
            "id": row["id"],
            "fp": row.get("fp", ""),
            "nombre": row.get("nombre", ""),
            "destinatario": row.get("destinatario", ""),
            "correo": row.get("correo", ""),
            "departamento": row.get("departamento", ""),
            "tipo_solicitud": row.get("tipo_solicitud", ""),
            "tipo_pago": row.get("tipo_pago", ""),
            "monto": row.get("monto", 0),
            "estado": row.get("estado", ""),
            "fecha": row.get("fecha", ""),
            "fecha_limite": row.get("fecha_limite", ""),
            "archivo_adjunto": row.get("archivo_adjunto", ""),

            "banco": row.get("banco") or "",
            "clabe": row.get("clabe") or "",
            "referencia": row.get("referencia") or "",
            "descripcion": row.get("descripcion") or "",
            "datos_deposito": row.get("datos_deposito") or "",

            # anticipo
            "anticipo": row.get("anticipo"),
            "porcentaje_anticipo": row.get("porcentaje_anticipo"),
            "monto_anticipo": row.get("monto_anticipo"),
            "monto_restante": row.get("monto_restante"),
            "tipo_anticipo": row.get("tipo_anticipo"),

            # comisión
            "tiene_comision": row.get("tiene_comision"),
            "porcentaje_comision": row.get("porcentaje_comision"),
            "monto_sin_comision": row.get("monto_sin_comision"),
            "monto_comision": row.get("monto_comision"),
        }
        items.append(item)

        if (_empty(item["banco"]) or _empty(item["clabe"]) or item["referencia"] is None
                or _empty(item["descripcion"]) or _empty(item["datos_deposito"])
                or item["anticipo"] is None or item["porcentaje_anticipo"] is None
                or item["monto_anticipo"] is None or item["monto_restante"] is None
                or item["tipo_anticipo"] is None):
            missing_ids.append(item["id"])

    if missing_ids:
        try:
            conn = get_db_connection()
            qmarks = ",".join("?" for _ in missing_ids)
            rows = conn.execute(
                f"""
                SELECT id, banco, clabe, referencia, descripcion, datos_deposito,
                       anticipo, porcentaje_anticipo, monto_anticipo, monto_restante, tipo_anticipo,
                       tiene_comision, porcentaje_comision, monto_sin_comision, monto_comision, monto
                FROM solicitudes WHERE id IN ({qmarks})
                """,
                missing_ids
            ).fetchall()
            conn.close()
            extra = {row["id"]: dict(row) for row in rows}

            for it in items:
                ex = extra.get(it["id"])
                if not ex:
                    continue
                for k in ("banco","clabe","referencia","descripcion","datos_deposito",
                          "anticipo","porcentaje_anticipo","monto_anticipo","monto_restante","tipo_anticipo",
                          "tiene_comision","porcentaje_comision","monto_sin_comision","monto_comision","monto"):
                    if it.get(k) in (None, ""):
                        it[k] = ex.get(k)
        except Exception:
            pass

    return jsonify({
        "items": items,
        "page": data.get("page", page),
        "page_size": data.get("page_size", page_size),
        "total_pages": data.get("total_pages", 1),
        "total_count": data.get("total_count", len(items)),
        "counts": data.get("counts", {})
    })


@app.route("/ver_historial")
def ver_historial():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    # Obtener el criterio de búsqueda y el término a buscar (por defecto "todos")
    criterio = request.args.get("criterio", "todos").strip().lower()
    busqueda = request.args.get("busqueda", "").strip().lower()

    conn = get_db_connection()
    solicitudes_db = conn.execute("SELECT * FROM solicitudes ORDER BY fecha DESC").fetchall()
    conn.close()
    solicitudes_list = [dict(row) for row in solicitudes_db]

    # Si se ingresó un término de búsqueda, filtrar según el criterio
    if busqueda:
        if criterio == "fp":
            solicitudes_list = [s for s in solicitudes_list if busqueda in s["fp"].strip().lower()]
        elif criterio == "monto":
            # Convertir el monto a cadena para buscar la subcadena
            solicitudes_list = [s for s in solicitudes_list if busqueda in str(s["monto"]).strip().lower()]
        elif criterio == "nombre":
            solicitudes_list = [s for s in solicitudes_list if busqueda in s["nombre"].strip().lower()]
        else:  # Buscar en FP, Nombre y Monto
            solicitudes_list = [s for s in solicitudes_list if
                                busqueda in s["fp"].strip().lower() or
                                busqueda in s["nombre"].strip().lower() or
                                busqueda in str(s["monto"]).strip().lower()]

    return render_template("admin_historial.html", solicitudes=solicitudes_list)

@app.route("/solicitudes/<int:sid>.json")
def solicitud_json(sid):
    if not session.get("admin_logged_in"):
        return jsonify({"success": False, "error": "No autorizado"}), 403
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM solicitudes WHERE id=?", (sid,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"success": False, "error": "No encontrado"}), 404
    data = dict(row)
    # Normalizamos historial a lista
    try:
        data["historial_estados"] = json.loads(data.get("historial_estados") or "[]")
    except Exception:
        data["historial_estados"] = []
    return jsonify({"success": True, "solicitud": data})


@app.route("/actualizar_estado/<int:solicitud_id>", methods=["POST"])
def actualizar_estado(solicitud_id):
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    nuevo_estado = request.form.get("nuevo_estado", "Pendiente")
    fecha_cambio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db_connection()

    # Obtener el estado actual y el historial
    solicitud_actual = conn.execute(
        "SELECT estado, historial_estados FROM solicitudes WHERE id = ?",
        (solicitud_id,)
    ).fetchone()

    if not solicitud_actual:
        conn.close()
        flash("Solicitud no encontrada.", "error")
        return redirect(url_for("admin_dashboard"))

    estado_anterior = solicitud_actual['estado']

    # Parsear historial existente
    try:
        historial = json.loads(solicitud_actual['historial_estados'] or '[]')
    except:
        historial = []

    # Agregar nuevo registro al historial
    historial.append({
        'estado_anterior': estado_anterior,
        'estado_nuevo': nuevo_estado.capitalize(),
        'fecha': fecha_cambio,
        'usuario': session.get('username', session.get('role', 'admin'))
    })

    # Construir la consulta de actualización base
    update_query = "UPDATE solicitudes SET estado = ?, fecha_ultimo_cambio = ?, historial_estados = ?"
    params = [nuevo_estado.capitalize(), fecha_cambio, json.dumps(historial)]

    # Registrar fecha específica según el tipo de cambio
    if nuevo_estado.lower() in ["aprobado", "aprobado con anticipo"]:
        update_query += ", fecha_aprobado = ?"
        params.append(fecha_cambio)
    elif nuevo_estado.lower() in ["liquidado", "liquidado con anticipo", "liquidacion total"]:
        update_query += ", fecha_liquidado = ?"
        params.append(fecha_cambio)

    # Completar la consulta
    update_query += " WHERE id = ?"
    params.append(solicitud_id)

    # Ejecutar actualización
    conn.execute(update_query, params)
    conn.commit()

    # Obtener solicitud actualizada para emails
    solicitud = conn.execute("SELECT * FROM solicitudes WHERE id = ?", (solicitud_id,)).fetchone()
    conn.close()

    flash(f"Solicitud {solicitud_id} actualizada a {nuevo_estado}.", "success")

    # Convertir a diccionario para los emails
    solicitud_dict = dict(solicitud)

    # Enviar emails según el nuevo estado
    if nuevo_estado.lower() == "aprobado":
        send_approval_email(solicitud_dict)
    elif nuevo_estado.lower() == "aprobado con anticipo":
        send_approval_anticipo_email(solicitud_dict)
    elif nuevo_estado.lower() == "declinada":
        send_declined_email(solicitud_dict)
    elif nuevo_estado.lower() == "liquidado":
        attachment = request.files.get("liquidado_file")
        send_liquidado_email(solicitud_dict, attachment)
    elif nuevo_estado.lower() == "liquidado con anticipo":
        attachment = request.files.get("liquidado_file")
        send_liquidado_anticipo_email(solicitud_dict, attachment)
    elif nuevo_estado.lower() == "liquidacion total":
        attachment = request.files.get("liquidado_file")
        send_liquidacion_total_email(solicitud_dict, attachment)

    # Sincronizar con base de datos remota si está habilitado
    try:
        threading.Thread(target=sync_solicitudes_to_remote, daemon=True).start()
    except Exception as e:
        print(f"Error al iniciar sincronización: {e}")

    return redirect(url_for("admin_dashboard"))

@app.route("/eliminar_solicitud/<int:solicitud_id>", methods=["POST"])
def eliminar_solicitud(solicitud_id):
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))
    conn = get_db_connection()
    conn.execute("DELETE FROM solicitudes WHERE id = ?", (solicitud_id,))
    conn.commit()
    conn.close()
    flash(f"Solicitud {solicitud_id} eliminada.", "success")
    return redirect(url_for("admin_dashboard"))

@app.route("/exportar_solicitudes")
def exportar_solicitudes():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))
    conn = get_db_connection()
    solicitudes_db = conn.execute("SELECT * FROM solicitudes").fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["FP", "Nombre", "Destinatario", "Correo", "Departamento", "Tipo Solicitud", "Tipo Pago", "Descripción",
                     "Datos de Depósito", "Banco", "CLABE", "Monto", "Anticipo", "Porcentaje Anticipo", "Monto Restante", "Estado", "Fecha", "Fecha Límite"])
    for sol in solicitudes_db:
        writer.writerow([
            sol["fp"], sol["nombre"], sol["destinatario"], sol["correo"], sol["departamento"], sol["tipo_solicitud"],
            sol["tipo_pago"], sol["descripcion"], sol["datos_deposito"], sol["banco"], sol["clabe"],
            sol["monto"], sol["anticipo"], sol["porcentaje_anticipo"], sol["monto_restante"], sol["estado"], sol["fecha"], sol["fecha_limite"]
        ])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode("utf-8")),
                     mimetype="text/csv",
                     as_attachment=True,
                     attachment_filename="solicitudes.csv")




@app.route("/estadisticas")
def estadisticas():
    """Dashboard ejecutivo de estadísticas con gráficas interactivas"""

    # --- Permisos ---
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a las estadísticas.", "error")
        return redirect(url_for("admin_dashboard"))

    # --- Parámetros de filtrado ---
    periodo = request.args.get('periodo', 'mensual')
    proveedor_selected = request.args.get('proveedor', '').strip()
    view = request.args.get('view', '')

    # --- Obtener datos base ---
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM solicitudes ORDER BY fecha DESC").fetchall()
    conn.close()

    if not rows:
        return render_template(
            'estadisticas.html',
            capital_total=capital_total,
            total_solicitudes=0,
            remaining=capital_total,
            summary_data=[],
            summary_by_type=[],
            proveedores=[],
            current_year=datetime.now().year,
            periodo=periodo,
            view=view,
            proveedor_selected=proveedor_selected,
            provider_total=0,
            provider_summary_by_type=[],
            provider_records=[],
            period_options=[],
            fecha_periodo_selected='',
            periodo_proveedor='total',
            # Datos adicionales para gráficos
            estados_data={},
            departamentos_data={},
            tipo_pago_data={},
            top_proveedores=[],
            timeline_data=[]
        )

    data = [dict(r) for r in rows]
    df = pd.DataFrame(data)

    # --- Conversiones de tipos ---
    df['monto'] = pd.to_numeric(df.get('monto', 0), errors='coerce').fillna(0)
    df['fecha'] = pd.to_datetime(df.get('fecha'), format="%Y-%m-%d %H:%M:%S", errors='coerce')
    df['fecha_limite_dt'] = pd.to_datetime(df.get('fecha_limite'), format="%Y-%m-%d", errors='coerce')

    # --- Cálculos globales ---
    total_solicitudes = df['monto'].sum()
    remaining = capital_total - total_solicitudes

    # --- Lista de proveedores únicos ---
    nombres_prov = sorted(df['destinatario'].dropna().unique())
    proveedores = [{"nombre": n} for n in nombres_prov]

    # --- Vista por proveedor (si aplica) ---
    provider_total = None
    provider_summary_by_type = []
    provider_records = []
    period_options = []
    fecha_periodo_selected = ''
    periodo_proveedor = request.args.get('periodo_proveedor', 'total')

    if view == 'proveedor' and proveedor_selected:
        df_prov = df[df['destinatario'] == proveedor_selected].copy()

        # Opciones de período
        if periodo_proveedor == 'diario':
            period_options = sorted(df_prov['fecha'].dt.date.astype(str).unique())
        elif periodo_proveedor == 'semanal':
            period_options = sorted(df_prov['fecha'].dt.strftime('%Y-W%U').unique())
        elif periodo_proveedor == 'mensual':
            period_options = sorted(df_prov['fecha'].dt.to_period('M').astype(str).unique())
        elif periodo_proveedor == 'anual':
            period_options = sorted(df_prov['fecha'].dt.year.astype(str).unique())
        else:
            period_options = ['total']

        # Filtrar por fecha si se seleccionó
        fecha_periodo = request.args.get('fecha_periodo', '').strip()
        fecha_periodo_selected = fecha_periodo

        df_filt = df_prov
        if periodo_proveedor != 'total' and fecha_periodo:
            if periodo_proveedor == 'diario':
                df_filt = df_prov[df_prov['fecha'].dt.date.astype(str) == fecha_periodo]
            elif periodo_proveedor == 'semanal':
                df_filt = df_prov[df_prov['fecha'].dt.strftime('%Y-W%U') == fecha_periodo]
            elif periodo_proveedor == 'mensual':
                df_filt = df_prov[df_prov['fecha'].dt.to_period('M').astype(str) == fecha_periodo]
            elif periodo_proveedor == 'anual':
                df_filt = df_prov[df_prov['fecha'].dt.year.astype(str) == fecha_periodo]

        # Cálculos para el proveedor
        provider_total = df_filt['monto'].sum()

        grp = (
            df_filt
            .groupby('tipo_solicitud')['monto']
            .sum()
            .reset_index()
            .rename(columns={'tipo_solicitud': 'Tipo de Solicitud', 'monto': 'Total'})
        )
        provider_summary_by_type = grp.to_dict(orient='records')

        # Historial
        provider_records = (
            df_filt[['fecha', 'tipo_solicitud', 'monto', 'descripcion']]
            .sort_values('fecha', ascending=False)
            .to_dict(orient='records')
        )

    # --- Agrupación global por período ---
    if periodo == 'diario':
        df_g = df.groupby(df['fecha'].dt.date).agg({
            'monto': 'sum',
            'id': 'count'
        }).reset_index()
        df_g.columns = ['Periodo', 'Total', 'Cantidad']
    elif periodo == 'semanal':
        df_g = df.groupby(df['fecha'].dt.strftime('%Y-W%U')).agg({
            'monto': 'sum',
            'id': 'count'
        }).reset_index()
        df_g.columns = ['Periodo', 'Total', 'Cantidad']
    elif periodo == 'mensual':
        df_g = df.groupby(df['fecha'].dt.to_period('M').astype(str)).agg({
            'monto': 'sum',
            'id': 'count'
        }).reset_index()
        df_g.columns = ['Periodo', 'Total', 'Cantidad']
    elif periodo == 'anual':
        df_g = df.groupby(df['fecha'].dt.year).agg({
            'monto': 'sum',
            'id': 'count'
        }).reset_index()
        df_g.columns = ['Periodo', 'Total', 'Cantidad']
    else:
        df_g = df.groupby(df['fecha'].dt.date).agg({
            'monto': 'sum',
            'id': 'count'
        }).reset_index()
        df_g.columns = ['Periodo', 'Total', 'Cantidad']

    summary_data = df_g.to_dict(orient='records')

    # --- Resumen por tipo de solicitud ---
    df_tipo = df.groupby('tipo_solicitud').agg({
        'monto': 'sum',
        'id': 'count'
    }).reset_index()
    df_tipo.columns = ['Tipo de Solicitud', 'Total', 'Cantidad']
    summary_by_type = df_tipo.to_dict(orient='records')

    # --- DATOS ADICIONALES PARA GRÁFICOS ---

    # 1. Estados (para gráfico polar)
    estados_count = df.groupby('estado').agg({
        'monto': 'sum',
        'id': 'count'
    }).reset_index()
    estados_data = {
        'labels': estados_count['estado'].tolist(),
        'totales': estados_count['monto'].tolist(),
        'cantidades': estados_count['id'].tolist()
    }

    # 2. Departamentos (para gráfico de barras)
    if 'departamento' in df.columns:
        dept_data = df.groupby('departamento')['monto'].sum().reset_index()
        dept_data = dept_data.sort_values('monto', ascending=False).head(10)
        departamentos_data = {
            'labels': dept_data['departamento'].tolist(),
            'totales': dept_data['monto'].tolist()
        }
    else:
        departamentos_data = {'labels': [], 'totales': []}

    # 3. Tipo de Pago (para gráfico radar)
    if 'tipo_pago' in df.columns:
        tipo_pago_count = df.groupby('tipo_pago').agg({
            'monto': 'sum',
            'id': 'count'
        }).reset_index()
        tipo_pago_data = {
            'labels': tipo_pago_count['tipo_pago'].tolist(),
            'totales': tipo_pago_count['monto'].tolist(),
            'cantidades': tipo_pago_count['id'].tolist()
        }
    else:
        tipo_pago_data = {'labels': [], 'totales': [], 'cantidades': []}

    # 4. Top 10 Proveedores
    if 'destinatario' in df.columns:
        top_prov = df.groupby('destinatario')['monto'].sum().reset_index()
        top_prov = top_prov.sort_values('monto', ascending=False).head(10)
        top_proveedores = {
            'labels': top_prov['destinatario'].tolist(),
            'totales': top_prov['monto'].tolist()
        }
    else:
        top_proveedores = {'labels': [], 'totales': []}

    # 5. Timeline (últimos 12 meses con estados)
    now = datetime.now()
    last_12_months = pd.date_range(
        end=now,
        periods=12,
        freq='MS'  # Month Start
    )

    timeline_data = []
    for month in last_12_months:
        month_str = month.strftime('%Y-%m')
        df_month = df[df['fecha'].dt.to_period('M').astype(str) == month_str]

        liquidado = df_month[df_month['estado'].str.lower().str.contains('liquidado', na=False)]['monto'].sum()
        pendiente = df_month[df_month['estado'].str.lower() == 'pendiente']['monto'].sum()
        aprobado = df_month[df_month['estado'].str.lower().str.contains('aprobado', na=False)]['monto'].sum()

        timeline_data.append({
            'periodo': month.strftime('%b %Y'),
            'liquidado': float(liquidado),
            'pendiente': float(pendiente),
            'aprobado': float(aprobado)
        })

    # --- Renderizar template ---
    return render_template(
        'estadisticas.html',
        # Datos básicos
        capital_total=capital_total,
        total_solicitudes=total_solicitudes,
        remaining=remaining,
        summary_data=summary_data,
        summary_by_type=summary_by_type,
        proveedores=proveedores,
        current_year=datetime.now().year,

        # Filtros
        periodo=periodo,
        view=view,
        proveedor_selected=proveedor_selected,
        periodo_proveedor=periodo_proveedor,
        fecha_periodo_selected=fecha_periodo_selected,
        period_options=period_options,

        # Datos de proveedor
        provider_total=provider_total or 0,
        provider_summary_by_type=provider_summary_by_type,
        provider_records=provider_records,

        # Datos para gráficos
        estados_data=estados_data,
        departamentos_data=departamentos_data,
        tipo_pago_data=tipo_pago_data,
        top_proveedores=top_proveedores,
        timeline_data=timeline_data
    )



@app.route("/calendario")
def calendario():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))
    conn = get_db_connection()
    events_db = conn.execute("SELECT * FROM solicitudes WHERE LOWER(estado) LIKE 'aprobado%'").fetchall()
    conn.close()
    events = [dict(row) for row in events_db]
    fullcalendar_events = []
    for event in events:
        fullcalendar_events.append({
            "title": f"FP: {event['fp']} - {event['destinatario']}",
            "start": event['fecha_limite'],
            "extendedProps": {
                "correo": event["correo"],
                "departamento": event["departamento"],
                "tipo_solicitud": event["tipo_solicitud"],
                "tipo_pago": event["tipo_pago"],
                "descripcion": event["descripcion"],
                "datos_deposito": event["datos_deposito"],
                "banco": event["banco"],
                "clabe": event["clabe"],
                "monto": event["monto"],
                "fecha": event["fecha"],
                "estado": event["estado"]
            }
        })
    events_json = json.dumps(fullcalendar_events)
    return render_template("calendario.html", events_json=events_json)


@app.route("/editar_solicitud/<int:solicitud_id>", methods=["GET", "POST"])
def editar_solicitud(solicitud_id):
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    conn = get_db_connection()
    solicitud = conn.execute("SELECT * FROM solicitudes WHERE id = ?", (solicitud_id,)).fetchone()

    if solicitud is None:
        conn.close()
        flash("Solicitud no encontrada.", "error")
        return redirect(url_for("admin_dashboard"))

    solicitud_dict = dict(solicitud)

    # Cargar conceptos indirectos para el selector de categoría administrativa
    conceptos_indirectos = []
    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Costos",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor(dictionary=True)
        cursor.execute("""SELECT regID as id, concepto
                          FROM AD17_Costos.Conceptos_Indirectos
                          WHERE habilitado = 1
                          ORDER BY concepto ASC""")
        conceptos_indirectos = cursor.fetchall()
        cursor.close()
        remote_conn.close()
    except Exception as e:
        print("Error al obtener conceptos indirectos:", e)
        conceptos_indirectos = []

    if request.method == "POST":
        # Obtener datos básicos del formulario
        fp = request.form.get("fp")
        nombre = request.form.get("nombre")
        destinatario = request.form.get("destinatario")

        # Procesar múltiples correos
        correo_list = request.form.getlist("correo")
        correo = ", ".join([c.strip() for c in correo_list if c.strip()])

        departamento = request.form.get("departamento")
        tipo_solicitud = request.form.get("tipo_solicitud")
        tipo_pago = request.form.get("tipo_pago")
        descripcion = request.form.get("descripcion")
        datos_deposito = request.form.get("datos_deposito")
        banco = request.form.get("banco")
        clabe = request.form.get("clabe")
        fecha_limite = request.form.get("fecha_limite")
        estado = request.form.get("estado")
        referencia = request.form.get("referencia", "")
        categoria_administrativa = request.form.get("categoria_administrativa", "")

        # -------- montos y comisión (misma lógica que solicitar_pago) --------
        import re as _re
        def _to_float_edit(s):
            if s is None:
                return 0.0
            s = str(s)
            s = _re.sub(r"[^\d,.\-]", "", s)
            s = s.replace(",", "")
            try:
                return float(s)
            except ValueError:
                return 0.0

        monto_sin_comision = _to_float_edit(request.form.get("monto"))
        tiene_comision = 1 if tipo_pago == "BBVA Sin factura" else 0
        porcentaje_comision = 6.0 if tiene_comision else 0.0
        if tiene_comision:
            # total bruto necesario para que neto sea monto_sin_comision
            monto = round(monto_sin_comision / 0.94, 2)
            monto_comision = round(monto - monto_sin_comision, 2)
        else:
            monto = round(monto_sin_comision, 2)
            monto_comision = 0.0

        # -------- ARCHIVOS ADJUNTOS (3 archivos + legacy) --------
        archivo_adjunto = solicitud_dict.get("archivo_adjunto", "")
        archivo_factura = solicitud_dict.get("archivo_factura", "")
        archivo_recibo = solicitud_dict.get("archivo_recibo", "")
        archivo_orden_compra = solicitud_dict.get("archivo_orden_compra", "")

        # Procesar archivo adjunto principal (legacy)
        adjunto_file = request.files.get("adjunto_file")
        filepath = None
        if adjunto_file and adjunto_file.filename != "":
            filename = secure_filename(adjunto_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            adjunto_file.save(filepath)
            archivo_adjunto = filename

        # Procesar factura
        factura_file = request.files.get("factura_file")
        if factura_file and factura_file.filename != "":
            filename = secure_filename(factura_file.filename)
            factura_filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            factura_file.save(factura_filepath)
            archivo_factura = filename

        # Procesar recibo
        recibo_file = request.files.get("recibo_file")
        if recibo_file and recibo_file.filename != "":
            filename = secure_filename(recibo_file.filename)
            recibo_filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            recibo_file.save(recibo_filepath)
            archivo_recibo = filename

        # Procesar orden de compra
        orden_file = request.files.get("orden_compra_file")
        if orden_file and orden_file.filename != "":
            filename = secure_filename(orden_file.filename)
            orden_filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            orden_file.save(orden_filepath)
            archivo_orden_compra = filename

        # Procesar anticipo mejorado
        anticipo = request.form.get("anticipo")
        tipo_anticipo = request.form.get("tipo_anticipo", "porcentaje")

        if anticipo:
            anticipo_val = "Si"

            if tipo_anticipo == "porcentaje":
                porcentaje_str = request.form.get("porcentaje_anticipo", "0")
                try:
                    porcentaje_anticipo = float(porcentaje_str)
                except ValueError:
                    porcentaje_anticipo = 0.0
                monto_anticipo = monto * (porcentaje_anticipo / 100)
            else:  # tipo_anticipo == "cantidad"
                monto_anticipo_str = request.form.get("monto_anticipo", "0")
                try:
                    monto_anticipo = float(monto_anticipo_str)
                except ValueError:
                    monto_anticipo = 0.0
                # Calcular el porcentaje basado en la cantidad
                if monto > 0:
                    porcentaje_anticipo = (monto_anticipo / monto) * 100
                else:
                    porcentaje_anticipo = 0.0
        else:
            anticipo_val = "No"
            porcentaje_anticipo = 0.0
            monto_anticipo = 0.0
            tipo_anticipo = "porcentaje"

        monto_restante = monto - monto_anticipo

        # Verificar si es programada (mantener el valor si no se envía en el formulario)
        es_programada = solicitud_dict.get("es_programada", 0)

        # LOGGING para debugging
        print(f"DEBUG - Actualización de solicitud {solicitud_id}:")
        print(f"  Tipo pago nuevo: {tipo_pago}")
        print(f"  Tiene comisión: {tiene_comision}")
        print(f"  Monto sin comisión: {monto_sin_comision}")
        print(f"  Monto comisión: {monto_comision}")
        print(f"  Monto total final: {monto}")

        # Actualizar la base de datos con todos los campos nuevos
        conn.execute("""
            UPDATE solicitudes SET
                fp = ?, nombre = ?, destinatario = ?, correo = ?, departamento = ?,
                tipo_solicitud = ?, tipo_pago = ?, descripcion = ?, datos_deposito = ?,
                banco = ?, clabe = ?, referencia = ?, monto = ?, estado = ?, fecha_limite = ?,
                archivo_adjunto = ?, archivo_factura = ?, archivo_recibo = ?, archivo_orden_compra = ?,
                anticipo = ?, porcentaje_anticipo = ?, monto_restante = ?,
                tiene_comision = ?, porcentaje_comision = ?, monto_comision = ?, monto_sin_comision = ?,
                tipo_anticipo = ?, monto_anticipo = ?, categoria_administrativa = ?
            WHERE id = ?
        """, (
            fp, nombre, destinatario, correo, departamento, tipo_solicitud, tipo_pago,
            descripcion, datos_deposito, banco, clabe, referencia, monto, estado, fecha_limite,
            archivo_adjunto, archivo_factura, archivo_recibo, archivo_orden_compra,
            anticipo_val, porcentaje_anticipo, monto_restante,
            tiene_comision, porcentaje_comision, monto_comision, monto_sin_comision,
            tipo_anticipo, monto_anticipo, categoria_administrativa, solicitud_id
        ))

        # Si el estado cambió, actualizar historial
        if estado != solicitud_dict.get("estado"):
            fecha_cambio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Parsear historial existente
            try:
                historial = json.loads(solicitud_dict.get('historial_estados', '[]'))
            except:
                historial = []

            # Agregar nuevo registro al historial
            historial.append({
                'estado_anterior': solicitud_dict.get("estado"),
                'estado_nuevo': estado,
                'fecha': fecha_cambio,
                'usuario': session.get('username', session.get('role', 'admin'))
            })

            # Actualizar campos de fecha según el estado
            update_query_parts = ["fecha_ultimo_cambio = ?", "historial_estados = ?"]
            update_params = [fecha_cambio, json.dumps(historial)]

            if estado.lower() in ["aprobado", "aprobado con anticipo"]:
                update_query_parts.append("fecha_aprobado = ?")
                update_params.append(fecha_cambio)
            elif estado.lower() in ["liquidado", "liquidado con anticipo", "liquidacion total"]:
                update_query_parts.append("fecha_liquidado = ?")
                update_params.append(fecha_cambio)

            update_params.append(solicitud_id)

            conn.execute(f"""
                UPDATE solicitudes
                SET {', '.join(update_query_parts)}
                WHERE id = ?
            """, update_params)

        conn.commit()
        conn.close()

        # Si el estado es Liquidado y se adjuntó un archivo, enviar correo
        if estado.lower() in ["liquidado", "liquidado con anticipo", "liquidacion total"] and filepath:
            solicitud_updated = {
                "fp": fp,
                "nombre": nombre,
                "destinatario": destinatario,
                "correo": correo,
                "departamento": departamento,
                "tipo_solicitud": tipo_solicitud,
                "tipo_pago": tipo_pago,
                "descripcion": descripcion,
                "datos_deposito": datos_deposito,
                "banco": banco,
                "clabe": clabe,
                "monto": monto,
                "monto_sin_comision": monto_sin_comision,
                "tiene_comision": tiene_comision,
                "porcentaje_comision": porcentaje_comision,
                "monto_comision": monto_comision,
                "estado": estado,
                "fecha_limite": fecha_limite,
                "archivo_adjunto": archivo_adjunto,
                "fecha": solicitud_dict.get("fecha"),
                "anticipo": anticipo_val,
                "porcentaje_anticipo": porcentaje_anticipo,
                "monto_anticipo": monto_anticipo,
                "monto_restante": monto_restante,
                "tipo_anticipo": tipo_anticipo
            }

            with open(filepath, "rb") as f:
                if estado.lower() == "liquidacion total":
                    send_liquidacion_total_email(solicitud_updated, attachment_file=f)
                elif anticipo_val == "Si":
                    send_liquidado_anticipo_email(solicitud_updated, attachment_file=f)
                else:
                    send_liquidado_email(solicitud_updated, attachment_file=f)

        # Enviar otros correos según el estado
        elif estado.lower() == "aprobado":
            solicitud_updated = {
                "fp": fp,
                "nombre": nombre,
                "destinatario": destinatario,
                "correo": correo,
                "departamento": departamento,
                "tipo_solicitud": tipo_solicitud,
                "tipo_pago": tipo_pago,
                "descripcion": descripcion,
                "banco": banco,
                "clabe": clabe,
                "monto": monto,
                "fecha": solicitud_dict.get("fecha"),
                "fecha_limite": fecha_limite
            }
            send_approval_email(solicitud_updated)

        elif estado.lower() == "aprobado con anticipo":
            solicitud_updated = {
                "fp": fp,
                "nombre": nombre,
                "destinatario": destinatario,
                "correo": correo,
                "porcentaje_anticipo": porcentaje_anticipo,
                "monto": monto,
                "monto_restante": monto_restante,
                "fecha": solicitud_dict.get("fecha"),
                "fecha_limite": fecha_limite,
                "banco": banco,
                "clabe": clabe,
                "destinatario": destinatario
            }
            send_approval_anticipo_email(solicitud_updated)

        elif estado.lower() == "declinada":
            solicitud_updated = {
                "fp": fp,
                "nombre": nombre,
                "correo": correo,
                "tipo_solicitud": tipo_solicitud,
                "tipo_pago": tipo_pago,
                "monto": monto,
                "fecha": solicitud_dict.get("fecha"),
                "fecha_limite": fecha_limite
            }
            send_declined_email(solicitud_updated)

        # Sincronizar con base de datos remota
        try:
            threading.Thread(target=sync_solicitudes_to_remote, daemon=True).start()
        except Exception as e:
            print(f"Error al iniciar sincronización: {e}")

        flash("Registro actualizado correctamente.", "success")
        return redirect(url_for("admin_dashboard"))

    conn.close()
    return render_template("editar_solicitud.html", solicitud=solicitud_dict, conceptos_indirectos=conceptos_indirectos)

@app.route("/descargar_adjunto/<filename>")
def descargar_adjunto(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename, as_attachment=True)
@app.route("/exportar_reporte_semanal")
def exportar_reporte_semanal():
    # — Verifica permisos —
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para exportar reportes.", "error")
        return redirect(url_for("admin_dashboard"))

    # — Determinar la fecha de inicio de la semana —
    fecha_inicio_str = request.args.get('fecha_inicio')

    if fecha_inicio_str:
        try:
            # Convertir el string de fecha a objeto datetime
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
            # Asegurarse que sea un lunes (día de semana = 0)
            if fecha_inicio.weekday() != 0:  # 0 = lunes en datetime.weekday()
                # Ajustar al lunes de esa semana
                fecha_inicio = fecha_inicio - timedelta(days=fecha_inicio.weekday())
        except ValueError:
            flash("Formato de fecha inválido. Usando semana actual.", "warning")
            # Usar la semana actual
            hoy = datetime.now()
            fecha_inicio = hoy - timedelta(days=hoy.weekday())
    else:
        # Si no se proporciona fecha, usar la semana actual
        hoy = datetime.now()
        fecha_inicio = hoy - timedelta(days=hoy.weekday())

    # Establece el fin de la semana (domingo 23:59:59)
    fecha_fin = fecha_inicio + timedelta(days=6, hours=23, minutes=59, seconds=59)

    # — Configura el título del reporte —
    fecha_inicio_str = fecha_inicio.strftime("%d-%m-%Y")
    fecha_fin_str = fecha_fin.strftime("%d-%m-%Y")
    titulo_reporte = f"Reporte Semanal: {fecha_inicio_str} al {fecha_fin_str}"

    try:
        # — Carga y prepara DataFrame —
        conn = get_db_connection()
        # Consulta explícitamente todas las columnas disponibles
        query = """
            SELECT id, fecha, monto, tipo_solicitud, descripcion, estado, fecha_limite,
                   destinatario, fp, nombre, correo, departamento, tipo_pago,
                   datos_deposito, banco, clabe, archivo_adjunto, anticipo,
                   porcentaje_anticipo, monto_restante
            FROM solicitudes
        """
        rows = conn.execute(query).fetchall()
        conn.close()

        if not rows:
            flash("No hay datos para generar el reporte.", "warning")
            return redirect(url_for("estadisticas"))

        df = pd.DataFrame([dict(r) for r in rows])

        # — Convertir columnas a tipos adecuados —
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        # Reemplaza la línea problemática por:
        df['fecha_limite_dt'] = pd.to_datetime(df['fecha_limite'], format="%Y-%m-%d", errors='coerce')

        df['monto'] = pd.to_numeric(df['monto'], errors='coerce').fillna(0)

        # — Asegura que todas las columnas necesarias existan —
        required_columns = ['destinatario', 'monto', 'fp', 'descripcion', 'fecha_limite', 'estado', 'tipo_solicitud',
                           'nombre', 'departamento', 'tipo_pago', 'datos_deposito', 'banco', 'anticipo', 'monto_restante']
        for col in required_columns:
            if col not in df.columns:
                df[col] = None

        # — Filtrar por fecha límite en la semana seleccionada —
        mask = (df['fecha_limite'] >= fecha_inicio) & (df['fecha_limite'] <= fecha_fin)
        df_sem = df.loc[mask].copy()

        if df_sem.empty:
            flash(f"No hay datos para la semana del {fecha_inicio_str} al {fecha_fin_str}.", "warning")
            return redirect(url_for("estadisticas"))

        # — 1) Historial de liquidados (por fecha_limite) —
        df_liq = df_sem[df_sem['estado'].str.lower() == 'liquidado'].copy() if not df_sem.empty else pd.DataFrame()
        if not df_liq.empty:
            df_liq = df_liq[['destinatario', 'monto', 'fp', 'descripcion', 'fecha_limite', 'tipo_solicitud',
                            'departamento', 'tipo_pago', 'banco', 'anticipo', 'monto_restante']]
            df_liq = df_liq.rename(columns={
                'destinatario': 'Proveedor',
                'monto': 'Monto Total',
                'fp': 'FP',
                'descripcion': 'Detalle',
                'fecha_limite': 'Fecha Límite',
                'tipo_solicitud': 'Tipo de Solicitud',
                'departamento': 'Departamento',
                'tipo_pago': 'Tipo de Pago',
                'banco': 'Banco',
                'anticipo': 'Anticipo',
                'monto_restante': 'Monto Restante'
            })
            # Formatear la fecha
            df_liq['Fecha Límite'] = df_liq['Fecha Límite'].dt.strftime('%d-%m-%Y')

        # — 2) Pagos pendientes (fecha límite esta semana pero no liquidado) —
        df_pen = df_sem[~df_sem['estado'].str.lower().eq('liquidado')].copy() if not df_sem.empty else pd.DataFrame()
        if not df_pen.empty:
            df_pen = df_pen[['destinatario', 'monto', 'fp', 'descripcion', 'fecha_limite', 'estado', 'tipo_solicitud',
                           'departamento', 'tipo_pago', 'banco', 'anticipo', 'monto_restante']]
            df_pen = df_pen.rename(columns={
                'destinatario': 'Proveedor',
                'monto': 'Monto Total',
                'fp': 'FP',
                'descripcion': 'Detalle',
                'fecha_limite': 'Fecha Límite',
                'estado': 'Estado',
                'tipo_solicitud': 'Tipo de Solicitud',
                'departamento': 'Departamento',
                'tipo_pago': 'Tipo de Pago',
                'banco': 'Banco',
                'anticipo': 'Anticipo',
                'monto_restante': 'Monto Restante'
            })
            # Formatear la fecha
            df_pen['Fecha Límite'] = df_pen['Fecha Límite'].dt.strftime('%d-%m-%Y')

        # — 3) Resumen por tipo de solicitud (de todos los de la semana) —
        if not df_sem.empty:
            # Asegurarse de que hay valores en tipo_solicitud para mejor visualización
            df_sem['tipo_solicitud'] = df_sem['tipo_solicitud'].fillna('No especificado')

            # Primero un resumen por proveedor
            df_prov = df_sem.groupby('destinatario').agg({
                'monto': 'sum',
                'anticipo': lambda x: x.sum() if 'anticipo' in df_sem.columns else 0,
                'monto_restante': lambda x: x.sum() if 'monto_restante' in df_sem.columns else 0
            }).reset_index()

            df_prov = df_prov.rename(columns={
                'destinatario': 'Proveedor',
                'monto': 'Total',
                'anticipo': 'Anticipo Total',
                'monto_restante': 'Restante Total'
            }).sort_values('Total', ascending=False)

            # Luego un resumen por tipo de solicitud - asegurar que todos los tipos se muestren
            df_res = df_sem.groupby('tipo_solicitud').agg({
                'monto': 'sum',
                'anticipo': lambda x: x.sum() if 'anticipo' in df_sem.columns else 0,
                'monto_restante': lambda x: x.sum() if 'monto_restante' in df_sem.columns else 0,
                'id': 'count'  # Contar registros
            }).reset_index()

            df_res = df_res.rename(columns={
                'tipo_solicitud': 'Tipo de Solicitud',
                'monto': 'Total',
                'anticipo': 'Anticipo Total',
                'monto_restante': 'Restante Total',
                'id': 'Cantidad'
            }).sort_values('Total', ascending=False)

            # Resumen por tipo de pago - IMPORTANTE: asegurar que todos los tipos se muestran
            if 'tipo_pago' in df_sem.columns:
                df_sem['tipo_pago'] = df_sem['tipo_pago'].fillna('No especificado')
                df_tipo_pago = df_sem.groupby('tipo_pago')['monto'].agg(['sum', 'count']).reset_index()
                df_tipo_pago = df_tipo_pago.rename(columns={
                    'tipo_pago': 'Tipo de Pago',
                    'sum': 'Total',
                    'count': 'Cantidad'
                }).sort_values('Total', ascending=False)
            else:
                df_tipo_pago = pd.DataFrame()

            # Resumen por departamento
            if 'departamento' in df_sem.columns:
                df_sem['departamento'] = df_sem['departamento'].fillna('No especificado')
                df_dept = df_sem.groupby('departamento').agg({
                    'monto': 'sum',
                    'id': 'count'
                }).reset_index()

                df_dept = df_dept.rename(columns={
                    'departamento': 'Departamento',
                    'monto': 'Total',
                    'id': 'Cantidad'
                }).sort_values('Total', ascending=False)
            else:
                df_dept = pd.DataFrame()

            # Resumen de estado
            df_estado = df_sem.groupby('estado')['monto'].agg(['sum', 'count']).reset_index()
            df_estado = df_estado.rename(columns={
                'estado': 'Estado',
                'sum': 'Total',
                'count': 'Cantidad'
            }).sort_values('Total', ascending=False)
        else:
            df_prov = pd.DataFrame()
            df_res = pd.DataFrame()
            df_estado = pd.DataFrame()
            df_dept = pd.DataFrame()
            df_tipo_pago = pd.DataFrame()

        # — Genera Excel con formato profesional —
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Formatos comunes mejorados
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter',
                'font_color': '#FFFFFF',
                'bg_color': '#FF9800',
                'border': 1,
                'border_color': '#E65100'
            })

            subtitle_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'align': 'left',
                'valign': 'vcenter',
                'font_color': '#FFFFFF',
                'bg_color': '#FF9800',
                'border': 1,
                'border_color': '#E65100'
            })

            header_format = workbook.add_format({
                'bold': True,
                'font_color': '#FFFFFF',
                'bg_color': '#FF9800',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True
            })

            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'top',
                'text_wrap': True
            })

            money_format = workbook.add_format({
                'num_format': '$#,##0.00',
                'border': 1,
                'align': 'right',
                'valign': 'top'
            })

            date_format = workbook.add_format({
                'num_format': 'dd/mm/yyyy',
                'border': 1,
                'align': 'center',
                'valign': 'top'
            })

            total_format = workbook.add_format({
                'bold': True,
                'num_format': '$#,##0.00',
                'border': 1,
                'align': 'right',
                'valign': 'top',
                'bg_color': '#FFF3E0',
                'font_color': '#E65100'
            })

            count_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'right',
                'valign': 'top',
                'bg_color': '#FFF3E0'
            })

            total_row_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'bg_color': '#FFF3E0',
                'font_color': '#E65100'
            })

            # Función para aplicar formatos comunes y mejorados a todas las hojas
            def format_sheet(worksheet, df, start_row=3, has_totals=False):
                # Si no hay datos, mostrar mensaje
                if df.empty:
                    worksheet.write(start_row, 0, "No hay datos disponibles para este período",
                                   workbook.add_format({'italic': True, 'align': 'center'}))
                    worksheet.merge_range(start_row, 0, start_row, 5, "No hay datos disponibles para este período",
                                          workbook.add_format({'italic': True, 'align': 'center'}))
                    return start_row + 2

                # Escribir encabezados
                for col_num, col_name in enumerate(df.columns):
                    worksheet.write(start_row, col_num, col_name, header_format)

                # Aplicar formato a cada celda según el tipo de dato
                for row_idx, row in df.iterrows():
                    for col_idx, col_name in enumerate(df.columns):
                        value = row[col_name]

                        # Aplicar formato según el tipo de columna
                        if 'monto' in col_name.lower() or 'total' in col_name.lower():
                            fmt = money_format
                        elif 'fecha' in col_name.lower():
                            fmt = date_format
                        else:
                            fmt = cell_format

                        worksheet.write(start_row + 1 + row_idx, col_idx, value, fmt)

                # Ajustar anchos para mejor visualización
                for i, col in enumerate(df.columns):
                    # Determinar el ancho óptimo para la columna
                    header_len = len(col)
                    max_data_len = df[col].astype(str).str.len().max() if not df.empty else 0

                    # Determinar ancho de columna con límites razonables
                    if col.lower() in ['descripcion', 'detalle']:
                        # Para descripciones largas, limitar para que no se expandan demasiado
                        col_width = min(60, max(25, max_data_len + 2))
                    elif 'proveedor' in col.lower() or 'destinatario' in col.lower():
                        # Para nombres de proveedores, dar suficiente espacio pero no excesivo
                        col_width = min(40, max(15, max_data_len + 2))
                    else:
                        # Para el resto de columnas, calcular automáticamente
                        col_width = min(30, max(header_len + 2, max_data_len + 2))

                    worksheet.set_column(i, i, col_width)

                # Si hay datos, agregar un total al final
                if has_totals:
                    last_row = start_row + len(df) + 1
                    worksheet.write(last_row, 0, 'TOTAL', total_row_format)

                    for col_idx, col_name in enumerate(df.columns):
                        if col_name.lower() in ['monto total', 'total', 'anticipo total', 'restante total', 'anticipo', 'monto restante']:
                            total_formula = f'=SUM({xl_col_to_name(col_idx)}{start_row + 1}:{xl_col_to_name(col_idx)}{last_row - 1})'
                            worksheet.write_formula(last_row, col_idx, total_formula, total_format)
                        elif col_name.lower() == 'cantidad':
                            count_formula = f'=SUM({xl_col_to_name(col_idx)}{start_row + 1}:{xl_col_to_name(col_idx)}{last_row - 1})'
                            worksheet.write_formula(last_row, col_idx, count_formula, count_format)

                # Autofiltro para facilitar la navegación
                if not df.empty:
                    worksheet.autofilter(start_row, 0, start_row + len(df), len(df.columns) - 1)

                # Freeze panes para mejor navegación
                worksheet.freeze_panes(start_row + 1, 0)

                return start_row + (len(df) + 3 if not df.empty else 2)

            # --- RESUMEN EJECUTIVO ---
            worksheet = workbook.add_worksheet('Resumen Ejecutivo')

            # Título del reporte más visible
            worksheet.merge_range('A1:H1', titulo_reporte, title_format)
            worksheet.set_row(0, 30)  # Altura para el título

            # Información general
            row = 2
            worksheet.write(row, 0, "Generado:", subtitle_format)
            worksheet.write(row, 1, datetime.now().strftime("%d/%m/%Y %H:%M"), cell_format)
            row += 2

            # Estadísticas principales con mejor formato
            worksheet.merge_range(f'A{row}:B{row}', "Estadísticas Generales:", subtitle_format)
            row += 1

            total_liquidado = df_liq['Monto Total'].sum() if not df_liq.empty else 0
            total_pendiente = df_pen['Monto Total'].sum() if not df_pen.empty else 0
            total_general = total_liquidado + total_pendiente

            stats_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter',
                'bg_color': '#FFF3E0'
            })

            value_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'bold': True
            })

            stats_data = [
                ["Total Liquidado", f"${total_liquidado:,.2f}"],
                ["Total Pendiente", f"${total_pendiente:,.2f}"],
                ["Total General", f"${total_general:,.2f}"],
                ["Cantidad de Solicitudes", str(len(df_sem)) if not df_sem.empty else "0"]
            ]

            # Crear tabla de estadísticas con mejor formato
            for i, (label, value) in enumerate(stats_data):
                worksheet.write(row + i, 0, label, stats_format)
                worksheet.write(row + i, 1, value, value_format)

            row += len(stats_data) + 2

            # Agregar resumen por tipo de solicitud
            if not df_res.empty:
                worksheet.merge_range(f'A{row}:D{row}', "Resumen por Tipo de Solicitud", subtitle_format)
                row += 1

                # Convertir a columnas para mostrar
                df_res_display = df_res[['Tipo de Solicitud', 'Total', 'Cantidad']]

                # Escribir encabezados
                for col_num, col_name in enumerate(df_res_display.columns):
                    worksheet.write(row, col_num, col_name, header_format)

                # Escribir datos
                for data_row_idx, data_row in df_res_display.iterrows():
                    for col_idx, col_name in enumerate(df_res_display.columns):
                        value = data_row[col_name]

                        if col_name == 'Total':
                            fmt = money_format
                        else:
                            fmt = cell_format

                        worksheet.write(row + 1 + data_row_idx, col_idx, value, fmt)

                # Ajustar anchos
                worksheet.set_column(0, 0, 20)  # Tipo de solicitud
                worksheet.set_column(1, 1, 15)  # Total
                worksheet.set_column(2, 2, 10)  # Cantidad

                row += len(df_res_display) + 3

            # Agregar resumen por tipo de pago - IMPORTANTE: Asegurar visualización de todos los tipos
            if not df_tipo_pago.empty:
                worksheet.merge_range(f'A{row}:D{row}', "Resumen por Tipo de Pago", subtitle_format)
                row += 1

                # Escribir encabezados
                for col_num, col_name in enumerate(df_tipo_pago.columns):
                    worksheet.write(row, col_num, col_name, header_format)

                # Escribir datos
                for data_row_idx, data_row in df_tipo_pago.iterrows():
                    for col_idx, col_name in enumerate(df_tipo_pago.columns):
                        value = data_row[col_name]

                        if col_name == 'Total':
                            fmt = money_format
                        else:
                            fmt = cell_format

                        worksheet.write(row + 1 + data_row_idx, col_idx, value, fmt)

                # Ajustar anchos
                worksheet.set_column(0, 0, 20)  # Tipo de pago
                worksheet.set_column(1, 1, 15)  # Total
                worksheet.set_column(2, 2, 10)  # Cantidad

                row += len(df_tipo_pago) + 3

            # Gráfico de torta - AHORA COLOCADO DEBAJO DE LAS TABLAS
            if not df_res.empty:
                chart = workbook.add_chart({'type': 'pie'})
                chart.add_series({
                    'name': 'Distribución por Tipo',
                    'categories': ['Resumen por Tipo', 1, 0, len(df_res), 0],
                    'values': ['Resumen por Tipo', 1, 1, len(df_res), 1],
                    'data_labels': {
                        'percentage': True,
                        'category': True,
                        'position': 'outside_end',
                        'font': {'bold': True}
                    },
                    'points': [
                        {'fill': {'color': '#FF9800'}},
                        {'fill': {'color': '#4CAF50'}},
                        {'fill': {'color': '#2196F3'}},
                        {'fill': {'color': '#9C27B0'}},
                        {'fill': {'color': '#F44336'}},
                        {'fill': {'color': '#009688'}},
                        {'fill': {'color': '#795548'}},
                        {'fill': {'color': '#607D8B'}}
                    ]
                })
                chart.set_title({'name': 'Distribución por Tipo de Solicitud', 'font': {'size': 14, 'bold': True}})
                chart.set_style(10)

                # POSICIONAR GRÁFICO DEBAJO DE LAS TABLAS
                worksheet.insert_chart(f'A{row}', chart, {
                    'x_offset': 25,
                    'y_offset': 10,
                    'x_scale': 1.5,
                    'y_scale': 1.5
                })

                # Gráfico de tipo de pagos
                if not df_tipo_pago.empty and len(df_tipo_pago) > 1:
                    chart_tp = workbook.add_chart({'type': 'pie'})
                    chart_tp.add_series({
                        'name': 'Distribución por Tipo de Pago',
                        'categories': ['Resumen Ejecutivo', row - len(df_tipo_pago) - 1, 0, row - 2, 0],
                        'values': ['Resumen Ejecutivo', row - len(df_tipo_pago) - 1, 1, row - 2, 1],
                        'data_labels': {
                            'percentage': True,
                            'category': True,
                            'position': 'outside_end',
                            'font': {'bold': True}
                        },
                        'points': [
                            {'fill': {'color': '#2196F3'}},
                            {'fill': {'color': '#4CAF50'}},
                            {'fill': {'color': '#FF9800'}},
                            {'fill': {'color': '#9C27B0'}},
                            {'fill': {'color': '#F44336'}},
                            {'fill': {'color': '#009688'}}
                        ]
                    })

                    chart_tp.set_title({'name': 'Distribución por Tipo de Pago', 'font': {'size': 14, 'bold': True}})
                    chart_tp.set_style(10)
                    worksheet.insert_chart(f'E{row}', chart_tp, {
                        'x_offset': 25,
                        'y_offset': 10,
                        'x_scale': 1.5,
                        'y_scale': 1.5
                    })

            # --- LIQUIDADOS ---
            if not df_liq.empty:
                worksheet = workbook.add_worksheet('Liquidados')
                num_columns = len(df_liq.columns)
                worksheet.merge_range(f'A1:{xl_col_to_name(num_columns-1)}1', f"Solicitudes Liquidadas - {fecha_inicio_str} al {fecha_fin_str}", title_format)
                worksheet.set_row(0, 30)  # Altura para el título
                worksheet.write(2, 0, f"Total Liquidado: ${total_liquidado:,.2f}", subtitle_format)

                # Usar nuestra función mejorada para dar formato a la hoja
                format_sheet(worksheet, df_liq, start_row=3, has_totals=True)

                # Establecer altura de fila para encabezados
                worksheet.set_row(3, 40)  # Mayor altura para facilitar lectura

            # --- PENDIENTES ---
            if not df_pen.empty:
                worksheet = workbook.add_worksheet('Pendientes')
                num_columns = len(df_pen.columns)
                worksheet.merge_range(f'A1:{xl_col_to_name(num_columns-1)}1', f"Solicitudes Pendientes - {fecha_inicio_str} al {fecha_fin_str}", title_format)
                worksheet.set_row(0, 30)  # Altura para el título
                worksheet.write(2, 0, f"Total Pendiente: ${total_pendiente:,.2f}", subtitle_format)

                # Aplicar formatos optimizados
                format_sheet(worksheet, df_pen, start_row=3, has_totals=True)

                # Establecer altura de fila para encabezados
                worksheet.set_row(3, 40)  # Mayor altura para facilitar lectura

            # --- RESUMEN POR TIPO ---
            if not df_res.empty:
                worksheet = workbook.add_worksheet('Resumen por Tipo')
                worksheet.merge_range('A1:D1', f"Resumen por Tipo de Solicitud - {fecha_inicio_str} al {fecha_fin_str}", title_format)
                worksheet.set_row(0, 30)  # Altura para el título

                # Aplicar formatos optimizados
                format_sheet(worksheet, df_res, start_row=3, has_totals=True)

                # Establecer altura de fila para encabezados
                worksheet.set_row(3, 40)  # Mayor altura para facilitar lectura

                # Gráfico para visualización de datos
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Total por Tipo',
                    'categories': ['Resumen por Tipo', 4, 0, 3 + len(df_res), 0],
                    'values': ['Resumen por Tipo', 4, 1, 3 + len(df_res), 1],
                    'data_labels': {'value': True},
                    'fill': {'color': '#FF9800'}
                })
                chart.set_title({'name': 'Total por Tipo de Solicitud'})
                chart.set_y_axis({'num_format': '$#,##0.00'})
                chart.set_style(10)

                # Posicionar gráfico después de la tabla
                row_graph = 5 + len(df_res)
                worksheet.insert_chart(f'A{row_graph}', chart, {'x_scale': 1.5, 'y_scale': 1.5})

            # --- RESUMEN POR TIPO DE PAGO ---
            if not df_tipo_pago.empty:
                worksheet = workbook.add_worksheet('Resumen por Tipo de Pago')
                worksheet.merge_range('A1:D1', f"Resumen por Tipo de Pago - {fecha_inicio_str} al {fecha_fin_str}", title_format)
                worksheet.set_row(0, 30)  # Altura para el título

                # Aplicar formatos optimizados
                format_sheet(worksheet, df_tipo_pago, start_row=3, has_totals=True)

                # Establecer altura de fila para encabezados
                worksheet.set_row(3, 40)  # Mayor altura para facilitar lectura

                # Gráfico para visualización de datos
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Total por Tipo de Pago',
                    'categories': ['Resumen por Tipo de Pago', 4, 0, 3 + len(df_tipo_pago), 0],
                    'values': ['Resumen por Tipo de Pago', 4, 1, 3 + len(df_tipo_pago), 1],
                    'data_labels': {'value': True},
                    'fill': {'color': '#2196F3'}
                })
                chart.set_title({'name': 'Total por Tipo de Pago'})
                chart.set_y_axis({'num_format': '$#,##0.00'})
                chart.set_style(10)

                # Posicionar gráfico después de la tabla
                row_graph = 5 + len(df_tipo_pago)
                worksheet.insert_chart(f'A{row_graph}', chart, {'x_scale': 1.5, 'y_scale': 1.5})
                        # --- RESUMEN POR PROVEEDOR ---
            if not df_prov.empty:
                worksheet = workbook.add_worksheet('Resumen por Proveedor')
                worksheet.merge_range(
                    'A1:D1',
                    f"Resumen por Proveedor - {fecha_inicio_str} al {fecha_fin_str}",
                    title_format
                )
                worksheet.set_row(0, 30)
                format_sheet(worksheet, df_prov, start_row=3, has_totals=True)
                worksheet.set_row(3, 40)

                # Top-10 proveedores (gráfico)
                top_prov = df_prov.nlargest(10, 'Total')
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Top-10 Proveedores',
                    'categories': ['Resumen por Proveedor', 4, 0, 3 + len(top_prov), 0],
                    'values':     ['Resumen por Proveedor', 4, 1, 3 + len(top_prov), 1],
                    'data_labels': {'value': True},
                    'fill': {'color': '#4CAF50'}
                })
                chart.set_title({'name': 'Top-10 Proveedores por Monto'})
                chart.set_y_axis({'num_format': '$#,##0.00'})
                chart.set_style(10)
                worksheet.insert_chart('E5', chart, {'x_scale': 1.5, 'y_scale': 1.5})

            # --- RESUMEN POR DEPARTAMENTO ---
            if not df_dept.empty:
                worksheet = workbook.add_worksheet('Resumen por Departamento')
                worksheet.merge_range(
                    'A1:D1',
                    f"Resumen por Departamento - {fecha_inicio_str} al {fecha_fin_str}",
                    title_format
                )
                worksheet.set_row(0, 30)
                format_sheet(worksheet, df_dept, start_row=3, has_totals=True)
                worksheet.set_row(3, 40)

                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Total por Departamento',
                    'categories': ['Resumen por Departamento', 4, 0, 3 + len(df_dept), 0],
                    'values':     ['Resumen por Departamento', 4, 1, 3 + len(df_dept), 1],
                    'data_labels': {'value': True},
                    'fill': {'color': '#9C27B0'}
                })
                chart.set_title({'name': 'Total por Departamento'})
                chart.set_y_axis({'num_format': '$#,##0.00'})
                chart.set_style(10)
                worksheet.insert_chart('E5', chart, {'x_scale': 1.5, 'y_scale': 1.5})

            # --- RESUMEN POR ESTADO ---
            if not df_estado.empty:
                worksheet = workbook.add_worksheet('Resumen por Estado')
                worksheet.merge_range(
                    'A1:D1',
                    f"Resumen por Estado - {fecha_inicio_str} al {fecha_fin_str}",
                    title_format
                )
                worksheet.set_row(0, 30)
                format_sheet(worksheet, df_estado, start_row=3, has_totals=True)
                worksheet.set_row(3, 40)

        # ────────────────────────────────
        # Enviar archivo al usuario
        # ────────────────────────────────
        output.seek(0)
        nombre_archivo = (
            f"reporte_semanal_{fecha_inicio.strftime('%Y%m%d')}"
            f"_a_{fecha_fin.strftime('%Y%m%d')}.xlsx"
        )
        return send_file(
            output,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        app.logger.error(f"Error al generar el reporte semanal: {e}")
        flash("Ocurrió un error al generar el reporte.", "error")
        return redirect(url_for("estadisticas"))


from xlsxwriter.utility import xl_col_to_name  # <- usado sólo en esta función

@app.route("/exportar_reporte_anual")
def exportar_reporte_anual():
    # Permisos
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para exportar reportes.", "error")
        return redirect(url_for("admin_dashboard"))

    # Año a exportar (default: año actual)
    anio_str = request.args.get("anio")
    try:
        anio = int(anio_str) if anio_str else datetime.now().year
    except ValueError:
        anio = datetime.now().year

    try:
        # --- Cargar datos base ---
        conn = get_db_connection()
        rows = conn.execute("""
            SELECT id, fp, nombre, destinatario, departamento, tipo_solicitud, tipo_pago,
                   banco, clabe, monto, anticipo, porcentaje_anticipo, monto_anticipo, monto_restante,
                   estado, fecha, fecha_limite, descripcion
            FROM solicitudes
        """).fetchall()
        conn.close()

        if not rows:
            flash("No hay datos para generar el reporte.", "warning")
            return redirect(url_for("estadisticas"))

        df = pd.DataFrame([dict(r) for r in rows])
        # Tipos
        df["monto"] = pd.to_numeric(df.get("monto", 0), errors="coerce").fillna(0)
        df["fecha_limite_dt"] = pd.to_datetime(df.get("fecha_limite"), format="%Y-%m-%d", errors="coerce")

        # Filtrar por año seleccionado (usamos fecha_límite como referencia temporal)
        df_year = df[df["fecha_limite_dt"].dt.year == anio].copy()

        if df_year.empty:
            flash(f"No hay datos en el año {anio}.", "warning")
            return redirect(url_for("estadisticas"))

        # ---------- RESÚMENES ----------
        # Por proveedor
        df_prov = (
            df_year.groupby("destinatario", dropna=False)
                   .agg(Total=("monto", "sum"), Cantidad=("id", "count"))
                   .reset_index()
                   .rename(columns={"destinatario": "Proveedor"})
                   .sort_values("Total", ascending=False)
        )

        # Por tipo de pago
        df_tipo_pago = (
            df_year.groupby("tipo_pago", dropna=False)
                   .agg(Total=("monto", "sum"), Cantidad=("id", "count"))
                   .reset_index()
                   .rename(columns={"tipo_pago": "Tipo de Pago"})
                   .sort_values("Total", ascending=False)
        )

        # Por tipo de solicitud (útil para el resumen)
        df_tipo_sol = (
            df_year.groupby("tipo_solicitud", dropna=False)
                   .agg(Total=("monto", "sum"), Cantidad=("id", "count"))
                   .reset_index()
                   .rename(columns={"tipo_solicitud": "Tipo de Solicitud"})
                   .sort_values("Total", ascending=False)
        )

        # Mensual (totales)
        meses = list(range(1, 13))
        mes_nombre = {
            1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
        }
        df_mes = (
            df_year.groupby(df_year["fecha_limite_dt"].dt.month)
                   .agg(Total=("monto", "sum"), Cantidad=("id", "count"))
                   .reindex(meses, fill_value=0)
                   .reset_index()
                   .rename(columns={"index": "Mes", "fecha_limite_dt": "Mes"})
        )
        df_mes["Mes"] = df_mes["index"].map(mes_nombre) if "index" in df_mes.columns else df_mes["Mes"].map(mes_nombre)
        if "index" in df_mes.columns:  # normalizar nombre de columna si aparece "index"
            df_mes = df_mes.drop(columns=["index"]).rename(columns={"Mes":"Mes", "Total":"Total", "Cantidad":"Cantidad"})

        # Pivot: Monto por Mes x Tipo de Pago
        pivot_pago_mes = (
            df_year.pivot_table(values="monto",
                                index=df_year["fecha_limite_dt"].dt.month,
                                columns="tipo_pago",
                                aggfunc="sum",
                                fill_value=0)
                  .reindex(meses, fill_value=0)
        )
        pivot_pago_mes.index = [mes_nombre[m] for m in pivot_pago_mes.index]
        pivot_pago_mes = pivot_pago_mes.reset_index().rename(columns={"index":"Mes"})

        # Pivot: Monto por Mes x Tipo de Solicitud
        pivot_sol_mes = (
            df_year.pivot_table(values="monto",
                                index=df_year["fecha_limite_dt"].dt.month,
                                columns="tipo_solicitud",
                                aggfunc="sum",
                                fill_value=0)
                  .reindex(meses, fill_value=0)
        )
        pivot_sol_mes.index = [mes_nombre[m] for m in pivot_sol_mes.index]
        pivot_sol_mes = pivot_sol_mes.reset_index().rename(columns={"index":"Mes"})

        # Detalle (limpio para Excel)
        df_det = df_year[[
            "fp","nombre","destinatario","departamento","tipo_solicitud","tipo_pago",
            "banco","clabe","monto","anticipo","porcentaje_anticipo","monto_anticipo",
            "monto_restante","estado","fecha","fecha_limite","descripcion"
        ]].copy()
        df_det = df_det.rename(columns={
            "fp":"FP","nombre":"Solicitante","destinatario":"Proveedor","departamento":"Departamento",
            "tipo_solicitud":"Tipo de Solicitud","tipo_pago":"Tipo de Pago","banco":"Banco","clabe":"CLABE",
            "monto":"Monto","anticipo":"Anticipo","porcentaje_anticipo":"% Anticipo","monto_anticipo":"Monto Anticipo",
            "monto_restante":"Monto Restante","estado":"Estado","fecha":"Fecha Solicitud","fecha_limite":"Fecha Límite",
            "descripcion":"Descripción"
        })

        # ---------- EXCEL ----------
        titulo = f"Reporte Anual {anio}"
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book

            # Formatos (paleta “naranja” como tu semanal)
            title_fmt = wb.add_format({'bold':True,'font_size':18,'align':'center','valign':'vcenter',
                                       'font_color':'#FFFFFF','bg_color':'#FF9800','border':1,'border_color':'#E65100'})
            header_fmt = wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#FF9800',
                                        'align':'center','valign':'vcenter','border':1,'text_wrap':True})
            cell_fmt   = wb.add_format({'border':1,'valign':'top'})
            money_fmt  = wb.add_format({'num_format':'$#,##0.00','border':1,'align':'right'})
            date_fmt   = wb.add_format({'num_format':'dd/mm/yyyy','border':1,'align':'center'})
            total_fmt  = wb.add_format({'bold':True,'num_format':'$#,##0.00','border':1,'bg_color':'#FFF3E0','font_color':'#E65100','align':'right'})

            # Helper para volcar un DF con encabezado y tamaños
            def write_df(ws_name, df_to_write, title=None, start_row=0, totals_cols=None):
                ws = wb.add_worksheet(ws_name)
                if title:
                    last_col = xl_col_to_name(max(0, len(df_to_write.columns)-1))
                    ws.merge_range(f"A1:{last_col}1", title, title_fmt)
                    ws.set_row(0, 28)
                    r = 2
                else:
                    r = 0
                # headers
                for c, col in enumerate(df_to_write.columns):
                    ws.write(r, c, col, header_fmt)
                # data rows
                for i, row in df_to_write.iterrows():
                    for c, col in enumerate(df_to_write.columns):
                        v = row[col]
                        fmt = money_fmt if any(k in col.lower() for k in ["monto","total"]) else cell_fmt
                        ws.write(r+1+i, c, v, fmt)
                # autosize
                for c, col in enumerate(df_to_write.columns):
                    width = min(45, max(len(str(col))+2, int(df_to_write[col].astype(str).str.len().max() if len(df_to_write)>0 else 0)+2))
                    ws.set_column(c, c, width)
                # totals row
                if totals_cols and len(df_to_write)>0:
                    tr = r + 1 + len(df_to_write)
                    ws.write(tr, 0, "TOTAL", total_fmt)
                    for c, col in enumerate(df_to_write.columns):
                        if col in totals_cols:
                            # =SUM(B3:Bn)
                            top = r+2
                            bottom = tr
                            ws.write_formula(tr, c, f"=SUM({xl_col_to_name(c)}{top}:{xl_col_to_name(c)}{bottom})", total_fmt)
                return ws

            # --- Resumen Ejecutivo ---
            ws = wb.add_worksheet("Resumen Ejecutivo")
            ws.merge_range("A1:H1", titulo, title_fmt); ws.set_row(0, 28)

            total_anual = float(df_year["monto"].sum())
            cantidad = int(df_year.shape[0])
            promedio_mensual = total_anual/12.0

            stats = [
                ("Total del Año", total_anual),
                ("Solicitudes (cantidad)", cantidad),
                ("Promedio mensual", promedio_mensual)
            ]
            ws.write(2, 0, "Métrica", header_fmt)
            ws.write(2, 1, "Valor", header_fmt)
            for i,(k,v) in enumerate(stats):
                ws.write(3+i, 0, k, cell_fmt)
                ws.write(3+i, 1, v, money_fmt if i!=1 else cell_fmt)

            # mini tablas
            start = 8
            # Por Tipo de Pago
            if not df_tipo_pago.empty:
                ws.write(start, 0, "Resumen por Tipo de Pago", header_fmt); start+=1
                ws.write_row(start, 0, list(df_tipo_pago.columns), header_fmt); start+=1
                for _,r in df_tipo_pago.iterrows():
                    ws.write_row(start, 0, r.tolist(), cell_fmt);
                    ws.write(start, 1, r["Total"], money_fmt)
                    start+=1
                start+=1
            # Por Tipo de Solicitud
            if not df_tipo_sol.empty:
                ws.write(start, 0, "Resumen por Tipo de Solicitud", header_fmt); start+=1
                ws.write_row(start, 0, list(df_tipo_sol.columns), header_fmt); start+=1
                for _,r in df_tipo_sol.iterrows():
                    ws.write_row(start, 0, r.tolist(), cell_fmt)
                    ws.write(start, 1, r["Total"], money_fmt)
                    start+=1

            # Gráfico mensual (columna)
            if not df_mes.empty:
                # Pasamos df_mes a hoja aparte y graficamos desde ahí
                df_mes.to_excel(writer, sheet_name="Por Mes (Totales)", index=False, startrow=2)
                ws_mes = writer.sheets["Por Mes (Totales)"]
                ws_mes.merge_range("A1:C1", f"Totales Mensuales {anio}", title_fmt)
                for c,_ in enumerate(df_mes.columns): ws_mes.write(2, c, df_mes.columns[c], header_fmt)
                # chart
                chart = wb.add_chart({'type':'column'})
                chart.add_series({
                    'name': 'Total',
                    'categories': ['Por Mes (Totales)', 3, 0, 3+len(df_mes)-1, 0],
                    'values':     ['Por Mes (Totales)', 3, 1, 3+len(df_mes)-1, 1],
                    'data_labels': {'value': True}
                })
                chart.set_title({'name':'Monto por Mes'})
                chart.set_y_axis({'num_format':'$#,##0.00'})
                ws.insert_chart("E3", chart, {'x_scale':1.5,'y_scale':1.4})

            # --- Por Proveedor ---
            if not df_prov.empty:
                write_df("Por Proveedor", df_prov, f"Resumen por Proveedor {anio}", totals_cols={"Total"})
                # Top-10 chart
                ws_p = writer.sheets["Por Proveedor"]
                top = df_prov.head(10)
                start_row = 3 + len(df_prov) + 3
                ws_p.write(start_row, 0, "Top 10 Proveedores por Monto", header_fmt)
                chartp = wb.add_chart({'type':'column'})
                chartp.add_series({
                    'name':'Total',
                    'categories':['Por Proveedor', 4, 0, 3+len(top), 0],
                    'values':    ['Por Proveedor', 4, 1, 3+len(top), 1],
                    'data_labels': {'value': True}
                })
                chartp.set_y_axis({'num_format':'$#,##0.00'})
                ws_p.insert_chart(start_row+1, 0, chartp, {'x_scale':1.5,'y_scale':1.4})

            # --- Por Tipo de Pago ---
            if not df_tipo_pago.empty:
                write_df("Por Tipo de Pago", df_tipo_pago, f"Resumen por Tipo de Pago {anio}", totals_cols={"Total"})
                # Pie chart
                ws_tp = writer.sheets["Por Tipo de Pago"]
                charttp = wb.add_chart({'type':'pie'})
                charttp.add_series({
                    'name':'Distribución',
                    'categories':['Por Tipo de Pago', 3, 0, 3+len(df_tipo_pago)-1, 0],
                    'values':    ['Por Tipo de Pago', 3, 1, 3+len(df_tipo_pago)-1, 1],
                    'data_labels': {'percentage':True, 'category':True}
                })
                charttp.set_title({'name':'Distribución por Tipo de Pago'})
                ws_tp.insert_chart("E3", charttp, {'x_scale':1.4,'y_scale':1.4})

            # --- Por Mes x Tipo de Pago ---
            if not pivot_pago_mes.empty:
                write_df("Mes x TipoPago", pivot_pago_mes, f"Mes x Tipo de Pago {anio}")

            # --- Por Mes x Tipo de Solicitud ---
            if not pivot_sol_mes.empty:
                write_df("Mes x TipoSol", pivot_sol_mes, f"Mes x Tipo de Solicitud {anio}")

            # --- Detalle del Año ---
            # Convertir las fechas a datetime para formato de Excel
            for col in ["Fecha Solicitud","Fecha Límite"]:
                df_det[col] = pd.to_datetime(df_det[col], errors="coerce")
            df_det.to_excel(writer, sheet_name=f"Detalle {anio}", index=False, startrow=2)
            ws_det = writer.sheets[f"Detalle {anio}"]
            ws_det.merge_range(0, 0, 0, len(df_det.columns)-1, f"Detalle de Solicitudes {anio}", title_fmt)
            # headers bonitos
            for c, col in enumerate(df_det.columns):
                ws_det.write(2, c, col, header_fmt)
                width = min(45, max(len(col)+2, int(df_det[col].astype(str).str.len().max() if len(df_det)>0 else 0)+2))
                ws_det.set_column(c, c, width)
            # formatear columnas de dinero y fechas
            for r in range(len(df_det)):
                for c, col in enumerate(df_det.columns):
                    val = df_det.iloc[r, c]
                    if col.lower().startswith("monto"):
                        ws_det.write(3+r, c, val, money_fmt)
                    elif "fecha" in col.lower():
                        ws_det.write(3+r, c, val, date_fmt)
                    else:
                        ws_det.write(3+r, c, val, cell_fmt)

        output.seek(0)
        filename = f"reporte_anual_{anio}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        app.logger.error(f"Error en reporte anual: {e}")
        flash("Ocurrió un error al generar el reporte anual.", "error")
        return redirect(url_for("estadisticas"))


# Vista principal de créditos con manejo correcto de variables

@app.route("/creditos")
def creditos_dashboard():
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    # Inicializa las variables con valores predeterminados
    total_creditos = 0
    monto_total_creditos = 0
    total_activos = 0
    monto_activos = 0
    pago_mensual_total = 0
    creditos_list = []

    # Filtro para créditos
    estado_filtro = request.args.get("estado", "todos")
    busqueda = request.args.get("busqueda", "").strip().lower()

    try:
        conn = get_db_connection()

        # Obtener las estadísticas generales de créditos
        total_creditos = conn.execute("SELECT COUNT(*) FROM creditos").fetchone()[0] or 0
        monto_total_creditos = conn.execute("SELECT SUM(monto_total) FROM creditos").fetchone()[0] or 0
        total_activos = conn.execute("SELECT COUNT(*) FROM creditos WHERE estado = 'Activo'").fetchone()[0] or 0
        monto_activos = conn.execute("SELECT SUM(monto_total) FROM creditos WHERE estado = 'Activo'").fetchone()[0] or 0

        # Pago mensual total
        pago_mensual_total = conn.execute("""
            SELECT SUM(pago_mensual) FROM creditos WHERE estado = 'Activo'
        """).fetchone()[0] or 0

        # Construir la consulta según los filtros
        query = "SELECT * FROM creditos"
        params = []

        if estado_filtro and estado_filtro != "todos":
            query += " WHERE estado = ?"
            params.append(estado_filtro)

        if busqueda:
            if "WHERE" in query:
                query += " AND (LOWER(nombre) LIKE ? OR LOWER(entidad) LIKE ? OR LOWER(descripcion) LIKE ?)"
            else:
                query += " WHERE (LOWER(nombre) LIKE ? OR LOWER(entidad) LIKE ? OR LOWER(descripcion) LIKE ?)"
            busqueda_param = f"%{busqueda}%"
            params.extend([busqueda_param, busqueda_param, busqueda_param])

        query += " ORDER BY fecha_registro DESC"

        if params:
            creditos_db = conn.execute(query, params).fetchall()
        else:
            creditos_db = conn.execute(query).fetchall()

        # Para cada crédito, calcular el monto pagado y el saldo restante
        creditos_list = []
        for credito in creditos_db:
            credito_dict = dict(credito)

            # Calcular monto pagado hasta la fecha
            pagado = conn.execute(
                "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
                (credito_dict["id"],)
            ).fetchone()[0] or 0

            credito_dict["monto_pagado"] = pagado
            credito_dict["saldo_pendiente"] = credito_dict["monto_total"] - pagado

            # Calcular porcentaje de avance
            if credito_dict["monto_total"] > 0:
                credito_dict["porcentaje_pagado"] = (pagado / credito_dict["monto_total"]) * 100
            else:
                credito_dict["porcentaje_pagado"] = 0

            creditos_list.append(credito_dict)

        conn.close()
    except Exception as e:
        # Manejar cualquier error que ocurra
        print(f"Error al cargar datos de créditos: {e}")
        flash(f"Error al cargar datos: {e}", "error")

    return render_template(
        "dashboard.html",
        creditos=creditos_list,
        total_creditos=total_creditos,
        monto_total_creditos=monto_total_creditos,
        total_activos=total_activos,
        monto_activos=monto_activos,
        pago_mensual_total=pago_mensual_total,
        estado_filtro=estado_filtro,
        busqueda=busqueda
    )
@app.route("/creditos/nuevo", methods=["GET", "POST"])
def nuevo_credito():
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        # Obtener datos del formulario
        nombre = request.form.get("nombre")
        entidad = request.form.get("entidad")
        descripcion = request.form.get("descripcion")
        monto_total_str = request.form.get("monto_total", "0").replace(",", ".")
        tasa_interes_str = request.form.get("tasa_interes", "0").replace(",", ".")
        fecha_inicio = request.form.get("fecha_inicio")
        fecha_final = request.form.get("fecha_final")
        plazo_meses = request.form.get("plazo_meses", "0")
        estado = request.form.get("estado", "Activo")
        numero_cuenta = request.form.get("numero_cuenta")
        tipo_credito = request.form.get("tipo_credito")
        pago_mensual_str = request.form.get("pago_mensual", "0").replace(",", ".")
        contacto = request.form.get("contacto")
        notas = request.form.get("notas")

        # Convertir a tipos adecuados
        try:
            monto_total = float(monto_total_str)
            tasa_interes = float(tasa_interes_str)
            plazo_meses = int(plazo_meses)
            pago_mensual = float(pago_mensual_str)
        except ValueError:
            flash("Por favor, ingresa valores numéricos válidos.", "error")
            return redirect(url_for("nuevo_credito"))

        # Validación básica
        if not nombre or not entidad or not fecha_inicio or not fecha_final:
            flash("Por favor, completa todos los campos obligatorios.", "error")
            return redirect(url_for("nuevo_credito"))

        # Registrar en la base de datos
        fecha_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        conn = get_db_connection()
        conn.execute("""
            INSERT INTO creditos (
                nombre, entidad, descripcion, monto_total, tasa_interes,
                fecha_inicio, fecha_final, plazo_meses, estado, fecha_registro,
                numero_cuenta, tipo_credito, pago_mensual, contacto, notas
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            nombre, entidad, descripcion, monto_total, tasa_interes,
            fecha_inicio, fecha_final, plazo_meses, estado, fecha_registro,
            numero_cuenta, tipo_credito, pago_mensual, contacto, notas
        ))
        conn.commit()
        conn.close()

        flash("Crédito registrado exitosamente.", "success")
        return redirect(url_for("creditos_dashboard"))

    return render_template("nuevo_credito.html")

# En tu app.py, modifica la función editar_credito:
@app.route("/creditos/editar/<int:credito_id>", methods=["GET", "POST"])
def editar_credito(credito_id):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db_connection()
    credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()

    if credito is None:
        conn.close()
        flash("Crédito no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    if request.method == "POST":
        # Obtener datos del formulario
        nombre = request.form.get("nombre")
        entidad = request.form.get("entidad")
        descripcion = request.form.get("descripcion")
        monto_total_str = request.form.get("monto_total", "0").replace(",", ".")
        tasa_interes_str = request.form.get("tasa_interes", "0").replace(",", ".")
        fecha_inicio = request.form.get("fecha_inicio")
        fecha_final = request.form.get("fecha_final")
        plazo_meses = request.form.get("plazo_meses", "0")
        estado = request.form.get("estado", "Activo")
        numero_cuenta = request.form.get("numero_cuenta")
        tipo_credito = request.form.get("tipo_credito")
        pago_mensual_str = request.form.get("pago_mensual", "0").replace(",", ".")
        contacto = request.form.get("contacto")
        notas = request.form.get("notas")
        motivo_cambio = request.form.get("motivo_cambio", "")

        # Convertir a tipos adecuados
        try:
            monto_total = float(monto_total_str)
            tasa_interes = float(tasa_interes_str)
            plazo_meses = int(plazo_meses)
            pago_mensual = float(pago_mensual_str)
        except ValueError:
            flash("Por favor, ingresa valores numéricos válidos.", "error")
            conn.close()
            return redirect(url_for("editar_credito", credito_id=credito_id))

        # Validación básica
        if not nombre or not entidad or not fecha_inicio or not fecha_final:
            flash("Por favor, completa todos los campos obligatorios.", "error")
            conn.close()
            return redirect(url_for("editar_credito", credito_id=credito_id))

        # Obtener el monto anterior para comparar
        monto_anterior = credito['monto_total']

        # Si el monto cambió, guardar en historial
        if monto_anterior != monto_total:
            if not motivo_cambio:
                motivo_cambio = "Modificación del monto del crédito"

            usuario = session.get("admin_name", "Administrador")
            fecha_cambio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            conn.execute("""
                INSERT INTO historial_monto_credito (
                    credito_id, monto_anterior, monto_nuevo, fecha_cambio, motivo, usuario
                ) VALUES (?, ?, ?, ?, ?, ?)
            """, (credito_id, monto_anterior, monto_total, fecha_cambio, motivo_cambio, usuario))

        # Actualizar el crédito
        conn.execute("""
            UPDATE creditos SET
                nombre = ?, entidad = ?, descripcion = ?, monto_total = ?, tasa_interes = ?,
                fecha_inicio = ?, fecha_final = ?, plazo_meses = ?, estado = ?,
                numero_cuenta = ?, tipo_credito = ?, pago_mensual = ?, contacto = ?, notas = ?
            WHERE id = ?
        """, (
            nombre, entidad, descripcion, monto_total, tasa_interes,
            fecha_inicio, fecha_final, plazo_meses, estado,
            numero_cuenta, tipo_credito, pago_mensual, contacto, notas,
            credito_id
        ))

        conn.commit()
        conn.close()

        if monto_anterior != monto_total:
            flash(f"Crédito actualizado exitosamente. Monto modificado de ${monto_anterior:,.2f} a ${monto_total:,.2f}", "success")
        else:
            flash("Crédito actualizado exitosamente.", "success")

        return redirect(url_for("detalle_credito", credito_id=credito_id))

    conn.close()
    return render_template("editar_credito.html", credito=dict(credito))

@app.route("/creditos/detalle/<int:credito_id>")
def detalle_credito(credito_id):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db_connection()
    credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()

    if credito is None:
        conn.close()
        flash("Crédito no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    # Obtener pagos del crédito
    pagos = conn.execute(
        "SELECT * FROM pagos_credito WHERE credito_id = ? ORDER BY fecha DESC",
        (credito_id,)
    ).fetchall()

    # Obtener historial de cambios de monto
    historial_monto = conn.execute("""
        SELECT * FROM historial_monto_credito
        WHERE credito_id = ?
        ORDER BY fecha_cambio DESC
    """, (credito_id,)).fetchall()

    # Calcular estadísticas
    total_pagado = conn.execute(
        "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
        (credito_id,)
    ).fetchone()[0] or 0

    credito_dict = dict(credito)
    credito_dict["monto_pagado"] = total_pagado
    credito_dict["saldo_pendiente"] = credito_dict["monto_total"] - total_pagado

    # Calcular porcentaje de avance
    if credito_dict["monto_total"] > 0:
        credito_dict["porcentaje_pagado"] = (total_pagado / credito_dict["monto_total"]) * 100
    else:
        credito_dict["porcentaje_pagado"] = 0

    pagos_list = [dict(pago) for pago in pagos]
    historial_monto_list = [dict(h) for h in historial_monto]
    conn.close()

    # Para visualización de progreso
    meses_transcurridos = calcular_meses_transcurridos(credito_dict["fecha_inicio"])
    meses_restantes = max(0, credito_dict["plazo_meses"] - meses_transcurridos)

    # Generar etiquetas para el gráfico de amortización
    labels_meses = []
    fecha_inicio = datetime.strptime(credito_dict["fecha_inicio"], "%Y-%m-%d")
    for i in range(credito_dict["plazo_meses"]):
        fecha_mes = fecha_inicio + timedelta(days=30*i)
        labels_meses.append(fecha_mes.strftime("%b %Y"))

    return render_template(
        "detalle_credito.html",
        credito=credito_dict,
        pagos=pagos_list,
        historial_monto=historial_monto_list,
        meses_transcurridos=meses_transcurridos,
        meses_restantes=meses_restantes,
        labels_meses=labels_meses
    )

# Función auxiliar para calcular meses transcurridos
def calcular_meses_transcurridos(fecha_inicio_str):
    """Calcula los meses transcurridos desde la fecha de inicio"""
    try:
        fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
        fecha_actual = datetime.now()

        meses = (fecha_actual.year - fecha_inicio.year) * 12
        meses += fecha_actual.month - fecha_inicio.month

        return max(0, meses)
    except:
        return 0

@app.route("/nuevo_pago/<int:credito_id>", methods=["GET", "POST"])
def nuevo_pago(credito_id):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db_connection()
    credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()

    if credito is None:
        conn.close()
        flash("Crédito no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    credito_dict = dict(credito)

    if request.method == "POST":
        # Obtener datos del formulario
        monto_str = request.form.get("monto", "0").replace(",", ".")
        fecha = request.form.get("fecha")
        referencia = request.form.get("referencia")
        descripcion = request.form.get("descripcion")
        tipo_pago = request.form.get("tipo_pago")

        # Procesar archivo adjunto
        comprobante_file = request.files.get("comprobante_file")
        comprobante_filename = ""
        if comprobante_file and comprobante_file.filename != "":
            comprobante_filename = secure_filename(comprobante_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], comprobante_filename)
            comprobante_file.save(filepath)

        # Convertir a tipos adecuados
        try:
            monto = float(monto_str)
        except ValueError:
            flash("Por favor, ingresa un monto válido.", "error")
            return redirect(url_for("nuevo_pago", credito_id=credito_id))

        # Validación básica
        if not fecha or monto <= 0:
            flash("Por favor, completa todos los campos obligatorios.", "error")
            return redirect(url_for("nuevo_pago", credito_id=credito_id))

        # Registrar en la base de datos
        conn.execute("""
            INSERT INTO pagos_credito (
                credito_id, monto, fecha, referencia, descripcion, comprobante, tipo_pago
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            credito_id, monto, fecha, referencia, descripcion, comprobante_filename, tipo_pago
        ))
        conn.commit()

        # Actualizar estado del crédito si ya está pagado completamente
        total_pagado = conn.execute(
            "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
            (credito_id,)
        ).fetchone()[0] or 0

        if total_pagado >= credito_dict["monto_total"] and credito_dict["estado"] == "Activo":
            conn.execute(
                "UPDATE creditos SET estado = 'Liquidado' WHERE id = ?",
                (credito_id,)
            )
            conn.commit()
            flash("Crédito completamente liquidado.", "success")

        conn.close()
        flash("Pago registrado exitosamente.", "success")
        return redirect(url_for("detalle_credito", credito_id=credito_id))

    conn.close()
    return render_template("nuevo_pago.html", credito=credito_dict)

@app.route("/creditos/editar_pago/<int:pago_id>", methods=["GET", "POST"])
def editar_pago(pago_id):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db_connection()
    pago = conn.execute("SELECT * FROM pagos_credito WHERE id = ?", (pago_id,)).fetchone()

    if pago is None:
        conn.close()
        flash("Pago no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    pago_dict = dict(pago)
    credito_id = pago_dict["credito_id"]

    credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()
    if credito is None:
        conn.close()
        flash("Crédito no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    credito_dict = dict(credito)

    if request.method == "POST":
        # Obtener datos del formulario
        monto_str = request.form.get("monto", "0").replace(",", ".")
        fecha = request.form.get("fecha")
        referencia = request.form.get("referencia")
        descripcion = request.form.get("descripcion")
        tipo_pago = request.form.get("tipo_pago")

        # Procesar archivo adjunto
        comprobante_file = request.files.get("comprobante_file")
        comprobante_filename = pago_dict["comprobante"] or ""
        if comprobante_file and comprobante_file.filename != "":
            comprobante_filename = secure_filename(comprobante_file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], comprobante_filename)
            comprobante_file.save(filepath)

        # Convertir a tipos adecuados
        try:
            monto = float(monto_str)
        except ValueError:
            flash("Por favor, ingresa un monto válido.", "error")
            return redirect(url_for("editar_pago", pago_id=pago_id))

        # Validación básica
        if not fecha or monto <= 0:
            flash("Por favor, completa todos los campos obligatorios.", "error")
            return redirect(url_for("editar_pago", pago_id=pago_id))

        # Actualizar en la base de datos
        conn.execute("""
            UPDATE pagos_credito SET
                monto = ?, fecha = ?, referencia = ?, descripcion = ?, comprobante = ?, tipo_pago = ?
            WHERE id = ?
        """, (
            monto, fecha, referencia, descripcion, comprobante_filename, tipo_pago, pago_id
        ))
        conn.commit()

        # Actualizar estado del crédito según los pagos
        total_pagado = conn.execute(
            "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
            (credito_id,)
        ).fetchone()[0] or 0

        if total_pagado >= credito_dict["monto_total"] and credito_dict["estado"] == "Activo":
            conn.execute(
                "UPDATE creditos SET estado = 'Liquidado' WHERE id = ?",
                (credito_id,)
            )
            conn.commit()
        elif total_pagado < credito_dict["monto_total"] and credito_dict["estado"] == "Liquidado":
            conn.execute(
                "UPDATE creditos SET estado = 'Activo' WHERE id = ?",
                (credito_id,)
            )
            conn.commit()

        conn.close()
        flash("Pago actualizado exitosamente.", "success")
        return redirect(url_for("detalle_credito", credito_id=credito_id))

    conn.close()
    return render_template("editar_pago.html", pago=pago_dict, credito=credito_dict)

@app.route("/creditos/eliminar_pago/<int:pago_id>", methods=["POST"])
def eliminar_pago(pago_id):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db_connection()
    pago = conn.execute("SELECT * FROM pagos_credito WHERE id = ?", (pago_id,)).fetchone()

    if pago is None:
        conn.close()
        flash("Pago no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    credito_id = pago["credito_id"]
    credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()

    # Eliminar el pago
    conn.execute("DELETE FROM pagos_credito WHERE id = ?", (pago_id,))
    conn.commit()

    # Recalcular estado del crédito
    if credito:
        total_pagado = conn.execute(
            "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
            (credito_id,)
        ).fetchone()[0] or 0

        if total_pagado >= credito["monto_total"] and credito["estado"] == "Activo":
            conn.execute(
                "UPDATE creditos SET estado = 'Liquidado' WHERE id = ?",
                (credito_id,)
            )
        elif total_pagado < credito["monto_total"] and credito["estado"] == "Liquidado":
            conn.execute(
                "UPDATE creditos SET estado = 'Activo' WHERE id = ?",
                (credito_id,)
            )
        conn.commit()

    conn.close()
    flash("Pago eliminado exitosamente.", "success")
    return redirect(url_for("detalle_credito", credito_id=credito_id))

@app.route("/creditos/eliminar/<int:credito_id>", methods=["POST"])
def eliminar_credito(credito_id):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db_connection()
    credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()

    if credito is None:
        conn.close()
        flash("Crédito no encontrado.", "error")
        return redirect(url_for("creditos_dashboard"))

    # Primero eliminar todos los pagos asociados
    conn.execute("DELETE FROM pagos_credito WHERE credito_id = ?", (credito_id,))

    # Luego eliminar el crédito
    conn.execute("DELETE FROM creditos WHERE id = ?", (credito_id,))
    conn.commit()
    conn.close()

    flash("Crédito y todos sus pagos eliminados exitosamente.", "success")
    return redirect(url_for("creditos_dashboard"))

@app.route("/creditos/estadisticas")
def estadisticas_creditos():
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    # Inicializar todas las variables con valores predeterminados
    total_creditos = 0
    monto_total_creditos = 0
    total_activos = 0
    monto_activos = 0
    total_liquidados = 0
    monto_liquidados = 0
    total_pagado = 0
    pago_mensual_total = 0
    tipos_credito = []
    entidades = []
    pagos_por_mes = []
    proyeccion_pagos = []
    grafico_tipo = None
    grafico_pagos = None

    try:
        conn = get_db_connection()

        # Estadísticas generales
        total_creditos = conn.execute("SELECT COUNT(*) FROM creditos").fetchone()[0] or 0
        monto_total_creditos = conn.execute("SELECT SUM(monto_total) FROM creditos").fetchone()[0] or 0
        total_activos = conn.execute("SELECT COUNT(*) FROM creditos WHERE estado = 'Activo'").fetchone()[0] or 0
        monto_activos = conn.execute("SELECT SUM(monto_total) FROM creditos WHERE estado = 'Activo'").fetchone()[0] or 0
        total_liquidados = conn.execute("SELECT COUNT(*) FROM creditos WHERE estado = 'Liquidado'").fetchone()[0] or 0
        monto_liquidados = conn.execute("SELECT SUM(monto_total) FROM creditos WHERE estado = 'Liquidado'").fetchone()[0] or 0

        # Total pagado de todos los créditos
        total_pagado = conn.execute("SELECT SUM(monto) FROM pagos_credito").fetchone()[0] or 0

        # Pago mensual total
        pago_mensual_total = conn.execute("""
            SELECT SUM(pago_mensual) FROM creditos WHERE estado = 'Activo'
        """).fetchone()[0] or 0

        # Desglose por tipo de crédito
        tipos_credito_rows = conn.execute("""
            SELECT tipo_credito, COUNT(*) as cantidad, SUM(monto_total) as monto
            FROM creditos
            GROUP BY tipo_credito
            ORDER BY monto DESC
        """).fetchall()

        tipos_credito = [dict(row) for row in tipos_credito_rows]

        # Desglose por entidad financiera
        entidades_rows = conn.execute("""
            SELECT entidad, COUNT(*) as cantidad, SUM(monto_total) as monto
            FROM creditos
            GROUP BY entidad
            ORDER BY monto DESC
        """).fetchall()

        entidades = [dict(row) for row in entidades_rows]

        # Historial de pagos por mes (últimos 12 meses)
        pagos_por_mes = []
        for i in range(12, 0, -1):
            fecha_inicio = (datetime.now() - timedelta(days=30*i)).strftime("%Y-%m-01")
            fecha_fin = (datetime.now() - timedelta(days=30*(i-1))).strftime("%Y-%m-01")
            monto = conn.execute("""
                SELECT SUM(monto) FROM pagos_credito
                WHERE fecha >= ? AND fecha < ?
            """, (fecha_inicio, fecha_fin)).fetchone()[0] or 0

            mes = (datetime.now() - timedelta(days=30*i)).strftime("%b %Y")
            pagos_por_mes.append({"mes": mes, "monto": monto})

        # Proyección de pagos futuros (próximos 12 meses)
        proyeccion_pagos = []
        fecha_actual = datetime.now()
        for i in range(12):
            fecha_proyeccion = (fecha_actual + timedelta(days=30*i)).strftime("%b %Y")
            proyeccion_pagos.append({
                "mes": fecha_proyeccion,
                "monto": pago_mensual_total
            })

        # Generar gráficos
        # Gráfico de distribución por tipo de crédito
        tipo_labels = [row["tipo_credito"] for row in tipos_credito]
        tipo_montos = [row["monto"] for row in tipos_credito]

        if tipo_labels and tipo_montos:
            try:
                fig1, ax1 = plt.subplots(figsize=(10, 6))
                ax1.pie(tipo_montos, labels=tipo_labels, autopct='%1.1f%%', startangle=90)
                ax1.axis('equal')
                plt.title("Distribución por Tipo de Crédito")

                buf1 = BytesIO()
                plt.savefig(buf1, format="png", dpi=80)
                buf1.seek(0)
                grafico_tipo = base64.b64encode(buf1.read()).decode("utf-8")
                plt.close(fig1)
            except Exception as e:
                print(f"Error al generar gráfico de tipo: {e}")

        # Gráfico de pagos mensuales (histórico)
        meses = [item["mes"] for item in pagos_por_mes]
        montos = [item["monto"] for item in pagos_por_mes]

        if meses and montos:
            try:
                fig2, ax2 = plt.subplots(figsize=(12, 6))
                ax2.bar(meses, montos)
                ax2.set_title("Historial de Pagos Mensuales")
                ax2.set_xlabel("Mes")
                ax2.set_ylabel("Monto Pagado")
                ax2.set_xticklabels(meses, rotation=45)
                plt.tight_layout()

                buf2 = BytesIO()
                plt.savefig(buf2, format="png", dpi=80)
                buf2.seek(0)
                grafico_pagos = base64.b64encode(buf2.read()).decode("utf-8")
                plt.close(fig2)
            except Exception as e:
                print(f"Error al generar gráfico de pagos: {e}")

        conn.close()
    except Exception as e:
        print(f"Error al obtener estadísticas de créditos: {e}")
        flash(f"Error al obtener estadísticas: {e}", "error")

    return render_template(
        "estadisticasC.html",
        total_creditos=total_creditos,
        monto_total_creditos=monto_total_creditos,
        total_activos=total_activos,
        monto_activos=monto_activos,
        total_liquidados=total_liquidados,
        monto_liquidados=monto_liquidados,
        total_pagado=total_pagado,
        pago_mensual_total=pago_mensual_total,
        tipos_credito=tipos_credito,
        entidades=entidades,
        pagos_por_mes=pagos_por_mes,
        proyeccion_pagos=proyeccion_pagos,
        grafico_tipo=grafico_tipo,
        grafico_pagos=grafico_pagos
    )

# Funciones auxiliares para el módulo de créditos
# Añadir estas funciones a tu archivo app.py

def calcular_meses_transcurridos(fecha_inicio_str):
    """
    Calcula los meses transcurridos desde la fecha de inicio hasta la fecha actual
    """
    try:
        fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
        fecha_actual = datetime.now()

        # Calcular diferencia en meses
        meses = (fecha_actual.year - fecha_inicio.year) * 12 + (fecha_actual.month - fecha_inicio.month)

        # Ajustar si aún no ha pasado un mes completo
        if fecha_actual.day < fecha_inicio.day:
            meses -= 1

        return max(0, meses)
    except Exception as e:
        print(f"Error al calcular meses transcurridos: {e}")
        return 0

def calcular_meses_restantes(fecha_inicio_str, plazo_meses):
    """
    Calcula los meses restantes del crédito
    """
    try:
        meses_transcurridos = calcular_meses_transcurridos(fecha_inicio_str)
        return max(0, plazo_meses - meses_transcurridos)
    except Exception:
        return 0

def generar_labels_meses(fecha_inicio_str, plazo_meses):
    """
    Genera etiquetas de meses para gráficos de amortización
    """
    try:
        fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
        labels = []

        for i in range(plazo_meses):
            fecha_mes = fecha_inicio + timedelta(days=30*i)
            labels.append(fecha_mes.strftime("%b %Y"))

        return labels
    except Exception:
        return []

def obtener_total_pagado_credito(credito_id):
    """
    Obtiene el total pagado de un crédito específico
    """
    try:
        conn = get_db_connection()
        total = conn.execute(
            "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
            (credito_id,)
        ).fetchone()[0] or 0
        conn.close()
        return total
    except Exception:
        return 0

def actualizar_estado_credito(credito_id):
    """
    Actualiza automáticamente el estado del crédito basado en los pagos
    """
    try:
        conn = get_db_connection()

        # Obtener información del crédito
        credito = conn.execute("SELECT * FROM creditos WHERE id = ?", (credito_id,)).fetchone()
        if not credito:
            conn.close()
            return False

        # Calcular total pagado
        total_pagado = conn.execute(
            "SELECT SUM(monto) FROM pagos_credito WHERE credito_id = ?",
            (credito_id,)
        ).fetchone()[0] or 0

        # Determinar nuevo estado
        nuevo_estado = credito["estado"]
        if total_pagado >= credito["monto_total"] and credito["estado"] == "Activo":
            nuevo_estado = "Liquidado"
        elif total_pagado < credito["monto_total"] and credito["estado"] == "Liquidado":
            nuevo_estado = "Activo"

        # Actualizar estado si es necesario
        if nuevo_estado != credito["estado"]:
            conn.execute(
                "UPDATE creditos SET estado = ? WHERE id = ?",
                (nuevo_estado, credito_id)
            )
            conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f"Error al actualizar estado del crédito: {e}")
        return False

def formatear_moneda(monto):
    """
    Formatea un monto como moneda
    """
    try:
        return f"${monto:,.2f}"
    except:
        return "$0.00"

def validar_archivo_comprobante(archivo):
    """
    Valida que el archivo de comprobante sea del tipo correcto y tamaño adecuado
    """
    if not archivo or archivo.filename == "":
        return True, ""  # No hay archivo, está bien

    # Verificar extensión
    extensiones_permitidas = {'.pdf', '.jpg', '.jpeg', '.png'}
    nombre = archivo.filename.lower()
    extension = '.' + nombre.rsplit('.', 1)[1] if '.' in nombre else ''

    if extension not in extensiones_permitidas:
        return False, "Formato de archivo no permitido. Use PDF, JPG o PNG."

    # Verificar tamaño (5MB máximo)
    archivo.seek(0, 2)  # Ir al final del archivo
    tamaño = archivo.tell()
    archivo.seek(0)  # Volver al inicio

    if tamaño > 5 * 1024 * 1024:  # 5MB
        return False, "El archivo es demasiado grande. Máximo 5MB."

    return True, ""

def generar_numero_fp():
    """
    Genera un número de FP único para créditos (si lo necesitas)
    """
    import random
    import string

    timestamp = int(datetime.now().timestamp())
    random_str = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    return f"CR-{timestamp}-{random_str}"

# Función mejorada para el manejo de errores en las vistas
def manejar_error_vista(error, mensaje_usuario="Error interno del servidor"):
    """
    Maneja errores de manera consistente en las vistas
    """
    print(f"Error en vista: {error}")
    flash(mensaje_usuario, "error")
    return redirect(url_for("creditos_dashboard"))

# Función para validar datos de crédito
def validar_datos_credito(datos):
    """
    Valida los datos de un crédito antes de guardar
    """
    errores = []

    # Validar campos requeridos
    campos_requeridos = ['nombre', 'entidad', 'tipo_credito', 'monto_total',
                        'tasa_interes', 'plazo_meses', 'pago_mensual',
                        'fecha_inicio', 'fecha_final']

    for campo in campos_requeridos:
        if not datos.get(campo):
            errores.append(f"El campo {campo.replace('_', ' ')} es requerido.")

    # Validar tipos numéricos
    try:
        monto_total = float(datos.get('monto_total', 0))
        if monto_total <= 0:
            errores.append("El monto total debe ser mayor a 0.")
    except (ValueError, TypeError):
        errores.append("El monto total debe ser un número válido.")

    try:
        tasa_interes = float(datos.get('tasa_interes', 0))
        if tasa_interes < 0:
            errores.append("La tasa de interés no puede ser negativa.")
    except (ValueError, TypeError):
        errores.append("La tasa de interés debe ser un número válido.")

    try:
        plazo_meses = int(datos.get('plazo_meses', 0))
        if plazo_meses <= 0:
            errores.append("El plazo debe ser mayor a 0 meses.")
    except (ValueError, TypeError):
        errores.append("El plazo debe ser un número entero válido.")

    try:
        pago_mensual = float(datos.get('pago_mensual', 0))
        if pago_mensual <= 0:
            errores.append("El pago mensual debe ser mayor a 0.")
    except (ValueError, TypeError):
        errores.append("El pago mensual debe ser un número válido.")

    # Validar fechas
    try:
        fecha_inicio = datetime.strptime(datos.get('fecha_inicio', ''), "%Y-%m-%d")
        fecha_final = datetime.strptime(datos.get('fecha_final', ''), "%Y-%m-%d")

        if fecha_final <= fecha_inicio:
            errores.append("La fecha final debe ser posterior a la fecha de inicio.")
    except ValueError:
        errores.append("Las fechas deben tener un formato válido (YYYY-MM-DD).")

    return errores

# Función para validar datos de pago
def validar_datos_pago(datos):
    """
    Valida los datos de un pago antes de guardar
    """
    errores = []

    # Validar campos requeridos
    if not datos.get('monto'):
        errores.append("El monto es requerido.")

    if not datos.get('fecha'):
        errores.append("La fecha es requerida.")

    if not datos.get('tipo_pago'):
        errores.append("El tipo de pago es requerido.")

    # Validar monto
    try:
        monto = float(datos.get('monto', 0))
        if monto <= 0:
            errores.append("El monto debe ser mayor a 0.")
    except (ValueError, TypeError):
        errores.append("El monto debe ser un número válido.")

    # Validar fecha
    try:
        fecha = datetime.strptime(datos.get('fecha', ''), "%Y-%m-%d")
        # Verificar que la fecha no sea muy futura (más de 1 año)
        if fecha > datetime.now() + timedelta(days=365):
            errores.append("La fecha del pago no puede ser más de un año en el futuro.")
    except ValueError:
        errores.append("La fecha debe tener un formato válido (YYYY-MM-DD).")

    return errores

try:
    start_scheduler()
except Exception as e:
    print(f"Error iniciando programador: {e}")

# AGREGAR ESTAS RUTAS A TU app.py DESPUÉS DE LAS RUTAS DE CRÉDITOS

# === GESTIÓN DE PROVEEDORES ===
def get_audit_metadata():
    """
    Obtiene metadatos completos para auditoría: IP pública, IP local, MAC, User-Agent
    """
    # IP del cliente (puede ser privada si está detrás de un proxy)
    ip_cliente = request.remote_addr
    if request.headers.getlist("X-Forwarded-For"):
        ip_cliente = request.headers.getlist("X-Forwarded-For")[0]

    # IP pública real (intentar obtenerla de servicios externos)
    public_ip = ip_cliente  # Por defecto usar la IP del cliente
    try:
        # Si estamos en un servidor, intentar obtener la IP pública real
        import requests
        response = requests.get('https://api.ipify.org?format=text', timeout=2)
        if response.status_code == 200:
            public_ip = response.text.strip()
    except:
        # Si falla, usar la IP del cliente
        public_ip = ip_cliente

    # IP local del servidor
    import socket
    try:
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
    except:
        local_ip = "127.0.0.1"

    # User agent
    user_agent = request.headers.get('User-Agent', 'Unknown')

    # MAC address del servidor
    mac_address = "00:00:00:00:00:00"
    try:
        import uuid
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff)
                       for elements in range(0,2*6,2)][::-1])
        mac_address = mac
    except:
        pass

    return {
        'ip': ip_cliente,
        'publicIP': public_ip,
        'localIP': local_ip,
        'mac': mac_address,
        'user_agent': user_agent
    }

# Función para obtener campos requeridos de una tabla
def get_required_fields(cursor, table_name):
    """
    Obtiene los campos requeridos de una tabla
    """
    cursor.execute(f"SHOW COLUMNS FROM {table_name}")
    required_fields = {}
    all_fields = {}

    for col in cursor.fetchall():
        col_name = col[0]
        col_type = col[1]
        col_null = col[2]
        col_default = col[4]

        all_fields[col_name] = col_type

        # Si el campo no acepta NULL y no tiene valor por defecto, es requerido
        if col_null == 'NO' and col_default is None:
            required_fields[col_name] = col_type

    return required_fields, all_fields
@app.route("/proveedores")
def proveedores_dashboard():
    """
    Vista principal de proveedores
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    # Filtros
    busqueda = request.args.get("busqueda", "").strip()

    try:
        # Conectar a la base de datos remota de proveedores
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor(dictionary=True)

        # Query base para obtener proveedores con sus datos más recientes
        query = """
            SELECT
                i.id AS id,
                d.regID AS datos_regID,
                d.nombre AS nombre,
                d.rfc AS rfc,
                d.direccion AS direccion,
                d.referencia AS referencia,

                c.regID AS contacto_regID,
                c.contacto AS contacto,
                c.telefono AS telefono,
                c.email AS email,
                COUNT(DISTINCT p.regID) as num_metodos_pago
            FROM AD17_Proveedores.ID AS i
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.Datos
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.Datos GROUP BY provID
                )
            ) AS d ON d.provID = i.id
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.Contactos
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.Contactos GROUP BY provID
                )
            ) AS c ON c.provID = i.id
            LEFT JOIN AD17_Proveedores.MetodosDePago AS p ON p.provID = i.id
        """

        # Aplicar filtro de búsqueda si existe
        if busqueda:
            query += """
                WHERE d.nombre LIKE %s
                OR d.rfc LIKE %s
                OR c.email LIKE %s
                OR c.telefono LIKE %s
                OR d.referencia LIKE %s
            """
            busqueda_param = f"%{busqueda}%"
            params = [busqueda_param] * 5
        else:
            params = []

        query += " GROUP BY i.id ORDER BY d.nombre"

        cursor.execute(query, params)
        proveedores = cursor.fetchall()

        cursor.close()
        remote_conn.close()

    except Exception as e:
        print(f"Error al obtener proveedores: {e}")
        flash(f"Error al cargar proveedores: {e}", "error")
        proveedores = []

    return render_template(
        "proveedores_dashboard.html",
        proveedores=proveedores,
        busqueda=busqueda
    )

@app.route("/proveedores/detalle/<string:proveedor_id>")
def detalle_proveedor(proveedor_id):
    """
    Vista detallada de un proveedor con todos sus métodos de pago
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor(dictionary=True)

        # Obtener datos del proveedor
        cursor.execute("""
            SELECT
                i.id AS id,
                d.regID AS datos_regID,
                d.nombre AS nombre,
                d.rfc AS rfc,
                d.direccion AS direccion,
                d.referencia AS referencia,

                c.regID AS contacto_regID,
                c.contacto AS contacto,
                c.telefono AS telefono,
                c.email AS email
            FROM AD17_Proveedores.ID AS i
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.Datos
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.Datos GROUP BY provID
                )
            ) AS d ON d.provID = i.id
            LEFT JOIN (
                SELECT * FROM AD17_Proveedores.Contactos
                WHERE regID IN (
                    SELECT max(regID) FROM AD17_Proveedores.Contactos GROUP BY provID
                )
            ) AS c ON c.provID = i.id
            WHERE i.id = %s
        """, (proveedor_id,))

        proveedor = cursor.fetchone()

        if not proveedor:
            cursor.close()
            remote_conn.close()
            flash("Proveedor no encontrado.", "error")
            return redirect(url_for("proveedores_dashboard"))

        # Obtener todos los métodos de pago del proveedor
        cursor.execute("""
            SELECT
                p.regID AS metodo_regID,
                p.metodo AS metodo_id,
                m.forma AS metodo_nombre,
                p.banco AS banco,
                p.beneficiario AS beneficiario,
                p.clabe AS clabe
            FROM AD17_Proveedores.MetodosDePago AS p
            LEFT JOIN AD17_Proveedores.Metodos AS m ON m.regID = p.metodo
            WHERE p.provID = %s
            ORDER BY p.regID DESC
        """, (proveedor_id,))

        metodos_pago = cursor.fetchall()

        # Obtener lista de tipos de métodos de pago disponibles
        cursor.execute("SELECT regID, forma FROM AD17_Proveedores.Metodos ORDER BY forma")
        tipos_metodos = cursor.fetchall()

        cursor.close()
        remote_conn.close()

    except Exception as e:
        print(f"Error al obtener detalle del proveedor: {e}")
        flash(f"Error al cargar información: {e}", "error")
        return redirect(url_for("proveedores_dashboard"))

    return render_template(
        "detalle_proveedor.html",
        proveedor=proveedor,
        metodos_pago=metodos_pago,
        tipos_metodos=tipos_metodos
    )

# FUNCIÓN PARA VERIFICAR EL TAMAÑO DEL CAMPO ID
@app.route("/proveedores/check_id_field")
def check_id_field_size():
    """
    Verifica el tamaño y tipo del campo ID en la base de datos
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para realizar esta acción.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # Verificar estructura del campo ID
        cursor.execute("DESCRIBE AD17_Proveedores.ID")
        columns = cursor.fetchall()

        id_info = None
        for col in columns:
            if col[0] == 'id':
                id_info = col
                break

        # Obtener algunos IDs existentes como ejemplo
        cursor.execute("SELECT id FROM AD17_Proveedores.ID LIMIT 10")
        existing_ids = cursor.fetchall()

        cursor.close()
        remote_conn.close()

        return f"""
        <html>
        <head>
            <title>Información del campo ID</title>
            <style>
                body {{ background: #000; color: #fff; font-family: monospace; padding: 20px; }}
                pre {{ background: #111; padding: 20px; border: 1px solid #444; }}
            </style>
        </head>
        <body>
            <h2>Información del campo ID</h2>
            <pre>
Campo ID:
- Nombre: {id_info[0] if id_info else 'No encontrado'}
- Tipo: {id_info[1] if id_info else 'No encontrado'}
- Null: {id_info[2] if id_info else 'No encontrado'}
- Key: {id_info[3] if id_info else 'No encontrado'}
- Default: {id_info[4] if id_info else 'No encontrado'}

IDs existentes (ejemplos):
{chr(10).join([f"- {id[0]}" for id in existing_ids])}
            </pre>
            <a href="{url_for('proveedores_dashboard')}">← Volver</a>
        </body>
        </html>
        """
    except Exception as e:
        return f"Error: {e}"

# FUNCIÓN MEJORADA PARA GENERAR IDs COMPATIBLES
def generate_provider_id(cursor):
    """
    Genera un ID único compatible con la estructura de la base de datos
    """
    # Primero, verificar el tipo y tamaño del campo ID
    cursor.execute("DESCRIBE AD17_Proveedores.ID")
    id_field_info = None
    for col in cursor.fetchall():
        if col[0] == 'id':
            id_field_info = col[1]  # Tipo de dato (ej: varchar(5), int, etc.)
            break

    print(f"Tipo de campo ID: {id_field_info}")

    # Si es un campo numérico (INT, BIGINT, etc.)
    if id_field_info and any(x in id_field_info.upper() for x in ['INT', 'NUMERIC', 'DECIMAL']):
        # Obtener el máximo ID actual y sumar 1
        cursor.execute("SELECT MAX(CAST(id AS UNSIGNED)) FROM AD17_Proveedores.ID")
        max_id = cursor.fetchone()[0]
        if max_id is None:
            return "1"
        else:
            return str(int(max_id) + 1)

    # Si es VARCHAR, verificar el tamaño
    elif id_field_info and 'VARCHAR' in id_field_info.upper():
        import re
        # Extraer el tamaño del varchar (ej: varchar(5) -> 5)
        match = re.search(r'VARCHAR\((\d+)\)', id_field_info.upper())
        if match:
            max_length = int(match.group(1))

            # Si el campo es muy pequeño (ej: varchar(5)), usar números
            if max_length <= 5:
                cursor.execute("SELECT MAX(CAST(id AS UNSIGNED)) FROM AD17_Proveedores.ID WHERE id REGEXP '^[0-9]+$'")
                max_id = cursor.fetchone()[0]
                if max_id is None:
                    return "1"
                else:
                    next_id = str(int(max_id) + 1)
                    if len(next_id) > max_length:
                        raise ValueError(f"No hay más IDs disponibles (máximo {max_length} dígitos)")
                    return next_id
            else:
                # Para campos más grandes, usar formato con prefijo
                import random
                import string
                # Generar ID con formato: P#### (donde # son números)
                while True:
                    random_num = ''.join(random.choices(string.digits, k=min(4, max_length-1)))
                    new_id = f"P{random_num}"[:max_length]  # Asegurar que no exceda el límite

                    # Verificar que no exista
                    cursor.execute("SELECT COUNT(*) FROM AD17_Proveedores.ID WHERE id = %s", (new_id,))
                    if cursor.fetchone()[0] == 0:
                        return new_id

    # Por defecto, generar un ID corto
    import random
    return str(random.randint(1000, 9999))
@app.route("/proveedores/nuevo", methods=["GET", "POST"])
def nuevo_proveedor():
    """
    Crear un nuevo proveedor con ID apropiado
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        # Obtener datos del formulario
        nombre = request.form.get("nombre", "").strip()
        rfc = request.form.get("rfc", "").strip()
        direccion = request.form.get("direccion", "").strip()  # AQUÍ ESTABA EL ERROR - faltaban los paréntesis
        referencia = request.form.get("referencia", "").strip()

        contacto = request.form.get("contacto", "").strip()
        telefono = request.form.get("telefono", "").strip()
        email = request.form.get("email", "").strip()

        # NUEVOS CAMPOS para método de pago inicial
        banco = request.form.get("banco", "").strip()
        clabe = request.form.get("clabe", "").strip()

        # Validación básica
        if not nombre:
            flash("El nombre del proveedor es requerido.", "error")
            return redirect(url_for("nuevo_proveedor"))

        try:
            remote_conn = mysql.connector.connect(
                host="ad17solutions.dscloud.me",
                port=3307,
                user="IvanUriel",
                password="iuOp20!!25",
                database="AD17_Proveedores",
                charset='utf8mb4'
            )
            cursor = remote_conn.cursor()

            # Generar ID apropiado
            try:
                nuevo_id = generate_provider_id(cursor)
                print(f"ID generado: {nuevo_id}")
            except ValueError as e:
                flash(str(e), "error")
                cursor.close()
                remote_conn.close()
                return redirect(url_for("nuevo_proveedor"))

            # Obtener metadatos de auditoría
            audit_data = get_audit_metadata()

            # === INSERTAR EN TABLA ID ===
            required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.ID")

            # Preparar datos para ID
            id_data = {'id': nuevo_id}

            # Agregar campos de auditoría si existen en la tabla
            audit_mapping = {
                'ip': audit_data.get('ip', ''),
                'publicIP': audit_data.get('publicIP', ''),
                'localIP': audit_data.get('localIP', ''),
                'mac': audit_data.get('mac', ''),
                'user_agent': audit_data.get('user_agent', '')[:255]  # Limitar longitud
            }

            for field, value in audit_mapping.items():
                if field in all_fields:
                    id_data[field] = value

            # Construir y ejecutar query para ID
            fields = list(id_data.keys())
            values = list(id_data.values())
            placeholders = ', '.join(['%s'] * len(fields))

            id_query = f"INSERT INTO AD17_Proveedores.ID ({', '.join(fields)}) VALUES ({placeholders})"
            print(f"Query ID: {id_query}")
            print(f"Values: {values}")
            cursor.execute(id_query, values)

            # === INSERTAR EN TABLA DATOS ===
            required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.Datos")

            # Preparar datos básicos
            datos_data = {
                'provID': nuevo_id,
                'nombre': nombre[:255],  # Limitar longitud
                'rfc': (rfc or '')[:13],  # RFC máximo 13 caracteres
                'direccion': (direccion or '')[:500],  # Limitar longitud
                'referencia': (referencia or '')[:500]
            }

            # Agregar campos de auditoría
            for field, value in audit_mapping.items():
                if field in all_fields:
                    datos_data[field] = value

            # Construir y ejecutar query para Datos
            fields = list(datos_data.keys())
            values = list(datos_data.values())
            placeholders = ', '.join(['%s'] * len(fields))

            datos_query = f"INSERT INTO AD17_Proveedores.Datos ({', '.join(fields)}) VALUES ({placeholders})"
            cursor.execute(datos_query, values)

            # === INSERTAR EN TABLA CONTACTOS (si hay datos) ===
            if contacto or telefono or email:
                required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.Contactos")

                contactos_data = {
                    'provID': nuevo_id,
                    'contacto': (contacto or '')[:255],
                    'telefono': (telefono or '')[:20],
                    'email': (email or '')[:255]
                }

                # Agregar campos de auditoría
                for field, value in audit_mapping.items():
                    if field in all_fields:
                        contactos_data[field] = value

                # Construir y ejecutar query para Contactos
                fields = list(contactos_data.keys())
                values = list(contactos_data.values())
                placeholders = ', '.join(['%s'] * len(fields))

                contactos_query = f"INSERT INTO AD17_Proveedores.Contactos ({', '.join(fields)}) VALUES ({placeholders})"
                cursor.execute(contactos_query, values)

            # === INSERTAR MÉTODO DE PAGO INICIAL (si se proporcionaron banco y CLABE) ===
            if banco or clabe:
                required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.MetodosDePago")

                # Por defecto usar el método 1 (Transferencia) si no se especifica
                metodo_pago_data = {
                    'provID': nuevo_id,
                    'metodo': '1',  # ID del método de transferencia
                    'banco': (banco or '')[:255],
                    'beneficiario': nombre[:255],  # Usar el nombre del proveedor como beneficiario
                    'clabe': (clabe or '')[:20]
                }

                # Agregar campos de auditoría
                for field, value in audit_mapping.items():
                    if field in all_fields:
                        metodo_pago_data[field] = value

                # Construir y ejecutar query para MetodosDePago
                fields = list(metodo_pago_data.keys())
                values = list(metodo_pago_data.values())
                placeholders = ', '.join(['%s'] * len(fields))

                metodo_query = f"INSERT INTO AD17_Proveedores.MetodosDePago ({', '.join(fields)}) VALUES ({placeholders})"
                cursor.execute(metodo_query, values)

            # Confirmar transacción
            remote_conn.commit()
            cursor.close()
            remote_conn.close()

            flash(f"Proveedor creado exitosamente con ID: {nuevo_id}", "success")
            return redirect(url_for("detalle_proveedor", proveedor_id=nuevo_id))

        except Exception as e:
            print(f"Error detallado al crear proveedor: {e}")
            if 'remote_conn' in locals():
                remote_conn.rollback()
            flash(f"Error al crear proveedor: {e}", "error")
            return redirect(url_for("nuevo_proveedor"))

    return render_template("nuevo_proveedor.html")

@app.route("/proveedores/editar/<string:proveedor_id>", methods=["GET", "POST"])
def editar_proveedor(proveedor_id):
    """
    Editar información de un proveedor con auditoría completa
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor(dictionary=True)

        if request.method == "GET":
            # Obtener datos actuales
            cursor.execute("""
                SELECT
                    i.id AS id,
                    d.nombre AS nombre,
                    d.rfc AS rfc,
                    d.direccion AS direccion,
                    d.referencia AS referencia,

                    c.contacto AS contacto,
                    c.telefono AS telefono,
                    c.email AS email
                FROM AD17_Proveedores.ID AS i
                LEFT JOIN (
                    SELECT * FROM AD17_Proveedores.Datos
                    WHERE regID IN (
                        SELECT max(regID) FROM AD17_Proveedores.Datos GROUP BY provID
                    )
                ) AS d ON d.provID = i.id
                LEFT JOIN (
                    SELECT * FROM AD17_Proveedores.Contactos
                    WHERE regID IN (
                        SELECT max(regID) FROM AD17_Proveedores.Contactos GROUP BY provID
                    )
                ) AS c ON c.provID = i.id
                WHERE i.id = %s
            """, (proveedor_id,))

            proveedor = cursor.fetchone()
            cursor.close()
            remote_conn.close()

            if not proveedor:
                flash("Proveedor no encontrado.", "error")
                return redirect(url_for("proveedores_dashboard"))

            return render_template("editar_proveedor.html", proveedor=proveedor)

        else:  # POST
            # Obtener datos del formulario
            nombre = request.form.get("nombre", "").strip()
            rfc = request.form.get("rfc", "").strip()
            direccion = request.form.get("direccion", "").strip()
            referencia = request.form.get("referencia", "").strip()
            contacto = request.form.get("contacto", "").strip()
            telefono = request.form.get("telefono", "").strip()
            email = request.form.get("email", "").strip()

            # Validación básica
            if not nombre:
                flash("El nombre del proveedor es requerido.", "error")
                return redirect(url_for("editar_proveedor", proveedor_id=proveedor_id))

            # Obtener metadatos de auditoría
            audit_data = get_audit_metadata()
            audit_mapping = {
                'ip': audit_data.get('ip', ''),
                'publicIP': audit_data.get('publicIP', ''),
                'localIP': audit_data.get('localIP', ''),
                'mac': audit_data.get('mac', ''),
                'user_agent': audit_data.get('user_agent', '')[:255]
            }

            # === ACTUALIZAR DATOS ===
            required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.Datos")

            datos_data = {
                'provID': proveedor_id,
                'nombre': nombre[:255],
                'rfc': rfc[:13] if rfc else '',
                'direccion': direccion[:500] if direccion else ''
            }
            # Agregar campos de auditoría disponibles
            for field, value in audit_mapping.items():
                if field in all_fields:
                    datos_data[field] = value

            # Insertar nuevo registro en Datos (mantener historial)
            fields = list(datos_data.keys())
            values = list(datos_data.values())
            placeholders = ', '.join(['%s'] * len(fields))

            datos_query = f"INSERT INTO AD17_Proveedores.Datos ({', '.join(fields)}) VALUES ({placeholders})"

            print(f"Query Datos: {datos_query}")
            print(f"Valores Datos: {values}")

            cursor.execute(datos_query, values)
            log_database_operation("INSERT", "Datos", datos_data)

            # === ACTUALIZAR CONTACTOS ===
            required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.Contactos")

            contactos_data = {
                'provID': proveedor_id,
                'contacto': contacto[:255] if contacto else '',
                'telefono': telefono[:20] if telefono else '',
                'email': email[:255] if email else ''
            }

            # Agregar campos de auditoría disponibles
            for field, value in audit_mapping.items():
                if field in all_fields:
                    contactos_data[field] = value

            # Insertar nuevo registro en Contactos
            fields = list(contactos_data.keys())
            values = list(contactos_data.values())
            placeholders = ', '.join(['%s'] * len(fields))

            contactos_query = f"INSERT INTO AD17_Proveedores.Contactos ({', '.join(fields)}) VALUES ({placeholders})"

            print(f"Query Contactos: {contactos_query}")
            print(f"Valores Contactos: {values}")

            cursor.execute(contactos_query, values)
            log_database_operation("INSERT", "Contactos", contactos_data)

            remote_conn.commit()
            cursor.close()
            remote_conn.close()

            flash("Proveedor actualizado exitosamente.", "success")
            return redirect(url_for("detalle_proveedor", proveedor_id=proveedor_id))

    except Exception as e:
        print(f"Error al editar proveedor: {e}")
        log_database_operation("UPDATE", "Proveedor", {}, error=str(e))
        if 'remote_conn' in locals():
            remote_conn.rollback()
        flash(f"Error al editar proveedor: {e}", "error")
        return redirect(url_for("proveedores_dashboard"))


@app.route("/proveedores/agregar_metodo_pago/<string:proveedor_id>", methods=["POST"])
def agregar_metodo_pago(proveedor_id):
    """
    Agregar un nuevo método de pago a un proveedor con auditoría completa
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    # Obtener datos del formulario
    metodo_id = request.form.get("metodo_id")
    banco = request.form.get("banco", "")
    beneficiario = request.form.get("beneficiario", "")
    clabe = request.form.get("clabe", "")

    # Obtener metadatos de auditoría
    audit_data = get_audit_metadata()

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # Obtener campos requeridos y disponibles
        required_fields, all_fields = get_required_fields(cursor, "AD17_Proveedores.MetodosDePago")

        # Preparar datos base
        metodo_data = {
            'provID': proveedor_id,
            'metodo': metodo_id,
            'banco': banco,
            'beneficiario': beneficiario,
            'clabe': clabe
        }

        # Mapeo de campos de auditoría
        audit_mapping = {
            'ip': audit_data.get('ip', ''),
            'publicIP': audit_data.get('publicIP', ''),
            'localIP': audit_data.get('localIP', ''),
            'mac': audit_data.get('mac', ''),
            'user_agent': audit_data.get('user_agent', '')[:255]  # Limitar longitud
        }

        # Agregar campos de auditoría disponibles
        for field, value in audit_mapping.items():
            if field in all_fields:
                metodo_data[field] = value

        # Construir query dinámicamente
        fields = list(metodo_data.keys())
        values = list(metodo_data.values())
        placeholders = ', '.join(['%s'] * len(fields))

        query = f"INSERT INTO AD17_Proveedores.MetodosDePago ({', '.join(fields)}) VALUES ({placeholders})"

        print(f"Query para método de pago: {query}")
        print(f"Valores: {values}")

        cursor.execute(query, values)
        remote_conn.commit()

        log_database_operation("INSERT", "MetodosDePago", metodo_data)
        flash("Método de pago agregado exitosamente.", "success")

        cursor.close()
        remote_conn.close()

    except Exception as e:
        print(f"Error al agregar método de pago: {e}")
        log_database_operation("INSERT", "MetodosDePago", metodo_data, error=str(e))
        flash(f"Error al agregar método de pago: {e}", "error")
        if 'remote_conn' in locals():
            remote_conn.rollback()

    return redirect(url_for("detalle_proveedor", proveedor_id=proveedor_id))

# También actualiza la función de verificación de estructura para incluir columnas de auditoría
def verify_and_add_audit_columns():
    """
    Verifica y agrega columnas de auditoría a las tablas de proveedores si no existen
    """
    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # Tablas a verificar
        tables = ['ID', 'Datos', 'Contactos', 'MetodosDePago']

        for table in tables:
            try:
                # Obtener columnas existentes
                cursor.execute(f"SHOW COLUMNS FROM AD17_Proveedores.{table}")
                existing_columns = [col[0] for col in cursor.fetchall()]

                # Columnas de auditoría necesarias
                audit_columns = {
                    'ip': "VARCHAR(45)",
                    'mac': "VARCHAR(17)",
                    'user_agent': "TEXT",
                    'fecha_registro': "DATETIME DEFAULT CURRENT_TIMESTAMP"
                }

                # Agregar columnas faltantes
                for column, definition in audit_columns.items():
                    if column not in existing_columns:
                        print(f"Agregando columna {column} a tabla {table}")
                        cursor.execute(f"ALTER TABLE AD17_Proveedores.{table} ADD COLUMN {column} {definition}")

            except Exception as e:
                print(f"Error procesando tabla {table}: {e}")
                continue

        remote_conn.commit()
        cursor.close()
        remote_conn.close()

        print("Verificación de columnas de auditoría completada")
        return True

    except Exception as e:
        print(f"Error al verificar/agregar columnas de auditoría: {e}")
        return False

# Llamar esta función al iniciar la aplicación
# Agrega esto después de start_scheduler() al final de tu archivo
try:
    verify_and_add_audit_columns()
except Exception as e:
    print(f"Error verificando columnas de auditoría: {e}")


@app.route("/proveedores/editar_metodo_pago/<int:metodo_regID>", methods=["POST"])
def editar_metodo_pago(metodo_regID):
    """
    Editar un método de pago existente con auditoría
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    # Obtener datos del formulario
    proveedor_id = request.form.get("proveedor_id")
    metodo_id = request.form.get("metodo_id")
    banco = request.form.get("banco", "")
    beneficiario = request.form.get("beneficiario", "")
    clabe = request.form.get("clabe", "")

    # Obtener metadatos de auditoría
    audit_data = get_audit_metadata()

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # Verificar qué campos están disponibles para actualización
        cursor.execute("SHOW COLUMNS FROM AD17_Proveedores.MetodosDePago")
        available_columns = [col[0] for col in cursor.fetchall()]

        # Preparar datos de actualización
        update_data = {
            'metodo': metodo_id,
            'banco': banco,
            'beneficiario': beneficiario,
            'clabe': clabe
        }

        # Agregar campos de auditoría si están disponibles
        audit_mapping = {
            'ip': audit_data.get('ip', ''),
            'publicIP': audit_data.get('publicIP', ''),
            'localIP': audit_data.get('localIP', ''),
            'mac': audit_data.get('mac', ''),
            'user_agent': audit_data.get('user_agent', '')[:255]
        }

        for field, value in audit_mapping.items():
            if field in available_columns:
                update_data[field] = value

        # Construir query de actualización
        set_clause = ', '.join([f"{field} = %s" for field in update_data.keys()])
        values = list(update_data.values()) + [metodo_regID]

        query = f"UPDATE AD17_Proveedores.MetodosDePago SET {set_clause} WHERE regID = %s"

        print(f"Query de actualización: {query}")
        print(f"Valores: {values}")

        cursor.execute(query, values)
        remote_conn.commit()

        log_database_operation("UPDATE", "MetodosDePago", update_data)
        flash("Método de pago actualizado exitosamente.", "success")

        cursor.close()
        remote_conn.close()

    except Exception as e:
        print(f"Error al actualizar método de pago: {e}")
        log_database_operation("UPDATE", "MetodosDePago", update_data, error=str(e))
        flash(f"Error al actualizar método de pago: {e}", "error")
        if 'remote_conn' in locals():
            remote_conn.rollback()

    return redirect(url_for("detalle_proveedor", proveedor_id=proveedor_id))

@app.route("/proveedores/eliminar_metodo_pago/<int:metodo_regID>", methods=["POST"])
def eliminar_metodo_pago(metodo_regID):
    """
    Eliminar un método de pago
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    proveedor_id = request.form.get("proveedor_id")

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # Eliminar método de pago
        cursor.execute(
            "DELETE FROM AD17_Proveedores.MetodosDePago WHERE regID = %s",
            (metodo_regID,)
        )

        remote_conn.commit()
        cursor.close()
        remote_conn.close()

        flash("Método de pago eliminado exitosamente.", "success")

    except Exception as e:
        print(f"Error al eliminar método de pago: {e}")
        flash(f"Error al eliminar método de pago: {e}", "error")

    return redirect(url_for("detalle_proveedor", proveedor_id=proveedor_id))

@app.route("/proveedores/eliminar/<string:proveedor_id>", methods=["POST"])
def eliminar_proveedor(proveedor_id):
    """
    Eliminar un proveedor (marcar como inactivo)
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a esta sección.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # En lugar de eliminar, podríamos marcar como inactivo
        # Por ahora, eliminamos todos los registros asociados
        cursor.execute("DELETE FROM AD17_Proveedores.MetodosDePago WHERE provID = %s", (proveedor_id,))
        cursor.execute("DELETE FROM AD17_Proveedores.Contactos WHERE provID = %s", (proveedor_id,))
        cursor.execute("DELETE FROM AD17_Proveedores.Datos WHERE provID = %s", (proveedor_id,))
        cursor.execute("DELETE FROM AD17_Proveedores.ID WHERE id = %s", (proveedor_id,))

        remote_conn.commit()
        cursor.close()
        remote_conn.close()

        flash("Proveedor eliminado exitosamente.", "success")

    except Exception as e:
        print(f"Error al eliminar proveedor: {e}")
        flash(f"Error al eliminar proveedor: {e}", "error")

    return redirect(url_for("proveedores_dashboard"))


# También actualiza la función de verificación de estructura para incluir columnas de auditoría
def verify_and_add_audit_columns():
    """
    Verifica y agrega columnas de auditoría a las tablas de proveedores si no existen
    """
    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        # Tablas a verificar
        tables = ['ID', 'Datos', 'Contactos', 'MetodosDePago']

        for table in tables:
            try:
                # Obtener columnas existentes
                cursor.execute(f"SHOW COLUMNS FROM AD17_Proveedores.{table}")
                existing_columns = [col[0] for col in cursor.fetchall()]

                # Columnas de auditoría necesarias
                audit_columns = {
                    'ip': "VARCHAR(45)",
                    'mac': "VARCHAR(17)",
                    'user_agent': "TEXT",
                    'fecha_registro': "DATETIME DEFAULT CURRENT_TIMESTAMP"
                }

                # Agregar columnas faltantes
                for column, definition in audit_columns.items():
                    if column not in existing_columns:
                        print(f"Agregando columna {column} a tabla {table}")
                        cursor.execute(f"ALTER TABLE AD17_Proveedores.{table} ADD COLUMN {column} {definition}")

            except Exception as e:
                print(f"Error procesando tabla {table}: {e}")
                continue

        remote_conn.commit()
        cursor.close()
        remote_conn.close()

        print("Verificación de columnas de auditoría completada")
        return True

    except Exception as e:
        print(f"Error al verificar/agregar columnas de auditoría: {e}")
        return False

# Llamar esta función al iniciar la aplicación
# Agrega esto después de start_scheduler() al final de tu archivo
try:
    verify_and_add_audit_columns()
except Exception as e:
    print(f"Error verificando columnas de auditoría: {e}")


@app.route("/proveedores/check_structure")
def check_table_structure():
    """
    Verifica la estructura completa de las tablas de proveedores
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para realizar esta acción.", "error")
        return redirect(url_for("admin_dashboard"))

    structure_info = []

    try:
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()

        tables = ['ID', 'Datos', 'Contactos', 'MetodosDePago']

        for table in tables:
            try:
                cursor.execute(f"SHOW COLUMNS FROM {table}")
                columns = cursor.fetchall()
                structure_info.append(f"\n=== TABLA: {table} ===")

                for col in columns:
                    col_name = col[0]
                    col_type = col[1]
                    col_null = col[2]
                    col_key = col[3]
                    col_default = col[4]
                    col_extra = col[5]

                    info = f"{col_name} - {col_type}"
                    if col_null == 'NO':
                        info += " [REQUERIDO]"
                    if col_default is not None:
                        info += f" (default: {col_default})"
                    if col_key:
                        info += f" [{col_key}]"

                    structure_info.append(f"  • {info}")

            except Exception as e:
                structure_info.append(f"\n✗ Error al verificar tabla {table}: {e}")

        cursor.close()
        remote_conn.close()

    except Exception as e:
        structure_info.append(f"\n✗ Error de conexión: {e}")

    return """
    <html>
    <head>
        <title>Estructura de Tablas - Proveedores</title>
        <style>
            body { background: #000; color: #fff; font-family: monospace; padding: 20px; }
            pre { background: #111; padding: 20px; border: 1px solid #444; overflow: auto; }
            .required { color: #ff5252; font-weight: bold; }
            a { color: #ff9800; }
        </style>
    </head>
    <body>
        <h2>Estructura de las Tablas de Proveedores</h2>
        <p>Los campos marcados como <span class="required">[REQUERIDO]</span> deben tener un valor al insertar.</p>
        <pre>{}</pre>
        <br>
        <a href="{}">← Volver al Dashboard</a>
    </body>
    </html>
    """.format(
        '\n'.join(structure_info).replace('[REQUERIDO]', '<span class="required">[REQUERIDO]</span>'),
        url_for('proveedores_dashboard')
    )

@app.route("/proveedores/test_connection")
def test_proveedores_connection():
    """
    Endpoint para probar la conexión y estructura de la base de datos de proveedores
    """
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para realizar esta acción.", "error")
        return redirect(url_for("admin_dashboard"))

    debug_info = []

    try:
        # Probar conexión
        remote_conn = mysql.connector.connect(
            host="ad17solutions.dscloud.me",
            port=3307,
            user="IvanUriel",
            password="iuOp20!!25",
            database="AD17_Proveedores",
            charset='utf8mb4'
        )
        cursor = remote_conn.cursor()
        debug_info.append("✓ Conexión exitosa a la base de datos")

        # Verificar tablas existentes
        cursor.execute("SHOW TABLES")
        tables = cursor.fetchall()
        debug_info.append(f"Tablas encontradas: {[t[0] for t in tables]}")

        # Verificar estructura de cada tabla
        for table_name in ['ID', 'Datos', 'Contactos', 'MetodosDePago']:
            try:
                cursor.execute(f"SHOW COLUMNS FROM {table_name}")
                columns = cursor.fetchall()
                debug_info.append(f"\nTabla {table_name}:")
                for col in columns:
                    debug_info.append(f"  - {col[0]} ({col[1]})")
            except Exception as e:
                debug_info.append(f"\n✗ Error al verificar tabla {table_name}: {e}")

        # Probar inserción de prueba
        try:
            import uuid
            test_id = f"TEST_{str(uuid.uuid4())[:8]}"

            # Obtener metadatos
            audit_data = get_audit_metadata()
            debug_info.append(f"\nMetadatos de auditoría:")
            debug_info.append(f"  - IP: {audit_data['ip']}")
            debug_info.append(f"  - MAC: {audit_data['mac']}")
            debug_info.append(f"  - User-Agent: {audit_data['user_agent'][:50]}...")

            # Intentar inserción en tabla ID
            cursor.execute("INSERT INTO ID (id) VALUES (%s)", (test_id,))
            debug_info.append(f"\n✓ Inserción de prueba en tabla ID exitosa (ID: {test_id})")

            # Verificar si se insertó
            cursor.execute("SELECT * FROM ID WHERE id = %s", (test_id,))
            result = cursor.fetchone()
            if result:
                debug_info.append("✓ Registro verificado en base de datos")
            else:
                debug_info.append("✗ Registro no encontrado después de inserción")

            # Limpiar registro de prueba
            cursor.execute("DELETE FROM ID WHERE id = %s", (test_id,))
            remote_conn.commit()
            debug_info.append("✓ Registro de prueba eliminado")

        except Exception as e:
            debug_info.append(f"\n✗ Error en prueba de inserción: {e}")
            remote_conn.rollback()

        cursor.close()
        remote_conn.close()

    except Exception as e:
        debug_info.append(f"\n✗ Error de conexión: {e}")

    # Mostrar información de depuración
    return """
    <html>
    <head>
        <title>Test de Conexión - Proveedores</title>
        <style>
            body { background: #000; color: #fff; font-family: monospace; padding: 20px; }
            pre { background: #111; padding: 20px; border: 1px solid #444; overflow: auto; }
            .success { color: #4CAF50; }
            .error { color: #f44336; }
            a { color: #ff9800; }
        </style>
    </head>
    <body>
        <h2>Resultado del Test de Conexión - Base de Datos de Proveedores</h2>
        <pre>{}</pre>
        <br>
        <a href="{}">← Volver al Dashboard</a>
    </body>
    </html>
    """.format(
        '\n'.join(debug_info).replace('✓', '<span class="success">✓</span>').replace('✗', '<span class="error">✗</span>'),
        url_for('proveedores_dashboard')
    )

# MODIFICACIÓN ADICIONAL: Agregar logs detallados en las funciones de inserción
# Modifica la función nuevo_proveedor para agregar más logging:

def log_database_operation(operation, table, data, error=None):
    """
    Registra operaciones de base de datos para depuración
    """
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {operation} en {table}"

    if error:
        log_entry += f" - ERROR: {error}"
        print(f"❌ {log_entry}")
    else:
        log_entry += f" - ÉXITO"
        print(f"✅ {log_entry}")

    # También puedes guardar en un archivo de log
    try:
        with open("proveedores_operations.log", "a", encoding="utf-8") as f:
            f.write(f"{log_entry}\n")
            if data:
                f.write(f"   Datos: {data}\n")
    except:
        pass
# 4. Agregar nueva ruta para el reporte consolidado:
# 4. Agregar nueva ruta para el reporte consolidado:
@app.route("/exportar_reporte_consolidado")
def exportar_reporte_consolidado():
    """
    Genera un reporte consolidado optimizado con semáforo de colores y análisis de créditos en hoja separada
    """
    # Verifica permisos
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para exportar reportes.", "error")
        return redirect(url_for("admin_dashboard"))

    # Determinar la fecha de inicio de la semana
    fecha_inicio_str = request.args.get('fecha_inicio')

    if fecha_inicio_str:
        try:
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
            if fecha_inicio.weekday() != 0:
                fecha_inicio = fecha_inicio - timedelta(days=fecha_inicio.weekday())
        except ValueError:
            flash("Formato de fecha inválido. Usando semana actual.", "warning")
            hoy = datetime.now()
            fecha_inicio = hoy - timedelta(days=hoy.weekday())
    else:
        hoy = datetime.now()
        fecha_inicio = hoy - timedelta(days=hoy.weekday())

    fecha_fin = fecha_inicio + timedelta(days=6, hours=23, minutes=59, seconds=59)

    # Configurar título
    fecha_inicio_str = fecha_inicio.strftime("%d-%m-%Y")
    fecha_fin_str = fecha_fin.strftime("%d-%m-%Y")
    titulo_reporte = f"Reporte Consolidado: {fecha_inicio_str} al {fecha_fin_str}"

    try:
        # Obtener TODOS los registros de solicitudes
        conn = get_db_connection()

        # Verificar columnas disponibles
        cursor = conn.execute("PRAGMA table_info(solicitudes)")
        columns = [row["name"] for row in cursor.fetchall()]

        # Query principal para obtener todas las solicitudes
        query = """
            SELECT *,
                   CASE
                       WHEN estado IN ('Liquidado', 'Liquidado con anticipo', 'Liquidacion total')
                            AND fecha_liquidado IS NOT NULL THEN fecha_liquidado
                       WHEN estado IN ('Aprobado', 'Aprobado con anticipo')
                            AND fecha_aprobado IS NOT NULL THEN fecha_aprobado
                       WHEN estado = 'Programado' THEN fecha_limite
                       WHEN fecha_ultimo_cambio IS NOT NULL THEN fecha_ultimo_cambio
                       ELSE fecha_limite
                   END as fecha_relevante
            FROM solicitudes
            ORDER BY fecha_limite, estado
        """

        if 'fecha_liquidado' not in columns or 'fecha_aprobado' not in columns:
            query = """
                SELECT *, fecha_limite as fecha_relevante
                FROM solicitudes
                ORDER BY fecha_limite, estado
            """

        rows = conn.execute(query).fetchall()

        # Obtener TODOS los pendientes (incluye pendientes con anticipo)
        all_pendientes = conn.execute("""
            SELECT * FROM solicitudes
            WHERE LOWER(estado) IN ('pendiente', 'pendiente con anticipo')
            ORDER BY fecha_limite
        """).fetchall()

        # Obtener datos de créditos para la segunda hoja
        creditos = conn.execute("""
            SELECT c.*,
                   (SELECT SUM(monto) FROM pagos_credito WHERE credito_id = c.id) as total_pagado
            FROM creditos c
            ORDER BY c.estado, c.fecha_registro DESC
        """).fetchall()

        # Obtener pagos de créditos del período con información del beneficiario
        pagos_creditos_periodo = conn.execute("""
            SELECT pc.*,
                   c.nombre as credito_nombre,
                   c.entidad,
                   c.tipo_credito,
                   c.numero_cuenta,
                   c.contacto,
                   COALESCE(pc.descripcion, '') as descripcion_pago,
                   CASE
                       WHEN pc.descripcion LIKE '%a:%' THEN SUBSTR(pc.descripcion, INSTR(pc.descripcion, 'a:') + 2)
                       WHEN pc.descripcion LIKE '%para:%' THEN SUBSTR(pc.descripcion, INSTR(pc.descripcion, 'para:') + 5)
                       WHEN c.contacto IS NOT NULL AND c.contacto != '' THEN c.contacto
                       ELSE c.entidad
                   END as beneficiario
            FROM pagos_credito pc
            JOIN creditos c ON pc.credito_id = c.id
            WHERE date(pc.fecha) >= ? AND date(pc.fecha) <= ?
            ORDER BY pc.fecha DESC
        """, (fecha_inicio.strftime("%Y-%m-%d"), fecha_fin.strftime("%Y-%m-%d"))).fetchall()

        conn.close()

        if not rows:
            flash("No hay datos para generar el reporte.", "warning")
            return redirect(url_for("admin_dashboard"))

        # Convertir a DataFrame
        df = pd.DataFrame([dict(r) for r in rows])
        df_all_pendientes = pd.DataFrame([dict(r) for r in all_pendientes])

        # Preparar campo de monto como numérico
        df['monto'] = pd.to_numeric(df['monto'], errors='coerce').fillna(0)
        if not df_all_pendientes.empty:
            df_all_pendientes['monto'] = pd.to_numeric(df_all_pendientes['monto'], errors='coerce').fillna(0)

        # Convertir fecha_relevante a datetime
        df['fecha_relevante'] = pd.to_datetime(df['fecha_relevante'], errors='coerce')
# Reemplaza la línea problemática por:
        df['fecha_limite_dt'] = pd.to_datetime(df['fecha_limite'], format="%Y-%m-%d", errors='coerce')

        # Si no hay fecha_relevante válida, usar fecha_limite
        df.loc[df['fecha_relevante'].isna(), 'fecha_relevante'] = df.loc[df['fecha_relevante'].isna(), 'fecha_limite_dt']

        # === FILTRADO MEJORADO POR FECHA ===

        # 1. Para liquidados
        mask_liquidados = (
            (df['estado'].str.lower().isin(['liquidado', 'liquidado con anticipo', 'liquidacion total'])) &
            (df['fecha_relevante'] >= fecha_inicio) &
            (df['fecha_relevante'] <= fecha_fin)
        )

        # 2. Para aprobados
        mask_aprobados = (
            (df['estado'].str.lower().isin(['aprobado', 'aprobado con anticipo'])) &
            (df['fecha_relevante'] >= fecha_inicio) &
            (df['fecha_relevante'] <= fecha_fin)
        )

        # 3. Para programados
        mask_programados = (
            (df['estado'].str.lower() == 'programado') &
            (df['fecha_limite_dt'] >= fecha_inicio) &
            (df['fecha_limite_dt'] <= fecha_fin)
        )

        # 4. Para pendientes del período
        mask_pendientes_periodo = (
            (df['estado'].str.lower().isin(['pendiente', 'pendiente con anticipo'])) &
            (df['fecha_limite_dt'] >= fecha_inicio) &
            (df['fecha_limite_dt'] <= fecha_fin)
        )

        # Aplicar filtros
        df_liquidados = df.loc[mask_liquidados].copy()
        df_aprobados = df.loc[mask_aprobados].copy()
        df_programados = df.loc[mask_programados].copy()
        df_pendientes_periodo = df.loc[mask_pendientes_periodo].copy()

        # Excluir pendientes del período de los pendientes totales
        if not df_pendientes_periodo.empty and not df_all_pendientes.empty:
            ids_pendientes_periodo = df_pendientes_periodo['id'].tolist()
            df_all_pendientes = df_all_pendientes[~df_all_pendientes['id'].isin(ids_pendientes_periodo)].copy()

        # === VERIFICACIÓN DE DATOS ===
        print(f"DEBUG - Total solicitudes: {len(df)}")
        print(f"DEBUG - Liquidados período: {len(df_liquidados)}")
        print(f"DEBUG - Aprobados período: {len(df_aprobados)}")
        print(f"DEBUG - Programados período: {len(df_programados)}")
        print(f"DEBUG - Pendientes período: {len(df_pendientes_periodo)}")
        print(f"DEBUG - Pendientes históricos (sin duplicados): {len(df_all_pendientes)}")
        estados_unicos = df['estado'].unique()
        print(f"DEBUG - Estados únicos en BD: {estados_unicos}")

        pendientes_con_anticipo = df[
            (df['estado'].str.lower().isin(['pendiente', 'pendiente con anticipo'])) &
            (df['anticipo'].str.lower() == 'si')
        ]
        print(f"DEBUG - Pendientes con anticipo total: {len(pendientes_con_anticipo)}")

        # Generar Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Formatos
            header_format = workbook.add_format({
                'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#2B579A',
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 11
            })
            title_format = workbook.add_format({
                'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter',
                'font_color': '#FFFFFF', 'bg_color': '#1F4E79', 'border': 2
            })
            section_liquidados_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': '#FFFFFF',
                'bg_color': '#70AD47', 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            section_aprobados_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': '#FFFFFF',
                'bg_color': '#5B9BD5', 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            section_programados_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': '#000000',
                'bg_color': '#FFC000', 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            section_pendientes_periodo_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': '#FFFFFF',
                'bg_color': '#ED7D31', 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            section_pendientes_todos_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': '#FFFFFF',
                'bg_color': '#C00000', 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            section_info_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'font_color': '#FFFFFF',
                'bg_color': '#7C7C7C', 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            money_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1, 'align': 'right'})
            cell_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
            date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})
            datetime_format = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm', 'border': 1, 'align': 'center'})
            total_format = workbook.add_format({
                'bold': True, 'num_format': '$#,##0.00', 'border': 1, 'bg_color': '#E7E6E6', 'align': 'right'
            })
            total_importante_format = workbook.add_format({
                'bold': True, 'num_format': '$#,##0.00', 'border': 2, 'bg_color': '#FFE6E6',
                'font_color': '#C00000', 'align': 'right', 'font_size': 12
            })
            anticipo_format = workbook.add_format({
                'border': 1, 'align': 'left', 'valign': 'top', 'bg_color': '#FFF2CC', 'text_wrap': True
            })

            # === HOJA 1: PAGOS ===
            worksheet = workbook.add_worksheet('Pagos')
            worksheet.merge_range('A1:M1', titulo_reporte, title_format)
            worksheet.set_row(0, 35)
            worksheet.write('A3', 'Generado:', cell_format)
            worksheet.write('B3', datetime.now().strftime("%d/%m/%Y %H:%M"), cell_format)
            worksheet.write('A4', 'Período:', cell_format)
            worksheet.write('B4', f'{fecha_inicio_str} al {fecha_fin_str}', cell_format)

            current_row = 6

            # Función de sección (usa merge_range numérico para evitar off-by-one)
            def write_section_mejorada(df_section, section_title, start_row, section_format_to_use, include_anticipo_info=True):
                if df_section.empty:
                    worksheet.merge_range(start_row, 0, start_row, 12, f'{section_title} - Sin registros', section_format_to_use)
                    return start_row + 2

                # Título de sección
                worksheet.merge_range(start_row, 0, start_row, 12, section_title, section_format_to_use)
                start_row += 2

                # Encabezados
                headers = ['FP', 'Solicitante', 'Destinatario', 'Departamento',
                           'Tipo Solicitud', 'Tipo Pago', 'Banco', 'CLABE',
                           'Monto', 'Anticipo', 'Monto Anticipo', 'Fecha Solicitud', 'Fecha Límite']
                for col, header in enumerate(headers):
                    worksheet.write(start_row, col, header, header_format)

                # Datos
                data_start_row = start_row + 1
                total = 0
                total_anticipos = 0

                for idx, (_, record) in enumerate(df_section.iterrows()):
                    r = data_start_row + idx
                    tiene_anticipo = str(record.get('anticipo', '')).lower() == 'si'
                    cell_fmt = anticipo_format if tiene_anticipo else cell_format

                    worksheet.write(r, 0, record.get('fp', ''), cell_fmt)
                    worksheet.write(r, 1, record.get('nombre', ''), cell_fmt)
                    worksheet.write(r, 2, record.get('destinatario', ''), cell_fmt)
                    worksheet.write(r, 3, record.get('departamento', ''), cell_fmt)
                    worksheet.write(r, 4, record.get('tipo_solicitud', ''), cell_fmt)
                    worksheet.write(r, 5, record.get('tipo_pago', ''), cell_fmt)
                    worksheet.write(r, 6, record.get('banco', ''), cell_fmt)
                    worksheet.write(r, 7, record.get('clabe', ''), cell_fmt)
                    worksheet.write(r, 8, record.get('monto', 0), money_format)

                    anticipo_text = 'Sí' if tiene_anticipo else 'No'
                    if tiene_anticipo:
                        porcentaje = record.get('porcentaje_anticipo', 0)
                        anticipo_text += f' ({porcentaje}%)'
                    worksheet.write(r, 9, anticipo_text, cell_fmt)

                    monto_anticipo = record.get('monto_anticipo', 0) if tiene_anticipo else 0
                    worksheet.write(r, 10, monto_anticipo, money_format)
                    worksheet.write(r, 11, record.get('fecha', ''), datetime_format)
                    worksheet.write(r, 12, record.get('fecha_limite', ''), date_format)

                    total += record.get('monto', 0)
                    total_anticipos += monto_anticipo

                # Fila total (justo debajo de los datos)
                total_row = data_start_row + len(df_section)
                worksheet.merge_range(total_row, 0, total_row, 7, 'TOTAL', total_format)
                worksheet.write(total_row, 8, total, total_format)
                worksheet.write(total_row, 9, '', total_format)
                worksheet.write(total_row, 10, total_anticipos, total_format)
                worksheet.write(total_row, 11, '', total_format)
                worksheet.write(total_row, 12, '', total_format)

                return total_row + 3

            # Secciones
            current_row = write_section_mejorada(
                df_liquidados,
                f'💰 PAGOS LIQUIDADOS (Semana {fecha_inicio_str} al {fecha_fin_str}) - {len(df_liquidados)} registros',
                current_row,
                section_liquidados_format
            )
            current_row = write_section_mejorada(
                df_aprobados,
                f'✅ PAGOS APROBADOS (Semana {fecha_inicio_str} al {fecha_fin_str}) - {len(df_aprobados)} registros',
                current_row,
                section_aprobados_format
            )
            current_row = write_section_mejorada(
                df_programados,
                f'📅 PAGOS PROGRAMADOS (Semana {fecha_inicio_str} al {fecha_fin_str}) - {len(df_programados)} registros',
                current_row,
                section_programados_format
            )
            current_row = write_section_mejorada(
                df_pendientes_periodo,
                f'⏳ PAGOS PENDIENTES (Semana {fecha_inicio_str} al {fecha_fin_str}) - {len(df_pendientes_periodo)} registros',
                current_row,
                section_pendientes_periodo_format
            )
            current_row = write_section_mejorada(
                df_all_pendientes,
                f'🚨 OTROS PAGOS PENDIENTES HISTÓRICOS (FUERA DEL PERÍODO ACTUAL) - {len(df_all_pendientes)} registros',
                current_row,
                section_pendientes_todos_format
            )

            # Resumen final
            current_row += 2
            worksheet.merge_range(current_row, 0, current_row, 12, 'RESUMEN GENERAL DE PAGOS', section_info_format)
            current_row += 2

            total_liquidados = df_liquidados['monto'].sum() if not df_liquidados.empty else 0
            total_aprobados = df_aprobados['monto'].sum() if not df_aprobados.empty else 0
            total_programados = df_programados['monto'].sum() if not df_programados.empty else 0
            total_pendientes_periodo = df_pendientes_periodo['monto'].sum() if not df_pendientes_periodo.empty else 0
            total_pendientes_todos = df_all_pendientes['monto'].sum() if not df_all_pendientes.empty else 0

            total_anticipos_liquidados = df_liquidados[df_liquidados['anticipo'].str.lower() == 'si']['monto_anticipo'].sum() if not df_liquidados.empty else 0
            total_anticipos_aprobados = df_aprobados[df_aprobados['anticipo'].str.lower() == 'si']['monto_anticipo'].sum() if not df_aprobados.empty else 0
            total_anticipos_pendientes = df_all_pendientes[df_all_pendientes['anticipo'].str.lower() == 'si']['monto_anticipo'].sum() if not df_all_pendientes.empty else 0

            gran_total_periodo = total_liquidados + total_aprobados + total_pendientes_periodo + total_programados

            headers_resumen = ['Estado', 'Cantidad', 'Monto Total', 'Con Anticipo', 'Monto Anticipos']
            for col, header in enumerate(headers_resumen):
                worksheet.write(current_row, col, header, header_format)
            current_row += 1

            liquidados_con_anticipo = len(df_liquidados[df_liquidados['anticipo'].str.lower() == 'si']) if not df_liquidados.empty else 0
            aprobados_con_anticipo = len(df_aprobados[df_aprobados['anticipo'].str.lower() == 'si']) if not df_aprobados.empty else 0
            programados_con_anticipo = len(df_programados[df_programados['anticipo'].str.lower() == 'si']) if not df_programados.empty else 0
            pendientes_periodo_con_anticipo = len(df_pendientes_periodo[df_pendientes_periodo['anticipo'].str.lower() == 'si']) if not df_pendientes_periodo.empty else 0
            pendientes_todos_con_anticipo = len(df_all_pendientes[df_all_pendientes['anticipo'].str.lower() == 'si']) if not df_all_pendientes.empty else 0

            resumen_data = [
                ['✅ Liquidados (semana)', len(df_liquidados), total_liquidados, liquidados_con_anticipo, total_anticipos_liquidados, section_liquidados_format],
                ['✅ Aprobados (semana)', len(df_aprobados), total_aprobados, aprobados_con_anticipo, total_anticipos_aprobados, section_aprobados_format],
                ['📅 Programados (semana)', len(df_programados), total_programados, programados_con_anticipo, 0, section_programados_format],
                ['⏳ Pendientes (semana)', len(df_pendientes_periodo), total_pendientes_periodo, pendientes_periodo_con_anticipo, 0, section_pendientes_periodo_format],
                ['SUBTOTAL SEMANA', len(df_liquidados) + len(df_aprobados) + len(df_programados) + len(df_pendientes_periodo), gran_total_periodo,
                 liquidados_con_anticipo + aprobados_con_anticipo + programados_con_anticipo + pendientes_periodo_con_anticipo,
                 total_anticipos_liquidados + total_anticipos_aprobados, total_format],
                ['', '', '', '', '', None],
                ['🚨 PENDIENTES HISTÓRICOS (fuera del período)', len(df_all_pendientes), total_pendientes_todos,
                 pendientes_todos_con_anticipo, total_anticipos_pendientes, total_importante_format]
            ]
            for i, (estado, cantidad, monto, con_anticipo, monto_anticipos, formato_fila) in enumerate(resumen_data):
                row = current_row + i
                if i == 5:  # Línea vacía
                    for col in range(5):
                        worksheet.write(row, col, '', cell_format)
                else:
                    if formato_fila and i not in [4, 6]:
                        for col, valor in enumerate([estado, cantidad, monto, con_anticipo, monto_anticipos]):
                            worksheet.write(row, col, valor, money_format if col in (2, 4) and formato_fila == total_format else formato_fila)
                    elif i == 6:
                        for col, valor in enumerate([estado, cantidad, monto, con_anticipo, monto_anticipos]):
                            worksheet.write(row, col, valor, total_importante_format)
                    else:
                        for col, valor in enumerate([estado, cantidad, monto, con_anticipo, monto_anticipos]):
                            worksheet.write(row, col, valor, money_format if col in (2, 4) else total_format)

            # Ajustes de columnas
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 25)
            worksheet.set_column('C:C', 25)
            worksheet.set_column('D:D', 20)
            worksheet.set_column('E:E', 18)
            worksheet.set_column('F:F', 18)
            worksheet.set_column('G:G', 20)
            worksheet.set_column('H:H', 20)
            worksheet.set_column('I:I', 15)
            worksheet.set_column('J:J', 15)
            worksheet.set_column('K:K', 15)
            worksheet.set_column('L:L', 18)
            worksheet.set_column('M:M', 18)

            # === HOJA 2: ANÁLISIS DE CRÉDITOS ===
            worksheet_creditos = workbook.add_worksheet('Análisis de Créditos')
            worksheet_creditos.merge_range('A1:J1', f'Análisis de Créditos - {fecha_inicio_str} al {fecha_fin_str}', title_format)
            worksheet_creditos.set_row(0, 35)

            current_row = 3
            creditos_list = [dict(c) for c in creditos]
            if creditos_list:
                total_creditos = len(creditos_list)
                creditos_activos = [c for c in creditos_list if c['estado'] == 'Activo']
                creditos_liquidados = [c for c in creditos_list if c['estado'] == 'Liquidado']
                monto_total_creditos = sum(c['monto_total'] for c in creditos_list)
                monto_activos = sum(c['monto_total'] for c in creditos_activos)
                monto_liquidados = sum(c['monto_total'] for c in creditos_liquidados)
                total_pagado_global = sum(c['total_pagado'] or 0 for c in creditos_list)
                saldo_pendiente_global = monto_activos - sum(c['total_pagado'] or 0 for c in creditos_activos)

                worksheet_creditos.merge_range(current_row, 0, current_row, 9, 'RESUMEN GENERAL DE CRÉDITOS', section_info_format)
                current_row += 2

                stats_headers = ['Concepto', 'Cantidad', 'Monto']
                for col, header in enumerate(stats_headers):
                    worksheet_creditos.write(current_row, col, header, header_format)
                current_row += 1

                stats_data = [
                    ['Total de Créditos', total_creditos, monto_total_creditos],
                    ['Créditos Activos', len(creditos_activos), monto_activos],
                    ['Créditos Liquidados', len(creditos_liquidados), monto_liquidados],
                    ['Total Pagado (histórico)', '-', total_pagado_global],
                    ['Saldo Pendiente Total', '-', saldo_pendiente_global]
                ]
                for stat in stats_data:
                    worksheet_creditos.write(current_row, 0, stat[0], cell_format)
                    worksheet_creditos.write(current_row, 1, stat[1] if stat[1] != '-' else '-', cell_format)
                    worksheet_creditos.write(current_row, 2, stat[2], money_format)
                    current_row += 1

                current_row += 2

                if creditos_activos:
                    worksheet_creditos.merge_range(current_row, 0, current_row, 9, 'CRÉDITOS ACTIVOS', section_aprobados_format)
                    current_row += 2

                    headers_creditos = ['Nombre', 'Entidad', 'Tipo', 'Monto Total',
                                        'Tasa %', 'Pago Mensual', 'Total Pagado',
                                        'Saldo Pendiente','Fecha Inicio', 'Fecha Final']
                    for col, header in enumerate(headers_creditos):
                        worksheet_creditos.write(current_row, col, header, header_format)
                    current_row += 1

                    for credito in creditos_activos:
                        saldo = credito['monto_total'] - (credito['total_pagado'] or 0)
                        worksheet_creditos.write(current_row, 0, credito['nombre'], cell_format)
                        worksheet_creditos.write(current_row, 1, credito['entidad'], cell_format)
                        worksheet_creditos.write(current_row, 2, credito['tipo_credito'], cell_format)
                        worksheet_creditos.write(current_row, 3, credito['monto_total'], money_format)
                        worksheet_creditos.write(current_row, 4, credito['tasa_interes'], cell_format)
                        worksheet_creditos.write(current_row, 5, credito['pago_mensual'], money_format)
                        worksheet_creditos.write(current_row, 6, credito['total_pagado'] or 0, money_format)
                        worksheet_creditos.write(current_row, 7, saldo, money_format)
                        worksheet_creditos.write(current_row, 8, credito['fecha_inicio'], date_format)
                        worksheet_creditos.write(current_row, 9, credito['fecha_final'], date_format)
                        current_row += 1

                current_row += 2

                if pagos_creditos_periodo:
                    worksheet_creditos.merge_range(
                        current_row, 0, current_row, 10,
                        f'PAGOS DE CRÉDITOS REALIZADOS ({fecha_inicio_str} al {fecha_fin_str})',
                        section_liquidados_format
                    )
                    current_row += 2

                    headers_pagos = ['Fecha', 'Crédito', 'Entidad', 'Beneficiario/Pagado a',
                                     'Tipo', 'Monto', 'Referencia', 'Tipo Pago', 'Descripción']
                    for col, header in enumerate(headers_pagos):
                        worksheet_creditos.write(current_row, col, header, header_format)
                    current_row += 1

                    total_pagos_periodo = 0
                    for pago in pagos_creditos_periodo:
                        pago_dict = dict(pago)
                        worksheet_creditos.write(current_row, 0, pago_dict['fecha'], date_format)
                        worksheet_creditos.write(current_row, 1, pago_dict['credito_nombre'], cell_format)
                        worksheet_creditos.write(current_row, 2, pago_dict['entidad'], cell_format)
                        worksheet_creditos.write(current_row, 3, pago_dict.get('beneficiario', ''), cell_format)
                        worksheet_creditos.write(current_row, 4, pago_dict['tipo_credito'], cell_format)
                        worksheet_creditos.write(current_row, 5, pago_dict['monto'], money_format)
                        worksheet_creditos.write(current_row, 6, pago_dict.get('referencia', ''), cell_format)
                        worksheet_creditos.write(current_row, 7, pago_dict['tipo_pago'], cell_format)
                        worksheet_creditos.write(current_row, 8, pago_dict.get('descripcion_pago', ''), cell_format)
                        total_pagos_periodo += pago_dict['monto']
                        current_row += 1

                    worksheet_creditos.merge_range(current_row, 0, current_row, 4, 'TOTAL PAGOS DEL PERÍODO', total_format)
                    worksheet_creditos.write(current_row, 5, total_pagos_periodo, total_format)
                    worksheet_creditos.merge_range(current_row, 6, current_row, 8, '', total_format)

                worksheet_creditos.set_column('A:A', 25)
                worksheet_creditos.set_column('B:B', 20)
                worksheet_creditos.set_column('C:C', 20)
                worksheet_creditos.set_column('D:D', 25)
                worksheet_creditos.set_column('E:E', 15)
                worksheet_creditos.set_column('F:F', 15)
                worksheet_creditos.set_column('G:G', 15)
                worksheet_creditos.set_column('H:H', 15)
                worksheet_creditos.set_column('I:I', 12)
                worksheet_creditos.set_column('J:J', 12)
                worksheet_creditos.set_column('K:K', 30)

            else:
                worksheet_creditos.merge_range(3, 0, 3, 9, 'No hay créditos registrados', section_info_format)

            # === HOJA 3: ANÁLISIS DETALLADO DE ANTICIPOS ===
            worksheet_anticipos = workbook.add_worksheet('Análisis de Anticipos')
            worksheet_anticipos.merge_range('A1:L1', f'Análisis Detallado de Anticipos - {fecha_inicio_str} al {fecha_fin_str}', title_format)
            worksheet_anticipos.set_row(0, 35)

            current_row = 3
            solicitudes_con_anticipo = df[df['anticipo'].str.lower() == 'si'].copy()

            if not solicitudes_con_anticipo.empty:
                anticipos_pendientes = solicitudes_con_anticipo[
                    solicitudes_con_anticipo['estado'].str.lower().isin(['pendiente', 'pendiente con anticipo'])
                ].copy()
                anticipos_aprobados = solicitudes_con_anticipo[
                    solicitudes_con_anticipo['estado'].str.lower().isin(['aprobado', 'aprobado con anticipo'])
                ].copy()
                anticipos_liquidados = solicitudes_con_anticipo[
                    solicitudes_con_anticipo['estado'].str.lower().isin(['liquidado', 'liquidado con anticipo', 'liquidacion total'])
                ].copy()

                def write_anticipo_section(df_anticipo, title, start_row, format_style):
                    if df_anticipo.empty:
                        worksheet_anticipos.merge_range(start_row, 0, start_row, 11, f'{title} - Sin registros', format_style)
                        return start_row + 2

                    worksheet_anticipos.merge_range(start_row, 0, start_row, 11, f'{title} - {len(df_anticipo)} registros', format_style)
                    start_row += 2

                    headers_anticipo = ['FP', 'Solicitante', 'Destinatario', 'Monto Total',
                                        'Tipo Anticipo', '% Anticipo', 'Monto Anticipo',
                                        'Monto Restante', 'Estado', 'Fecha Límite',
                                        'Departamento', 'Tipo Solicitud']
                    for col, header in enumerate(headers_anticipo):
                        worksheet_anticipos.write(start_row, col, header, header_format)

                    data_start_row = start_row + 1
                    total_anticipos = 0
                    total_restantes = 0
                    total_general = 0

                    for idx, (_, record) in enumerate(df_anticipo.iterrows()):
                        r = data_start_row + idx
                        worksheet_anticipos.write(r, 0, record.get('fp', ''), cell_format)
                        worksheet_anticipos.write(r, 1, record.get('nombre', ''), cell_format)
                        worksheet_anticipos.write(r, 2, record.get('destinatario', ''), cell_format)
                        worksheet_anticipos.write(r, 3, record.get('monto', 0), money_format)
                        worksheet_anticipos.write(r, 4, record.get('tipo_anticipo', 'porcentaje'), cell_format)
                        worksheet_anticipos.write(r, 5, f"{record.get('porcentaje_anticipo', 0)}%", cell_format)

                        monto_anticipo = record.get('monto_anticipo', 0)
                        monto_restante = record.get('monto_restante', 0)
                        monto_total = record.get('monto', 0)

                        worksheet_anticipos.write(r, 6, monto_anticipo, money_format)
                        worksheet_anticipos.write(r, 7, monto_restante, money_format)
                        worksheet_anticipos.write(r, 8, record.get('estado', ''), cell_format)
                        worksheet_anticipos.write(r, 9, record.get('fecha_limite', ''), date_format)
                        worksheet_anticipos.write(r, 10, record.get('departamento', ''), cell_format)
                        worksheet_anticipos.write(r, 11, record.get('tipo_solicitud', ''), cell_format)

                        total_anticipos += monto_anticipo
                        total_restantes += monto_restante
                        total_general += monto_total

                    total_row = data_start_row + len(df_anticipo)
                    worksheet_anticipos.merge_range(total_row, 0, total_row, 2, 'TOTALES', total_format)
                    worksheet_anticipos.write(total_row, 3, total_general, total_format)
                    worksheet_anticipos.write(total_row, 4, '', total_format)
                    worksheet_anticipos.write(total_row, 5, '', total_format)
                    worksheet_anticipos.write(total_row, 6, total_anticipos, total_format)
                    worksheet_anticipos.write(total_row, 7, total_restantes, total_format)
                    worksheet_anticipos.write(total_row, 8, '', total_format)
                    worksheet_anticipos.write(total_row, 9, '', total_format)
                    worksheet_anticipos.write(total_row, 10, '', total_format)
                    worksheet_anticipos.write(total_row, 11, '', total_format)

                    return total_row + 3

                current_row = write_anticipo_section(
                    anticipos_pendientes,
                    '🚨 ANTICIPOS PENDIENTES (REQUIEREN ATENCIÓN URGENTE)',
                    current_row,
                    section_pendientes_todos_format
                )
                current_row = write_anticipo_section(
                    anticipos_aprobados,
                    '✅ ANTICIPOS APROBADOS (EN PROCESO DE PAGO)',
                    current_row,
                    section_aprobados_format
                )
                current_row = write_anticipo_section(
                    anticipos_liquidados,
                    '💰 ANTICIPOS LIQUIDADOS (COMPLETADOS)',
                    current_row,
                    section_liquidados_format
                )

                current_row += 2
                worksheet_anticipos.merge_range(current_row, 0, current_row, 11, 'RESUMEN GENERAL DE ANTICIPOS', section_info_format)
                current_row += 2

                total_anticipos_pendientes = anticipos_pendientes['monto_anticipo'].sum() if not anticipos_pendientes.empty else 0
                total_anticipos_aprobados = anticipos_aprobados['monto_anticipo'].sum() if not anticipos_aprobados.empty else 0
                total_anticipos_liquidados = anticipos_liquidados['monto_anticipo'].sum() if not anticipos_liquidados.empty else 0
                total_restantes_pendientes = anticipos_pendientes['monto_restante'].sum() if not anticipos_pendientes.empty else 0
                total_restantes_aprobados = anticipos_aprobados['monto_restante'].sum() if not anticipos_aprobados.empty else 0
                total_restantes_liquidados = anticipos_liquidados['monto_restante'].sum() if not anticipos_liquidados.empty else 0

                resumen_headers = ['Estado', 'Cantidad', 'Total Anticipos', 'Total Restantes', 'Total General']
                for col, header in enumerate(resumen_headers):
                    worksheet_anticipos.write(current_row, col, header, header_format)
                current_row += 1

                resumen_anticipos = [
                    ['🚨 Pendientes', len(anticipos_pendientes), total_anticipos_pendientes, total_restantes_pendientes,
                     total_anticipos_pendientes + total_restantes_pendientes],
                    ['✅ Aprobados', len(anticipos_aprobados), total_anticipos_aprobados, total_restantes_aprobados,
                     total_anticipos_aprobados + total_restantes_aprobados],
                    ['💰 Liquidados', len(anticipos_liquidados), total_anticipos_liquidados, total_restantes_liquidados,
                     total_anticipos_liquidados + total_restantes_liquidados],
                    ['TOTAL GLOBAL', len(solicitudes_con_anticipo),
                     total_anticipos_pendientes + total_anticipos_aprobados + total_anticipos_liquidados,
                     total_restantes_pendientes + total_restantes_aprobados + total_restantes_liquidados,
                     (total_anticipos_pendientes + total_anticipos_aprobados + total_anticipos_liquidados +
                      total_restantes_pendientes + total_restantes_aprobados + total_restantes_liquidados)]
                ]
                for i, (estado, cantidad, anticipos, restantes, total) in enumerate(resumen_anticipos):
                    formato = total_importante_format if i == 3 else cell_format
                    worksheet_anticipos.write(current_row, 0, estado, formato)
                    worksheet_anticipos.write(current_row, 1, cantidad, formato)
                    worksheet_anticipos.write(current_row, 2, anticipos, money_format if i != 3 else total_importante_format)
                    worksheet_anticipos.write(current_row, 3, restantes, money_format if i != 3 else total_importante_format)
                    worksheet_anticipos.write(current_row, 4, total, money_format if i != 3 else total_importante_format)
                    current_row += 1

                worksheet_anticipos.set_column('A:A', 15)
                worksheet_anticipos.set_column('B:B', 25)
                worksheet_anticipos.set_column('C:C', 25)
                worksheet_anticipos.set_column('D:D', 15)
                worksheet_anticipos.set_column('E:E', 15)
                worksheet_anticipos.set_column('F:F', 12)
                worksheet_anticipos.set_column('G:G', 15)
                worksheet_anticipos.set_column('H:H', 15)
                worksheet_anticipos.set_column('I:I', 20)
                worksheet_anticipos.set_column('J:J', 15)
                worksheet_anticipos.set_column('K:K', 20)
                worksheet_anticipos.set_column('L:L', 18)
            else:
                worksheet_anticipos.merge_range(3, 0, 3, 11, 'No hay solicitudes con anticipo registradas', section_info_format)

            # Configurar impresión
            for ws in [worksheet, worksheet_creditos, worksheet_anticipos]:
                ws.set_landscape()
                ws.set_paper(9)
                ws.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
                ws.fit_to_pages(1, 0)

        # Enviar archivo
        output.seek(0)
        nombre_archivo = f"reporte_consolidado_optimizado_{fecha_inicio.strftime('%Y%m%d')}_a_{fecha_fin.strftime('%Y%m%d')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        app.logger.error(f"Error al generar el reporte consolidado optimizado: {e}")
        flash(f"Ocurrió un error al generar el reporte: {str(e)}", "error")
        return redirect(url_for("admin_dashboard"))

# Función auxiliar mejorada para convertir números de columna a letra (Excel)
def xl_col_to_name(col_idx):
    """Convierte un índice de columna a letra estilo Excel (0->A, 1->B, etc.)"""
    name = ""
    while col_idx >= 0:
        name = chr(col_idx % 26 + ord('A')) + name
        col_idx = col_idx // 26 - 1
    return name
# 5. Función auxiliar para obtener el historial de estados
def obtener_historial_estados(solicitud_id):
    """
    Obtiene el historial completo de cambios de estado de una solicitud
    """
    conn = get_db_connection()
    solicitud = conn.execute("SELECT historial_estados FROM solicitudes WHERE id = ?", (solicitud_id,)).fetchone()
    conn.close()

    if solicitud and solicitud['historial_estados']:
        try:
            return json.loads(solicitud['historial_estados'])
        except:
            return []
    return []

# === AGREGAR ESTAS RUTAS A TU app.py ===

@app.route("/actualizar_estado_masivo", methods=["POST"])
def actualizar_estado_masivo():
    """
    Actualiza el estado de múltiples solicitudes a la vez
    """
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    try:
        # Obtener datos del formulario
        solicitudes_ids = request.form.getlist("solicitudes_ids[]")
        nuevo_estado = request.form.get("nuevo_estado_masivo")
        fecha_cambio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not solicitudes_ids or not nuevo_estado:
            flash("No se seleccionaron solicitudes o estado.", "error")
            return redirect(url_for("admin_dashboard"))

        conn = get_db_connection()
        solicitudes_actualizadas = []
        solicitudes_con_anticipo = []

        # Procesar cada solicitud
        for solicitud_id in solicitudes_ids:
            # Obtener información actual de la solicitud
            solicitud_actual = conn.execute(
                "SELECT * FROM solicitudes WHERE id = ?",
                (solicitud_id,)
            ).fetchone()

            if not solicitud_actual:
                continue

            estado_anterior = solicitud_actual['estado']

            # Parsear historial existente
            try:
                historial = json.loads(solicitud_actual['historial_estados'] or '[]')
            except:
                historial = []

            # Agregar nuevo registro al historial
            historial.append({
                'estado_anterior': estado_anterior,
                'estado_nuevo': nuevo_estado.capitalize(),
                'fecha': fecha_cambio,
                'usuario': session.get('username', session.get('role', 'admin')),
                'tipo': 'masivo'
            })

            # Construir la consulta de actualización
            update_query = "UPDATE solicitudes SET estado = ?, fecha_ultimo_cambio = ?, historial_estados = ?"
            params = [nuevo_estado.capitalize(), fecha_cambio, json.dumps(historial)]

            # Registrar fecha específica según el tipo de cambio
            if nuevo_estado.lower() in ["aprobado", "aprobado con anticipo"]:
                update_query += ", fecha_aprobado = ?"
                params.append(fecha_cambio)
            elif nuevo_estado.lower() in ["liquidado", "liquidado con anticipo", "liquidacion total"]:
                update_query += ", fecha_liquidado = ?"
                params.append(fecha_cambio)

            update_query += " WHERE id = ?"
            params.append(solicitud_id)

            # Ejecutar actualización
            conn.execute(update_query, params)

            # Guardar información para emails
            solicitud_dict = dict(solicitud_actual)
            solicitud_dict['estado'] = nuevo_estado
            solicitudes_actualizadas.append(solicitud_dict)

            # Verificar si tiene anticipo para procesar archivos después
            if solicitud_actual['anticipo'] == 'Si':
                solicitudes_con_anticipo.append({
                    'id': solicitud_id,
                    'fp': solicitud_actual['fp'],
                    'nombre': solicitud_actual['nombre']
                })

        conn.commit()
        conn.close()

        # Procesar archivos adjuntos si es liquidación
        archivos_procesados = {}
        if nuevo_estado.lower() in ["liquidado", "liquidado con anticipo", "liquidacion total"]:
            for key in request.files:
                if key.startswith('liquidado_file_'):
                    solicitud_id = key.replace('liquidado_file_', '')
                    archivo = request.files[key]
                    if archivo and archivo.filename:
                        filename = secure_filename(archivo.filename)
                        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                        archivo.save(filepath)
                        archivos_procesados[solicitud_id] = archivo

        # Enviar emails según el nuevo estado
        for solicitud in solicitudes_actualizadas:
            solicitud_id_str = str(solicitud['id'])

            if nuevo_estado.lower() == "aprobado":
                send_approval_email(solicitud)
            elif nuevo_estado.lower() == "aprobado con anticipo":
                send_approval_anticipo_email(solicitud)
            elif nuevo_estado.lower() == "declinada":
                send_declined_email(solicitud)
            elif nuevo_estado.lower() == "liquidado":
                if solicitud_id_str in archivos_procesados:
                    archivo = archivos_procesados[solicitud_id_str]
                    archivo.seek(0)
                    send_liquidado_email(solicitud, archivo)
                else:
                    send_liquidado_email(solicitud, None)
            elif nuevo_estado.lower() == "liquidado con anticipo":
                if solicitud_id_str in archivos_procesados:
                    archivo = archivos_procesados[solicitud_id_str]
                    archivo.seek(0)
                    send_liquidado_anticipo_email(solicitud, archivo)
                else:
                    send_liquidado_anticipo_email(solicitud, None)
            elif nuevo_estado.lower() == "liquidacion total":
                if solicitud_id_str in archivos_procesados:
                    archivo = archivos_procesados[solicitud_id_str]
                    archivo.seek(0)
                    send_liquidacion_total_email(solicitud, archivo)
                else:
                    send_liquidacion_total_email(solicitud, None)

        # Sincronizar con base de datos remota
        try:
            threading.Thread(target=sync_solicitudes_to_remote, daemon=True).start()
        except Exception as e:
            print(f"Error al iniciar sincronización: {e}")

        flash(f"Se actualizaron {len(solicitudes_actualizadas)} solicitudes a {nuevo_estado}.", "success")

        # Si hay solicitudes con anticipo, devolver información para mostrar modal de archivos
        if solicitudes_con_anticipo and nuevo_estado.lower() in ["liquidado", "liquidado con anticipo", "liquidacion total"]:
            return jsonify({
                'success': True,
                'solicitudes_con_anticipo': solicitudes_con_anticipo,
                'message': f'Se actualizaron {len(solicitudes_actualizadas)} solicitudes.'
            })

        return jsonify({
            'success': True,
            'message': f'Se actualizaron {len(solicitudes_actualizadas)} solicitudes.'
        })

    except Exception as e:
        print(f"Error en actualización masiva: {e}")
        return jsonify({
            'success': False,
            'message': f'Error al actualizar: {str(e)}'
        }), 500

@app.route("/obtener_solicitudes_seleccionadas", methods=["POST"])
def obtener_solicitudes_seleccionadas():
    """
    Obtiene la información de las solicitudes seleccionadas
    """
    if not session.get("admin_logged_in"):
        return jsonify({'error': 'No autorizado'}), 401

    try:
        solicitudes_ids = request.json.get("solicitudes_ids", [])

        if not solicitudes_ids:
            return jsonify({'error': 'No se seleccionaron solicitudes'}), 400

        conn = get_db_connection()
        placeholders = ','.join('?' * len(solicitudes_ids))
        query = f"""
            SELECT id, fp, nombre, destinatario, monto, estado, anticipo, tipo_solicitud
            FROM solicitudes
            WHERE id IN ({placeholders})
        """

        solicitudes = conn.execute(query, solicitudes_ids).fetchall()
        conn.close()

        solicitudes_data = []
        for sol in solicitudes:
            solicitudes_data.append({
                'id': sol['id'],
                'fp': sol['fp'],
                'nombre': sol['nombre'],
                'destinatario': sol['destinatario'],
                'monto': sol['monto'],
                'estado': sol['estado'],
                'anticipo': sol['anticipo'],
                'tipo_solicitud': sol['tipo_solicitud']
            })

        return jsonify({
            'success': True,
            'solicitudes': solicitudes_data
        })

    except Exception as e:
        print(f"Error obteniendo solicitudes: {e}")
        return jsonify({'error': str(e)}), 500

# ====== ALERTAS A CREADORES DE SOLICITUDES (ADMIN/COORD) ======
# Crea columnas auxiliares, detecta el creador, monitorea cambios y envía email.

from flask import g
from typing import Optional

# ---------- Utilidades de DB para este módulo ----------
def _exec_scalar(conn, query, params=()):
    cur = conn.execute(query, params)
    row = cur.fetchone()
    return row[0] if row and isinstance(row, tuple) else (row[0] if row else None)

def ensure_alerts_columns():
    """
    Añade columnas en SQLite (solicitudes):
      - creado_por TEXT NOT NULL DEFAULT ''
      - ultimo_estado_alertado TEXT
      - alertar_creador INTEGER NOT NULL DEFAULT 1
    """
    conn = get_db_connection()
    try:
        cur = conn.execute("PRAGMA table_info(solicitudes)")
        cols = [r["name"] for r in cur.fetchall()]
        alter_statements = []
        if "creado_por" not in cols:
            alter_statements.append("ALTER TABLE solicitudes ADD COLUMN creado_por TEXT NOT NULL DEFAULT ''")
        if "ultimo_estado_alertado" not in cols:
            alter_statements.append("ALTER TABLE solicitudes ADD COLUMN ultimo_estado_alertado TEXT")
        if "alertar_creador" not in cols:
            alter_statements.append("ALTER TABLE solicitudes ADD COLUMN alertar_creador INTEGER NOT NULL DEFAULT 1")
        for stmt in alter_statements:
            conn.execute(stmt)
        if alter_statements:
            conn.commit()

        # Inicializa ultimo_estado_alertado para registros existentes si está null
        conn.execute("""
            UPDATE solicitudes
               SET ultimo_estado_alertado = estado
             WHERE ultimo_estado_alertado IS NULL OR ultimo_estado_alertado = ''
        """)
        conn.commit()
    finally:
        conn.close()

# Ejecuta la migración de columnas al arrancar el proceso
try:
    ensure_alerts_columns()
except Exception as _e:
    print(f"[alerts] No se pudieron asegurar columnas: {_e}")

# ---------- Sesión/usuarios ----------
def _user_email_from_username(username: str) -> Optional[str]:
    if not username:
        return None
    user = username.strip()
    if not user:
        return None
    return f"{user}@ad17solutions.com"

def _normalize_username(name: Optional[str]) -> str:
    return (name or "").strip()

# Guarda en g si intentan login (para capturar username tras redirect)
@app.before_request
def _capture_login_username():
    try:
        if request.path == "/admin_login" and request.method == "POST":
            # Guardamos el username posteado; lo validará la vista original.
            g._posted_username = (request.form.get("username") or "").strip()
    except Exception:
        pass

# Después de cada request:
# - Si hubo login exitoso (redirect al dashboard), fija session['username']
# - Si un admin/coordinador creó una solicitud (POST a /solicitar_pago),
#   marca 'creado_por' con ese username.
@app.after_request
def _post_hooks(resp):
    try:
        # 1) Fijar session['username'] si el login fue exitoso
        if request.path == "/admin_login" and request.method == "POST":
            # Éxito = respuesta de redirección (302) y ya se setearon flags por tu vista original
            if resp.status_code in (301, 302) and session.get("admin_logged_in"):
                if getattr(g, "_posted_username", None):
                    # Guardamos el username real de quien inició sesión
                    session["username"] = _normalize_username(g._posted_username)

        # 2) Estampar 'creado_por' cuando un admin/coordinador envía /solicitar_pago
        if request.path == "/solicitar_pago" and request.method == "POST":
            # Sólo si hay admin/coordinador logueado
            if session.get("admin_logged_in") and session.get("role") in ("admin", "coordinador"):
                creador = _normalize_username(session.get("username") or session.get("role"))
                # Intentamos identificar el FP posteado para marcar correctamente
                fp_form = (request.form.get("fp") or "").strip()
                conn = get_db_connection()
                try:
                    if fp_form:
                        # Marca por FP
                        conn.execute("""
                            UPDATE solicitudes
                               SET creado_por = CASE WHEN creado_por IS NULL OR creado_por = '' THEN ? ELSE creado_por END
                             WHERE fp = ?
                        """, (creador, fp_form))
                    else:
                        # Caso de respaldo: marca el registro más reciente del creador/correo actual por fecha
                        conn.execute("""
                            UPDATE solicitudes
                               SET creado_por = CASE WHEN creado_por IS NULL OR creado_por = '' THEN ? ELSE creado_por END
                             WHERE id = (SELECT id FROM solicitudes ORDER BY id DESC LIMIT 1)
                        """, (creador,))
                    conn.commit()
                finally:
                    conn.close()
    except Exception as e:
        print(f"[alerts after_request] {e}")
    return resp

# ---------- Email de alerta ----------
def send_creator_alert_email(fp: str, creado_por: str, estado_anterior: str, estado_nuevo: str, cambiado_por: str):
    """Email de alerta al creador de la solicitud"""

    to_email = _user_email_from_username(creado_por)
    if not to_email:
        return False

    # Determinar color según el tipo de cambio
    if estado_nuevo.lower() in ['aprobado', 'aprobado con anticipo', 'liquidado', 'liquidado con anticipo', 'liquidacion total']:
        bg_color = "linear-gradient(135deg, #E8F5E9 0%, #C8E6C9 100%)"
        border_color = "#2E7D32"
        text_color = "#2E7D32"
        icon = "✅"
    elif estado_nuevo.lower() == 'declinada':
        bg_color = "linear-gradient(135deg, #FFEBEE 0%, #FFCDD2 100%)"
        border_color = "#C62828"
        text_color = "#C62828"
        icon = "❌"
    else:
        bg_color = "linear-gradient(135deg, #FFF3E0 0%, #FFE0B2 100%)"
        border_color = "#FF9800"
        text_color = "#E65100"
        icon = "🔄"

    highlight_html = f"""
    <div class="highlight-box" style="background: {bg_color}; border-left-color: {border_color};">
        <h2 style="color: {text_color};">{icon} Cambio de Estado en tu FP</h2>
        <div class="critical-info">
            <div class="info-item" style="border-color: {border_color};">
                <div class="info-label">📋 FP</div>
                <div class="info-value" style="color: {text_color};">{fp}</div>
            </div>
            <div class="info-item" style="border-color: {border_color};">
                <div class="info-label">📊 Estado Anterior</div>
                <div class="info-value" style="color: {text_color};">{estado_anterior}</div>
            </div>
            <div class="info-item" style="border-color: {border_color};">
                <div class="info-label">📊 Estado Nuevo</div>
                <div class="info-value" style="color: {text_color}; font-size: 24px;">{estado_nuevo}</div>
            </div>
            <div class="info-item" style="border-color: {border_color};">
                <div class="info-label">👤 Actualizado por</div>
                <div class="info-value" style="color: {text_color};">{cambiado_por or '(no identificado)'}</div>
            </div>
        </div>
    </div>
    """

    content_html = f"""
        <h3>{icon} Actualización de tu Solicitud</h3>
        <p>Hola <strong>{creado_por}</strong>,</p>
        <p>El registro que creaste ha cambiado de estado:</p>

        <table class="details-table">
            <tr>
                <td>📅 Fecha del Cambio:</td>
                <td><strong>{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</strong></td>
            </tr>
            <tr>
                <td>🔄 Cambio:</td>
                <td>
                    <span class="status-badge status-{estado_anterior.lower().replace(' ', '-')}">{estado_anterior}</span>
                    <strong style="font-size: 18px; margin: 0 10px;">→</strong>
                    <span class="status-badge status-{estado_nuevo.lower().replace(' ', '-')}">{estado_nuevo}</span>
                </td>
            </tr>
        </table>

        <div class="alert-box">
            <p><strong>ℹ️ Información</strong></p>
            <p>Este aviso se genera automáticamente porque iniciaste sesión al crear el registro.</p>
            <p>Si no deseas recibir más alertas de este FP, contacta con el administrador del sistema.</p>
        </div>
    """

    html_content = get_email_html_template(
        title="Cambio de Estado en tu FP",
        content_html=content_html,
        highlight_section=highlight_html
    )

    msg = EmailMessage()
    msg['Subject'] = f"{icon} [AD17] Cambio de estado en tu FP {fp}: {estado_anterior} → {estado_nuevo}"
    msg['From'] = "ad17solutionsbot@gmail.com"
    msg['To'] = to_email
    msg.set_content("Por favor, habilita la visualización HTML en tu cliente de correo.")
    msg.add_alternative(html_content, subtype='html')

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login("ad17solutionsbot@gmail.com", "misvtfhrnwbmiptb")
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"[alerts email] Error enviando alerta a {to_email}: {e}")
        return False
# ---------- Monitor de cambios de estado ----------
def _get_last_change_user(hist_json: str) -> Optional[str]:
    """
    Extrae el 'usuario' del último cambio desde historial_estados (JSON).
    """
    try:
        hist = json.loads(hist_json or "[]")
        if not hist:
            return None
        last = hist[-1]
        return _normalize_username(last.get("usuario"))
    except Exception:
        return None

def monitor_state_changes_and_notify():
    """
    Revisa solicitudes y, si cambió el estado respecto a 'ultimo_estado_alertado',
    notifica al creador cuando el cambio NO lo hizo el mismo creador.
    """
    try:
        conn = get_db_connection()
        rows = conn.execute("""
            SELECT id, fp, estado, creado_por, ultimo_estado_alertado, historial_estados
              FROM solicitudes
             WHERE alertar_creador = 1
               AND (creado_por IS NOT NULL AND creado_por <> '')
        """).fetchall()

        to_update = []
        for r in rows:
            d = dict(r)
            fp = d.get("fp") or ""
            estado_actual = (d.get("estado") or "").strip()
            ultimo_alertado = (d.get("ultimo_estado_alertado") or "").strip()
            creado_por = _normalize_username(d.get("creado_por"))
            if not fp or not creado_por:
                continue

            # Sólo si cambió el estado real
            if not estado_actual or estado_actual == ultimo_alertado:
                continue

            # ¿Quién realizó el último cambio?
            cambiado_por = _get_last_change_user(d.get("historial_estados") or "")
            # Si no hay historial, no alertamos (evita falsos positivos).
            if not cambiado_por:
                continue

            # Evita alertar si el creador fue quien cambió
            if cambiado_por.lower() == creado_por.lower():
                # Aún así, avanzamos el último estado alertado para no repetir
                to_update.append((estado_actual, d["id"]))
                continue

            # Envía la alerta
            ok = send_creator_alert_email(
                fp=fp,
                creado_por=creado_por,
                estado_anterior=ultimo_alertado or "(desconocido)",
                estado_nuevo=estado_actual,
                cambiado_por=cambiado_por
            )
            if ok:
                to_update.append((estado_actual, d["id"]))

        # Actualiza ultimo_estado_alertado en lote
        if to_update:
            for nuevo_estado, _id in to_update:
                conn.execute(
                    "UPDATE solicitudes SET ultimo_estado_alertado = ? WHERE id = ?",
                    (nuevo_estado, _id)
                )
            conn.commit()
        conn.close()
    except Exception as e:
        print(f"[alerts monitor] {e}")

def get_email_html_template(title, content_html, highlight_section=""):
    """
    Template base HTML para todos los correos con diseño profesional
    """
    return f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background-color: #f4f4f4;
            padding: 20px;
            line-height: 1.6;
        }}
        .email-container {{
            max-width: 650px;
            margin: 0 auto;
            background-color: #ffffff;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }}
        .header {{
            background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
            color: white;
            padding: 30px 20px;
            text-align: center;
        }}
        .header h1 {{
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 5px;
        }}
        .header p {{
            font-size: 14px;
            opacity: 0.95;
        }}
        .highlight-box {{
            background: linear-gradient(135deg, #FFF3E0 0%, #FFE0B2 100%);
            border-left: 5px solid #FF9800;
            margin: 25px 20px;
            padding: 25px;
            border-radius: 8px;
        }}
        .highlight-box h2 {{
            color: #E65100;
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .critical-info {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }}
        .info-item {{
            background: white;
            padding: 15px;
            border-radius: 8px;
            border: 2px solid #FF9800;
        }}
        .info-label {{
            font-size: 12px;
            color: #666;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 5px;
        }}
        .info-value {{
            font-size: 20px;
            color: #E65100;
            font-weight: 700;
        }}
        .info-value.monto {{
            font-size: 28px;
        }}
        .content {{
            padding: 25px;
            color: #333;
        }}
        .content h3 {{
            color: #FF9800;
            font-size: 18px;
            margin-bottom: 15px;
            border-bottom: 2px solid #FFE0B2;
            padding-bottom: 10px;
        }}
        .details-table {{
            width: 100%;
            margin: 20px 0;
            border-collapse: collapse;
        }}
        .details-table tr {{
            border-bottom: 1px solid #f0f0f0;
        }}
        .details-table tr:last-child {{
            border-bottom: none;
        }}
        .details-table td {{
            padding: 12px 8px;
            vertical-align: top;
        }}
        .details-table td:first-child {{
            font-weight: 600;
            color: #666;
            width: 40%;
        }}
        .details-table td:last-child {{
            color: #333;
        }}
        .status-badge {{
            display: inline-block;
            padding: 6px 16px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .status-pendiente {{ background-color: #FFF3E0; color: #F57C00; }}
        .status-aprobado {{ background-color: #E8F5E9; color: #2E7D32; }}
        .status-liquidado {{ background-color: #E3F2FD; color: #1565C0; }}
        .status-declinado {{ background-color: #FFEBEE; color: #C62828; }}
        .alert-box {{
            background-color: #FFF3E0;
            border-left: 4px solid #FF9800;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
        }}
        .alert-box p {{
            margin: 5px 0;
            color: #E65100;
        }}
        .footer {{
            background-color: #f9f9f9;
            padding: 25px;
            text-align: center;
            border-top: 3px solid #FF9800;
        }}
        .footer p {{
            color: #666;
            font-size: 13px;
            margin: 8px 0;
        }}
        .footer strong {{
            color: #FF9800;
        }}
        @media only screen and (max-width: 600px) {{
            .critical-info {{
                grid-template-columns: 1fr;
            }}
            .info-value {{
                font-size: 18px;
            }}
            .info-value.monto {{
                font-size: 24px;
            }}
            .header h1 {{
                font-size: 24px;
            }}
        }}
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>💼 AD17 Solutions</h1>
            <p>Sistema de Gestión de Costos</p>
        </div>

        {highlight_section}

        <div class="content">
            {content_html}
        </div>

        <div class="footer">
            <p><strong>AD17 Solutions</strong></p>
            <p>Sistema de Gestión de Costos</p>
            <p style="color: #999; font-size: 11px; margin-top: 15px;">
                Este es un correo automático, por favor no responder.<br>
                Para cualquier duda, contacta con el departamento administrativo.
            </p>
        </div>
    </div>
</body>
</html>
"""


# ---------- Scheduler dedicado de alertas ----------
def start_alerts_scheduler():
    """
    Inicia un scheduler dedicado a las alertas (independiente del que ya tienes para sync),
    evitando doble arranque por worker.
    """
    if getattr(app, "_alerts_scheduler_started", False):
        return
    try:
        scheduler = BackgroundScheduler()
        # Revisa cada 2 minutos
        scheduler.add_job(
            func=monitor_state_changes_and_notify,
            trigger="interval",
            minutes=2,
            id="creator_alerts_monitor",
            replace_existing=True
        )
        scheduler.start()
        app._alerts_scheduler_started = True
        print("[alerts] Scheduler de alertas iniciado (cada 2 min)")
        atexit.register(lambda: scheduler.shutdown())
    except Exception as e:
        print(f"[alerts] No se pudo iniciar scheduler: {e}")




# Arranca el scheduler de alertas en una inicialización perezosa paralela a la tuya
@app.before_request
def _start_alerts_once():
    try:
        start_alerts_scheduler()
    except Exception as e:
        print(f"[alerts] error start once: {e}")



if __name__ == "__main__":
    app.run(debug=False, threaded=True)