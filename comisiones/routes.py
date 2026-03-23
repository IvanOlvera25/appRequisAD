# comisiones/routes.py
import os
import mimetypes
from collections import defaultdict
from datetime import datetime
from flask import (
    Blueprint, render_template, request, redirect, url_for, flash, jsonify,
    session, send_from_directory, current_app
)
from werkzeug.utils import secure_filename
from .mail import send_email_basic, build_comision_body, get_recipients_from_string

comisiones_bp = Blueprint("comisiones", __name__)

ALLOWED_COMM_FILES = {".pdf", ".png", ".jpg", ".jpeg"}

# ================= Helpers base =================

def _allowed_comm(filename: str) -> bool:
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_COMM_FILES

def _uploads_folder() -> str:
    folder = current_app.config.get(
        "COMM_UPLOAD_FOLDER",
        os.path.join(current_app.root_path, "uploads", "comisiones")
    )
    os.makedirs(folder, exist_ok=True)
    return folder

def _db():
    return current_app.config["GET_DB_CONNECTION"]()

def _employees():
    return current_app.config["READ_EMPLOYEES"]()

def _to_float(s) -> float:
    if s is None:
        return 0.0
    try:
        s = str(s).strip().replace(",", "")
        return float(s or 0)
    except Exception:
        return 0.0

def _schema_cols():
    """Devuelve set con los nombres de columna actuales de la tabla comisiones."""
    conn = _db()
    cur = conn.execute("PRAGMA table_info(comisiones)")
    cols = {row["name"] if isinstance(row, dict) or hasattr(row, "keys") else row[1] for row in cur.fetchall()}
    conn.close()
    return cols

def _email_col(cols:set) -> str | None:
    """Detecta la columna de correo en DB."""
    if "correo" in cols:
        return "correo"
    if "correos" in cols:
        return "correos"
    return None

def _has(cols:set, *names) -> bool:
    return all(n in cols for n in names)

def _normalize_row(row: dict, cols:set) -> dict:
    """
    Normaliza un registro de comisión para que el template siempre tenga:
     - comision_monto
     - comision_tipo  ('porcentaje'|'monto')
     - comision_valor (número)
     - correo (aunque la columna sea 'correos')
     - comprobante_archivo/comprobante_mime (si no existen, dejar vacío)
    """
    d = dict(row)

    # Normalizar correo
    if "correo" not in d and "correos" in d:
        d["correo"] = d.get("correos", "")

    # Normalizar tipo/valor
    tipo = d.get("comision_tipo")
    valor = d.get("comision_valor")

    # Posible esquema viejo
    if not tipo and "tipo_calculo" in d:
        tipo = d.get("tipo_calculo", "porcentaje")
    if (valor is None) and "porcentaje" in d:
        # si el tipo fuera 'monto', usar monto_comision como valor
        if (tipo or "").lower() == "monto" and "monto_comision" in d and d["monto_comision"] is not None:
            valor = float(d["monto_comision"])
        else:
            valor = float(d.get("porcentaje") or 0)

    # Asegurar minúsculas
    tipo = (tipo or "porcentaje").lower()
    d["comision_tipo"] = tipo
    try:
        d["comision_valor"] = float(valor or 0)
    except Exception:
        d["comision_valor"] = 0.0

    # Normalizar comision_monto
    cm = d.get("comision_monto", None)
    if cm is None:
        # Esquema viejo: quizá hay monto_calculado
        if "monto_calculado" in d and d["monto_calculado"] is not None:
            cm = float(d["monto_calculado"])
        elif (tipo == "monto") and ("monto_comision" in d) and (d["monto_comision"] is not None):
            cm = float(d["monto_comision"])
        else:
            # Calcular: porcentaje sobre monto_proyecto
            mp = _to_float(d.get("monto_proyecto"))
            cm = round(mp * (d["comision_valor"] / 100.0), 2) if mp else 0.0
    try:
        d["comision_monto"] = float(cm)
    except Exception:
        d["comision_monto"] = 0.0

    # Archivos (si la DB no tiene columnas, expón vacíos para el template)
    if "comprobante_archivo" not in d:
        d["comprobante_archivo"] = ""
    if "comprobante_mime" not in d:
        d["comprobante_mime"] = ""

    # Defaults seguros
    d.setdefault("estado", "Pendiente")
    d.setdefault("fecha", datetime.now().strftime("%Y-%m-%d"))
    d.setdefault("trabajador", "")
    d.setdefault("proyecto", "")
    d.setdefault("monto_proyecto", 0.0)

    return d


# ================= Rutas =================

@comisiones_bp.route("/", methods=["GET"])
def comisiones():
    """Pantalla principal: alta + tabla + estadísticas (robusta a esquemas viejos)."""
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso para acceder a Comisiones.", "error")
        return redirect(url_for("admin_dashboard"))

    employees = _employees()

    conn = _db()
    rows = conn.execute("SELECT * FROM comisiones ORDER BY date(fecha) DESC, id DESC").fetchall()
    conn.close()

    cols = _schema_cols()
    items = [_normalize_row(dict(r), cols) for r in rows]

    # Estadísticas en Python (evitamos referenciar columnas que no existan en SQL)
    stats_map = defaultdict(lambda: {"trabajador": "", "cantidad": 0, "total": 0.0})
    for it in items:
        t = it.get("trabajador") or "(sin nombre)"
        st = stats_map[t]
        st["trabajador"] = t
        st["cantidad"] += 1
        st["total"] += float(it.get("comision_monto") or 0.0)
    stats = sorted(stats_map.values(), key=lambda x: x["total"], reverse=True)

    return render_template(
        "comisiones.html",
        employees=employees,
        comisiones=items,
        stats=stats
    )


@comisiones_bp.route("/nueva", methods=["POST"])
def comision_nueva():
    """Crear una nueva comisión (compatible con esquema viejo y nuevo)."""
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))

    f = request.form

    # Trabajador al estilo 'solicitar_pago' (sin departamento)
    selected_nombre = (f.get("selected_nombre") or "").strip()
    trabajador = (f.get("nombre_otro") or "").strip() if selected_nombre == "otro" else selected_nombre

    proyecto = (f.get("proyecto") or "").strip()
    monto_proyecto = _to_float(f.get("monto_proyecto"))
    comision_tipo = (f.get("comision_tipo") or "porcentaje").strip().lower()  # 'porcentaje'|'monto'
    comision_valor = _to_float(f.get("comision_valor"))
    correo_in = (f.get("correo") or "").strip()
    estado = (f.get("estado") or "Pendiente").strip()
    fecha = (f.get("fecha") or "").strip() or datetime.now().strftime("%Y-%m-%d")
    fecha_creacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Cálculo de monto de comisión
    comision_monto = comision_valor if comision_tipo == "monto" else round(monto_proyecto * (comision_valor / 100.0), 2)

    cols = _schema_cols()
    email_col = _email_col(cols)

    # Construir INSERT según esquema disponible
    conn = _db()
    try:
        if _has(cols, "comision_tipo", "comision_valor", "comision_monto"):
            # Esquema nuevo
            fields = ["trabajador", "proyecto", "monto_proyecto",
                      "comision_tipo", "comision_valor", "comision_monto",
                      "estado", "fecha", "fecha_creacion"]
            values = [trabajador, proyecto, monto_proyecto,
                      comision_tipo, comision_valor, comision_monto,
                      estado, fecha, fecha_creacion]
            if email_col:
                fields.insert(6, email_col)  # antes de estado
                values.insert(6, correo_in)

            q = f"INSERT INTO comisiones ({', '.join(fields)}) VALUES ({', '.join(['?']*len(fields))})"
            conn.execute(q, values)

        else:
            # Esquema viejo (tipo_calculo/porcentaje/monto_comision/monto_calculado)
            tipo_calculo = "monto" if comision_tipo == "monto" else "porcentaje"
            porcentaje = 0.0 if comision_tipo == "monto" else comision_valor
            monto_comision = comision_valor if comision_tipo == "monto" else 0.0
            monto_calculado = comision_monto

            fields = ["trabajador", "proyecto", "monto_proyecto",
                      "tipo_calculo", "porcentaje", "monto_comision", "monto_calculado",
                      "estado", "fecha", "fecha_creacion"]
            values = [trabajador, proyecto, monto_proyecto,
                      tipo_calculo, porcentaje, monto_comision, monto_calculado,
                      estado, fecha, fecha_creacion]
            if email_col:
                fields.insert(6, email_col)  # después de monto_calculado? lo ponemos justo antes de estado
                values.insert(6, correo_in)

            q = f"INSERT INTO comisiones ({', '.join(fields)}) VALUES ({', '.join(['?']*len(fields))})"
            conn.execute(q, values)

        conn.commit()
    finally:
        conn.close()

    # Notificación inicial opcional
    if correo_in:
        try:
            rcpts = get_recipients_from_string(correo_in)
            if rcpts:
                subject = f"Nueva comisión registrada - {trabajador}"
                body = build_comision_body({
                    "trabajador": trabajador,
                    "proyecto": proyecto,
                    "monto_proyecto": monto_proyecto,
                    "comision_tipo": comision_tipo,
                    "comision_valor": comision_valor,
                    "comision_monto": comision_monto,
                    "correo": correo_in,
                    "estado": estado,
                    "fecha": fecha,
                })
                send_email_basic(subject, body, rcpts)
        except Exception as e:
            print("Error enviando correo inicial de comisión:", e)

    flash("Comisión creada.", "success")
    return redirect(url_for("comisiones.comisiones"))


@comisiones_bp.route("/<int:cid>.json", methods=["GET"])
def comision_json(cid: int):
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        return jsonify({"error": "No autorizado"}), 403

    cols = _schema_cols()
    conn = _db()
    row = conn.execute("SELECT * FROM comisiones WHERE id=?", (cid,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"error": "No encontrado"}), 404

    return jsonify(_normalize_row(dict(row), cols))


@comisiones_bp.route("/<int:cid>/editar", methods=["POST"])
def comision_editar(cid: int):
    """Editar comisión respetando el esquema existente."""
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))

    f = request.form
    sel = (f.get("selected_trabajador") or "").strip()
    trabajador = (f.get("trabajador_otro") or "").strip() if sel == "otro" else sel

    proyecto = (f.get("proyecto") or "").strip()
    fecha = (f.get("fecha") or "").strip() or datetime.now().strftime("%Y-%m-%d")
    monto_proyecto = _to_float(f.get("monto_proyecto"))

    # En edición, el modal viejo podría mandar tipo_calculo/porcentaje/monto_comision
    # y el nuevo comision_tipo/comision_valor. Detectamos ambos:
    comision_tipo = (f.get("comision_tipo") or f.get("tipo_calculo") or "porcentaje").strip().lower()
    porcentaje = _to_float(f.get("porcentaje"))
    monto_comision = _to_float(f.get("monto_comision"))
    comision_valor = _to_float(f.get("comision_valor"))
    if comision_tipo == "monto":
        val = monto_comision if monto_comision else comision_valor
        comision_monto = round(val, 2)
        comision_val = val
    else:
        val = porcentaje if porcentaje else comision_valor
        comision_monto = round(monto_proyecto * (val / 100.0), 2)
        comision_val = val

    notas = f.get("notas") or ""
    correo_in = (f.get("correos") or f.get("correo") or "").strip()

    cols = _schema_cols()
    email_col = _email_col(cols)
    conn = _db()
    try:
        if _has(cols, "comision_tipo", "comision_valor", "comision_monto"):
            # Esquema nuevo
            q = """
                UPDATE comisiones SET
                  trabajador=?, proyecto=?, monto_proyecto=?,
                  comision_tipo=?, comision_valor=?, comision_monto=?,
                  fecha=?, notas=? {EMAIL}
                WHERE id=?
            """
            email_sql = ""
            params = [trabajador, proyecto, monto_proyecto,
                      comision_tipo, comision_val, comision_monto,
                      fecha, notas]
            if email_col:
                email_sql = f", {email_col}=?"
                params.append(correo_in)
            params.append(cid)
            conn.execute(q.replace("{EMAIL}", email_sql), params)
        else:
            # Esquema viejo
            tipo_calculo = "monto" if comision_tipo == "monto" else "porcentaje"
            porcentaje_val = 0.0 if tipo_calculo == "monto" else comision_val
            monto_com_val = comision_val if tipo_calculo == "monto" else 0.0
            q = """
                UPDATE comisiones SET
                  trabajador=?, proyecto=?, monto_proyecto=?,
                  tipo_calculo=?, porcentaje=?, monto_comision=?, monto_calculado=?,
                  fecha=?, notas=? {EMAIL}
                WHERE id=?
            """
            email_sql = ""
            params = [trabajador, proyecto, monto_proyecto,
                      tipo_calculo, porcentaje_val, monto_com_val, comision_monto,
                      fecha, notas]
            if email_col:
                email_sql = f", {email_col}=?"
                params.append(correo_in)
            params.append(cid)
            conn.execute(q.replace("{EMAIL}", email_sql), params)

        conn.commit()
    finally:
        conn.close()

    flash("Comisión actualizada.", "success")
    return redirect(url_for("comisiones.comisiones"))

@comisiones_bp.route("/<int:cid>/estado", methods=["POST"])
def comision_cambiar_estado(cid: int):
    """Cambiar estado; si es 'Pagada', adjuntar comprobante (si hay columnas)."""
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))

    # Acepta múltiples nombres de campo para 'estado'
    def _get_estado(form):
        candidates = [
            "estado", "nuevo_estado", "estado_nuevo",
            "estado_c", "estado_cambio", "status", "state"
        ]
        for k in candidates:
            v = (form.get(k) or "").strip()
            if v:
                return v
        # fallback: cualquier clave que contenga 'estado' o 'status'
        for k in form.keys():
            if "estado" in k.lower() or "status" in k.lower():
                v = (form.get(k) or "").strip()
                if v:
                    return v
        return ""

    nuevo_estado = _get_estado(request.form)
    if not nuevo_estado:
        flash("Selecciona un estado.", "error")
        return redirect(url_for("comisiones.comisiones"))

    cols = _schema_cols()
    has_comprobante = _has(cols, "comprobante_archivo", "comprobante_mime")

    conn = _db()
    row = conn.execute("SELECT * FROM comisiones WHERE id=?", (cid,)).fetchone()
    if not row:
        conn.close()
        flash("Registro no encontrado.", "error")
        return redirect(url_for("comisiones.comisiones"))

    drow = dict(row)
    filename = drow.get("comprobante_archivo", "") if has_comprobante else ""
    mime = drow.get("comprobante_mime", "") if has_comprobante else ""

    # Si la forma permite comprobante, asegúrate que el form tenga enctype="multipart/form-data"
    comprobante_file = request.files.get("comprobante")
    if has_comprobante and nuevo_estado.lower() == "pagada" and comprobante_file and comprobante_file.filename:
        if not _allowed_comm(comprobante_file.filename):
            conn.close()
            flash("Formato no permitido. Usa PDF/JPG/PNG.", "error")
            return redirect(url_for("comisiones.comisiones"))
        safe = secure_filename(comprobante_file.filename)
        dest = os.path.join(_uploads_folder(), f"{cid}_{safe}")
        comprobante_file.save(dest)
        filename = os.path.basename(dest)
        mime = mimetypes.guess_type(dest)[0] or "application/octet-stream"

    # Persistir cambio
    if has_comprobante:
        conn.execute(
            "UPDATE comisiones SET estado=?, comprobante_archivo=?, comprobante_mime=? WHERE id=?",
            (nuevo_estado, filename, mime, cid)
        )
    else:
        conn.execute("UPDATE comisiones SET estado=? WHERE id=?", (nuevo_estado, cid))
    conn.commit()

    # Recargar para correo
    row = conn.execute("SELECT * FROM comisiones WHERE id=?", (cid,)).fetchone()
    conn.close()

    c = _normalize_row(dict(row), cols)
    rcpts = get_recipients_from_string(c.get("correo", ""))

    if rcpts:
        subject = f"Actualización de comisión: {c.get('estado','')}"
        body = build_comision_body(c)
        attachment_path = None
        if has_comprobante and c.get("estado", "").lower() == "pagada" and c.get("comprobante_archivo"):
            attachment_path = os.path.join(_uploads_folder(), c["comprobante_archivo"])
            body += "\nSe adjunta comprobante de pago.\n"
        try:
            send_email_basic(subject, body, rcpts, attachment_path=attachment_path)
        except Exception as e:
            print("Error enviando correo de cambio de estado:", e)

    flash("Estado actualizado.", "success")
    return redirect(url_for("comisiones.comisiones"))


@comisiones_bp.route("/<int:cid>/eliminar", methods=["POST"])
def comision_eliminar(cid: int):
    """Eliminar comisión. Si existen columnas de comprobante, elimina el archivo también."""
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No tienes permiso.", "error")
        return redirect(url_for("admin_dashboard"))

    cols = _schema_cols()
    has_comprobante = _has(cols, "comprobante_archivo")

    conn = _db()
    if has_comprobante:
        row = conn.execute("SELECT comprobante_archivo FROM comisiones WHERE id=?", (cid,)).fetchone()
        if row and row["comprobante_archivo"]:
            try:
                os.remove(os.path.join(_uploads_folder(), row["comprobante_archivo"]))
            except Exception:
                pass
    conn.execute("DELETE FROM comisiones WHERE id=?", (cid,))
    conn.commit()
    conn.close()

    flash("Comisión eliminada.", "success")
    return redirect(url_for("comisiones.comisiones"))


@comisiones_bp.route("/comprobante/<int:cid>")
def comision_ver_comprobante(cid: int):
    """Previsualizar comprobante si existen columnas para ello."""
    if not session.get("admin_logged_in") or session.get("role") != "admin":
        flash("No autorizado.", "error")
        return redirect(url_for("admin_dashboard"))

    cols = _schema_cols()
    if not _has(cols, "comprobante_archivo"):
        flash("Este esquema no soporta comprobantes.", "error")
        return redirect(url_for("comisiones.comisiones"))

    conn = _db()
    row = conn.execute("SELECT comprobante_archivo FROM comisiones WHERE id=?", (cid,)).fetchone()
    conn.close()
    if not row or not row["comprobante_archivo"]:
        flash("Sin comprobante.", "error")
        return redirect(url_for("comisiones.comisiones"))

    return send_from_directory(_uploads_folder(), row["comprobante_archivo"], as_attachment=False)