try:
    import cgi
except Exception:
    cgi = None
import csv
import hashlib
import hmac
import io
import json
import os
import shutil
import sqlite3
import threading
from datetime import date, datetime, timedelta
from pathlib import Path
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
except Exception:
    colors = None
    letter = None
    getSampleStyleSheet = None
    Image = None
    Paragraph = None
    SimpleDocTemplate = None
    Spacer = None
    Table = None
    TableStyle = None


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", str(APP_DIR))).resolve()
DB_PATH = DATA_DIR / "bdp_crm.sqlite"
BACKUP_DIR = DATA_DIR / "backups"
QUOTE_DIR = DATA_DIR / "cotizaciones_pdf_actuales"
QUOTE_WON_DIR = DATA_DIR / "cotizaciones_logradas"
SOURCE_XLSX = Path(r"C:\Users\Equipo\Documents\ANALISIS Y REPORTES 2026 BDP\control_visitas_red_distribuidores_nayarit.xlsx")


def bootstrap_persistent_data():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    seed_files = [
        (APP_DIR / "bdp_crm.sqlite", DB_PATH),
    ]
    for src, dst in seed_files:
        if src.exists() and not dst.exists():
            shutil.copy2(src, dst)
    for folder_name in ("backups", "cotizaciones_pdf_actuales", "cotizaciones_logradas"):
        src_dir = APP_DIR / folder_name
        dst_dir = DATA_DIR / folder_name
        dst_dir.mkdir(parents=True, exist_ok=True)
        if src_dir.exists():
            for src in src_dir.iterdir():
                dst = dst_dir / src.name
                if src.is_file() and not dst.exists():
                    shutil.copy2(src, dst)

REGIONES = {
    "Costa Sur": {
        "color": "#1f8f73",
        "descripcion": "Region de alto desarrollo turistico y hotelero internacional.",
        "municipios": ["Bahía de Banderas", "Compostela"],
    },
    "Costa Centro y Valles del Norte": {
        "color": "#2d76a3",
        "descripcion": "Zonas costeras, agricolas y pesqueras del centro-norte.",
        "municipios": ["San Blas", "Santiago Ixcuintla", "Tuxpan", "Ruiz"],
    },
    "Costa Norte y Frontera": {
        "color": "#d17a21",
        "descripcion": "Llanura costera del norte, pesca de camaron y agricultura.",
        "municipios": ["Tecuala", "Acaponeta", "Huajicori", "Rosamorada"],
    },
    "Centro y Sierra Centro": {
        "color": "#7a5ac8",
        "descripcion": "Valle central, capital y area conurbada.",
        "municipios": ["Tepic", "Xalisco", "Santa María del Oro", "San Pedro Lagunillas"],
    },
    "Sierra y Sur": {
        "color": "#8b6f3d",
        "descripcion": "Sierra Madre Occidental, minas, volcanes y corredor sur.",
        "municipios": ["Del Nayar", "La Yesca", "Jala", "Ahuacatlán", "Ixtlán del Río", "Amatlán de Cañas"],
    },
}

MUNICIPIO_REGION = {municipio: region for region, data in REGIONES.items() for municipio in data["municipios"]}

NAYARIT_MUNICIPIOS = {
    "Acaponeta": {"cabecera": "Acaponeta", "zona": "Costa Norte y Frontera"},
    "Ahuacatlán": {"cabecera": "Ahuacatlán", "zona": "Sierra y Sur"},
    "Amatlán de Cañas": {"cabecera": "Amatlán de Cañas", "zona": "Sierra y Sur"},
    "Bahía de Banderas": {"cabecera": "Valle de Banderas", "zona": "Costa Sur"},
    "Compostela": {"cabecera": "Compostela", "zona": "Costa Sur"},
    "Del Nayar": {"cabecera": "Jesús María", "zona": "Sierra y Sur"},
    "Huajicori": {"cabecera": "Huajicori", "zona": "Costa Norte y Frontera"},
    "Ixtlán del Río": {"cabecera": "Ixtlán del Río", "zona": "Sierra y Sur"},
    "Jala": {"cabecera": "Jala", "zona": "Sierra y Sur"},
    "La Yesca": {"cabecera": "La Yesca", "zona": "Sierra y Sur"},
    "Rosamorada": {"cabecera": "Rosamorada", "zona": "Costa Norte y Frontera"},
    "Ruiz": {"cabecera": "Ruiz", "zona": "Costa Centro y Valles del Norte"},
    "San Blas": {"cabecera": "San Blas", "zona": "Costa Centro y Valles del Norte"},
    "San Pedro Lagunillas": {"cabecera": "San Pedro Lagunillas", "zona": "Centro y Sierra Centro"},
    "Santa María del Oro": {"cabecera": "Santa María del Oro", "zona": "Centro y Sierra Centro"},
    "Santiago Ixcuintla": {"cabecera": "Santiago Ixcuintla", "zona": "Costa Centro y Valles del Norte"},
    "Tecuala": {"cabecera": "Tecuala", "zona": "Costa Norte y Frontera"},
    "Tepic": {"cabecera": "Tepic", "zona": "Centro y Sierra Centro"},
    "Tuxpan": {"cabecera": "Tuxpan", "zona": "Costa Centro y Valles del Norte"},
    "Xalisco": {"cabecera": "Xalisco", "zona": "Centro y Sierra Centro"},
}

PRODUCTOS = {
    "Block 10 × 14 × 28 cm": 5.50,
    "Block 12 × 20 × 40 cm": None,
    "Block 14 × 20 × 40 cm": 14.56,
    "Bovedilla prefabricada": None,
    "Adoquín tipo romano": None,
}

TARIFAS = {
    "Costa Sur": 28.40,
    "Costa Centro y Valles del Norte": 28.40,
    "Costa Norte y Frontera": 28.40,
    "Centro y Sierra Centro": 26.00,
    "Sierra y Sur": 28.40,
}


DEFAULT_SETTINGS = {
    "user_name": "José Luis",
    "user_role": "Administrador",
    "company_name": "Bazar de Prefabricados",
    "annual_goal": "25000000",
    "monthly_goal_default": "2000000",
    "profile_photo": "",
    "material_cemento_ton": "4300",
    "material_jal_m3": "1600",
    "mo_block_10_14_28": "0",
    "mo_block_14_20_40": "0",
}

MESES_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

def ensure_business_tables(db):
    db.executescript("""
        CREATE TABLE IF NOT EXISTS app_settings (
          key TEXT PRIMARY KEY,
          value TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS monthly_sales (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          year INTEGER NOT NULL,
          month INTEGER NOT NULL,
          amount REAL DEFAULT 0,
          goal REAL DEFAULT 0,
          updated_at TEXT DEFAULT '',
          UNIQUE(year, month)
        );
        CREATE TABLE IF NOT EXISTS product_prices (
          producto TEXT PRIMARY KEY,
          precio REAL,
          updated_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS freight_rates (
          zona TEXT PRIMARY KEY,
          tarifa REAL,
          updated_at TEXT DEFAULT ''
        );
    """)
    for k, v in DEFAULT_SETTINGS.items():
        db.execute("INSERT OR IGNORE INTO app_settings(key,value) VALUES(?,?)", (k, str(v)))
    year = date.today().year
    default_goal = float(DEFAULT_SETTINGS["monthly_goal_default"])
    for m in range(1,13):
        db.execute("INSERT OR IGNORE INTO monthly_sales(year, month, amount, goal, updated_at) VALUES(?,?,?,?,?)", (year, m, 0, default_goal, now_iso()))
    for prod, precio in PRODUCTOS.items():
        if precio is not None:
            db.execute("INSERT OR IGNORE INTO product_prices(producto, precio, updated_at) VALUES(?,?,?)", (prod, precio, now_iso()))
    for zona, tarifa in TARIFAS.items():
        db.execute("INSERT OR IGNORE INTO freight_rates(zona, tarifa, updated_at) VALUES(?,?,?)", (zona, tarifa, now_iso()))


def get_settings_payload(db, year=None):
    year = int(year or date.today().year)
    settings = {r["key"]: r["value"] for r in db.execute("SELECT key,value FROM app_settings").fetchall()}
    sales = rows_to_dicts(db.execute("SELECT year, month, amount, goal FROM monthly_sales WHERE year=? ORDER BY month", (year,)).fetchall())
    product_rows = db.execute("SELECT producto, precio FROM product_prices ORDER BY producto").fetchall()
    freight_rows = db.execute("SELECT zona, tarifa FROM freight_rates ORDER BY zona").fetchall()
    products = dict(PRODUCTOS)
    for r in product_rows:
        products[r["producto"]] = r["precio"]
    tarifas = dict(TARIFAS)
    for r in freight_rows:
        tarifas[r["zona"]] = r["tarifa"]
    return {"settings": settings, "monthlySales": sales, "productos": products, "tarifas": tarifas, "year": year, "months": MESES_ES}



def parse_upload_file(body, content_type):
    """Small multipart/form-data parser used when cgi.FieldStorage is unavailable."""
    if "boundary=" not in content_type:
        return None, "Carga manual"
    boundary = content_type.split("boundary=", 1)[1].strip().strip('"')
    marker = ("--" + boundary).encode("utf-8")
    for part in body.split(marker):
        if b"Content-Disposition" not in part:
            continue
        header, sep, payload = part.partition(b"\r\n\r\n")
        if not sep:
            continue
        header_text = header.decode("utf-8", errors="ignore")
        if 'name="file"' not in header_text:
            continue
        filename = "Carga manual"
        if 'filename="' in header_text:
            filename = header_text.split('filename="', 1)[1].split('"', 1)[0] or filename
        payload = payload.rstrip(b"\r\n")
        if payload.endswith(b"--"):
            payload = payload[:-2].rstrip(b"\r\n")
        return payload, filename
    return None, "Carga manual"

def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def now_iso():
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def normalize_date(value):
    if not value:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value).split(" ")[0]


def zone_for(municipio):
    return NAYARIT_MUNICIPIOS.get(municipio or "", {}).get("zona", "")


def sync_regions(db):
    for municipio, data in NAYARIT_MUNICIPIOS.items():
        db.execute(
            "UPDATE clients SET zona=?, cabecera=? WHERE municipio=?",
            (data["zona"], data["cabecera"], municipio),
        )


def migrate_existing_quote_pdfs(db):
    for folder, status, prefix in ((QUOTE_DIR, "Elaborada", "/cotizaciones_pdf/"), (QUOTE_WON_DIR, "Lograda", "/cotizaciones_logradas/")):
        for pdf in folder.glob("*.pdf"):
            pdf_url = prefix + pdf.name
            exists = db.execute("SELECT id FROM quotes WHERE pdf LIKE ?", (f"%/{pdf.name}",)).fetchone()
            if exists:
                continue
            parts = pdf.stem.split("_")
            client_id = parts[1] if len(parts) > 2 and parts[1].isdigit() else ""
            client = db.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone() if client_id else None
            raw_date = parts[2] if len(parts) > 3 else ""
            fecha = f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:8]}" if len(raw_date) == 8 else date.today().isoformat()
            db.execute(
                """
                INSERT INTO quotes
                (client_id, fecha, cliente, producto, status, pdf, created_at, updated_at)
                VALUES (?,?,?,?,?,?,?,?)
                """,
                (
                    int(client_id) if client_id else 0,
                    fecha,
                    client["nombre_comercial"] if client else pdf.stem,
                    "",
                    status,
                    pdf_url,
                    now_iso(),
                    now_iso(),
                ),
            )


def init_db():
    BACKUP_DIR.mkdir(exist_ok=True)
    QUOTE_DIR.mkdir(exist_ok=True)
    QUOTE_WON_DIR.mkdir(exist_ok=True)
    with conn() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS clients (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              nombre_comercial TEXT NOT NULL,
              razon_social TEXT DEFAULT '',
              tipo_cliente TEXT DEFAULT '',
              potencial TEXT DEFAULT '',
              etapa TEXT DEFAULT 'Prospecto nuevo',
              municipio TEXT DEFAULT '',
              localidad TEXT DEFAULT '',
              cabecera TEXT DEFAULT '',
              zona TEXT DEFAULT '',
              estado TEXT DEFAULT 'Nayarit',
              direccion TEXT DEFAULT '',
              contacto TEXT DEFAULT '',
              telefono TEXT DEFAULT '',
              whatsapp TEXT DEFAULT '',
              correo TEXT DEFAULT '',
              productos_interes TEXT DEFAULT '',
              consumo_mensual TEXT DEFAULT '',
              precio_actual_compra REAL,
              proveedor_actual TEXT DEFAULT '',
              precios_competencia TEXT DEFAULT '',
              observaciones TEXT DEFAULT '',
              info_enviada TEXT DEFAULT '',
              fecha_info_enviada TEXT DEFAULT '',
              created_at TEXT DEFAULT '',
              updated_at TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS interactions (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              client_id INTEGER NOT NULL,
              fecha TEXT NOT NULL,
              hora TEXT DEFAULT '',
              tipo_contacto TEXT DEFAULT '',
              comentarios TEXT DEFAULT '',
              productos_ofrecidos TEXT DEFAULT '',
              precio_ofrecido REAL,
              resultado TEXT DEFAULT '',
              proxima_accion TEXT DEFAULT '',
              fecha_seguimiento TEXT DEFAULT '',
              created_at TEXT DEFAULT '',
              FOREIGN KEY(client_id) REFERENCES clients(id)
            );
            CREATE TABLE IF NOT EXISTS imports (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              source TEXT,
              imported_at TEXT
            );
            CREATE TABLE IF NOT EXISTS quotes (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              client_id INTEGER NOT NULL,
              fecha TEXT DEFAULT '',
              cliente TEXT DEFAULT '',
              producto TEXT DEFAULT '',
              cantidad REAL DEFAULT 0,
              precio_unitario REAL DEFAULT 0,
              flete_total REAL DEFAULT 0,
              casetas REAL DEFAULT 0,
              total REAL DEFAULT 0,
              costo_individual REAL DEFAULT 0,
              status TEXT DEFAULT 'Elaborada',
              pdf TEXT DEFAULT '',
              created_at TEXT DEFAULT '',
              updated_at TEXT DEFAULT ''
            );
            """
        )
        ensure_business_tables(db)
        cols = {r["name"] for r in db.execute("PRAGMA table_info(clients)").fetchall()}
        if "localidad" not in cols:
            db.execute("ALTER TABLE clients ADD COLUMN localidad TEXT DEFAULT ''")
        if "info_enviada" not in cols:
            db.execute("ALTER TABLE clients ADD COLUMN info_enviada TEXT DEFAULT ''")
        if "fecha_info_enviada" not in cols:
            db.execute("ALTER TABLE clients ADD COLUMN fecha_info_enviada TEXT DEFAULT ''")
        sync_regions(db)
        migrate_existing_quote_pdfs(db)
        db.commit()


def client_by_name_phone(db, name, phone):
    name = (name or "").strip()
    phone = (phone or "").strip()
    if phone:
        row = db.execute(
            "SELECT * FROM clients WHERE lower(nombre_comercial)=lower(?) OR telefono=? OR whatsapp=? LIMIT 1",
            (name, phone, phone),
        ).fetchone()
    else:
        row = db.execute("SELECT * FROM clients WHERE lower(nombre_comercial)=lower(?) LIMIT 1", (name,)).fetchone()
    return row


def upsert_client(db, data):
    municipio = data.get("municipio", "")
    catalog = NAYARIT_MUNICIPIOS.get(municipio, {})
    zona = data.get("zona") or catalog.get("zona", "")
    cabecera = data.get("cabecera") or catalog.get("cabecera", municipio)
    nombre = (data.get("nombre_comercial") or "Cliente sin nombre").strip()
    phone = (data.get("telefono") or data.get("whatsapp") or "").strip()
    existing = None
    if data.get("id"):
        existing = db.execute("SELECT * FROM clients WHERE id=?", (data.get("id"),)).fetchone()
    if not existing:
        existing = client_by_name_phone(db, nombre, phone)
    payload = {
        "nombre_comercial": nombre,
        "razon_social": data.get("razon_social", ""),
        "tipo_cliente": data.get("tipo_cliente", ""),
        "potencial": data.get("potencial", ""),
        "etapa": data.get("etapa") or data.get("estatus") or "Prospecto nuevo",
        "municipio": municipio,
        "localidad": data.get("localidad", ""),
        "cabecera": cabecera,
        "zona": zona,
        "estado": data.get("estado") or "Nayarit",
        "direccion": data.get("direccion", ""),
        "contacto": data.get("contacto", ""),
        "telefono": data.get("telefono", ""),
        "whatsapp": data.get("whatsapp") or data.get("telefono", ""),
        "correo": data.get("correo", ""),
        "productos_interes": data.get("productos_interes", ""),
        "consumo_mensual": data.get("consumo_mensual", ""),
        "precio_actual_compra": data.get("precio_actual_compra"),
        "proveedor_actual": data.get("proveedor_actual", ""),
        "precios_competencia": data.get("precios_competencia", ""),
        "observaciones": data.get("observaciones", ""),
        "info_enviada": data.get("info_enviada", ""),
        "fecha_info_enviada": data.get("fecha_info_enviada", ""),
        "updated_at": now_iso(),
    }
    if existing:
        # Actualización robusta: cuando el usuario edita una ficha, se deben guardar
        # TODOS los campos enviados por el formulario, incluyendo cambios en
        # tipo_cliente y campos que intencionalmente se dejan vacíos.
        # Antes se omitían valores vacíos, lo que podía dejar información vieja.
        fields = list(payload.keys())
        sets = ", ".join([f"{k}=?" for k in fields])
        db.execute(f"UPDATE clients SET {sets} WHERE id=?", [payload[k] for k in fields] + [existing["id"]])
        return existing["id"]
    payload["created_at"] = now_iso()
    keys = list(payload.keys())
    db.execute(
        f"INSERT INTO clients ({','.join(keys)}) VALUES ({','.join(['?'] * len(keys))})",
        [payload[k] for k in keys],
    )
    return db.execute("SELECT last_insert_rowid()").fetchone()[0]


def add_interaction(db, client_id, data):
    fecha = data.get("fecha") or date.today().isoformat()
    hora = data.get("hora") or datetime.now().strftime("%H:%M")
    db.execute(
        """
        INSERT INTO interactions
        (client_id, fecha, hora, tipo_contacto, comentarios, productos_ofrecidos, precio_ofrecido,
         resultado, proxima_accion, fecha_seguimiento, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            client_id,
            fecha,
            hora,
            data.get("tipo_contacto", ""),
            data.get("comentarios", ""),
            data.get("productos_ofrecidos", ""),
            data.get("precio_ofrecido") or None,
            data.get("resultado", ""),
            data.get("proxima_accion", ""),
            data.get("fecha_seguimiento", ""),
            now_iso(),
        ),
    )
    if data.get("resultado"):
        db.execute("UPDATE clients SET updated_at=? WHERE id=?", (now_iso(), client_id))


def update_interaction(db, interaction_id, data):
    row = db.execute("SELECT * FROM interactions WHERE id=?", (interaction_id,)).fetchone()
    if not row:
        return None
    fields = [
        "fecha",
        "hora",
        "tipo_contacto",
        "comentarios",
        "productos_ofrecidos",
        "precio_ofrecido",
        "resultado",
        "proxima_accion",
        "fecha_seguimiento",
    ]
    values = {}
    for field in fields:
        if field in data:
            values[field] = data.get(field) or (None if field == "precio_ofrecido" else "")
    if values:
        sets = ", ".join([f"{key}=?" for key in values])
        db.execute(f"UPDATE interactions SET {sets} WHERE id=?", list(values.values()) + [interaction_id])
        db.execute("UPDATE clients SET updated_at=? WHERE id=?", (now_iso(), row["client_id"]))
    if data.get("etapa"):
        db.execute("UPDATE clients SET etapa=?, updated_at=? WHERE id=?", (data["etapa"], now_iso(), row["client_id"]))
    return row["client_id"]


def import_excel(path_or_bytes, source="Excel"):
    if openpyxl is None:
        return {"imported": 0, "error": "openpyxl no está disponible"}
    wb = openpyxl.load_workbook(path_or_bytes, data_only=True)
    if "Captura_Visitas" not in wb.sheetnames:
        return {"imported": 0, "error": "No se encontró la hoja Captura_Visitas"}
    ws = wb["Captura_Visitas"]
    headers = [ws.cell(row=4, column=i).value for i in range(1, ws.max_column + 1)]
    count = 0
    with conn() as db:
        for row in ws.iter_rows(min_row=5, values_only=True):
            record = dict(zip(headers, row))
            name = record.get("Cliente / negocio")
            if not name:
                continue
            estatus = record.get("Estatus") or "Prospecto nuevo"
            etapa_map = {
                "Nuevo prospecto": "Prospecto nuevo",
                "Contactado": "Contactado",
                "Cotización enviada": "Cotización enviada",
                "Seguimiento": "Negociación",
                "Distribuidor captado": "Cliente activo",
                "Cliente activo": "Cliente activo",
                "No interesado": "Cliente perdido",
            }
            potencial_map = {"Alto": "A", "Medio": "B", "Bajo": "C", "A": "A", "B": "B", "C": "C"}
            client_id = upsert_client(
                db,
                {
                    "nombre_comercial": name,
                    "contacto": record.get("Contacto") or "",
                    "telefono": str(record.get("Teléfono") or ""),
                    "whatsapp": str(record.get("Teléfono") or ""),
                    "tipo_cliente": record.get("Tipo cliente") or "",
                    "municipio": record.get("Municipio") or "",
                    "cabecera": record.get("Cabecera municipal") or "",
                    "zona": record.get("Zona") or "",
                    "productos_interes": record.get("Producto") or "",
                    "precio_actual_compra": record.get("Precio unitario planta"),
                    "observaciones": record.get("Observaciones") or "",
                    "potencial": potencial_map.get(str(record.get("Potencial") or ""), str(record.get("Potencial") or "")),
                    "etapa": etapa_map.get(str(estatus), str(estatus)),
                },
            )
            add_interaction(
                db,
                client_id,
                {
                    "fecha": normalize_date(record.get("Fecha visita")),
                    "tipo_contacto": "Visita",
                    "comentarios": record.get("Observaciones") or record.get("Próxima acción") or "",
                    "productos_ofrecidos": record.get("Producto") or "",
                    "precio_ofrecido": record.get("Costo individual ofrecido") or record.get("Precio ofrecido unitario"),
                    "resultado": str(estatus),
                    "proxima_accion": record.get("Próxima acción") or "",
                    "fecha_seguimiento": normalize_date(record.get("Fecha seguimiento")),
                },
            )
            count += 1
        db.execute("INSERT INTO imports (source, imported_at) VALUES (?, ?)", (source, now_iso()))
        db.commit()
    return {"imported": count}


def seed_initial_data():
    with conn() as db:
        existing = db.execute("SELECT COUNT(*) FROM clients").fetchone()[0]
    if existing == 0 and SOURCE_XLSX.exists():
        import_excel(SOURCE_XLSX, "Base inicial")


def auto_backup():
    if not DB_PATH.exists():
        return
    stamp = date.today().isoformat()
    target = BACKUP_DIR / f"bdp_crm_backup_{stamp}.sqlite"
    if not target.exists():
        shutil.copy2(DB_PATH, target)


def rows_to_dicts(rows):
    return [dict(r) for r in rows]


def export_xlsx():
    if openpyxl is None:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clientes"
    headers = [
        "ID", "Nombre comercial", "Razón social", "Tipo cliente", "Potencial", "Etapa", "Zona", "Municipio",
        "Localidad", "Estado", "Dirección", "Contacto", "Teléfono", "WhatsApp", "Correo", "Productos interés",
        "Consumo mensual", "Precio actual compra", "Proveedor actual", "Precios competencia", "Observaciones",
        "Info enviada", "Fecha info enviada",
    ]
    ws.append(headers)
    with conn() as db:
        for r in db.execute("SELECT * FROM clients ORDER BY nombre_comercial"):
            ws.append([
                r["id"], r["nombre_comercial"], r["razon_social"], r["tipo_cliente"], r["potencial"], r["etapa"],
                r["zona"], r["municipio"], r["localidad"], r["estado"], r["direccion"], r["contacto"], r["telefono"], r["whatsapp"],
                r["correo"], r["productos_interes"], r["consumo_mensual"], r["precio_actual_compra"],
                r["proveedor_actual"], r["precios_competencia"], r["observaciones"], r["info_enviada"], r["fecha_info_enviada"],
            ])
    ws2 = wb.create_sheet("Historial")
    ws2.append(["Cliente ID", "Cliente", "Fecha", "Hora", "Tipo contacto", "Comentarios", "Productos ofrecidos", "Precio ofrecido", "Resultado", "Próxima acción", "Fecha seguimiento"])
    with conn() as db:
        q = """
        SELECT i.*, c.nombre_comercial
        FROM interactions i JOIN clients c ON c.id=i.client_id
        ORDER BY i.fecha DESC, i.hora DESC
        """
        for r in db.execute(q):
            ws2.append([r["client_id"], r["nombre_comercial"], r["fecha"], r["hora"], r["tipo_contacto"], r["comentarios"], r["productos_ofrecidos"], r["precio_ofrecido"], r["resultado"], r["proxima_accion"], r["fecha_seguimiento"]])
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col[:80])
            sheet.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 38)
        sheet.freeze_panes = "A2"
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()


def money(value):
    try:
        return f"${float(value):,.2f}"
    except Exception:
        return "$0.00"


def _words_0_999(number):
    units = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"]
    teens = {
        10: "diez", 11: "once", 12: "doce", 13: "trece", 14: "catorce", 15: "quince",
        16: "dieciseis", 17: "diecisiete", 18: "dieciocho", 19: "diecinueve",
    }
    tens = ["", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"]
    hundreds = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"]
    number = int(number)
    if number == 0:
        return ""
    if number == 100:
        return "cien"
    if number < 10:
        return units[number]
    if number < 20:
        return teens[number]
    if number < 30:
        special = {20: "veinte", 21: "veintiuno", 22: "veintidos", 23: "veintitres", 24: "veinticuatro", 25: "veinticinco", 26: "veintiseis", 27: "veintisiete", 28: "veintiocho", 29: "veintinueve"}
        return special[number]
    if number < 100:
        ten, unit = divmod(number, 10)
        return tens[ten] if unit == 0 else f"{tens[ten]} y {units[unit]}"
    hundred, rest = divmod(number, 100)
    return hundreds[hundred] if rest == 0 else f"{hundreds[hundred]} {_words_0_999(rest)}"


def number_to_words_es(number):
    number = int(number)
    if number == 0:
        return "cero"
    if number < 1000:
        return _words_0_999(number)
    if number < 1_000_000:
        thousands, rest = divmod(number, 1000)
        prefix = "mil" if thousands == 1 else f"{_words_0_999(thousands)} mil"
        return prefix if rest == 0 else f"{prefix} {_words_0_999(rest)}"
    millions, rest = divmod(number, 1_000_000)
    prefix = "un millon" if millions == 1 else f"{number_to_words_es(millions)} millones"
    return prefix if rest == 0 else f"{prefix} {number_to_words_es(rest)}"


def amount_words_for_pesos(pesos):
    words = number_to_words_es(pesos)
    replacements = (
        ("veintiuno", "veintiun"),
        (" y uno", " y un"),
        (" uno", " un"),
    )
    if words == "uno":
        return "un"
    for old, new in replacements:
        if words.endswith(old):
            return words[: -len(old)] + new
    return words


def amount_to_words_mxn(value):
    amount = round(float(value or 0), 2)
    pesos = int(amount)
    cents = int(round((amount - pesos) * 100))
    if cents == 100:
        pesos += 1
        cents = 0
    currency = "PESO" if pesos == 1 else "PESOS"
    return f"SON: {amount_words_for_pesos(pesos).upper()} {currency} {cents:02d}/100 M.N."


def generate_quote_pdf(client, data):
    if SimpleDocTemplate is None:
        return None
    safe_name = "".join(ch for ch in client["nombre_comercial"] if ch.isalnum() or ch in (" ", "_", "-")).strip().replace(" ", "_")
    filename = f"cotizacion_{client['id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name or 'cliente'}.pdf"
    path = QUOTE_DIR / filename
    logo = APP_DIR / "assets" / "logo.png"
    doc = SimpleDocTemplate(str(path), pagesize=letter, rightMargin=42, leftMargin=42, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    if logo.exists():
        story.append(Image(str(logo), width=130, height=80, kind="proportional"))
    story.append(Paragraph("<b>Bazar de Prefabricados</b>", styles["Title"]))
    story.append(Paragraph("Cotizacion comercial - Soluciones en concreto", styles["Normal"]))
    story.append(Spacer(1, 14))
    story.append(Paragraph(f"<b>Cliente:</b> {client['nombre_comercial']}", styles["Normal"]))
    story.append(Paragraph(f"<b>Contacto:</b> {client['contacto'] or '-'} | <b>WhatsApp:</b> {client['whatsapp'] or client['telefono'] or '-'}", styles["Normal"]))
    story.append(Paragraph(f"<b>Municipio:</b> {client['municipio'] or '-'} | <b>Region:</b> {client['zona'] or '-'}", styles["Normal"]))
    story.append(Spacer(1, 14))
    rows = [
        ["Concepto", "Importe / dato"],
        ["Producto", data.get("producto") or data.get("productos_ofrecidos") or "-"],
        ["Cantidad de piezas", data.get("cantidad") or "-"],
        ["Precio unitario planta", money(data.get("precio_unitario"))],
        ["Kilometros flete", data.get("km_flete") or "0"],
        ["Tarifa flete por km", money(data.get("tarifa_km"))],
        ["Flete base", money(data.get("flete_base"))],
        ["Casetas de cobro", money(data.get("casetas"))],
        ["Flete total", money(data.get("flete_total"))],
        ["Flete por pieza", money(data.get("flete_unitario"))],
        ["Costo individual con flete", money(data.get("precio_ofrecido"))],
        ["Total cotizado", money(data.get("total"))],
    ]
    table = Table(rows, colWidths=[210, 270])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#c99740")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f7f3e8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("PADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(table)
    final_total = float(data.get("total") or 0)
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>COSTO FINAL: {money(final_total)}</b>", styles["Heading2"]))
    story.append(Paragraph(f"<b>{amount_to_words_mxn(final_total)}</b>", styles["Heading3"]))
    story.append(Spacer(1, 14))
    story.append(Paragraph("<b>Observaciones:</b>", styles["Heading3"]))
    story.append(Paragraph(data.get("comentarios", "Precio sujeto a confirmacion de volumen, ruta y disponibilidad."), styles["Normal"]))
    story.append(Spacer(1, 14))
    story.append(Paragraph("Contacto: 311 105 1368 | bdprefabricados@hotmail.com | Tepic, Nayarit", styles["Normal"]))
    doc.build(story)
    return filename


def auth_user():
    return os.environ.get("CRM_USER", "admin")


def auth_password():
    return os.environ.get("CRM_PASSWORD", "BDP2026")


def auth_secret():
    return os.environ.get("SESSION_SECRET", auth_password() + "_crm_bdp_2026")


def session_token():
    raw = f"{auth_user()}:{auth_password()}".encode("utf-8")
    return hmac.new(auth_secret().encode("utf-8"), raw, hashlib.sha256).hexdigest()


def login_page(error=""):
    message = f"<div class='error'>{error}</div>" if error else ""
    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Acceso CRM BDP</title>
  <style>
    body{{margin:0;min-height:100vh;display:grid;place-items:center;background:#111;color:#f6f0df;font-family:Arial,sans-serif}}
    .box{{width:min(420px,92vw);background:#1b1b1b;border:1px solid #8f7b52;border-radius:12px;padding:28px;box-shadow:0 22px 60px rgba(0,0,0,.45)}}
    h1{{margin:0 0 8px;color:#c9a15a}} p{{color:#d5d0c4}}
    label{{display:block;margin:14px 0 6px;font-size:13px;color:#c9a15a;font-weight:bold}}
    input{{width:100%;box-sizing:border-box;padding:13px;border-radius:8px;border:1px solid #444;background:#0d0d0d;color:white;font-size:16px}}
    button{{width:100%;margin-top:18px;padding:13px;border:0;border-radius:8px;background:#c9a15a;color:#111;font-size:16px;font-weight:800;cursor:pointer}}
    .error{{background:#3a1717;color:#ffdede;border:1px solid #9a4444;padding:10px;border-radius:8px;margin:12px 0}}
  </style>
</head>
<body>
  <form class="box" method="post" action="/api/login">
    <h1>CRM BDP 2026</h1>
    <p>Ingresa usuario y contraseña para acceder.</p>
    {message}
    <label>Usuario</label>
    <input name="user" autocomplete="username" required>
    <label>Contraseña</label>
    <input name="password" type="password" autocomplete="current-password" required>
    <button>Entrar</button>
  </form>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def _send(self, status=200, content_type="application/json", body=b"", headers=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.wfile.write(body)

    def json(self, obj, status=200):
        self._send(status, "application/json; charset=utf-8", json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def is_authenticated(self):
        cookie = self.headers.get("Cookie", "")
        expected = "crm_session=" + session_token()
        return expected in cookie

    def require_auth(self, path):
        if self.is_authenticated():
            return True
        if path.startswith("/api/"):
            self.json({"error": "No autorizado"}, 401)
        else:
            self._send(200, "text/html; charset=utf-8", login_page())
        return False

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)
        if path == "/login":
            return self._send(200, "text/html; charset=utf-8", login_page())
        if path.startswith("/assets/"):
            filename = Path(path).name
            target = APP_DIR / "assets" / filename
            if target.exists() and target.is_file():
                ext = target.suffix.lower()
                content_type = "application/octet-stream"
                if ext == ".png":
                    content_type = "image/png"
                elif ext in (".jpg", ".jpeg"):
                    content_type = "image/jpeg"
                elif ext == ".gif":
                    content_type = "image/gif"
                elif ext == ".svg":
                    content_type = "image/svg+xml"
                elif ext == ".css":
                    content_type = "text/css; charset=utf-8"
                elif ext == ".js":
                    content_type = "application/javascript; charset=utf-8"
                return self._send(200, content_type, target.read_bytes())
            return self.json({"error": "Asset no encontrado"}, 404)
        if not self.require_auth(path):
            return
        if path in ("/", "/index.html"):
            return self._send(200, "text/html; charset=utf-8", (APP_DIR / "index.html").read_bytes())
        if path.startswith("/assets/"):
            filename = Path(path).name
            target = APP_DIR / "assets" / filename
            if target.exists() and target.is_file():
                ext = target.suffix.lower()
                content_type = "application/octet-stream"
                if ext == ".png":
                    content_type = "image/png"
                elif ext in (".jpg", ".jpeg"):
                    content_type = "image/jpeg"
                elif ext == ".gif":
                    content_type = "image/gif"
                elif ext == ".svg":
                    content_type = "image/svg+xml"
                elif ext == ".css":
                    content_type = "text/css; charset=utf-8"
                elif ext == ".js":
                    content_type = "application/javascript; charset=utf-8"
                return self._send(200, content_type, target.read_bytes())
            return self.json({"error": "Asset no encontrado"}, 404)
        if path.startswith("/cotizaciones_pdf/"):
            filename = Path(path).name
            target = QUOTE_DIR / filename
            if target.exists() and target.is_file():
                return self._send(200, "application/pdf", target.read_bytes())
        if path.startswith("/cotizaciones_logradas/"):
            filename = Path(path).name
            target = QUOTE_WON_DIR / filename
            if target.exists() and target.is_file():
                return self._send(200, "application/pdf", target.read_bytes())
            return self.json({"error": "PDF no encontrado"}, 404)
        if path == "/api/options":
            
            with conn() as db:
                cfg = get_settings_payload(db)
            return self.json({"municipios": NAYARIT_MUNICIPIOS, "productos": cfg["productos"], "tarifas": cfg["tarifas"], "regiones": REGIONES})
        if path == "/api/clients":
            q = (qs.get("q", [""])[0] or "").lower()
            with conn() as db:
                if q:
                    like = f"%{q}%"
                    rows = db.execute(
                        """
                        SELECT * FROM clients
                        WHERE lower(nombre_comercial) LIKE ? OR lower(contacto) LIKE ? OR lower(municipio) LIKE ?
                           OR lower(tipo_cliente) LIKE ? OR lower(etapa) LIKE ? OR lower(potencial) LIKE ?
                           OR telefono LIKE ? OR whatsapp LIKE ?
                        ORDER BY updated_at DESC, nombre_comercial
                        """,
                        (like, like, like, like, like, like, like, like),
                    ).fetchall()
                else:
                    rows = db.execute(
                        """
                        SELECT * FROM clients
                        ORDER BY
                          CASE zona
                            WHEN 'Costa Sur' THEN 1
                            WHEN 'Costa Centro y Valles del Norte' THEN 2
                            WHEN 'Costa Norte y Frontera' THEN 3
                            WHEN 'Centro y Sierra Centro' THEN 4
                            WHEN 'Sierra y Sur' THEN 5
                            ELSE 6
                          END,
                          nombre_comercial COLLATE NOCASE
                        """
                    ).fetchall()
            return self.json(rows_to_dicts(rows))
        if path == "/api/client":
            cid = qs.get("id", [""])[0]
            with conn() as db:
                client = db.execute("SELECT * FROM clients WHERE id=?", (cid,)).fetchone()
                if not client:
                    return self.json({"error": "Cliente no encontrado"}, 404)
                hist = db.execute("SELECT * FROM interactions WHERE client_id=? ORDER BY fecha DESC, hora DESC, id DESC", (cid,)).fetchall()
            return self.json({"client": dict(client), "history": rows_to_dicts(hist)})
        if path == "/api/dashboard":
            today = date.today()
            overdue = today.isoformat()
            thirty = (today - timedelta(days=30)).isoformat()
            month_prefix = today.strftime("%Y-%m")
            with conn() as db:
                total = db.execute("SELECT COUNT(*) FROM clients").fetchone()[0]
                active = db.execute("SELECT COUNT(*) FROM clients WHERE etapa='Cliente activo'").fetchone()[0]
                pending_quotes = db.execute("SELECT COUNT(*) FROM clients WHERE etapa='Cotización enviada'").fetchone()[0]
                quotes_month = db.execute("SELECT COUNT(*) FROM quotes WHERE fecha LIKE ?", (month_prefix + "%",)).fetchone()[0]
                quotes_won = db.execute("SELECT COUNT(*) FROM quotes WHERE status='Lograda'").fetchone()[0]
                quotes_total_amount = db.execute("SELECT COALESCE(SUM(total),0) FROM quotes").fetchone()[0]
                quotes_won_amount = db.execute("SELECT COALESCE(SUM(total),0) FROM quotes WHERE status='Lograda'").fetchone()[0]
                follows = db.execute("SELECT COUNT(*) FROM interactions WHERE fecha_seguimiento<>'' AND fecha_seguimiento<=?", (overdue,)).fetchone()[0]
                by_stage = rows_to_dicts(db.execute("SELECT etapa, COUNT(*) total FROM clients GROUP BY etapa ORDER BY total DESC").fetchall())
                by_mun = rows_to_dicts(db.execute("SELECT municipio, zona, COUNT(*) total FROM clients GROUP BY municipio, zona ORDER BY total DESC").fetchall())
                by_zone_clients = rows_to_dicts(db.execute(
                    """
                    SELECT id, nombre_comercial, contacto, municipio, localidad, zona, etapa, potencial, tipo_cliente, whatsapp, telefono
                    FROM clients
                    ORDER BY
                      CASE zona
                        WHEN 'Costa Sur' THEN 1
                        WHEN 'Costa Centro y Valles del Norte' THEN 2
                        WHEN 'Costa Norte y Frontera' THEN 3
                        WHEN 'Centro y Sierra Centro' THEN 4
                        WHEN 'Sierra y Sur' THEN 5
                        ELSE 6
                      END,
                      municipio,
                      nombre_comercial
                    """
                ).fetchall())
                alerts = rows_to_dicts(db.execute(
                    """
                    SELECT c.id, c.nombre_comercial, c.municipio, c.localidad, c.zona, c.etapa, c.info_enviada,
                           MAX(i.fecha) ultima_fecha,
                           MIN(CASE WHEN i.fecha_seguimiento<>'' THEN i.fecha_seguimiento END) seguimiento
                    FROM clients c LEFT JOIN interactions i ON i.client_id=c.id
                    GROUP BY c.id
                    HAVING (seguimiento IS NOT NULL AND seguimiento<=?)
                       OR (ultima_fecha IS NOT NULL AND ultima_fecha<=?)
                       OR c.etapa='Cotización enviada'
                       OR c.info_enviada=''
                    ORDER BY seguimiento ASC
                    LIMIT 60
                    """,
                    (overdue, thirty),
                ).fetchall())
                volume = db.execute("SELECT COUNT(*) FROM clients WHERE potencial='A'").fetchone()[0]
                cfg = get_settings_payload(db)
                month_now = today.month
                sales_rows = cfg["monthlySales"]
                current_sales = next((r for r in sales_rows if int(r["month"]) == month_now), {"amount": 0, "goal": 0})
                annual_goal = float(cfg["settings"].get("annual_goal") or 0)
                annual_sales = sum(float(r.get("amount") or 0) for r in sales_rows)
            return self.json({"total": total, "active": active, "pendingQuotes": pending_quotes, "quotesMonth": quotes_month, "quotesWon": quotes_won, "quotesTotalAmount": quotes_total_amount, "quotesWonAmount": quotes_won_amount, "followUps": follows, "byStage": by_stage, "byMunicipio": by_mun, "byZoneClients": by_zone_clients, "alerts": alerts, "altoPotencial": volume, "settings": cfg["settings"], "monthlySales": sales_rows, "currentSales": current_sales, "annualSales": annual_sales, "annualGoal": annual_goal, "months": MESES_ES})
        if path == "/api/quotes":
            with conn() as db:
                rows = db.execute("""
                    SELECT q.*, c.municipio, c.zona, c.contacto, c.whatsapp, c.telefono
                    FROM quotes q LEFT JOIN clients c ON c.id=q.client_id
                    ORDER BY q.created_at DESC, q.id DESC
                """).fetchall()
            return self.json(rows_to_dicts(rows))
        if path == "/api/settings":
            year = qs.get("year", [str(date.today().year)])[0]
            with conn() as db:
                return self.json(get_settings_payload(db, year))
        if path == "/api/export.xlsx":
            data = export_xlsx()
            if data is None:
                return self.json({"error": "No se pudo exportar a Excel"}, 500)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=bdp_crm_export.xlsx")
            self.end_headers()
            self.wfile.write(data)
            return
        return self.json({"error": "Ruta no encontrada"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        ctype = self.headers.get("Content-Type", "")
        if parsed.path == "/api/login":
            raw = self.rfile.read(length) if length else b""
            form = parse_qs(raw.decode("utf-8"))
            user = (form.get("user", [""])[0] or "").strip()
            password = form.get("password", [""])[0] or ""
            if hmac.compare_digest(user, auth_user()) and hmac.compare_digest(password, auth_password()):
                return self._send(
                    303,
                    "text/plain; charset=utf-8",
                    "OK",
                    {
                        "Set-Cookie": f"crm_session={session_token()}; Path=/; HttpOnly; SameSite=Lax",
                        "Location": "/",
                    },
                )
            return self._send(401, "text/html; charset=utf-8", login_page("Usuario o contraseña incorrectos"))
        if parsed.path == "/api/logout":
            return self._send(
                303,
                "text/plain; charset=utf-8",
                "OK",
                {
                    "Set-Cookie": "crm_session=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax",
                    "Location": "/login",
                },
            )
        if not self.require_auth(parsed.path):
            return
        if parsed.path == "/api/import_excel":
            if cgi is not None:
                env = {"REQUEST_METHOD": "POST", "CONTENT_TYPE": ctype}
                form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=env)
                fileitem = form["file"] if "file" in form else None
                if not fileitem:
                    return self.json({"error": "No se recibió archivo"}, 400)
                data = fileitem.file.read()
                filename = fileitem.filename or "Carga manual"
            else:
                raw_upload = self.rfile.read(length) if length else b""
                data, filename = parse_upload_file(raw_upload, ctype)
                if not data:
                    return self.json({"error": "No se recibió archivo"}, 400)
            result = import_excel(io.BytesIO(data), filename)
            return self.json(result)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            data = json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            return self.json({"error": "JSON inválido"}, 400)
        if parsed.path == "/api/shutdown":
            self.json({"ok": True, "message": "CRM apagado"})
            threading.Timer(0.35, self.server.shutdown).start()
            return
        if parsed.path == "/api/client":
            with conn() as db:
                cid = upsert_client(db, data)
                db.commit()
            return self.json({"ok": True, "id": cid})
        if parsed.path == "/api/interaction":
            if data.get("id"):
                with conn() as db:
                    cid = update_interaction(db, data.get("id"), data)
                    if not cid:
                        return self.json({"error": "Historial no encontrado"}, 404)
                    db.commit()
                return self.json({"ok": True, "client_id": cid})
            cid = data.get("client_id")
            if not cid:
                return self.json({"error": "Falta cliente"}, 400)
            with conn() as db:
                add_interaction(db, cid, data)
                if data.get("etapa"):
                    db.execute("UPDATE clients SET etapa=?, updated_at=? WHERE id=?", (data["etapa"], now_iso(), cid))
                db.commit()
            return self.json({"ok": True})
        if parsed.path == "/api/quote":
            cid = data.get("client_id")
            if not cid:
                return self.json({"error": "Falta cliente"}, 400)
            cantidad = float(data.get("cantidad") or 0)
            precio = float(data.get("precio_unitario") or 0)
            km_flete = float(data.get("km_flete") or 0)
            tarifa_km = float(data.get("tarifa_km") or 0)
            casetas = float(data.get("casetas") or 0)
            flete_base = km_flete * tarifa_km
            flete_total = flete_base + casetas
            flete_unitario = flete_total / cantidad if cantidad else 0
            costo_individual = precio + flete_unitario
            total = (precio * cantidad) + flete_total
            data["flete_base"] = flete_base
            data["flete_total"] = flete_total
            data["flete_unitario"] = flete_unitario
            data["casetas"] = casetas
            data["precio_ofrecido"] = data.get("precio_ofrecido") or costo_individual
            data["total"] = total
            data["tipo_contacto"] = "Cotización"
            data["resultado"] = data.get("resultado") or "Cotización enviada"
            data["etapa"] = data.get("etapa") or "Cotización enviada"
            with conn() as db:
                client = db.execute("SELECT * FROM clients WHERE id=?", (cid,)).fetchone()
                if not client:
                    return self.json({"error": "Cliente no encontrado"}, 404)
                add_interaction(db, cid, data)
                db.execute("UPDATE clients SET etapa=?, updated_at=? WHERE id=?", ("Cotización enviada", now_iso(), cid))
                db.commit()
            filename = generate_quote_pdf(client, data)
            pdf_url = f"/cotizaciones_pdf/{filename}" if filename else ""
            with conn() as db:
                db.execute(
                    """
                    INSERT INTO quotes
                    (client_id, fecha, cliente, producto, cantidad, precio_unitario, flete_total, casetas, total, costo_individual, status, pdf, created_at, updated_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        cid,
                        data.get("fecha") or date.today().isoformat(),
                        client["nombre_comercial"],
                        data.get("producto") or data.get("productos_ofrecidos") or "",
                        cantidad,
                        precio,
                        flete_total,
                        casetas,
                        total,
                        costo_individual,
                        "Elaborada",
                        pdf_url,
                        now_iso(),
                        now_iso(),
                    ),
                )
                db.commit()
            return self.json({"ok": True, "pdf": pdf_url, "whatsapp": ""})

        if parsed.path == "/api/quote_status":
            qid = data.get("id")
            status = data.get("status") or "Elaborada"
            if not qid:
                return self.json({"error": "Falta cotizacion"}, 400)
            with conn() as db:
                quote = db.execute("SELECT * FROM quotes WHERE id=?", (qid,)).fetchone()
                if not quote:
                    return self.json({"error": "Cotizacion no encontrada"}, 404)
                pdf_url = quote["pdf"] or ""
                if status == "Lograda" and pdf_url.startswith("/cotizaciones_pdf/"):
                    source = QUOTE_DIR / Path(pdf_url).name
                    target = QUOTE_WON_DIR / source.name
                    if source.exists():
                        shutil.copy2(source, target)
                        pdf_url = f"/cotizaciones_logradas/{target.name}"
                    db.execute("UPDATE clients SET etapa=?, updated_at=? WHERE id=?", ("Cliente activo", now_iso(), quote["client_id"]))
                db.execute("UPDATE quotes SET status=?, pdf=?, updated_at=? WHERE id=?", (status, pdf_url, now_iso(), qid))
                db.commit()
            return self.json({"ok": True, "status": status, "pdf": pdf_url})

        if parsed.path == "/api/settings":
            with conn() as db:
                for k, v in (data.get("settings") or {}).items():
                    db.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))
                for prod, precio in (data.get("productos") or {}).items():
                    db.execute("INSERT INTO product_prices(producto,precio,updated_at) VALUES(?,?,?) ON CONFLICT(producto) DO UPDATE SET precio=excluded.precio, updated_at=excluded.updated_at", (prod, precio if precio != '' else None, now_iso()))
                for zona, tarifa in (data.get("tarifas") or {}).items():
                    db.execute("INSERT INTO freight_rates(zona,tarifa,updated_at) VALUES(?,?,?) ON CONFLICT(zona) DO UPDATE SET tarifa=excluded.tarifa, updated_at=excluded.updated_at", (zona, tarifa if tarifa != '' else None, now_iso()))
                db.commit()
            return self.json({"ok": True})
        if parsed.path == "/api/monthly_sales":
            year = int(data.get("year") or date.today().year)
            with conn() as db:
                for row in data.get("monthlySales", []):
                    month = int(row.get("month") or 0)
                    if 1 <= month <= 12:
                        amount = float(row.get("amount") or 0)
                        goal = float(row.get("goal") or 0)
                        db.execute("INSERT INTO monthly_sales(year,month,amount,goal,updated_at) VALUES(?,?,?,?,?) ON CONFLICT(year,month) DO UPDATE SET amount=excluded.amount, goal=excluded.goal, updated_at=excluded.updated_at", (year, month, amount, goal, now_iso()))
                db.commit()
            return self.json({"ok": True})
        if parsed.path == "/api/quick":
            with conn() as db:
                cid = upsert_client(db, data)
                add_interaction(db, cid, data)
                db.commit()
            return self.json({"ok": True, "id": cid})
        return self.json({"error": "Ruta no encontrada"}, 404)

    def do_DELETE(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        if not self.require_auth(parsed.path):
            return
        if parsed.path == "/api/client":
            cid = qs.get("id", [""])[0]
            if not cid:
                return self.json({"error": "Falta ID de cliente"}, 400)
            with conn() as db:
                db.execute("DELETE FROM interactions WHERE client_id=?", (cid,))
                cur = db.execute("DELETE FROM clients WHERE id=?", (cid,))
                db.commit()
            return self.json({"ok": True, "deleted": cur.rowcount})
        if parsed.path == "/api/interaction":
            iid = qs.get("id", [""])[0]
            if not iid:
                return self.json({"error": "Falta ID de historial"}, 400)
            with conn() as db:
                row = db.execute("SELECT client_id FROM interactions WHERE id=?", (iid,)).fetchone()
                if not row:
                    return self.json({"error": "Historial no encontrado"}, 404)
                cur = db.execute("DELETE FROM interactions WHERE id=?", (iid,))
                db.execute("UPDATE clients SET updated_at=? WHERE id=?", (now_iso(), row["client_id"]))
                db.commit()
            return self.json({"ok": True, "deleted": cur.rowcount, "client_id": row["client_id"]})
        return self.json({"error": "Ruta no encontrada"}, 404)


def main():
    bootstrap_persistent_data()
    init_db()
    seed_initial_data()
    auto_backup()
    host = "0.0.0.0"
    port = int(os.environ.get("PORT", "8765"))
    print(f"CRM Bazar de Prefabricados listo en http://127.0.0.1:{port}")
    print("Para celular: abre la IP de esta computadora en la misma red, puerto 8765.")
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == "__main__":
    main()
