# =============================================================================
#  DASHBOARD DE DEVOLUCIONES Y RECLAMACIONES — Seguimiento por etapas
#  -----------------------------------------------------------------------------
#  Aplicación Streamlit que gestiona el ciclo de vida completo de una
#  reclamación a través de 4 etapas encadenadas (máquina de estados):
#
#     1) Reporte de reclamo   (7 días)
#     2) Gestión              (30 días)
#     3) Disposición final    (recolección con contador de 20 días / destrucción)
#     4) Cuentas por pagar    (53 días) — cierra el proceso
#
#  Cada etapa registra fecha y usuario de cada estatus, observaciones, y al
#  cerrarse bloquea su pestaña y activa la siguiente. Las etapas se pueden
#  reabrir con la clave de reactivación. Los cambios se guardan en datos.xlsx
#  y se suben automáticamente a GitHub con GitPython.
#
#  El identificador único de cada reclamación es FOLIO REPORTE.
# =============================================================================

from __future__ import annotations

import os
from datetime import datetime, date, timedelta
try:
    from zoneinfo import ZoneInfo
    ZONA_MX = ZoneInfo("America/Mexico_City")
except Exception:
    ZONA_MX = None

import pandas as pd
import streamlit as st

# Altair viene incluido con Streamlit; se usa para dar formato a los ejes y a
# las etiquetas emergentes de las gráficas (separador de miles y moneda).
try:
    import altair as alt
    ALTAIR_OK = True
except Exception:
    ALTAIR_OK = False


def hoy_mx() -> date:
    """Fecha actual en la zona horaria de México (evita el desfase por UTC).

    Streamlit Cloud corre en UTC; sin esto, en las tardes/noches de México
    la app graba la fecha del día siguiente.
    """
    if ZONA_MX is not None:
        return datetime.now(ZONA_MX).date()
    # Respaldo: UTC menos 6 horas (horario del centro de México)
    return (datetime.utcnow() - timedelta(hours=6)).date()


def ahora_mx() -> datetime:
    """Fecha y hora actual en la zona horaria de México (sin tzinfo, para logs)."""
    if ZONA_MX is not None:
        return datetime.now(ZONA_MX).replace(tzinfo=None)
    return datetime.utcnow() - timedelta(hours=6)

# =============================================================================
# 1. CONFIGURACIÓN GENERAL Y CONSTANTES
# =============================================================================

st.set_page_config(
    page_title="Seguimiento a devoluciones",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_EXCEL = os.path.join(RUTA_BASE, "datos.xlsx")
NOMBRE_HOJA = "datos"

# Clave para las acciones protegidas (reactivar etapas, ajuste manual de fechas
# y corrección de modalidad).
#
# Se lee de los secretos de Streamlit: CLAVE_REACTIVACION. Si no está definida
# ahí, se usa la clave de respaldo de abajo. Para cambiarla sin tocar el código,
# agrégala en Streamlit Cloud: ⋮ → Settings → Secrets
#     CLAVE_REACTIVACION = "TuNuevaClave"
CLAVE_REACTIVACION_RESPALDO = "devoluciones2026"


def clave_autorizacion() -> str:
    """Devuelve la clave para desbloquear acciones protegidas."""
    return st.secrets.get("CLAVE_REACTIVACION", CLAVE_REACTIVACION_RESPALDO)

# --- Columnas base (encabezados normalizados a MAYÚSCULAS sin saltos) ---
COL_MES = "MES DE DEVOLUCION"
COL_MES_ETIQUETA = "MES (ETIQUETA)"
COL_FOLIO = "FOLIO REPORTE"          # identificador principal
COL_ID = "ID"
COL_PROVEEDOR = "PROVEEDOR"
COL_CARTA_FIRMADA = "CARTA FIRMADA"  # informativo
COL_IMPORTE = "IMPORTE (MXN)"
COL_COMPRADOR = "COMPRADOR"
COL_FECHA_CORTE = "FECHA CORTE"      # fecha de inicio para contar días
COL_NOTAS = "ACCIONES / NOTAS"

# --- Estado global de la reclamación ---
COL_ETAPA = "ETAPA ACTUAL"
COL_RESPUESTA_TIPO = "RESPUESTA PROVEEDOR TIPO"
COL_MODIFICADO_POR = "MODIFICADO POR"
COL_FECHA_MODIFICACION = "FECHA MODIFICACIÓN"

# Nombres de las etapas
ETAPA_1 = "Reporte de reclamo"
ETAPA_2 = "Gestión"
ETAPA_3 = "Cuentas por pagar"
ETAPA_4 = "Disposición final"

# Nombres antiguos que puedan existir en la base de datos, para migrarlos
# automáticamente al cargar (evita que aparezcan etiquetas obsoletas en
# tablas, filtros y gráficas).
NOMBRES_ETAPA_ANTIGUOS = {
    "Destino final": ETAPA_4,
    "DESTINO FINAL": ETAPA_4,
    "Destino Final": ETAPA_4,
}
ETAPA_FINAL = "FINALIZADO"

# Duraciones (días) de cada etapa.
DIAS_ETAPA_1 = 7    # desde FECHA CORTE
DIAS_ETAPA_2 = 30   # desde el vencimiento de la etapa 1
DIAS_ETAPA_3 = 53   # desde el vencimiento de la etapa 2
DIAS_VENCIMIENTO_TOTAL = 90  # desde FECHA CORTE: vencimiento global del reclamo
MSG_VENCIDO_90 = "RECLAMO VENCIDO SIN DEFINICIÓN, ENVIAR A DESTRUCCIÓN"
DIAS_RECOLECCION = 20  # desde que se define la recolección (etapa 4)

# Tipos de respuesta del proveedor (etapa 2)
RESP_RECOLECCION = "Recolección"
RESP_DESTRUCCION = "Destrucción"
RESP_SIN = "Sin respuesta"

# --- Pasos (estatus) de cada etapa, en orden ---
# Cada paso: (clave, etiqueta, col_fecha, col_usuario)
PASOS_E1 = [
    ("recepcion", "Recepción de Folio", "E1 FECHA RECEPCION FOLIO", "E1 USUARIO RECEPCION"),
    ("revision", "Revisión de folio", "E1 FECHA REVISION FOLIO", "E1 USUARIO REVISION"),
    ("envio", "Envío a proveedores", "E1 FECHA ENVIO PROVEEDORES", "E1 USUARIO ENVIO"),
]
PASOS_E2 = [
    ("enviado", "Enviado a proveedor", "E2 FECHA ENVIADO PROVEEDOR", "E2 USUARIO ENVIADO"),
    ("seguimiento", "Seguimiento", "E2 FECHA SEGUIMIENTO", "E2 USUARIO SEGUIMIENTO"),
    ("respuesta", "Respuesta de proveedor", "E2 FECHA RESPUESTA", "E2 USUARIO RESPUESTA"),
]
PASOS_E3 = [
    ("seguimiento", "Seguimiento", "E3 FECHA SEGUIMIENTO", "E3 USUARIO SEGUIMIENTO"),
    ("recepcion_nc", "Recepción de nota de crédito", "E3 FECHA RECEPCION NC", "E3 USUARIO RECEPCION NC"),
    ("aplicacion", "Aplicación de pago", "E3 FECHA APLICACION PAGO", "E3 USUARIO APLICACION"),
]
PASOS_E4_DESTRUCCION = [
    ("reporte_almacen", "Reporte al almacén de devoluciones", "E4 FECHA REPORTE ALMACEN", "E4 USUARIO REPORTE ALMACEN"),
    ("folio_ajuste", "Recepción de folio de ajuste", "E4 FECHA RECEPCION FOLIO AJUSTE", "E4 USUARIO FOLIO AJUSTE"),
]
PASOS_E4_RECOLECCION = [
    ("programacion", "Programación", "E4 FECHA PROGRAMACION", "E4 USUARIO PROGRAMACION"),
    ("recoleccion", "Recolección", "E4 FECHA RECOLECCION", "E4 USUARIO RECOLECCION"),
    ("folio_dev", "Recepción de folio de devolución", "E4 FECHA RECEPCION FOLIO DEV", "E4 USUARIO FOLIO DEV"),
]

# Columnas de control por etapa
COL_E1_ESTATUS, COL_E1_OBS = "E1 ESTATUS", "E1 OBSERVACIONES"
COL_E1_LIMITE, COL_E1_DIAS, COL_E1_TERM = "E1 FECHA LIMITE", "E1 DIAS TARDO", "E1 TERMINADA"
COL_E2_ESTATUS, COL_E2_OBS = "E2 ESTATUS", "E2 OBSERVACIONES"
COL_E2_COMPROMISO = "E2 FECHA COMPROMISO RECOLECCION"
COL_E2_LIMITE, COL_E2_DIAS, COL_E2_TERM = "E2 FECHA LIMITE", "E2 DIAS TARDO", "E2 TERMINADA"
COL_E3_ESTATUS, COL_E3_OBS = "E3 ESTATUS", "E3 OBSERVACIONES"
COL_E3_LIMITE, COL_E3_DIAS, COL_E3_TERM = "E3 FECHA LIMITE", "E3 DIAS TARDO", "E3 TERMINADA"
COL_E4_MODALIDAD, COL_E4_ESTATUS, COL_E4_OBS = "E4 MODALIDAD", "E4 ESTATUS", "E4 OBSERVACIONES"
COL_E4_LIMITE_REC, COL_E4_TERM = "E4 FECHA LIMITE RECOLECCION", "E4 TERMINADA"

# Semáforo
EN_TIEMPO = "🟢 EN TIEMPO"
POR_VENCERSE = "🟡 POR VENCERSE"
VENCIDO = "🔴 VENCIDO"
TERMINADO = "✅ TERMINADO"
SIN_DATO = "—"

# Todas las columnas que la app necesita (para crearlas si faltan).
COLUMNAS_REQUERIDAS = [
    COL_ETAPA, COL_RESPUESTA_TIPO,
    COL_E1_ESTATUS, COL_E1_OBS, COL_E1_LIMITE, COL_E1_DIAS, COL_E1_TERM,
    COL_E2_ESTATUS, COL_E2_OBS, COL_E2_COMPROMISO, COL_E2_LIMITE, COL_E2_DIAS, COL_E2_TERM,
    COL_E3_ESTATUS, COL_E3_OBS, COL_E3_LIMITE, COL_E3_DIAS, COL_E3_TERM,
    COL_E4_MODALIDAD, COL_E4_ESTATUS, COL_E4_OBS, COL_E4_LIMITE_REC, COL_E4_TERM,
    COL_MODIFICADO_POR, COL_FECHA_MODIFICACION,
]
for _pasos in (PASOS_E1, PASOS_E2, PASOS_E3, PASOS_E4_DESTRUCCION, PASOS_E4_RECOLECCION):
    for _clave, _etq, _cf, _cu in _pasos:
        COLUMNAS_REQUERIDAS.extend([_cf, _cu])


# =============================================================================
# 2. UTILIDADES DE FECHA Y SEMÁFORO
# =============================================================================

def _a_fecha(valor):
    """Convierte datetime, serial de Excel o texto a date. None si no es válido."""
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)):
        try:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(valor))).date()
        except Exception:
            return None
    try:
        return pd.to_datetime(str(valor), dayfirst=True, errors="raise").date()
    except Exception:
        return None


def _fmt_fecha(valor) -> str:
    """Formatea una fecha como DD/MM/AAAA, o '—' si no hay."""
    f = _a_fecha(valor)
    return f"{f:%d/%m/%Y}" if f else SIN_DATO


def semaforo_por_limite(fecha_limite, terminada: bool) -> str:
    """Semáforo de una etapa según su fecha límite.

    - Si la etapa ya terminó -> TERMINADO
    - Si no hay fecha límite  -> SIN_DATO
    - Vencida                 -> VENCIDO
    - 15 días o menos (incl. hoy) -> POR_VENCERSE
    - Más de 15 días          -> EN_TIEMPO
    """
    if terminada:
        return TERMINADO
    f = _a_fecha(fecha_limite)
    if f is None:
        return SIN_DATO
    hoy = hoy_mx()
    if f < hoy:
        return VENCIDO
    if f <= hoy + timedelta(days=15):
        return POR_VENCERSE
    return EN_TIEMPO


def es_verdadero(valor) -> bool:
    """Interpreta un valor de celda como booleano (terminada / sí)."""
    return str(valor).strip().upper() in ("SÍ", "SI", "TRUE", "1", "VERDADERO", "X")


# =============================================================================
# 3. MÓDULO DE DATOS
# =============================================================================

def _normalizar_encabezado(nombre: str) -> str:
    return " ".join(str(nombre).split()).upper()


@st.cache_data(show_spinner="Cargando base de datos…")
def cargar_datos(ruta: str, _version: int) -> pd.DataFrame:
    """Carga el Excel, normaliza encabezados y asegura las columnas nuevas."""
    df = pd.read_excel(ruta, sheet_name=NOMBRE_HOJA)
    df.columns = [_normalizar_encabezado(c) for c in df.columns]

    # Crear columnas faltantes (vacías)
    for col in COLUMNAS_REQUERIDAS:
        if col not in df.columns:
            df[col] = ""

    # Las columnas de control (estatus, terminada, observaciones, usuarios,
    # modalidad, etc.) deben ser de tipo texto. Si vienen vacías, pandas las
    # carga como float64 y rechazaría valores como "SÍ". Se excluyen las de
    # fecha y las numéricas de días.
    cols_fecha = set()
    for _pasos in (PASOS_E1, PASOS_E2, PASOS_E3, PASOS_E4_DESTRUCCION, PASOS_E4_RECOLECCION):
        for _c, _e, _cf, _cu in _pasos:
            cols_fecha.add(_cf)
    cols_fecha |= {COL_E1_LIMITE, COL_E2_LIMITE, COL_E3_LIMITE,
                   COL_E2_COMPROMISO, COL_E4_LIMITE_REC}
    cols_numericas = {COL_E1_DIAS, COL_E2_DIAS, COL_E3_DIAS}
    for col in COLUMNAS_REQUERIDAS:
        if col in df.columns and col not in cols_fecha and col not in cols_numericas:
            df[col] = df[col].fillna("").astype(str).replace("nan", "")
    # Las columnas de fecha a datetime SIN componente de hora (solo fecha).
    # .dt.normalize() pone la hora en 00:00 para que no se guarde hora alguna.
    for col in cols_fecha:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()
    # También normalizar FECHA CORTE (fecha base del proceso).
    if COL_FECHA_CORTE in df.columns:
        df[COL_FECHA_CORTE] = pd.to_datetime(df[COL_FECHA_CORTE], errors="coerce").dt.normalize()

    # Etapa inicial por defecto
    if COL_ETAPA in df.columns:
        df[COL_ETAPA] = df[COL_ETAPA].replace("", pd.NA).fillna(ETAPA_1)
        # Migración de nombres antiguos: los registros guardados antes del
        # cambio de nombre conservan "Destino final" en la base. Se renombran
        # a la etiqueta vigente para que tablas, filtros y gráficas coincidan.
        df[COL_ETAPA] = df[COL_ETAPA].astype(str).str.strip().replace(
            NOMBRES_ETAPA_ANTIGUOS
        )

    # Texto clave normalizado
    for col in (COL_PROVEEDOR, COL_COMPRADOR, COL_FOLIO):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # FOLIO REPORTE es el identificador principal y único.
    df["CLAVE"] = df[COL_FOLIO].astype(str)

    # Etiqueta legible del mes (para el filtro)
    if COL_MES in df.columns:
        meses_es = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo",
            6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre",
            10: "Octubre", 11: "Noviembre", 12: "Diciembre",
        }
        fm = pd.to_datetime(df[COL_MES], errors="coerce")
        df[COL_MES_ETIQUETA] = fm.apply(
            lambda d: f"{meses_es[d.month]} {d.year}" if pd.notna(d) else "Sin fecha"
        )
    else:
        df[COL_MES_ETIQUETA] = "Sin fecha"

    return df


def _columnas_fecha() -> set:
    """Conjunto de todas las columnas que deben tratarse como fecha."""
    cols = set()
    for _pasos in (PASOS_E1, PASOS_E2, PASOS_E3, PASOS_E4_DESTRUCCION, PASOS_E4_RECOLECCION):
        for _c, _e, _cf, _cu in _pasos:
            cols.add(_cf)
    cols |= {COL_E1_LIMITE, COL_E2_LIMITE, COL_E3_LIMITE,
             COL_E2_COMPROMISO, COL_E4_LIMITE_REC, COL_FECHA_CORTE}
    return cols


def _preparar_para_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Copia sin columnas auxiliares y con fechas como date puro (sin hora)."""
    salida = df.drop(columns=["CLAVE", COL_MES_ETIQUETA], errors="ignore").copy()
    for col in _columnas_fecha():
        if col in salida.columns:
            salida[col] = pd.to_datetime(salida[col], errors="coerce").dt.date
    return salida


def guardar_excel(df: pd.DataFrame, ruta: str) -> None:
    """Sobrescribe el Excel quitando columnas auxiliares y sin hora en las fechas."""
    salida = _preparar_para_excel(df)
    with pd.ExcelWriter(ruta, engine="openpyxl", datetime_format="DD/MM/YYYY",
                        date_format="DD/MM/YYYY") as writer:
        salida.to_excel(writer, sheet_name=NOMBRE_HOJA, index=False)


# =============================================================================
# 4. SINCRONIZACIÓN CON GITHUB
# =============================================================================

def subir_a_github(mensaje_commit: str) -> tuple[bool, str]:
    """Hace commit y push del Excel usando GitPython y st.secrets['GITHUB_TOKEN']."""
    try:
        from git import Repo
    except ImportError:
        return False, "GitPython no está instalado (revisa requirements.txt)."

    token = st.secrets.get("GITHUB_TOKEN")
    if not token:
        return False, (
            "No se encontró 'GITHUB_TOKEN' en los secretos. El cambio se guardó "
            "localmente, pero NO se subió a GitHub."
        )

    try:
        repo = Repo(RUTA_BASE, search_parent_directories=True)
        with repo.config_writer() as cw:
            cw.set_value("user", "name", "Dashboard Reclamaciones")
            cw.set_value("user", "email", "dashboard@reclamaciones.app")

        repo.index.add([RUTA_EXCEL])
        if not repo.index.diff("HEAD"):
            return True, "No había cambios nuevos que subir."

        repo.index.commit(mensaje_commit)

        origen = repo.remote(name="origin")
        url = origen.url
        if url.startswith("git@github.com:"):
            url = url.replace("git@github.com:", "https://github.com/")
        if not url.endswith(".git"):
            url = url + ".git"
        if "@" in url and url.startswith("https://"):
            url = "https://" + url.split("@", 1)[1]
        url_token = url.replace("https://", f"https://x-access-token:{token}@")

        entorno = {"GIT_TERMINAL_PROMPT": "0", "GIT_ASKPASS": "echo",
                   "GCM_INTERACTIVE": "never"}
        rama = repo.active_branch.name
        try:
            with repo.git.custom_environment(**entorno):
                repo.git.push(url_token, f"HEAD:{rama}")
        finally:
            origen.set_url(url)
        return True, "Cambios subidos a GitHub correctamente. ✅"
    except Exception as e:
        detalle = str(e)
        if "403" in detalle or "denied" in detalle.lower():
            ayuda = " — El token no tiene permiso de escritura (Contents: Read and write)."
        elif "could not read Password" in detalle or "Authentication" in detalle:
            ayuda = " — El token es inválido o expiró; genera uno nuevo."
        else:
            ayuda = ""
        return False, f"Error al subir a GitHub: {detalle}{ayuda}"


# =============================================================================
# 5. LÓGICA DE LA MÁQUINA DE ESTADOS
# =============================================================================

def calcular_fechas_limite(fila: pd.Series) -> dict:
    """Calcula las fechas límite de cada etapa a partir de FECHA CORTE.

    Flujo: Reporte (7d) → Gestión (30d) → Disposición final → Cuentas por pagar.
    Los plazos suman 90 días en total desde la fecha de corte:
      E1: FECHA CORTE + 7
      E2: límite E1 + 30   (día 37)
      E4: comparte el tramo final con E3 (día 90), pues va después de Gestión
      E3: límite E2 + 53   (día 90, cierre del proceso)
    Devuelve un dict con date o None por etapa.
    """
    corte = _a_fecha(fila.get(COL_FECHA_CORTE))
    limites = {"E1": None, "E2": None, "E3": None, "E4": None}
    if corte:
        limites["E1"] = corte + timedelta(days=DIAS_ETAPA_1)
        limites["E2"] = limites["E1"] + timedelta(days=DIAS_ETAPA_2)
        limites["E3"] = limites["E2"] + timedelta(days=DIAS_ETAPA_3)
        # Disposición final ocurre entre Gestión y Cuentas por pagar: su fecha
        # tope es la misma del cierre global (90 días desde el corte).
        limites["E4"] = limites["E3"]
    return limites


def fecha_vencimiento_90(fila: pd.Series):
    """Fecha de vencimiento global del reclamo: FECHA CORTE + 90 días."""
    corte = _a_fecha(fila.get(COL_FECHA_CORTE))
    return corte + timedelta(days=DIAS_VENCIMIENTO_TOTAL) if corte else None


def vencido_90_sin_definicion(fila: pd.Series) -> bool:
    """True si pasaron los 90 días desde FECHA CORTE y el reclamo NO está resuelto.

    'Sin definición' significa que el proceso aún no ha llegado a su fin
    (la etapa no es FINALIZADO), por lo que debe enviarse a destrucción.
    """
    etapa = str(fila.get(COL_ETAPA, "")).strip()
    if etapa == ETAPA_FINAL:
        return False
    limite = fecha_vencimiento_90(fila)
    return limite is not None and hoy_mx() > limite


def estado_alarma(fecha_limite, terminada: bool) -> tuple[str, str]:
    """Devuelve (nivel, mensaje) de alarma para una etapa activa.

    nivel: 'ok' | 'warn' | 'danger' | 'done' | 'none'
    """
    if terminada:
        return "done", "Etapa terminada."
    f = _a_fecha(fecha_limite)
    if f is None:
        return "none", "Sin fecha límite definida."
    hoy = hoy_mx()
    dias = (f - hoy).days
    if dias < 0:
        return "danger", f"VENCIDA hace {abs(dias)} día(s) (límite {f:%d/%m/%Y})."
    if dias <= 15:
        return "warn", f"Por vencerse: faltan {dias} día(s) (límite {f:%d/%m/%Y})."
    return "ok", f"En tiempo: faltan {dias} día(s) (límite {f:%d/%m/%Y})."


def construir_notificaciones(df: pd.DataFrame) -> list[dict]:
    """Genera la lista de alarmas (a punto de vencer / vencidas) de toda la base.

    La estructura está lista para, en el futuro, enviarse por correo:
    cada elemento trae folio, proveedor, etapa, nivel, mensaje y destinatario.
    """
    avisos = []
    for _, fila in df.iterrows():
        etapa = str(fila.get(COL_ETAPA, "")).strip()
        if etapa in ("", ETAPA_FINAL):
            continue

        # --- ALARMA CRÍTICA: vencimiento global de 90 días sin definición ---
        if vencido_90_sin_definicion(fila):
            lim90 = fecha_vencimiento_90(fila)
            dias_vencido = (hoy_mx() - lim90).days
            avisos.append({
                "folio": fila.get(COL_FOLIO, ""),
                "proveedor": fila.get(COL_PROVEEDOR, ""),
                "comprador": fila.get(COL_COMPRADOR, ""),
                "etapa": etapa,
                "nivel": "critico",
                "mensaje": (f"{MSG_VENCIDO_90} · Venció hace {dias_vencido} día(s) "
                            f"(90 días desde el corte: {lim90:%d/%m/%Y})."),
                "destinatario": "",
            })
            continue  # esta alarma reemplaza a la de etapa (es más grave)

        limites = calcular_fechas_limite(fila)
        # Determinar la fecha límite y bandera de terminada de la etapa activa
        if etapa == ETAPA_1:
            lim, term = limites["E1"], es_verdadero(fila.get(COL_E1_TERM))
        elif etapa == ETAPA_2:
            lim, term = limites["E2"], es_verdadero(fila.get(COL_E2_TERM))
        elif etapa == ETAPA_3:
            lim, term = limites["E3"], es_verdadero(fila.get(COL_E3_TERM))
        elif etapa == ETAPA_4:
            # El plazo de la etapa es el global; el de recolección (20 días) es
            # un contador aparte que se muestra dentro de la propia etapa.
            lim, term = limites["E4"], es_verdadero(fila.get(COL_E4_TERM))
        else:
            continue
        nivel, mensaje = estado_alarma(lim, term)
        if nivel in ("warn", "danger"):
            avisos.append({
                "folio": fila.get(COL_FOLIO, ""),
                "proveedor": fila.get(COL_PROVEEDOR, ""),
                "comprador": fila.get(COL_COMPRADOR, ""),
                "etapa": etapa,
                "nivel": nivel,
                "mensaje": mensaje,
                # Campo reservado para el futuro envío por correo:
                "destinatario": "",
            })
    # Orden: críticas (90 días) primero, luego vencidas, luego por vencerse
    prioridad = {"critico": 0, "danger": 1, "warn": 2}
    avisos.sort(key=lambda a: prioridad.get(a["nivel"], 3))
    return avisos


def aplicar_guardado(df: pd.DataFrame, clave: str, cambios: dict,
                     usuario: str, mensaje_log: str) -> str:
    """Aplica un diccionario de cambios a la fila y registra auditoría.

    Devuelve el mensaje de commit.
    """
    mascara = df["CLAVE"] == clave
    ahora = ahora_mx()
    for col, valor in cambios.items():
        df.loc[mascara, col] = valor
    # Bitácora acumulada en ACCIONES / NOTAS
    nota_previa = str(df.loc[mascara, COL_NOTAS].iloc[0] or "").strip()
    if nota_previa.lower() in ("nan", "none"):
        nota_previa = ""
    nueva = f"[{ahora:%d/%m/%Y %H:%M} · {usuario}] {mensaje_log}"
    df.loc[mascara, COL_NOTAS] = f"{nota_previa} | {nueva}" if nota_previa else nueva
    df.loc[mascara, COL_MODIFICADO_POR] = usuario
    df.loc[mascara, COL_FECHA_MODIFICACION] = f"{ahora:%d/%m/%Y %H:%M:%S}"
    folio = df.loc[mascara, COL_FOLIO].iloc[0]
    return f"Folio {folio}: {mensaje_log} (por {usuario}, {ahora:%d/%m/%Y %H:%M})"


def persistir_y_sincronizar(df: pd.DataFrame, mensaje_commit: str,
                            mensaje_ok: str, celebrar: bool = False) -> None:
    """Guarda el Excel, sube a GitHub, refresca la sesión y vuelve a renderizar.

    Si celebrar=True, se marca para mostrar los globos UNA sola vez tras el
    rerun (por ejemplo, al cerrar una etapa).
    """
    try:
        guardar_excel(df, RUTA_EXCEL)
    except Exception as e:
        st.error(f"No se pudo escribir el Excel: {e}")
        return
    with st.spinner("Subiendo cambios a GitHub…"):
        exito, mensaje = subir_a_github(mensaje_commit)

    cargar_datos.clear()
    st.session_state["version_datos"] += 1
    st.session_state["df"] = cargar_datos(RUTA_EXCEL, st.session_state["version_datos"])

    if exito:
        st.session_state["flash"] = ("success", f"✅ {mensaje_ok} {mensaje}")
    else:
        st.session_state["flash"] = ("warning", f"💾 {mensaje_ok} Guardado local, pero: {mensaje}")
    if celebrar:
        st.session_state["celebrar"] = True
    st.rerun()


# =============================================================================
# 6. AUTENTICACIÓN
# =============================================================================

def verificar_acceso() -> bool:
    """Pantalla de inicio de sesión con contraseña maestra (DASHBOARD_PASSWORD)."""
    if st.session_state.get("autenticado", False):
        return True

    st.markdown("<br><br>", unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        st.markdown("## 🔐 Seguimiento a devoluciones")
        st.caption("Acceso restringido. Ingresa la contraseña para continuar.")
        with st.form("form_login"):
            password = st.text_input("Contraseña", type="password", placeholder="••••••••")
            enviar = st.form_submit_button("Ingresar", use_container_width=True)
        if enviar:
            correcta = st.secrets.get("DASHBOARD_PASSWORD")
            if correcta is None:
                st.error("⚠️ No se encontró 'DASHBOARD_PASSWORD' en los secretos.")
            elif password == correcta:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
    return False


# =============================================================================
# 7. COMPONENTES DE UI COMPARTIDOS
# =============================================================================

def _mostrar_alarma(lim, terminada: bool) -> None:
    """Muestra el banner de alarma de la etapa activa."""
    nivel, mensaje = estado_alarma(lim, terminada)
    if nivel == "danger":
        st.error(f"🔴 {mensaje}")
    elif nivel == "warn":
        st.warning(f"🟡 {mensaje}")
    elif nivel == "ok":
        st.success(f"🟢 {mensaje}")
    elif nivel == "done":
        st.success(f"✅ {mensaje}")
    else:
        st.info(f"ℹ️ {mensaje}")


def _resumen_pasos(fila: pd.Series, pasos: list) -> None:
    """Muestra una línea por paso con su fecha y usuario registrados."""
    for _clave, etiqueta, cf, cu in pasos:
        fecha = _fmt_fecha(fila.get(cf))
        usuario = str(fila.get(cu, "") or "").strip() or "—"
        icono = "✅" if _a_fecha(fila.get(cf)) else "⬜"
        st.markdown(f"{icono} **{etiqueta}** — {fecha} · {usuario}")


def _boton_reactivar(etapa_cod: str, clave_registro: str) -> None:
    """Botón protegido por clave para reabrir una etapa terminada.

    Al reactivar una etapa, se limpia su marca de terminada y la de las etapas
    posteriores, para que el flujo se recorra de nuevo de forma limpia.
    """
    with st.expander("🔓 Reactivar esta etapa (requiere clave)"):
        st.caption(
            "Solo personal autorizado. Reactivar reabre la etapa para corregir "
            "información; las etapas posteriores se reinician y deberán completarse "
            "de nuevo."
        )
        c1, c2 = st.columns([2, 1])
        clave_in = c1.text_input(
            "Clave de reactivación", type="password",
            key=f"react_{etapa_cod}_{clave_registro}",
        )
        if c2.button("Reactivar", key=f"btn_react_{etapa_cod}_{clave_registro}",
                     use_container_width=True):
            if clave_in != clave_autorizacion():
                st.error("Clave incorrecta.")
                return
            df = st.session_state["df"]
            cambios = {COL_ETAPA: ETAPA_NOMBRE_POR_COD[etapa_cod]}
            # Reabrir esta etapa y todas las posteriores (limpiar terminada
            # y las fechas/usuarios de sus pasos, para recorrerlas de nuevo).
            orden = ["E1", "E2", "E4", "E3"]  # orden real del flujo
            pasos_por_cod = {
                "E1": PASOS_E1, "E2": PASOS_E2, "E3": PASOS_E3,
                "E4": PASOS_E4_DESTRUCCION + PASOS_E4_RECOLECCION,
            }
            desde = orden.index(etapa_cod)
            for cod in orden[desde:]:
                cambios[COL_TERM_POR_COD[cod]] = "NO"
                for _c, _e, cf, cu in pasos_por_cod[cod]:
                    cambios[cf] = pd.NaT
                    cambios[cu] = ""
            mensaje = aplicar_guardado(
                df, clave_registro, cambios, "Reactivación",
                f"Etapa '{ETAPA_NOMBRE_POR_COD[etapa_cod]}' reactivada",
            )
            persistir_y_sincronizar(df, mensaje, "Etapa reactivada.")


# Mapas auxiliares para reactivación
ETAPA_NOMBRE_POR_COD = {"E1": ETAPA_1, "E2": ETAPA_2, "E3": ETAPA_3, "E4": ETAPA_4}
COL_TERM_POR_COD = {"E1": COL_E1_TERM, "E2": COL_E2_TERM, "E3": COL_E3_TERM, "E4": COL_E4_TERM}
COL_DIAS_POR_COD = {"E1": COL_E1_DIAS, "E2": COL_E2_DIAS, "E3": COL_E3_DIAS}


def modalidad_destino_final(fila: pd.Series) -> str:
    """Determina la modalidad de la etapa 4 (Recolección o Destrucción).

    Prioridad:
      1. La columna E4 MODALIDAD, si está definida.
      2. Si no, se deriva de la respuesta del proveedor: solo 'Recolección'
         lleva a recolección; destrucción o sin respuesta llevan a destrucción.

    Esto evita que un registro con respuesta 'Recolección' aparezca como
    destrucción cuando la modalidad no quedó grabada.
    """
    modalidad = str(fila.get(COL_E4_MODALIDAD, "")).strip()
    if modalidad in (RESP_RECOLECCION, RESP_DESTRUCCION):
        return modalidad
    respuesta = str(fila.get(COL_RESPUESTA_TIPO, "")).strip()
    return RESP_RECOLECCION if respuesta == RESP_RECOLECCION else RESP_DESTRUCCION


def _pasos_de_etapa(cod: str, fila: pd.Series) -> list:
    """Devuelve la lista de pasos de una etapa. Para E4 depende de la modalidad."""
    if cod == "E1":
        return PASOS_E1
    if cod == "E2":
        return PASOS_E2
    if cod == "E3":
        return PASOS_E3
    modalidad = modalidad_destino_final(fila)
    return PASOS_E4_RECOLECCION if modalidad == RESP_RECOLECCION else PASOS_E4_DESTRUCCION


def _panel_ajuste_manual(etapa_cod: str, fila: pd.Series) -> None:
    """Panel protegido con clave para capturar/corregir fechas de pasos ya realizados.

    Sirve para registros que se completaron antes de usar la app: permite poner
    la fecha real de cada paso (retroactiva), quién lo hizo, y cerrar la etapa
    con esa fecha para que los días transcurridos se calculen correctamente.
    """
    clave = fila["CLAVE"]
    pasos = _pasos_de_etapa(etapa_cod, fila)
    nombre_etapa = ETAPA_NOMBRE_POR_COD[etapa_cod]

    with st.expander("🗓️ Ajuste manual de fechas (requiere clave)"):
        st.caption(
            "Usa esta opción cuando la etapa se realizó **antes** de usar la app y "
            "necesitas capturar las fechas reales. Puedes dejar en blanco los pasos "
            "que no apliquen."
        )
        clave_in = st.text_input(
            "Clave de autorización", type="password",
            key=f"aj_clave_{etapa_cod}_{clave}",
        )

        with st.form(f"form_ajuste_{etapa_cod}_{clave}"):
            st.markdown("##### Fecha y responsable de cada paso")
            valores = {}
            for pkey, etiqueta, cf, cu in pasos:
                c1, c2, c3 = st.columns([2, 2, 2])
                fecha_actual = _a_fecha(fila.get(cf))
                usar = c1.checkbox(
                    etiqueta, value=fecha_actual is not None,
                    key=f"aj_use_{etapa_cod}_{pkey}_{clave}",
                )
                fecha_in = c2.date_input(
                    "Fecha", value=fecha_actual or hoy_mx(), format="DD/MM/YYYY",
                    key=f"aj_fecha_{etapa_cod}_{pkey}_{clave}",
                    label_visibility="collapsed",
                )
                usuario_in = c3.text_input(
                    "Responsable",
                    value=str(fila.get(cu, "") or ""),
                    placeholder="Responsable",
                    key=f"aj_user_{etapa_cod}_{pkey}_{clave}",
                    label_visibility="collapsed",
                )
                valores[pkey] = (usar, fecha_in, usuario_in, cf, cu)

            st.divider()
            cerrar_etapa = st.checkbox(
                f"Marcar la etapa '{nombre_etapa}' como TERMINADA con estas fechas",
                value=es_verdadero(fila.get(COL_TERM_POR_COD[etapa_cod])),
                key=f"aj_cerrar_{etapa_cod}_{clave}",
                help="Al cerrarla se activa la siguiente etapa y se calculan los días "
                     "transcurridos con la fecha del último paso.",
            )

            # En la etapa 2, la respuesta del proveedor define hacia dónde avanza.
            respuesta_aj = None
            if etapa_cod == "E2":
                actual = str(fila.get(COL_RESPUESTA_TIPO, "")).strip()
                opciones = [RESP_RECOLECCION, RESP_DESTRUCCION, RESP_SIN]
                idx = opciones.index(actual) if actual in opciones else 0
                respuesta_aj = st.radio(
                    "Respuesta del proveedor (define la siguiente etapa)",
                    options=opciones, index=idx, horizontal=True,
                    key=f"aj_resp_{clave}",
                )

            firma = st.text_input(
                "Tu nombre / usuario (quien hace el ajuste)",
                key=f"aj_firma_{etapa_cod}_{clave}",
            )
            aplicar = st.form_submit_button(
                "💾 Aplicar ajuste manual", use_container_width=True
            )

        if not aplicar:
            return
        if clave_in != clave_autorizacion():
            st.error("🔒 Clave de autorización incorrecta.")
            return
        if not firma.strip():
            st.error("✍️ Firma con tu nombre antes de aplicar el ajuste.")
            return

        df = st.session_state["df"]
        cambios = {}
        fechas_marcadas = []
        for pkey, (usar, fecha_in, usuario_in, cf, cu) in valores.items():
            if usar:
                cambios[cf] = pd.Timestamp(fecha_in)
                cambios[cu] = usuario_in.strip() or firma.strip()
                fechas_marcadas.append(fecha_in)
            else:
                cambios[cf] = pd.NaT
                cambios[cu] = ""

        if etapa_cod == "E2" and respuesta_aj:
            cambios[COL_RESPUESTA_TIPO] = respuesta_aj

        mensaje_log = f"Ajuste manual de fechas en '{nombre_etapa}'"

        if cerrar_etapa:
            if not fechas_marcadas:
                st.error("Marca al menos un paso con su fecha antes de cerrar la etapa.")
                return
            cambios[COL_TERM_POR_COD[etapa_cod]] = "SÍ"
            fecha_cierre = max(fechas_marcadas)  # la fecha del último paso realizado

            # Días transcurridos, calculados con la fecha real de cierre.
            limites = calcular_fechas_limite(fila)
            if etapa_cod == "E1":
                corte = _a_fecha(fila.get(COL_FECHA_CORTE))
                if corte:
                    cambios[COL_E1_DIAS] = (fecha_cierre - corte).days
                cambios[COL_ETAPA] = ETAPA_2
                cambios[COL_E1_LIMITE] = (pd.Timestamp(limites["E1"])
                                          if limites["E1"] else pd.NaT)
            elif etapa_cod == "E2":
                if limites["E2"]:
                    cambios[COL_E2_DIAS] = (fecha_cierre - limites["E2"]).days + DIAS_ETAPA_2
                    cambios[COL_E2_LIMITE] = pd.Timestamp(limites["E2"])
                # Gestión SIEMPRE pasa a Disposición final.
                cambios[COL_ETAPA] = ETAPA_4
                if respuesta_aj == RESP_RECOLECCION:
                    cambios[COL_E4_MODALIDAD] = RESP_RECOLECCION
                    cambios[COL_E4_LIMITE_REC] = pd.Timestamp(
                        fecha_cierre + timedelta(days=DIAS_RECOLECCION)
                    )
                else:
                    cambios[COL_E4_MODALIDAD] = RESP_DESTRUCCION
            elif etapa_cod == "E4":
                # Disposición final SIEMPRE pasa a Cuentas por pagar.
                cambios[COL_ETAPA] = ETAPA_3
            else:  # E3 — última etapa: cierra el proceso
                if limites["E3"]:
                    cambios[COL_E3_DIAS] = (fecha_cierre - limites["E3"]).days + DIAS_ETAPA_3
                    cambios[COL_E3_LIMITE] = pd.Timestamp(limites["E3"])
                cambios[COL_ETAPA] = ETAPA_FINAL

            mensaje_log += f" · etapa TERMINADA con fecha {fecha_cierre:%d/%m/%Y}"
        else:
            # Si se desmarca el cierre, la etapa vuelve a quedar abierta.
            cambios[COL_TERM_POR_COD[etapa_cod]] = "NO"
            cambios[COL_ETAPA] = nombre_etapa

        mensaje = aplicar_guardado(df, clave, cambios, firma.strip(), mensaje_log)
        persistir_y_sincronizar(df, mensaje, "Ajuste manual aplicado.")


def _etapa_bloqueada(fila: pd.Series, etapa_cod: str) -> bool:
    """Una etapa está bloqueada para edición si ya fue terminada."""
    return es_verdadero(fila.get(COL_TERM_POR_COD[etapa_cod]))


def _paso_quedo_hecho(col_fecha: str, cambios: dict, fila: pd.Series) -> bool:
    """True si el paso (su fecha) quedó registrado: nuevo en cambios o ya en la fila."""
    if col_fecha in cambios:
        val = cambios[col_fecha]
        # Si se acaba de limpiar (NaT/""), no está hecho
        if val in ("", None) or (isinstance(val, float) and pd.isna(val)):
            return False
        try:
            if pd.isna(val):
                return False
        except (TypeError, ValueError):
            pass
        return True
    return _a_fecha(fila.get(col_fecha)) is not None


def _registrar_pasos_form(fila, pasos, prefijo, solo_lectura):
    """Renderiza un checkbox por paso. Devuelve (cambios, cols_usuario_nuevas).

    cambios: dict columna->valor a aplicar.
    cols_usuario_nuevas: lista de columnas de usuario a llenar con la firma.
    """
    cambios = {}
    cols_usuario_nuevas = []
    hoy = hoy_mx()
    for clave, etiqueta, cf, cu in pasos:
        ya_hecho = _a_fecha(fila.get(cf)) is not None
        col1, col2 = st.columns([3, 2])
        marcado = col1.checkbox(
            etiqueta, value=ya_hecho, disabled=solo_lectura,
            key=f"{prefijo}_{clave}_{fila['CLAVE']}",
        )
        if ya_hecho:
            col2.caption(f"{_fmt_fecha(fila.get(cf))} · {fila.get(cu, '') or '—'}")
        if marcado and not ya_hecho:
            cambios[cf] = pd.Timestamp(hoy)
            cols_usuario_nuevas.append(cu)
        elif not marcado and ya_hecho and not solo_lectura:
            cambios[cf] = pd.NaT
            cambios[cu] = ""
    return cambios, cols_usuario_nuevas


# =============================================================================
# 8. PESTAÑAS DE ETAPAS
# =============================================================================

def pestania_etapa1(fila: pd.Series) -> None:
    """Etapa 1 — Reporte de reclamo (7 días desde FECHA CORTE)."""
    clave = fila["CLAVE"]
    terminada = _etapa_bloqueada(fila, "E1")
    limites = calcular_fechas_limite(fila)

    st.markdown(f"#### 1️⃣ {ETAPA_1}")
    st.caption("Pasos: Recepción de Folio → Revisión de folio → Envío a proveedores. "
               "Plazo de 7 días desde la fecha de corte.")
    _mostrar_alarma(limites["E1"], terminada)

    if terminada:
        st.success("Etapa terminada y bloqueada. Continúa en la pestaña **Gestión**.")
        _resumen_pasos(fila, PASOS_E1)
        dias = fila.get(COL_E1_DIAS)
        if str(dias).strip():
            st.metric("Días que tardó la etapa", f"{float(dias):.0f}"
                      if str(dias).replace('.', '').replace('-', '').isdigit() else dias)
        if str(fila.get(COL_E1_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E1_OBS)}")
        _boton_reactivar("E1", clave)
        _panel_ajuste_manual("E1", fila)
        return

    with st.form(f"form_e1_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, PASOS_E1, "e1", solo_lectura=False)
        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E1_OBS, "") or ""),
                           placeholder="Notas sobre la recepción/revisión/envío…")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e1_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

    # Ajuste manual disponible también con la etapa abierta
    _panel_ajuste_manual("E1", fila)

    if not guardar:
        return
    if not usuario.strip():
        st.error("✍️ Firma con tu nombre antes de guardar.")
        return

    df = st.session_state["df"]
    for cu in cols_usuario:
        cambios[cu] = usuario.strip()
    cambios[COL_E1_OBS] = obs.strip()
    if limites["E1"]:
        cambios[COL_E1_LIMITE] = pd.Timestamp(limites["E1"])

    # ¿Se completó el último paso (Envío a proveedores)?
    envio_hecho = _paso_quedo_hecho(PASOS_E1[-1][2], cambios, fila)
    mensaje_log = "Avance en Reporte de reclamo"
    if envio_hecho:
        # Cerrar etapa 1: contar días desde FECHA CORTE y activar etapa 2
        corte = _a_fecha(fila.get(COL_FECHA_CORTE))
        if corte:
            cambios[COL_E1_DIAS] = (hoy_mx() - corte).days
        cambios[COL_E1_TERM] = "SÍ"
        cambios[COL_ETAPA] = ETAPA_2
        mensaje_log = "Etapa 1 (Reporte de reclamo) TERMINADA → pasa a Gestión"

    mensaje = aplicar_guardado(df, clave, cambios, usuario.strip(), mensaje_log)
    persistir_y_sincronizar(df, mensaje, "Avance guardado.", celebrar=envio_hecho)


def pestania_etapa2(fila: pd.Series) -> None:
    """Etapa 2 — Gestión (30 días desde el vencimiento de la etapa 1)."""
    clave = fila["CLAVE"]
    etapa_actual = str(fila.get(COL_ETAPA, "")).strip()
    disponible = etapa_actual in (ETAPA_2, ETAPA_3, ETAPA_4, ETAPA_FINAL) or _etapa_bloqueada(fila, "E1")
    terminada = _etapa_bloqueada(fila, "E2")
    limites = calcular_fechas_limite(fila)

    st.markdown(f"#### 2️⃣ {ETAPA_2}")
    st.caption("Pasos: Enviado a proveedor → Seguimiento → Respuesta de proveedor. "
               "Plazo de 30 días desde el vencimiento de la etapa anterior.")

    if not disponible:
        st.warning("🔒 Esta etapa se activa cuando termines la etapa **Reporte de reclamo**.")
        return

    _mostrar_alarma(limites["E2"], terminada)

    if terminada:
        st.success(f"Etapa terminada. Respuesta del proveedor: "
                   f"**{fila.get(COL_RESPUESTA_TIPO, '—')}**.")
        _resumen_pasos(fila, PASOS_E2)
        if str(fila.get(COL_E2_COMPROMISO, "")).strip():
            st.info(f"📅 Fecha compromiso de recolección: {_fmt_fecha(fila.get(COL_E2_COMPROMISO))}")
        if str(fila.get(COL_E2_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E2_OBS)}")
        _boton_reactivar("E2", clave)
        _panel_ajuste_manual("E2", fila)
        return

    with st.form(f"form_e2_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, PASOS_E2, "e2", solo_lectura=False)

        st.markdown("##### Respuesta del proveedor")
        respuesta = st.radio(
            "¿Cuál fue la respuesta?",
            options=[RESP_RECOLECCION, RESP_DESTRUCCION, RESP_SIN],
            horizontal=True, key=f"resp_{clave}",
        )
        fecha_compromiso = None
        if respuesta == RESP_RECOLECCION:
            fecha_compromiso = st.date_input(
                "Fecha compromiso de recolección", value=hoy_mx(),
                format="DD/MM/YYYY", key=f"comp_{clave}",
            )
            st.info(f"Al cerrar, se activará **{ETAPA_4}** en modalidad recolección, "
                    f"con un contador de {DIAS_RECOLECCION} días para recoger.")
        elif respuesta == RESP_DESTRUCCION:
            st.warning("Destrucción: se reportará a Devoluciones y se informará a CxP de la "
                       "nota de crédito. Al cerrar se activará **Cuentas por pagar**.")
        else:
            st.warning("Sin respuesta: al vencer el plazo se notifica a CxP para negociación. "
                       "Al cerrar se activará **Cuentas por pagar**.")

        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E2_OBS, "") or ""))
        st.caption("Al marcar 'Respuesta de proveedor' se cierra la etapa y se activa la "
                   "siguiente según la respuesta seleccionada.")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e2_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

    # Ajuste manual disponible también con la etapa abierta
    _panel_ajuste_manual("E2", fila)

    if not guardar:
        return
    if not usuario.strip():
        st.error("✍️ Firma con tu nombre antes de guardar.")
        return

    df = st.session_state["df"]
    for cu in cols_usuario:
        cambios[cu] = usuario.strip()
    cambios[COL_E2_OBS] = obs.strip()
    cambios[COL_RESPUESTA_TIPO] = respuesta
    if limites["E2"]:
        cambios[COL_E2_LIMITE] = pd.Timestamp(limites["E2"])
    if respuesta == RESP_RECOLECCION and fecha_compromiso:
        cambios[COL_E2_COMPROMISO] = pd.Timestamp(fecha_compromiso)

    # La etapa cierra cuando el último paso (Respuesta de proveedor) queda marcado.
    cerrar = _paso_quedo_hecho(PASOS_E2[-1][2], cambios, fila)
    mensaje_log = "Avance en Gestión"
    if cerrar:
        # Cerrar etapa 2: SIEMPRE pasa a Disposición final (flujo lineal).
        # La respuesta del proveedor ya no decide la etapa siguiente, solo
        # define los pasos que se mostrarán en Disposición final.
        for cu in [p[3] for p in PASOS_E2]:
            if not str(fila.get(cu, "")).strip() and cu not in cambios:
                cambios[cu] = usuario.strip()
        lim_e2 = limites["E2"]
        if lim_e2:
            cambios[COL_E2_DIAS] = (hoy_mx() - lim_e2).days + DIAS_ETAPA_2
        cambios[COL_E2_TERM] = "SÍ"
        cambios[COL_ETAPA] = ETAPA_4
        if respuesta == RESP_RECOLECCION:
            cambios[COL_E4_MODALIDAD] = RESP_RECOLECCION
            # Contador independiente: 20 días para recoger desde que se define.
            cambios[COL_E4_LIMITE_REC] = pd.Timestamp(hoy_mx() + timedelta(days=DIAS_RECOLECCION))
        else:
            cambios[COL_E4_MODALIDAD] = RESP_DESTRUCCION
        mensaje_log = (f"Etapa 2 TERMINADA · {respuesta} → pasa a {ETAPA_4}")

    mensaje = aplicar_guardado(df, clave, cambios, usuario.strip(), mensaje_log)
    persistir_y_sincronizar(df, mensaje, "Avance guardado.", celebrar=bool(cerrar))


def pestania_etapa3(fila: pd.Series) -> None:
    """Cuentas por pagar — última etapa del flujo, tras Disposición final."""
    clave = fila["CLAVE"]
    etapa_actual = str(fila.get(COL_ETAPA, "")).strip()
    # Se activa al cerrar Disposición final (todas las reclamaciones pasan por aquí).
    disponible = etapa_actual in (ETAPA_3, ETAPA_FINAL) or _etapa_bloqueada(fila, "E4")
    terminada = _etapa_bloqueada(fila, "E3")
    limites = calcular_fechas_limite(fila)

    st.markdown(f"#### 4️⃣ {ETAPA_3}")
    st.caption("Pasos: Seguimiento → Recepción de nota de crédito → Aplicación de pago. "
               "Es la última etapa: al cerrarla se da por terminado todo el proceso "
               f"(plazo total de {DIAS_VENCIMIENTO_TOTAL} días desde la fecha de corte).")

    if not disponible:
        st.warning(f"🔒 Esta etapa se activa cuando cierres **{ETAPA_4}**.")
        return

    _mostrar_alarma(limites["E3"], terminada)

    # Aviso por vencimiento: si venció sin avanzar, alarma de destrucción por vencimiento
    nivel, _ = estado_alarma(limites["E3"], terminada)
    if nivel == "danger" and not terminada:
        st.error("⚠️ Vencido sin respuesta del proveedor: informar al proveedor de la "
                 "**destrucción por vencimiento**. Estatus sugerido: 'Aplicación a factura'. "
                 "Notificar a Devoluciones para proceder con la destrucción.")

    if terminada:
        st.success("🎉 Proceso TERMINADO. Todas las etapas y fechas quedaron registradas.")
        _resumen_pasos(fila, PASOS_E3)
        if str(fila.get(COL_E3_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E3_OBS)}")
        _boton_reactivar("E3", clave)
        _panel_ajuste_manual("E3", fila)
        return

    with st.form(f"form_e3_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, PASOS_E3, "e3", solo_lectura=False)
        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E3_OBS, "") or ""))
        st.caption("Al registrar 'Aplicación de pago' se da por terminado TODO el "
                   "proceso de la reclamación.")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e3_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

    # Ajuste manual disponible también con la etapa abierta
    _panel_ajuste_manual("E3", fila)

    if not guardar:
        return
    if not usuario.strip():
        st.error("✍️ Firma con tu nombre antes de guardar.")
        return

    df = st.session_state["df"]
    for cu in cols_usuario:
        cambios[cu] = usuario.strip()
    cambios[COL_E3_OBS] = obs.strip()
    if limites["E3"]:
        cambios[COL_E3_LIMITE] = pd.Timestamp(limites["E3"])

    # ¿Se completó el último paso (Aplicación de pago)?
    aplicacion_hecha = _paso_quedo_hecho(PASOS_E3[-1][2], cambios, fila)
    mensaje_log = "Avance en Cuentas por pagar"
    if aplicacion_hecha:
        lim_e3 = limites["E3"]
        if lim_e3:
            cambios[COL_E3_DIAS] = (hoy_mx() - lim_e3).days + DIAS_ETAPA_3
        cambios[COL_E3_TERM] = "SÍ"
        # Cuentas por pagar es la ÚLTIMA etapa del flujo: cierra el proceso.
        cambios[COL_ETAPA] = ETAPA_FINAL
        mensaje_log = "Etapa Cuentas por pagar TERMINADA · PROCESO FINALIZADO"

    mensaje = aplicar_guardado(df, clave, cambios, usuario.strip(), mensaje_log)
    persistir_y_sincronizar(df, mensaje, "Avance guardado.", celebrar=aplicacion_hecha)


def pestania_etapa4(fila: pd.Series) -> None:
    """Disposición final — va después de Gestión y antes de Cuentas por pagar."""
    clave = fila["CLAVE"]
    etapa_actual = str(fila.get(COL_ETAPA, "")).strip()
    modalidad = modalidad_destino_final(fila)
    # Se activa al cerrar Gestión; sigue visible en las etapas posteriores.
    disponible = (etapa_actual in (ETAPA_4, ETAPA_3, ETAPA_FINAL)
                  or _etapa_bloqueada(fila, "E2"))
    terminada = _etapa_bloqueada(fila, "E4")

    st.markdown(f"#### 3️⃣ {ETAPA_4}")
    st.caption("Define el destino del producto. Los pasos dependen de la respuesta "
               f"del proveedor. Al cerrarla se activa **{ETAPA_3}**.")

    if not disponible:
        st.warning(f"🔒 Esta etapa se activa al cerrar **{ETAPA_2}**.")
        return

    st.info(f"Modalidad: **{modalidad}**")

    # Corrector de modalidad (por si quedó mal grabada en un registro).
    with st.expander("🔀 Corregir modalidad (requiere clave)"):
        st.caption("La modalidad debe coincidir con la respuesta del proveedor "
                   f"registrada en Gestión: **{fila.get(COL_RESPUESTA_TIPO, '') or '—'}**. "
                   "Usa esto solo si quedó grabada de forma incorrecta.")
        cm1, cm2, cm3 = st.columns([2, 2, 1])
        clave_mod = cm1.text_input("Clave", type="password", key=f"clave_mod_{clave}")
        nueva_mod = cm2.selectbox(
            "Modalidad correcta", options=[RESP_RECOLECCION, RESP_DESTRUCCION],
            index=0 if modalidad == RESP_RECOLECCION else 1,
            key=f"sel_mod_{clave}",
        )
        if cm3.button("Aplicar", key=f"btn_mod_{clave}", use_container_width=True):
            if clave_mod != clave_autorizacion():
                st.error("Clave incorrecta.")
            elif nueva_mod == modalidad:
                st.info("La modalidad ya es esa; no hay cambios.")
            else:
                df = st.session_state["df"]
                cambios = {COL_E4_MODALIDAD: nueva_mod}
                # Si pasa a recolección, fijar el límite de 20 días si no existe.
                if nueva_mod == RESP_RECOLECCION and _a_fecha(fila.get(COL_E4_LIMITE_REC)) is None:
                    cambios[COL_E4_LIMITE_REC] = pd.Timestamp(
                        hoy_mx() + timedelta(days=DIAS_RECOLECCION))
                mensaje = aplicar_guardado(
                    df, clave, cambios, "Corrección",
                    f"Modalidad de {ETAPA_4} corregida a '{nueva_mod}'",
                )
                persistir_y_sincronizar(df, mensaje, "Modalidad corregida.")

    # Alarma del plazo de la etapa (parte de los 90 días globales)
    limites = calcular_fechas_limite(fila)
    _mostrar_alarma(limites["E4"], terminada)

    # ----- RECOLECCIÓN -----
    if modalidad == RESP_RECOLECCION:
        # Contador INDEPENDIENTE: 20 días para que recojan el producto, contados
        # desde que se definió la recolección al cerrar Gestión.
        lim_rec = _a_fecha(fila.get(COL_E4_LIMITE_REC))
        if lim_rec:
            dias_rest = (lim_rec - hoy_mx()).days
            if terminada:
                st.info(f"📦 Recolección · fecha límite: {lim_rec:%d/%m/%Y}")
            elif dias_rest < 0:
                st.error(f"📦 **Recolección vencida** hace {abs(dias_rest)} día(s) "
                         f"(límite {lim_rec:%d/%m/%Y}). Procede a enviarlo a destrucción.")
                if st.button("➡️ Cambiar a destrucción por vencimiento",
                             key=f"to_destr_{clave}"):
                    df = st.session_state["df"]
                    cambios = {COL_E4_MODALIDAD: RESP_DESTRUCCION}
                    mensaje = aplicar_guardado(
                        df, clave, cambios, "Sistema",
                        f"Recolección vencida ({DIAS_RECOLECCION} días) → cambia a destrucción")
                    persistir_y_sincronizar(df, mensaje, "Cambiado a destrucción.")
            elif dias_rest <= 5:
                st.warning(f"📦 Recolección: quedan {dias_rest} día(s) "
                           f"(límite {lim_rec:%d/%m/%Y}).")
            else:
                st.info(f"📦 Recolección: quedan {dias_rest} día(s) de los "
                        f"{DIAS_RECOLECCION} (límite {lim_rec:%d/%m/%Y}).")

        st.caption("Recolección: Programación → Recolección → Recepción de folio "
                   "de devolución.")
        pasos = PASOS_E4_RECOLECCION
        ultimo_etiqueta = "Recepción de folio de devolución"
    # ----- DESTRUCCIÓN -----
    else:
        st.caption("Destrucción: Reporte al almacén de devoluciones → Recepción de "
                   "folio de ajuste.")
        pasos = PASOS_E4_DESTRUCCION
        ultimo_etiqueta = "Recepción de folio de ajuste"

    if terminada:
        st.success(f"Etapa terminada. Continúa en **{ETAPA_3}**.")
        _resumen_pasos(fila, pasos)
        if str(fila.get(COL_E4_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E4_OBS)}")
        _boton_reactivar("E4", clave)
        _panel_ajuste_manual("E4", fila)
        return

    with st.form(f"form_e4_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, pasos, "e4", solo_lectura=False)
        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E4_OBS, "") or ""))
        st.caption(f"Al registrar '{ultimo_etiqueta}' se cierra esta etapa y se "
                   f"activa **{ETAPA_3}**.")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e4_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

    # Ajuste manual disponible también con la etapa abierta
    _panel_ajuste_manual("E4", fila)

    if not guardar:
        return
    if not usuario.strip():
        st.error("✍️ Firma con tu nombre antes de guardar.")
        return

    df = st.session_state["df"]
    for cu in cols_usuario:
        cambios[cu] = usuario.strip()
    cambios[COL_E4_OBS] = obs.strip()

    ultimo_hecho = _paso_quedo_hecho(pasos[-1][2], cambios, fila)
    mensaje_log = f"Avance en {ETAPA_4}"
    if ultimo_hecho:
        cambios[COL_E4_TERM] = "SÍ"
        # Disposición final SIEMPRE pasa a Cuentas por pagar (flujo lineal).
        cambios[COL_ETAPA] = ETAPA_3
        mensaje_log = f"Etapa {ETAPA_4} TERMINADA → pasa a {ETAPA_3}"

    mensaje = aplicar_guardado(df, clave, cambios, usuario.strip(), mensaje_log)
    persistir_y_sincronizar(df, mensaje, "Avance guardado.", celebrar=ultimo_hecho)


# =============================================================================
# 9. VISTA DE EDICIÓN (las 4 etapas en pestañas)
# =============================================================================

def vista_editar(df: pd.DataFrame) -> None:
    st.subheader("✏️ Editar reclamación")
    if "flash" in st.session_state:
        tipo, texto = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(texto)
    # Globos SOLO al terminar una etapa (bandera de un solo uso).
    if st.session_state.pop("celebrar", False):
        st.balloons()

    if df.empty:
        st.info("No hay registros con los filtros actuales.")
        return

    df_sel = df.copy()
    df_sel["ETIQUETA"] = ("Folio " + df_sel[COL_FOLIO].astype(str)
                          + " · " + df_sel[COL_PROVEEDOR].str.slice(0, 40)
                          + " · [" + df_sel[COL_ETAPA].astype(str) + "]")
    etiqueta = st.selectbox("Selecciona la reclamación (por Folio)",
                            options=df_sel["ETIQUETA"].tolist())
    fila = df_sel[df_sel["ETIQUETA"] == etiqueta].iloc[0]

    # Ficha compacta del registro (una sola línea, no ocupa espacio)
    imp = pd.to_numeric(pd.Series([fila.get(COL_IMPORTE)]), errors="coerce").iloc[0]
    imp_txt = f"${imp:,.2f}" if pd.notna(imp) else "—"
    lim90 = fecha_vencimiento_90(fila)
    st.markdown(
        f"""<div style='display:flex;gap:1.2rem;flex-wrap:wrap;
                        padding:0.4rem 0.7rem;margin:0.2rem 0 0.5rem 0;
                        border:1px solid #e6e6e6;border-radius:8px;font-size:0.82rem'>
          <span><b>Folio:</b> {fila[COL_FOLIO]}</span>
          <span><b>Proveedor:</b> {fila[COL_PROVEEDOR]}</span>
          <span><b>Importe:</b> {imp_txt}</span>
          <span><b>Comprador:</b> {fila[COL_COMPRADOR] or '—'}</span>
          <span><b>Corte:</b> {_fmt_fecha(fila.get(COL_FECHA_CORTE))}</span>
          <span><b>Vence 90d:</b> {lim90:%d/%m/%Y}</span>
          <span><b>Etapa:</b> <code>{fila.get(COL_ETAPA, '—')}</code></span>
          <span><b>Respuesta:</b> {fila.get(COL_RESPUESTA_TIPO, '') or '—'}</span>
          <span><b>Carta:</b> {fila.get(COL_CARTA_FIRMADA, '—')}</span>
        </div>""" if lim90 else
        f"""<div style='display:flex;gap:1.2rem;flex-wrap:wrap;
                        padding:0.4rem 0.7rem;margin:0.2rem 0 0.5rem 0;
                        border:1px solid #e6e6e6;border-radius:8px;font-size:0.82rem'>
          <span><b>Folio:</b> {fila[COL_FOLIO]}</span>
          <span><b>Proveedor:</b> {fila[COL_PROVEEDOR]}</span>
          <span><b>Importe:</b> {imp_txt}</span>
          <span><b>Etapa:</b> <code>{fila.get(COL_ETAPA, '—')}</code></span>
        </div>""",
        unsafe_allow_html=True,
    )

    # Alarma crítica de 90 días para ESTE registro
    if vencido_90_sin_definicion(fila):
        dias_v = (hoy_mx() - lim90).days
        st.error(f"🚨 **{MSG_VENCIDO_90}** — Venció hace {dias_v} día(s) "
                 f"(límite: {lim90:%d/%m/%Y}).")

    # Orden del flujo: Reporte → Gestión → Disposición final → Cuentas por pagar
    t1, t2, t3, t4 = st.tabs([
        f"1️⃣ {ETAPA_1}", f"2️⃣ {ETAPA_2}", f"3️⃣ {ETAPA_4}", f"4️⃣ {ETAPA_3}",
    ])
    with t1:
        pestania_etapa1(fila)
    with t2:
        pestania_etapa2(fila)
    with t3:
        pestania_etapa4(fila)
    with t4:
        pestania_etapa3(fila)


# =============================================================================
# 10. TABLA, FILTROS, KPIs Y NOTIFICACIONES
# =============================================================================

def _aplicar_filtros(df: pd.DataFrame, seleccion: dict, excluir: str = None) -> pd.DataFrame:
    """Aplica todos los filtros de `seleccion`, opcionalmente omitiendo uno.

    Omitir un filtro permite calcular sus opciones disponibles según el resto
    de selecciones (cascada bidireccional): así cada lista muestra solo valores
    que existen en combinación con lo ya elegido en los demás filtros.
    """
    d = df
    if excluir != "folio" and seleccion.get("folio"):
        d = d[d[COL_FOLIO].str.contains(seleccion["folio"], case=False, na=False)]
    if excluir != "mes" and seleccion.get("mes"):
        d = d[d[COL_MES_ETIQUETA].isin(seleccion["mes"])]
    if excluir != "proveedor" and seleccion.get("proveedor"):
        d = d[d[COL_PROVEEDOR].isin(seleccion["proveedor"])]
    if excluir != "comprador" and seleccion.get("comprador"):
        d = d[d[COL_COMPRADOR].isin(seleccion["comprador"])]
    if excluir != "etapa" and seleccion.get("etapa"):
        d = d[d[COL_ETAPA].isin(seleccion["etapa"])]
    if excluir != "criticos" and seleccion.get("criticos") and not d.empty:
        d = d[d.apply(vencido_90_sin_definicion, axis=1)]
    return d


def construir_filtros(df: pd.DataFrame) -> pd.DataFrame:
    """Filtros en cascada bidireccional con selección múltiple.

    Cada filtro muestra únicamente las opciones que siguen siendo posibles según
    lo elegido en TODOS los demás filtros, en cualquier orden.

    Las claves de los widgets llevan un sufijo de versión (`_v{n}`). Para limpiar
    los filtros se incrementa esa versión: Streamlit ve widgets nuevos y los crea
    vacíos. Esto es necesario porque borrar la clave del session_state no basta
    —el widget se vuelve a dibujar con su valor anterior— y reasignarla lanzaría
    StreamlitAPIException.
    """
    st.sidebar.markdown("### 🔎 Filtros")

    v = st.session_state.get("filtros_version", 0)
    k_folio, k_mes = f"f_folio_v{v}", f"f_mes_v{v}"
    k_prov, k_comp = f"f_proveedor_v{v}", f"f_comprador_v{v}"
    k_etapa, k_crit = f"f_etapa_v{v}", f"f_criticos_v{v}"

    # Estado previo de las selecciones (lo que el usuario ya eligió)
    sel = {
        "folio": st.session_state.get(k_folio, "").strip(),
        "mes": st.session_state.get(k_mes, []),
        "proveedor": st.session_state.get(k_prov, []),
        "comprador": st.session_state.get(k_comp, []),
        "etapa": st.session_state.get(k_etapa, []),
        "criticos": st.session_state.get(k_crit, False),
    }

    # --- Búsqueda por folio ---
    st.sidebar.text_input("Buscar folio", placeholder="Ej. DC-MZ017", key=k_folio)

    # --- Mes: opciones según los demás filtros ---
    base_mes = _aplicar_filtros(df, sel, excluir="mes")
    opciones_mes = (base_mes[[COL_MES_ETIQUETA, COL_MES]].drop_duplicates()
                    .sort_values(COL_MES)[COL_MES_ETIQUETA].tolist())
    # Conservar valores ya elegidos aunque el resto los excluya (evita perder la selección)
    opciones_mes += [m for m in sel["mes"] if m not in opciones_mes]
    st.sidebar.multiselect("Mes", options=opciones_mes, placeholder="Todos", key=k_mes)

    # --- Proveedor ---
    base_prov = _aplicar_filtros(df, sel, excluir="proveedor")
    opciones_prov = sorted(x for x in base_prov[COL_PROVEEDOR].unique() if x)
    opciones_prov += [p for p in sel["proveedor"] if p not in opciones_prov]
    st.sidebar.multiselect("Proveedor", options=opciones_prov, placeholder="Todos",
                           key=k_prov)

    # --- Comprador ---
    base_comp = _aplicar_filtros(df, sel, excluir="comprador")
    opciones_comp = sorted(x for x in base_comp[COL_COMPRADOR].unique() if x)
    opciones_comp += [c for c in sel["comprador"] if c not in opciones_comp]
    st.sidebar.multiselect("Comprador", options=opciones_comp, placeholder="Todos",
                           key=k_comp)

    # --- Etapa ---
    base_etapa = _aplicar_filtros(df, sel, excluir="etapa")
    presentes = set(base_etapa[COL_ETAPA])
    opciones_etapa = [e for e in (ETAPA_1, ETAPA_2, ETAPA_3, ETAPA_4, ETAPA_FINAL)
                      if e in presentes]
    opciones_etapa += [e for e in sel["etapa"] if e not in opciones_etapa]
    st.sidebar.multiselect("Etapa", options=opciones_etapa, placeholder="Todas",
                           key=k_etapa)

    # --- Solo vencidos a 90 días ---
    st.sidebar.checkbox(f"🚨 Solo vencidos ({DIAS_VENCIMIENTO_TOTAL} días)",
                        help=MSG_VENCIDO_90, key=k_crit)

    # --- Resultado con TODAS las selecciones aplicadas ---
    seleccion_actual = {
        "folio": st.session_state.get(k_folio, "").strip(),
        "mes": st.session_state.get(k_mes, []),
        "proveedor": st.session_state.get(k_prov, []),
        "comprador": st.session_state.get(k_comp, []),
        "etapa": st.session_state.get(k_etapa, []),
        "criticos": st.session_state.get(k_crit, False),
    }
    dff = _aplicar_filtros(df, seleccion_actual)

    activos = sum([
        bool(seleccion_actual["folio"]), bool(seleccion_actual["mes"]),
        bool(seleccion_actual["proveedor"]), bool(seleccion_actual["comprador"]),
        bool(seleccion_actual["etapa"]), bool(seleccion_actual["criticos"]),
    ])

    # --- Limpiar (solo se ofrece si hay algo que limpiar) ---
    if st.sidebar.button("🧹 Limpiar filtros", use_container_width=True,
                         disabled=activos == 0):
        # Borrar los valores actuales y estrenar versión de claves.
        for k in (k_folio, k_mes, k_prov, k_comp, k_etapa, k_crit):
            st.session_state.pop(k, None)
        st.session_state["filtros_version"] = v + 1
        st.rerun()

    resumen = f"**{len(dff)}** de **{len(df)}** registros"
    if activos:
        resumen += f" · {activos} filtro(s) activo(s)"
    st.sidebar.caption(resumen)
    return dff


def mostrar_kpis(df: pd.DataFrame) -> None:
    """KPIs compactos en una sola línea."""
    total = len(df)
    importe = pd.to_numeric(df.get(COL_IMPORTE), errors="coerce").fillna(0).sum()
    finalizados = int((df[COL_ETAPA] == ETAPA_FINAL).sum())
    en_proceso = total - finalizados
    vencidos_90 = int(df.apply(vencido_90_sin_definicion, axis=1).sum()) if total else 0

    st.markdown(
        f"""<div style='display:flex;gap:1.6rem;flex-wrap:wrap;
                        padding:0.35rem 0.7rem;margin-bottom:0.4rem;
                        border:1px solid #e6e6e6;border-radius:8px;
                        font-size:0.85rem;align-items:center'>
          <span>📦 <b>{total:,}</b> reclamaciones</span>
          <span>💰 <b>${importe:,.2f}</b></span>
          <span>⏳ <b>{en_proceso}</b> en proceso</span>
          <span>✅ <b>{finalizados}</b> finalizadas</span>
          <span style='color:#c0392b'>🚨 <b>{vencidos_90}</b> vencidas
            ({DIAS_VENCIMIENTO_TOTAL} días)</span>
        </div>""",
        unsafe_allow_html=True,
    )


def mostrar_notificaciones(df: pd.DataFrame) -> None:
    avisos = construir_notificaciones(df)
    if not avisos:
        st.success("✅ No hay reclamaciones por vencerse o vencidas con los filtros actuales.")
        return

    criticos = [a for a in avisos if a["nivel"] == "critico"]
    vencidas = [a for a in avisos if a["nivel"] == "danger"]
    por_vencer = [a for a in avisos if a["nivel"] == "warn"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"🚨 Vencidas {DIAS_VENCIMIENTO_TOTAL} días", len(criticos))
    c2.metric("🔴 Etapa vencida", len(vencidas))
    c3.metric("🟡 Por vencerse", len(por_vencer))

    # --- Sección crítica: reclamos vencidos a 90 días ---
    if criticos:
        st.error(f"🚨 **{MSG_VENCIDO_90}**")
        for a in criticos:
            st.markdown(
                f"&nbsp;&nbsp;🚨 **Folio {a['folio']}** · {a['proveedor']} · "
                f"_{a['etapa']}_ — {a['mensaje']}"
            )
        st.divider()

    # --- Etapas vencidas y por vencerse ---
    for a in vencidas + por_vencer:
        icono = "🔴" if a["nivel"] == "danger" else "🟡"
        st.markdown(f"{icono} **Folio {a['folio']}** · {a['proveedor']} · "
                    f"_{a['etapa']}_ — {a['mensaje']} (Comprador: {a['comprador']})")


def mostrar_tabla(df: pd.DataFrame) -> None:
    vista = df.copy()
    # Columna de alerta: reclamos vencidos a 90 días sin definición
    COL_ALERTA = "⚠️ ALERTA"
    if not vista.empty:
        vista[COL_ALERTA] = vista.apply(
            lambda f: MSG_VENCIDO_90 if vencido_90_sin_definicion(f) else "", axis=1
        )
        vista["VENCE 90D"] = vista.apply(fecha_vencimiento_90, axis=1)
    else:
        vista[COL_ALERTA] = ""
        vista["VENCE 90D"] = pd.NaT

    columnas = [COL_FOLIO, COL_PROVEEDOR, COL_COMPRADOR, COL_IMPORTE, COL_FECHA_CORTE,
                "VENCE 90D", COL_ALERTA,
                COL_ETAPA, COL_RESPUESTA_TIPO, COL_E1_TERM, COL_E2_TERM, COL_E3_TERM,
                COL_E4_TERM, COL_E1_DIAS, COL_E2_DIAS, COL_E3_DIAS,
                COL_MODIFICADO_POR, COL_FECHA_MODIFICACION]
    columnas = [c for c in columnas if c in vista.columns]
    st.dataframe(
        vista[columnas], use_container_width=True, hide_index=True,
        column_config={
            # Importes en formato moneda; días como número entero.
            COL_IMPORTE: st.column_config.NumberColumn("Importe (MXN)", format="$%.2f"),
            COL_FECHA_CORTE: st.column_config.DateColumn(format="DD/MM/YYYY"),
            "VENCE 90D": st.column_config.DateColumn(format="DD/MM/YYYY"),
            COL_E1_DIAS: st.column_config.NumberColumn("Días E1", format="%d"),
            COL_E2_DIAS: st.column_config.NumberColumn("Días E2", format="%d"),
            COL_E3_DIAS: st.column_config.NumberColumn("Días E3", format="%d"),
        },
    )


# =============================================================================
# 11. PESTAÑA DE GUÍA DEL PROCESO
# =============================================================================

def vista_guia() -> None:
    st.subheader("📖 Guía del proceso")
    st.markdown(
        f"""
Esta aplicación da seguimiento a las reclamaciones a proveedores a lo largo de
**cuatro etapas encadenadas**. Cada reclamación se identifica por su **Folio de
reporte**. Al terminar la última fase de una etapa, esa etapa se **bloquea** y se
**activa** la siguiente automáticamente.

### 1️⃣ {ETAPA_1} · plazo {DIAS_ETAPA_1} días
Pasos: **Recepción de Folio → Revisión de folio → Envío a proveedores**.
El plazo se cuenta desde la **fecha de corte**. Al registrar *Envío a proveedores*
se cierra la etapa, se cuentan los días que tardó y se activa **Gestión**.

### 2️⃣ {ETAPA_2} · plazo {DIAS_ETAPA_2} días
Pasos: **Enviado a proveedor → Seguimiento → Respuesta de proveedor**.
Aquí se captura la **respuesta del proveedor** (recolección, destrucción o sin
respuesta). Al cerrar la etapa **siempre** se pasa a **{ETAPA_4}**; la respuesta
solo define qué pasos se mostrarán ahí.

### 3️⃣ {ETAPA_4}
Define el destino del producto, según la respuesta capturada en Gestión:
- **Recolección**: **Programación → Recolección → Recepción de folio de devolución**.
  Lleva un contador **independiente de {DIAS_RECOLECCION} días** para que el
  proveedor recoja, contados desde que se definió la recolección al cerrar
  Gestión. Si vence, se puede cambiar a destrucción con un botón.
- **Destrucción**: **Reporte al almacén de devoluciones → Recepción de folio de ajuste**.

Al registrar el último paso se cierra la etapa y **siempre** se activa **{ETAPA_3}**.

### 4️⃣ {ETAPA_3} · plazo {DIAS_ETAPA_3} días
Pasos: **Seguimiento → Recepción de nota de crédito → Aplicación de pago**.
Es la **última etapa**: al registrar *Aplicación de pago* se da por terminado
todo el proceso. Si vence sin respuesta del proveedor, se lanza la alarma de
destrucción por vencimiento y se notifica a Devoluciones.

> **Plazos:** 7 + 30 + 53 = **{DIAS_VENCIMIENTO_TOTAL} días** en total desde la
> fecha de corte. El contador de recolección corre por separado.

### 🔓 Reactivar etapas
Cada etapa cerrada tiene un botón **Reactivar** protegido con clave. Solo el
personal autorizado puede reabrir una etapa para corregir información.

### 🗓️ Ajuste manual de fechas (retroactivo)
En cada etapa hay un panel **Ajuste manual de fechas**, también protegido con la
misma clave. Sirve para los registros que se **completaron antes de usar la app**:
permite capturar la fecha real y el responsable de cada paso, y cerrar la etapa
con esa fecha. Los días transcurridos se calculan con la fecha real capturada,
no con la de hoy, para que las métricas del proceso sean correctas.

### 📈 Avance por comprador
La primera pestaña muestra tres gráficas: **distribución por etapa** (en qué fase
está el trabajo de cada comprador), **porcentaje finalizado** (qué proporción ya
cerró todo el proceso) y **estado de vencimiento** (vencidas, por vencerse o en
tiempo). Tiene su propio selector de mes y un interruptor para medir por
**cantidad de reclamaciones** o por **importe en pesos**. Respeta además los
filtros generales de la barra lateral.

### 🚨 Alarma de vencimiento total ({DIAS_VENCIMIENTO_TOTAL} días)
Si pasan **{DIAS_VENCIMIENTO_TOTAL} días desde la fecha de corte** y el reclamo
todavía no está finalizado, se dispara la alarma:
**{MSG_VENCIDO_90}**. Aparece en el encabezado, en la ficha del registro, en la
tabla y en la pestaña *Resumen y alarmas*. En la barra lateral hay un filtro para
ver solo estos reclamos.

### 🔔 Alarmas por etapa
En la pestaña **Resumen** se listan las reclamaciones *por vencerse* (≤ 15 días) y
*vencidas*. La estructura de avisos está preparada para enviarse por correo en el
futuro.

### ✍️ Registro
Cada paso guarda **fecha y usuario**; las observaciones quedan en la base de datos
y la bitácora completa se acumula en *Acciones / Notas* para su posterior análisis.
        """
    )


# =============================================================================
# 11a. GRÁFICAS DE AVANCE POR COMPRADOR
# =============================================================================

# Orden real del flujo: Reporte → Gestión → Disposición final → Cuentas por pagar
ORDEN_ETAPAS = [ETAPA_1, ETAPA_2, ETAPA_4, ETAPA_3, ETAPA_FINAL]


def _clasificar_urgencia(fila: pd.Series) -> str:
    """Clasifica una reclamación para la gráfica de vencimientos."""
    etapa = str(fila.get(COL_ETAPA, "")).strip()
    if etapa == ETAPA_FINAL:
        return "Finalizada"
    if vencido_90_sin_definicion(fila):
        return f"Vencida {DIAS_VENCIMIENTO_TOTAL} días"

    limites = calcular_fechas_limite(fila)
    if etapa == ETAPA_1:
        lim, term = limites["E1"], es_verdadero(fila.get(COL_E1_TERM))
    elif etapa == ETAPA_2:
        lim, term = limites["E2"], es_verdadero(fila.get(COL_E2_TERM))
    elif etapa == ETAPA_3:
        lim, term = limites["E3"], es_verdadero(fila.get(COL_E3_TERM))
    elif etapa == ETAPA_4:
        lim, term = _a_fecha(fila.get(COL_E4_LIMITE_REC)), es_verdadero(fila.get(COL_E4_TERM))
    else:
        return "En tiempo"

    nivel, _ = estado_alarma(lim, term)
    if nivel == "danger":
        return "Etapa vencida"
    if nivel == "warn":
        return "Por vencerse"
    return "En tiempo"


def _grafica_barras(tabla: pd.DataFrame, etiqueta_valor: str, es_dinero: bool,
                    orden_series: list = None, altura: int = 340) -> None:
    """Dibuja barras apiladas con formato numérico legible.

    Los números llevan separador de miles y, si son importes, símbolo de pesos
    con dos decimales. El formato se aplica tanto al eje como a la etiqueta
    emergente que aparece al pasar el cursor.
    """
    if tabla.empty:
        st.info("Sin datos para graficar.")
        return

    if not ALTAIR_OK:
        # Respaldo: gráfica nativa (sin formato personalizado).
        st.bar_chart(tabla, height=altura, stack=True, horizontal=True,
                     x_label=etiqueta_valor, y_label="Comprador")
        return

    # Formatos estilo D3: '$,.2f' → $1,234.56 · ',d' → 1,234
    fmt = "$,.2f" if es_dinero else ",d"
    nombre_valor = "Importe" if es_dinero else "Cantidad"

    largo = (tabla.reset_index()
             .melt(id_vars=tabla.index.name or "index",
                   var_name="Serie", value_name="Valor"))
    col_x = tabla.index.name or "index"
    largo = largo.rename(columns={col_x: "Comprador"})
    largo = largo[largo["Valor"] != 0]

    orden = orden_series or list(tabla.columns)
    # Barras HORIZONTALES: el valor va en el eje X y el comprador en el Y.
    grafica = (
        alt.Chart(largo)
        .mark_bar()
        .encode(
            y=alt.Y("Comprador:N", title=None,
                    sort=list(tabla.index)),
            x=alt.X("Valor:Q", title=etiqueta_valor,
                    axis=alt.Axis(format=fmt)),
            color=alt.Color("Serie:N", title="", sort=orden,
                            legend=alt.Legend(orient="bottom", columns=3)),
            order=alt.Order("Serie:N"),
            tooltip=[
                alt.Tooltip("Comprador:N", title="Comprador"),
                alt.Tooltip("Serie:N", title="Etapa"),
                alt.Tooltip("Valor:Q", title=nombre_valor, format=fmt),
            ],
        )
        .properties(height=altura)
    )
    st.altair_chart(grafica, use_container_width=True)


def _grafica_porcentaje(resumen: pd.DataFrame, altura: int = 320) -> None:
    """Barras del porcentaje finalizado por comprador, con formato '00.0 %'."""
    if resumen.empty:
        st.info("Sin datos para graficar.")
        return
    if not ALTAIR_OK:
        st.bar_chart(resumen[["% finalizado"]], height=altura, horizontal=True,
                     x_label="% finalizado", y_label="Comprador")
        return

    datos = resumen.reset_index().rename(columns={"_comprador": "Comprador"})
    # Barras HORIZONTALES
    grafica = (
        alt.Chart(datos)
        .mark_bar()
        .encode(
            y=alt.Y("Comprador:N", title=None, sort=list(resumen.index)),
            x=alt.X("% finalizado:Q", title="% finalizado",
                    scale=alt.Scale(domain=[0, 100]),
                    axis=alt.Axis(format=".0f")),
            tooltip=[
                alt.Tooltip("Comprador:N", title="Comprador"),
                alt.Tooltip("% finalizado:Q", title="Finalizado", format=".1f"),
            ],
        )
        .properties(height=altura)
    )
    st.altair_chart(grafica, use_container_width=True)


def _dias_por_etapa(fila: pd.Series) -> dict:
    """Días que tardó cada etapa en una reclamación.

    E1, E2 y E3 tienen su columna de días registrada. Disposición final (E4) no
    la tiene, así que se calcula desde el cierre de Gestión hasta la fecha del
    último paso realizado en esa etapa.
    Devuelve solo las etapas terminadas con un valor válido.
    """
    dias = {}

    def _num(valor):
        try:
            v = float(valor)
            return v if v == v else None  # descarta NaN
        except (TypeError, ValueError):
            return None

    if es_verdadero(fila.get(COL_E1_TERM)):
        v = _num(fila.get(COL_E1_DIAS))
        if v is not None:
            dias[ETAPA_1] = v
    if es_verdadero(fila.get(COL_E2_TERM)):
        v = _num(fila.get(COL_E2_DIAS))
        if v is not None:
            dias[ETAPA_2] = v
    if es_verdadero(fila.get(COL_E3_TERM)):
        v = _num(fila.get(COL_E3_DIAS))
        if v is not None:
            dias[ETAPA_3] = v

    # Disposición final: desde la respuesta del proveedor (cierre de Gestión)
    # hasta el último paso registrado en la etapa.
    if es_verdadero(fila.get(COL_E4_TERM)):
        inicio = _a_fecha(fila.get("E2 FECHA RESPUESTA"))
        pasos_e4 = (PASOS_E4_RECOLECCION
                    if modalidad_destino_final(fila) == RESP_RECOLECCION
                    else PASOS_E4_DESTRUCCION)
        fechas = [_a_fecha(fila.get(cf)) for _c, _e, cf, _cu in pasos_e4]
        fechas = [f for f in fechas if f]
        if inicio and fechas:
            d = (max(fechas) - inicio).days
            if d >= 0:
                dias[ETAPA_4] = float(d)
    return dias


def _grafica_dias_promedio(df: pd.DataFrame, altura: int = 360) -> None:
    """Días promedio por etapa y mes (barras horizontales agrupadas)."""
    filas = []
    for _, fila in df.iterrows():
        mes = fila.get(COL_MES_ETIQUETA, "Sin fecha")
        orden_mes = _a_fecha(fila.get(COL_MES))
        for etapa, dias in _dias_por_etapa(fila).items():
            filas.append({"Mes": mes, "_orden": orden_mes or date.min,
                          "Etapa": etapa, "Días": dias})

    if not filas:
        st.info("Todavía no hay etapas terminadas con días registrados. "
                "Este dato aparece conforme se vayan cerrando etapas.")
        return

    detalle = pd.DataFrame(filas)
    resumen = (detalle.groupby(["Mes", "_orden", "Etapa"], as_index=False)
               .agg(Promedio=("Días", "mean"), Casos=("Días", "size")))
    resumen["Promedio"] = resumen["Promedio"].round(1)
    resumen = resumen.sort_values("_orden")
    orden_meses = resumen.drop_duplicates("Mes")["Mes"].tolist()
    orden_etapas = [e for e in (ETAPA_1, ETAPA_2, ETAPA_4, ETAPA_3)
                    if e in set(resumen["Etapa"])]

    if not ALTAIR_OK:
        st.dataframe(
            resumen.pivot(index="Mes", columns="Etapa", values="Promedio"),
            use_container_width=True,
        )
        return

    grafica = (
        alt.Chart(resumen)
        .mark_bar()
        .encode(
            y=alt.Y("Etapa:N", title=None, sort=orden_etapas),
            x=alt.X("Promedio:Q", title="Días promedio",
                    axis=alt.Axis(format=".1f")),
            color=alt.Color("Etapa:N", title="", sort=orden_etapas,
                            legend=alt.Legend(orient="bottom", columns=2)),
            row=alt.Row("Mes:N", title=None, sort=orden_meses,
                        header=alt.Header(labelAngle=0, labelAlign="left")),
            tooltip=[
                alt.Tooltip("Mes:N", title="Mes"),
                alt.Tooltip("Etapa:N", title="Etapa"),
                alt.Tooltip("Promedio:Q", title="Días promedio", format=".1f"),
                alt.Tooltip("Casos:Q", title="Reclamaciones", format=",d"),
            ],
        )
        .properties(height=max(60, altura // max(1, len(orden_meses))))
    )
    st.altair_chart(grafica, use_container_width=True)

    with st.expander("📋 Ver detalle de días promedio"):
        tabla = resumen.pivot(index="Mes", columns="Etapa", values="Promedio")
        tabla = tabla.reindex(orden_meses)
        cols = [c for c in orden_etapas if c in tabla.columns]
        st.dataframe(
            tabla[cols], use_container_width=True,
            column_config={c: st.column_config.NumberColumn(format="%.1f")
                           for c in cols},
        )


def vista_graficas(df: pd.DataFrame) -> None:
    """Gráficas de avance por comprador, con filtro de mes y medida propios."""
    st.subheader("📈 Avance por comprador")

    if df.empty:
        st.info("No hay registros con los filtros actuales.")
        return

    # --- Controles propios de esta vista ---
    c1, c2 = st.columns([2, 2])
    meses_disp = (df[[COL_MES_ETIQUETA, COL_MES]].drop_duplicates()
                  .sort_values(COL_MES)[COL_MES_ETIQUETA].tolist())
    sel_meses = c1.multiselect(
        "Mes de reclamación", options=meses_disp,
        placeholder="Todos los meses", key="graf_meses",
    )
    medida = c2.radio(
        "Medir por", options=["Cantidad", "Importe (MXN)"],
        horizontal=True, key="graf_medida",
    )

    dfg = df[df[COL_MES_ETIQUETA].isin(sel_meses)] if sel_meses else df.copy()
    if dfg.empty:
        st.info("No hay registros para los meses seleccionados.")
        return

    # Columna de valor y formato según la medida elegida.
    # Cantidad → número entero con separador de miles (sin decimales).
    # Importe  → moneda en pesos con dos decimales.
    if medida == "Cantidad":
        dfg = dfg.assign(_valor=1.0)
        es_dinero = False
        formato = "{:,.0f}"
        etiqueta_valor = "Reclamaciones"
        # Formato para las columnas numéricas de las tablas
        fmt_columna = st.column_config.NumberColumn(format="%d")
    else:
        dfg = dfg.assign(
            _valor=pd.to_numeric(dfg[COL_IMPORTE], errors="coerce").fillna(0.0)
        )
        es_dinero = True
        formato = "${:,.2f}"
        etiqueta_valor = "Importe (MXN)"
        fmt_columna = st.column_config.NumberColumn(format="$%.2f")

    dfg["_comprador"] = dfg[COL_COMPRADOR].replace("", "SIN COMPRADOR")

    # =====================================================================
    # 1) Reclamaciones por etapa (barras apiladas)
    # =====================================================================
    st.markdown(f"##### 1. Distribución por etapa · {etiqueta_valor}")
    st.caption("Muestra en qué fase del proceso está el trabajo de cada comprador.")

    tabla_etapas = (dfg.pivot_table(index="_comprador", columns=COL_ETAPA,
                                    values="_valor", aggfunc="sum", fill_value=0))
    # Ordenar columnas por el flujo del proceso
    columnas_orden = [e for e in ORDEN_ETAPAS if e in tabla_etapas.columns]
    columnas_orden += [c for c in tabla_etapas.columns if c not in columnas_orden]
    tabla_etapas = tabla_etapas[columnas_orden]
    tabla_etapas = tabla_etapas.loc[tabla_etapas.sum(axis=1).sort_values(ascending=False).index]
    # Las cantidades son conteos: se muestran como enteros.
    if not es_dinero:
        tabla_etapas = tabla_etapas.round(0).astype(int)

    _grafica_barras(tabla_etapas, etiqueta_valor, es_dinero,
                    orden_series=columnas_orden, altura=340)

    # =====================================================================
    # 2) Porcentaje finalizado por comprador
    # =====================================================================
    st.markdown("##### 2. Porcentaje finalizado")
    st.caption("Qué proporción del trabajo de cada comprador ya cerró todo el proceso.")

    resumen = dfg.groupby("_comprador").agg(
        total=("_valor", "sum"),
        finalizado=("_valor", lambda s: s[dfg.loc[s.index, COL_ETAPA] == ETAPA_FINAL].sum()),
    )
    resumen["% finalizado"] = (resumen["finalizado"] / resumen["total"] * 100).round(1)
    resumen = resumen.sort_values("% finalizado", ascending=False)

    cg1, cg2 = st.columns([3, 2])
    with cg1:
        _grafica_porcentaje(resumen, altura=320)
    with cg2:
        tabla_pct = resumen.copy()
        if not es_dinero:
            tabla_pct["total"] = tabla_pct["total"].round(0).astype(int)
            tabla_pct["finalizado"] = tabla_pct["finalizado"].round(0).astype(int)
        tabla_pct = tabla_pct.rename(columns={
            "total": etiqueta_valor, "finalizado": "Finalizado"})
        st.dataframe(
            tabla_pct, use_container_width=True, height=320,
            column_config={
                etiqueta_valor: fmt_columna,
                "Finalizado": fmt_columna,
                "% finalizado": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )

    # =====================================================================
    # 3) Estado de vencimiento por comprador
    # =====================================================================
    st.markdown("##### 3. Estado de vencimiento")
    st.caption("Reclamaciones vencidas, por vencerse o en tiempo, según su etapa activa. "
               f"La categoría crítica es la de {DIAS_VENCIMIENTO_TOTAL} días sin definición.")

    dfg["_urgencia"] = dfg.apply(_clasificar_urgencia, axis=1)
    orden_urgencia = [f"Vencida {DIAS_VENCIMIENTO_TOTAL} días", "Etapa vencida",
                      "Por vencerse", "En tiempo", "Finalizada"]
    tabla_urg = (dfg.pivot_table(index="_comprador", columns="_urgencia",
                                 values="_valor", aggfunc="sum", fill_value=0))
    cols_urg = [c for c in orden_urgencia if c in tabla_urg.columns]
    cols_urg += [c for c in tabla_urg.columns if c not in cols_urg]
    tabla_urg = tabla_urg[cols_urg]
    # Ordenar por lo más urgente primero
    criticas = [c for c in cols_urg if "Vencida" in c or "vencida" in c]
    if criticas:
        tabla_urg = tabla_urg.loc[
            tabla_urg[criticas].sum(axis=1).sort_values(ascending=False).index]
    if not es_dinero:
        tabla_urg = tabla_urg.round(0).astype(int)

    _grafica_barras(tabla_urg, etiqueta_valor, es_dinero,
                    orden_series=cols_urg, altura=340)

    # --- Resumen numérico general ---
    with st.expander("📋 Ver detalle numérico por comprador"):
        detalle = tabla_etapas.copy()
        detalle["TOTAL"] = detalle.sum(axis=1)
        if es_dinero:
            detalle = detalle.round(2)
        st.dataframe(
            detalle, use_container_width=True,
            column_config={c: fmt_columna for c in detalle.columns},
        )

    # =====================================================================
    # 4) Días promedio por etapa y mes
    # =====================================================================
    st.markdown("##### 4. Días promedio por etapa")
    st.caption("Cuánto tardó en promedio cada etapa, agrupado por mes de "
               "reclamación. Solo considera etapas ya terminadas.")
    _grafica_dias_promedio(dfg, altura=360)


# =============================================================================
# 11b. GESTIÓN DE LA BASE DE DATOS (descargar / cargar)
# =============================================================================

def _excel_en_memoria(df: pd.DataFrame) -> bytes:
    """Genera el Excel completo (sin columnas auxiliares, fechas sin hora)."""
    import io
    buffer = io.BytesIO()
    df_export = _preparar_para_excel(df)
    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="DD/MM/YYYY",
                        date_format="DD/MM/YYYY") as writer:
        df_export.to_excel(writer, sheet_name=NOMBRE_HOJA, index=False)
    return buffer.getvalue()


def vista_base_datos(df: pd.DataFrame) -> None:
    st.subheader("🗄️ Base de datos")
    if "flash_bd" in st.session_state:
        tipo, texto = st.session_state.pop("flash_bd")
        (st.success if tipo == "success" else st.warning)(texto)

    # ----- DESCARGAR -----
    st.markdown("##### ⬇️ Descargar base de datos completa")
    st.caption("Descarga el Excel con todos los registros y columnas tal como están "
               "hoy. Puedes abrirlo, agregar reclamaciones nuevas o completar datos, y "
               "volver a cargarlo aquí.")
    hoy = ahora_mx().strftime("%Y%m%d_%H%M")
    st.download_button(
        "⬇️ Descargar datos.xlsx actualizado",
        data=_excel_en_memoria(df),
        file_name=f"datos_{hoy}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    # ----- CARGAR -----
    st.markdown("##### ⬆️ Cargar base de datos actualizada")
    st.caption("Sube un Excel (misma estructura de columnas) para reemplazar la base "
               "de datos. Se guardará y se subirá al repositorio para que la app tome "
               "la nueva información. **Recomendado:** parte del archivo que descargaste "
               "arriba para no perder columnas.")

    archivo = st.file_uploader("Selecciona el archivo datos.xlsx", type=["xlsx"])
    if archivo is None:
        return

    # Vista previa y validación
    try:
        df_nuevo = pd.read_excel(archivo, sheet_name=NOMBRE_HOJA)
    except Exception as e:
        st.error(f"No se pudo leer el archivo. ¿Tiene una hoja llamada '{NOMBRE_HOJA}'? "
                 f"Detalle: {e}")
        return

    df_nuevo.columns = [_normalizar_encabezado(c) for c in df_nuevo.columns]
    if COL_FOLIO not in df_nuevo.columns:
        st.error(f"El archivo no tiene la columna obligatoria '{COL_FOLIO}'. "
                 "Descarga la base actual y parte de ella.")
        return

    n_nuevos = len(df_nuevo)
    folios = df_nuevo[COL_FOLIO].astype(str).str.strip()
    duplicados = folios[folios.duplicated()].unique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Registros en el archivo", n_nuevos)
    c2.metric("Registros actuales", len(df))
    c3.metric("Diferencia", f"{n_nuevos - len(df):+d}")

    if len(duplicados) > 0:
        st.error(f"⚠️ Hay folios repetidos en el archivo (deben ser únicos): "
                 f"{', '.join(map(str, duplicados[:10]))}"
                 f"{'…' if len(duplicados) > 10 else ''}. Corrige antes de cargar.")
        return

    st.markdown("**Vista previa (primeras filas):**")
    st.dataframe(df_nuevo.head(10), use_container_width=True, hide_index=True)

    st.warning("Al confirmar, la base actual se reemplazará por completo con este "
               "archivo y se subirá al repositorio.")
    usuario = st.text_input("Tu nombre / usuario (para el registro)", key="user_carga_bd")
    confirmar = st.checkbox("Entiendo que se reemplazará toda la base de datos",
                            key="confirmar_carga_bd")

    if st.button("⬆️ Cargar y subir al repositorio", type="primary",
                 use_container_width=True, disabled=not confirmar):
        if not usuario.strip():
            st.error("✍️ Escribe tu nombre antes de cargar.")
            return
        try:
            # Escribir el archivo nuevo tal cual en datos.xlsx
            df_nuevo.to_excel(RUTA_EXCEL, sheet_name=NOMBRE_HOJA, index=False)
        except Exception as e:
            st.error(f"No se pudo escribir el Excel: {e}")
            return

        mensaje_commit = (f"Carga masiva de base de datos por {usuario.strip()} "
                          f"({n_nuevos} registros, {ahora_mx():%d/%m/%Y %H:%M})")
        with st.spinner("Subiendo la nueva base a GitHub…"):
            exito, mensaje = subir_a_github(mensaje_commit)

        cargar_datos.clear()
        st.session_state["version_datos"] += 1
        st.session_state["df"] = cargar_datos(RUTA_EXCEL, st.session_state["version_datos"])

        if exito:
            st.session_state["flash_bd"] = ("success", f"✅ Base cargada. {mensaje}")
        else:
            st.session_state["flash_bd"] = ("warning", f"💾 Guardado local, pero: {mensaje}")
        st.rerun()


# =============================================================================
# 13. FUNCIÓN PRINCIPAL
# =============================================================================

def main() -> None:
    if not verificar_acceso():
        st.stop()

    if "version_datos" not in st.session_state:
        st.session_state["version_datos"] = 0
    if "df" not in st.session_state:
        st.session_state["df"] = cargar_datos(RUTA_EXCEL, st.session_state["version_datos"])
    df = st.session_state["df"]

    # --- Encabezado compacto (una sola línea) ---
    st.markdown(
        "<div style='margin-bottom:0.2rem'>"
        "<span style='font-size:1.15rem;font-weight:700'>📋 Seguimiento a devoluciones</span>"
        "<span style='color:#888;font-size:0.85rem'> · Devoluciones y reclamaciones</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    # --- Barra lateral: control + filtros (colapsable, no ocupa el área principal) ---
    st.sidebar.markdown("### ⚙️ Panel")
    cb1, cb2 = st.sidebar.columns(2)
    if cb1.button("🔄 Recargar", use_container_width=True):
        cargar_datos.clear()
        st.session_state["version_datos"] += 1
        st.session_state["df"] = cargar_datos(RUTA_EXCEL, st.session_state["version_datos"])
        st.rerun()
    if cb2.button("🚪 Salir", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    df_filtrado = construir_filtros(df)

    # --- KPIs compactos en una línea ---
    mostrar_kpis(df_filtrado)

    # --- Aviso destacado de reclamos vencidos a 90 días ---
    criticos = [f for _, f in df_filtrado.iterrows() if vencido_90_sin_definicion(f)]
    if criticos:
        st.error(
            f"🚨 **{MSG_VENCIDO_90}** — {len(criticos)} reclamo(s) rebasaron los "
            f"{DIAS_VENCIMIENTO_TOTAL} días desde la fecha de corte. "
            "Revísalos en la pestaña *Resumen y alarmas*."
        )

    tab_graf, tab_editar, tab_tabla, tab_resumen, tab_bd, tab_guia = st.tabs(
        ["📈 Avance", "✏️ Editar reclamación", "📊 Tabla", "🔔 Resumen y alarmas",
         "🗄️ Base de datos", "📖 Guía"]
    )
    with tab_graf:
        vista_graficas(df_filtrado)
    with tab_editar:
        vista_editar(df_filtrado)
    with tab_tabla:
        mostrar_tabla(df_filtrado)
    with tab_resumen:
        mostrar_notificaciones(df_filtrado)
    with tab_bd:
        # La descarga/carga siempre opera sobre la base COMPLETA, no la filtrada.
        vista_base_datos(df)
    with tab_guia:
        vista_guia()


if __name__ == "__main__":
    main()
