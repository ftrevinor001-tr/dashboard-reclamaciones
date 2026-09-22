# =============================================================================
#  SEGUIMIENTO A DEVOLUCIONES — REDISEÑO 2026-09
#  -----------------------------------------------------------------------------
#  Aplicación Streamlit que sustituye por completo el modelo anterior por
#  etapas. Ahora:
#
#    - Cada folio se mide de extremo a extremo (recepción → nota de crédito)
#      con 90 días como plazo total.
#    - Un folio queda RESUELTO cuando se captura su NOTA DE CRÉDITO.
#    - Cuarentena automática para folios ≤ $300 (30 días); al vencer se
#      liberan solos y ahí empiezan los 90 días.
#    - Se agrega la pestaña "Notas de Crédito Pendientes" (medición: 20 días
#      desde la fecha de recepción del reporte hasta el envío a admin).
#    - Toda la gestión la lleva una sola persona → solo una contraseña de
#      acceso, sin claves para modificar/reactivar/liberar.
#
#  Fuentes de datos: un solo Excel con TRES hojas:
#     1) "datos- FOLIO DE GARANTIA"          → folios de garantía (Pestaña 2)
#     2) "FOLIOS DE DEVOLUCION-GARANTIA"     → folios sueltos de devolución
#     3) "NC PENDIENTES"                     → notas de crédito por conceptos
#                                              varios (Pestaña 3)
# =============================================================================

from __future__ import annotations

import os
import io
from datetime import datetime, date, timedelta

try:
    from zoneinfo import ZoneInfo
    ZONA_MX = ZoneInfo("America/Mexico_City")
except Exception:
    ZONA_MX = None

import numpy as np
import pandas as pd
import streamlit as st

try:
    import altair as alt
    ALTAIR_OK = True
except Exception:
    ALTAIR_OK = False


# =============================================================================
# 1. CONFIGURACIÓN GENERAL Y CONSTANTES
# =============================================================================

st.set_page_config(
    page_title="Seguimiento a devoluciones",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# CSS de compactación para aprovechar el ancho y reducir espacios en blanco
st.markdown("""
<style>
    /* Reducir padding superior y márgenes generales */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 1rem !important;
        max-width: 100% !important;
    }
    /* Encabezados más compactos */
    h1, h2, h3 { margin-top: 0.4rem !important; margin-bottom: 0.4rem !important; }
    h4 { margin-top: 0.6rem !important; margin-bottom: 0.3rem !important;
         font-size: 1.05rem !important; }
    /* Separadores más finos */
    hr { margin: 0.6rem 0 !important; }
    /* Tabs con menos padding */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        padding-top: 0.4rem !important;
        padding-bottom: 0.4rem !important;
    }
    /* Reducir margen de widgets */
    [data-testid="stVerticalBlock"] { gap: 0.4rem !important; }
    /* Captions más discretos */
    .stCaption { font-size: 0.78rem !important; color: #64748b !important; }
</style>
""", unsafe_allow_html=True)

RUTA_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_EXCEL = os.path.join(RUTA_BASE, "datos.xlsx")

# --- Hojas del Excel ---
HOJA_GARANTIAS = "datos- FOLIO DE GARANTIA"
HOJA_DEVOLUCIONES = "FOLIOS DE DEVOLUCION-GARANTIA"
HOJA_NC = "NC PENDIENTES"

# --- Umbrales del proceso ---
DIAS_PLAZO_TOTAL = 90     # plazo total de un folio de garantía
DIAS_PLAZO_NC = 20        # plazo para gestionar una nota de crédito pendiente

# Metas de cumplimiento e indicadores de bono
META_CUMPLIMIENTO_GARANTIA = 95.0   # % objetivo de folios cerrados en 90 días
META_CUMPLIMIENTO_NC = 90.0         # % objetivo de NC cerradas en 20 días
DIAS_ALERTA_ROJA = 120              # folio abierto > 120 días penaliza bono
DIAS_ALERTA_NC = 60                 # NC abierta > 60 días penaliza bono
DIAS_POR_VENCER = 15      # ventana amarilla antes del vencimiento
DIAS_POR_VENCER_NC = 5    # ventana amarilla para NC (plazo más corto)
UMBRAL_CUARENTENA = 300.0
DIAS_CUARENTENA = 30

MSG_VENCIDO = "RECLAMO VENCIDO SIN DEFINICIÓN, ENVIAR A DESTRUCCIÓN"

# =============================================================================
# 2. UTILIDADES DE FECHA Y CONVERSIÓN
# =============================================================================

def hoy_mx() -> date:
    """Fecha actual en la zona horaria de México (evita el desfase por UTC)."""
    if ZONA_MX is not None:
        return datetime.now(ZONA_MX).date()
    return (datetime.utcnow() - timedelta(hours=6)).date()


def ahora_mx() -> datetime:
    """Fecha y hora actual en la zona horaria de México (sin tzinfo)."""
    if ZONA_MX is not None:
        return datetime.now(ZONA_MX).replace(tzinfo=None)
    return datetime.utcnow() - timedelta(hours=6)


def _a_fecha(valor):
    """Convierte cualquier cosa razonable a date, o None si no es válido."""
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
    f = _a_fecha(valor)
    return f"{f:%d/%m/%Y}" if f else "—"


def _fmt_mxn(valor) -> str:
    try:
        return f"${float(valor):,.2f}"
    except (TypeError, ValueError):
        return "—"


MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}


def etiqueta_mes(fecha) -> str:
    """'Marzo 2026' a partir de una fecha o timestamp."""
    f = _a_fecha(fecha)
    if f is None:
        return "Sin fecha"
    return f"{MESES_ES[f.month]} {f.year}"


# =============================================================================
# 3. CARGA Y NORMALIZACIÓN DE DATOS
# =============================================================================

# --- Columnas de la hoja 1: FOLIOS DE GARANTÍA ---
COL_G_MES = "MES DE DEVOLUCION"
COL_G_FOLIO = "FOLIO REPORTE"
COL_G_ID = "ID"
COL_G_PROVEEDOR = "PROVEEDOR"
COL_G_CARTA = "CARTA FIRMADA"
COL_G_IMPORTE = "IMPORTE (MXN)"
COL_G_COMPRADOR = "COMPRADOR"
COL_G_FECHA_CORTE = "FECHA CORTE"
COL_G_FECHA_RECEPCION = "FECHA RECEPCION REPORTE"  # NUEVA — arranca el reloj
COL_G_FOLIO_DEV = "FOLIO DEVOLUCION"
COL_G_FOLIO_AJUSTE = "FOLIO AJUSTE"
COL_G_NOTA_CREDITO = "NOTA CREDITO"
COL_G_RESPUESTA = "RESPUESTA PROVEEDOR TIPO"
COL_G_NOTAS = "ACCIONES / NOTAS"
# Columnas de control interno
COL_G_ESTADO = "ESTADO"                     # activo / cuarentena / resuelto / cancelado
COL_G_ETAPA = "ETAPA"                       # estatus intermedio del flujo
COL_G_CUAR_INICIO = "CUARENTENA INICIO"     # fecha entró
COL_G_CUAR_FIN = "CUARENTENA FIN"           # fecha calculada de liberación
COL_G_FECHA_RESUELTO = "FECHA RESUELTO"     # cuándo se capturó la NC
COL_G_MODIFICADO = "ULTIMA MODIFICACION"

# Estados posibles de un folio de garantía
ESTADO_ACTIVO = "Activo"
ESTADO_CUARENTENA = "Cuarentena"
ESTADO_RESUELTO = "Resuelto"
ESTADO_CANCELADO = "Cancelado"

# Estatus intermedios (kanban) para folios de garantía activos
# Lista ordenada de izquierda a derecha en el tablero de flujo.
ETAPAS_GARANTIA = [
    "📝 Reporte recibido",
    "📤 Reportado al proveedor",
    "⏳ Esperando respuesta",
    "✉️ Respuesta recibida",
    "🚚 En recolección",
    "🗑️ En destrucción",
    "📄 Folio devolución/ajuste generado",
    "📨 Enviado a Cuentas por Pagar",
]
ETAPA_GARANTIA_INICIAL = ETAPAS_GARANTIA[0]

# Estatus intermedios para NC pendientes
ETAPAS_NC = [
    "📥 Reporte recibido",
    "📤 Solicitada al proveedor",
    "🔍 En revisión del proveedor",
    "✏️ Autorizada por proveedor",
    "📄 Emitida (pendiente timbrar)",
]
ETAPA_NC_INICIAL = ETAPAS_NC[0]

# --- Columnas de la hoja 2: FOLIOS DE DEVOLUCIÓN ---
COL_D_FECHA_REPORTE = "FECHA REPORTE"  # NUEVA — arranca el reloj de 20 días
COL_D_FOLIO = "Folio"
COL_D_FECHA = "Fecha"
COL_D_PROVEEDOR = "Proveedor"
COL_D_TOTAL = "Total"
COL_D_PENDIENTE = "PENDIENTE POR DESCONTAR"
COL_D_APLICADO_MXN = "APLICADO"
COL_D_APLICADO_EST = "Aplicado"
COL_D_RESOLUCION = "Resolucion"
COL_D_TIPO_CLIENTE = "Devolucion cliente saldo a favor"
COL_D_EJECUTIVO = "Ejecutivo"
COL_D_COMPRADOR = "Comprador"
COL_D_ESTADO = "ESTADO"  # activo / bloqueado / cancelado
COL_D_NOTAS = "NOTAS"

# --- Columnas de la hoja 3: NC PENDIENTES ---
COL_NC_FECHA_REPORTE = "FECHA REPORTE"  # NUEVA — arranca el reloj de 20 días
COL_NC_ENTRADA = "FOLIO DE ENTRADA"
COL_NC_FECHA_FACTURA = "FECHA FACTURA"
COL_NC_FISCAL = "FOLIO FISCAL"
COL_NC_PROVEEDOR = "PROVEEDOR"
COL_NC_NUM_FACTURA = "NUM. FACTURA"
COL_NC_IMP_FACTURA = "IMPORTE FACTURA"
COL_NC_IMP_PENDIENTE = "IMPORTE DE LA NC PENDIENTE POR TIMBRAR"
COL_NC_FECHA_NC = "FECHA DE LA NC"
COL_NC_FOLIO_NC = "FOLIO"
COL_NC_OBSERVACIONES = "OBSERVACIONES"
COL_NC_EJECUTIVA = "EJECUTIVA"
COL_NC_COMPRADOR = "COMPRADOR"
COL_NC_ESTADO = "ESTADO"  # pendiente / resuelto / cancelado
COL_NC_ETAPA = "ETAPA"    # estatus intermedio del flujo


def _normalizar_encabezados(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [" ".join(str(c).split()).strip() for c in df.columns]
    return df


def _renombrar(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """Renombra columnas usando el mapa {viejo: nuevo}; ignora las que faltan."""
    df = df.rename(columns={k: v for k, v in mapa.items() if k in df.columns})
    return df


@st.cache_data(show_spinner="Cargando datos…")
def cargar_datos(ruta: str, _version: int) -> dict:
    """Carga las tres hojas del Excel y devuelve un diccionario.

    _version se usa solo para invalidar el caché cuando queremos releer.
    """
    resultado = {"garantias": None, "devoluciones": None, "nc": None,
                 "error": None}
    if not os.path.exists(ruta):
        resultado["error"] = f"No se encontró el archivo: {ruta}"
        return resultado
    try:
        xl = pd.ExcelFile(ruta)
    except Exception as e:
        resultado["error"] = f"No se pudo abrir el Excel: {e}"
        return resultado

    # ------ HOJA 1: GARANTÍAS ------
    if HOJA_GARANTIAS in xl.sheet_names:
        g = pd.read_excel(xl, sheet_name=HOJA_GARANTIAS)
        g = _normalizar_encabezados(g)
        # Añadir columnas nuevas si faltan
        for col, default in [
            (COL_G_FECHA_RECEPCION, pd.NaT),
            (COL_G_ESTADO, ""),
            (COL_G_CUAR_INICIO, pd.NaT),
            (COL_G_CUAR_FIN, pd.NaT),
            (COL_G_FECHA_RESUELTO, pd.NaT),
            (COL_G_MODIFICADO, ""),
            (COL_G_ETAPA, ""),
        ]:
            if col not in g.columns:
                g[col] = default
        # Tipos
        for col in [COL_G_MES, COL_G_FECHA_CORTE, COL_G_FECHA_RECEPCION,
                    COL_G_CUAR_INICIO, COL_G_CUAR_FIN, COL_G_FECHA_RESUELTO]:
            g[col] = pd.to_datetime(g[col], errors="coerce").dt.normalize()
        g[COL_G_IMPORTE] = pd.to_numeric(g[COL_G_IMPORTE], errors="coerce").fillna(0.0)
        for col in [COL_G_FOLIO, COL_G_PROVEEDOR, COL_G_COMPRADOR,
                    COL_G_CARTA, COL_G_NOTAS, COL_G_RESPUESTA]:
            if col in g.columns:
                g[col] = g[col].fillna("").astype(str).str.strip()

        # Si no hay fecha de recepción del reporte, usar FECHA CORTE como valor
        # inicial (los folios viejos ya nacen con el reloj corriendo desde ahí).
        mask_sin_recep = g[COL_G_FECHA_RECEPCION].isna()
        g.loc[mask_sin_recep, COL_G_FECHA_RECEPCION] = g.loc[mask_sin_recep, COL_G_FECHA_CORTE]

        # Determinar ESTADO si viene vacío (por migración desde el modelo viejo)
        g[COL_G_ESTADO] = g[COL_G_ESTADO].fillna("").astype(str).str.strip()
        vacio = g[COL_G_ESTADO].isin(["", "nan", "None"])
        # Regla de migración desde el modelo por etapas:
        #  - ETAPA ACTUAL = "FINALIZADO"           → Resuelto
        #  - Si tiene NOTA CREDITO capturada       → Resuelto
        #  - ETAPA ACTUAL = "Cuarentena"           → Cuarentena
        #  - Si importe ≤ UMBRAL y no cabe arriba  → Cuarentena
        #  - Resto                                 → Activo
        etapa_vieja = g.get("ETAPA ACTUAL", pd.Series([""] * len(g))).fillna("").astype(str).str.strip()
        # Para detectar NC real, usar la columna ORIGINAL con pd.notna, no strings
        if COL_G_NOTA_CREDITO in g.columns:
            # Reconstruir booleano desde la fuente original (antes del cast a str)
            nc_original = pd.read_excel(ruta, sheet_name=HOJA_GARANTIAS,
                                        usecols=[COL_G_NOTA_CREDITO])[COL_G_NOTA_CREDITO]
            con_nc = nc_original.notna() & (nc_original.astype(str).str.strip() != "")
            con_nc = con_nc.reindex(g.index, fill_value=False)
        else:
            con_nc = pd.Series([False] * len(g), index=g.index)
        bajo = g[COL_G_IMPORTE] <= UMBRAL_CUARENTENA
        # Aplicar en orden de precedencia
        g.loc[vacio & (etapa_vieja == "FINALIZADO"), COL_G_ESTADO] = ESTADO_RESUELTO
        aun_vacio = g[COL_G_ESTADO].isin(["", "nan", "None"])
        g.loc[aun_vacio & con_nc, COL_G_ESTADO] = ESTADO_RESUELTO
        aun_vacio = g[COL_G_ESTADO].isin(["", "nan", "None"])
        g.loc[aun_vacio & (etapa_vieja == "Cuarentena"), COL_G_ESTADO] = ESTADO_CUARENTENA
        aun_vacio = g[COL_G_ESTADO].isin(["", "nan", "None"])
        g.loc[aun_vacio & bajo, COL_G_ESTADO] = ESTADO_CUARENTENA
        aun_vacio = g[COL_G_ESTADO].isin(["", "nan", "None"])
        g.loc[aun_vacio, COL_G_ESTADO] = ESTADO_ACTIVO

        # Limpiar las cadenas de folios/nc después de la migración
        for col in [COL_G_FOLIO_DEV, COL_G_FOLIO_AJUSTE, COL_G_NOTA_CREDITO]:
            if col in g.columns:
                # Convertir NaN a "" y limpiar
                g[col] = g[col].apply(
                    lambda v: "" if pd.isna(v) or str(v).strip().lower() in ("nan", "none")
                    else str(v).strip())
                # Quitar el ".0" que agrega pandas a los enteros de Excel
                g[col] = g[col].str.replace(r"\.0$", "", regex=True)

        # ETAPA (estatus intermedio del flujo). Solo aplica a folios Activos.
        # Si vacía, se infiere del avance ya capturado:
        #   tiene folio devolución/ajuste  → "Folio devolución/ajuste generado"
        #   respuesta = Destrucción         → "En destrucción"
        #   respuesta = Recolección         → "En recolección"
        #   respuesta = otra                → "Respuesta recibida"
        #   sin respuesta                    → "Reporte recibido" (inicio)
        g[COL_G_ETAPA] = g[COL_G_ETAPA].fillna("").astype(str).str.strip()
        vacia_etapa = g[COL_G_ETAPA] == ""
        activos = g[COL_G_ESTADO] == ESTADO_ACTIVO
        tiene_folio_dev = g[COL_G_FOLIO_DEV].astype(str).str.strip() != ""
        tiene_folio_aj = g[COL_G_FOLIO_AJUSTE].astype(str).str.strip() != ""
        respuesta = g[COL_G_RESPUESTA].astype(str).str.strip().str.lower()
        # Prioridad de inferencia
        g.loc[vacia_etapa & activos & (tiene_folio_dev | tiene_folio_aj),
              COL_G_ETAPA] = "📄 Folio devolución/ajuste generado"
        aun = g[COL_G_ETAPA] == ""
        g.loc[aun & activos & (respuesta == "destrucción"),
              COL_G_ETAPA] = "🗑️ En destrucción"
        aun = g[COL_G_ETAPA] == ""
        g.loc[aun & activos & (respuesta == "recolección"),
              COL_G_ETAPA] = "🚚 En recolección"
        aun = g[COL_G_ETAPA] == ""
        g.loc[aun & activos & respuesta.isin(["sin respuesta", ""]),
              COL_G_ETAPA] = ETAPA_GARANTIA_INICIAL
        aun = g[COL_G_ETAPA] == ""
        g.loc[aun & activos, COL_G_ETAPA] = "✉️ Respuesta recibida"

        # Inicio de cuarentena si le falta
        m_cuar = (g[COL_G_ESTADO] == ESTADO_CUARENTENA) & (g[COL_G_CUAR_INICIO].isna())
        g.loc[m_cuar, COL_G_CUAR_INICIO] = pd.Timestamp(hoy_mx())
        g.loc[m_cuar, COL_G_CUAR_FIN] = pd.Timestamp(hoy_mx() + timedelta(days=DIAS_CUARENTENA))

        # Autoliberar cuarentenas vencidas: pasan a Activo y su reloj de 90d
        # arranca AHORA (se ajusta FECHA RECEPCION al fin de cuarentena).
        hoy = pd.Timestamp(hoy_mx())
        m_venc = (g[COL_G_ESTADO] == ESTADO_CUARENTENA) & \
                 (g[COL_G_CUAR_FIN].notna()) & (g[COL_G_CUAR_FIN] <= hoy)
        if m_venc.any():
            g.loc[m_venc, COL_G_ESTADO] = ESTADO_ACTIVO
            # Al liberarse, comienza el reloj de 90 días
            g.loc[m_venc, COL_G_FECHA_RECEPCION] = g.loc[m_venc, COL_G_CUAR_FIN]

        # Etiqueta de mes
        g["MES ETIQUETA"] = g[COL_G_MES].apply(etiqueta_mes)

        resultado["garantias"] = g
    else:
        resultado["error"] = f"Falta la hoja '{HOJA_GARANTIAS}' en el Excel."

    # ------ HOJA 2: DEVOLUCIONES ------
    if HOJA_DEVOLUCIONES in xl.sheet_names:
        d = pd.read_excel(xl, sheet_name=HOJA_DEVOLUCIONES)
        d = _normalizar_encabezados(d)
        for col, default in [(COL_D_ESTADO, "Activo"), (COL_D_NOTAS, ""),
                              (COL_D_FECHA_REPORTE, pd.NaT)]:
            if col not in d.columns:
                d[col] = default
        # Tipos
        d[COL_D_FECHA] = pd.to_datetime(d[COL_D_FECHA], errors="coerce").dt.normalize()
        d[COL_D_FECHA_REPORTE] = pd.to_datetime(d[COL_D_FECHA_REPORTE],
                                                  errors="coerce").dt.normalize()
        for col in [COL_D_TOTAL, COL_D_PENDIENTE, COL_D_APLICADO_MXN]:
            if col in d.columns:
                d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0.0)
        for col in [COL_D_PROVEEDOR, COL_D_RESOLUCION, COL_D_TIPO_CLIENTE,
                    COL_D_EJECUTIVO, COL_D_COMPRADOR, COL_D_NOTAS,
                    COL_D_APLICADO_EST, COL_D_ESTADO]:
            if col in d.columns:
                d[col] = d[col].fillna("").astype(str).str.strip()
        # Estado por defecto
        d.loc[d[COL_D_ESTADO] == "", COL_D_ESTADO] = "Activo"
        # Marcar si es garantía (tipo cliente) o incidente con proveedor
        d["TIPO"] = d[COL_D_TIPO_CLIENTE].apply(
            lambda x: "Garantía (cliente)" if str(x).strip().lower() == "clientes"
            else "Incidente proveedor" if str(x).strip() == ""
            else "Otro")
        d["MES ETIQUETA"] = d[COL_D_FECHA].apply(etiqueta_mes)
        resultado["devoluciones"] = d

    # ------ HOJA 3: NC PENDIENTES ------
    if HOJA_NC in xl.sheet_names:
        nc = pd.read_excel(xl, sheet_name=HOJA_NC)
        nc = _normalizar_encabezados(nc)
        # Renombrar la columna con nombre largo si existe
        obs_larga = "OBSERVACIONES , QUE SE ESTA REALIZANDO PARA QUE NOS EMITAN LA NC"
        if obs_larga in nc.columns:
            nc = nc.rename(columns={obs_larga: COL_NC_OBSERVACIONES})
        # Migración: si viene la columna antigua "FECHA RECEPCION REPORTE"
        # y no la nueva "FECHA REPORTE", usarla como origen.
        if ("FECHA RECEPCION REPORTE" in nc.columns
                and COL_NC_FECHA_REPORTE not in nc.columns):
            nc = nc.rename(columns={"FECHA RECEPCION REPORTE": COL_NC_FECHA_REPORTE})
        # Si vienen las dos, priorizar la nueva y llenar huecos con la vieja
        if ("FECHA RECEPCION REPORTE" in nc.columns
                and COL_NC_FECHA_REPORTE in nc.columns):
            vieja = pd.to_datetime(nc["FECHA RECEPCION REPORTE"],
                                    errors="coerce").dt.normalize()
            nueva = pd.to_datetime(nc[COL_NC_FECHA_REPORTE],
                                     errors="coerce").dt.normalize()
            nc[COL_NC_FECHA_REPORTE] = nueva.fillna(vieja)
            nc = nc.drop(columns=["FECHA RECEPCION REPORTE"])
        for col, default in [
            (COL_NC_FECHA_REPORTE, pd.NaT),
            (COL_NC_ESTADO, ""),
            (COL_NC_ETAPA, ""),
        ]:
            if col not in nc.columns:
                nc[col] = default
        # Tipos: FECHA FACTURA puede venir como string por celdas mezcladas
        nc[COL_NC_FECHA_FACTURA] = pd.to_datetime(nc[COL_NC_FECHA_FACTURA],
                                                   errors="coerce").dt.normalize()
        nc[COL_NC_FECHA_REPORTE] = pd.to_datetime(nc[COL_NC_FECHA_REPORTE],
                                                   errors="coerce").dt.normalize()
        nc[COL_NC_FECHA_NC] = pd.to_datetime(nc[COL_NC_FECHA_NC],
                                              errors="coerce").dt.normalize()
        for col in [COL_NC_IMP_FACTURA, COL_NC_IMP_PENDIENTE]:
            if col in nc.columns:
                nc[col] = pd.to_numeric(nc[col], errors="coerce").fillna(0.0)
        for col in [COL_NC_PROVEEDOR, COL_NC_OBSERVACIONES, COL_NC_EJECUTIVA,
                    COL_NC_COMPRADOR, COL_NC_ESTADO]:
            if col in nc.columns:
                nc[col] = nc[col].fillna("").astype(str).str.strip()
        # Si no hay fecha de reporte, usar la fecha de la factura como respaldo
        m_sin_rep = nc[COL_NC_FECHA_REPORTE].isna()
        nc.loc[m_sin_rep, COL_NC_FECHA_REPORTE] = nc.loc[m_sin_rep, COL_NC_FECHA_FACTURA]
        # Estado por defecto: si tiene FECHA DE LA NC → Resuelto, si no → Pendiente
        vacio_e = nc[COL_NC_ESTADO] == ""
        con_nc_e = nc[COL_NC_FECHA_NC].notna()
        nc.loc[vacio_e & con_nc_e, COL_NC_ESTADO] = "Resuelto"
        nc.loc[vacio_e & ~con_nc_e, COL_NC_ESTADO] = "Pendiente"
        # ETAPA inicial: pendientes empiezan en el paso 1 si vacía
        nc[COL_NC_ETAPA] = nc[COL_NC_ETAPA].fillna("").astype(str).str.strip()
        vacia_nc = nc[COL_NC_ETAPA] == ""
        pendientes = nc[COL_NC_ESTADO] == "Pendiente"
        nc.loc[vacia_nc & pendientes, COL_NC_ETAPA] = ETAPA_NC_INICIAL
        nc["MES ETIQUETA"] = nc[COL_NC_FECHA_REPORTE].apply(etiqueta_mes)
        resultado["nc"] = nc

    return resultado


def guardar_excel(datos: dict, ruta: str) -> None:
    """Sobrescribe el Excel con las tres hojas.

    Si alguna hoja viene vacía (None), se escribe una hoja mínima con solo los
    encabezados para no romper el writer (openpyxl no admite libros vacíos).
    """
    with pd.ExcelWriter(ruta, engine="openpyxl",
                        datetime_format="DD/MM/YYYY",
                        date_format="DD/MM/YYYY") as writer:
        # Garantías (siempre escribir algo, aunque sea plantilla vacía)
        g = datos.get("garantias")
        if g is not None:
            g = g.copy().drop(columns=["MES ETIQUETA"], errors="ignore")
        else:
            g = pd.DataFrame(columns=[
                COL_G_MES, COL_G_FOLIO, COL_G_ID, COL_G_PROVEEDOR, COL_G_CARTA,
                COL_G_IMPORTE, COL_G_COMPRADOR, COL_G_FECHA_CORTE,
                COL_G_FECHA_RECEPCION, COL_G_FOLIO_DEV, COL_G_FOLIO_AJUSTE,
                COL_G_NOTA_CREDITO, COL_G_RESPUESTA, COL_G_NOTAS,
                COL_G_ESTADO, COL_G_ETAPA, COL_G_CUAR_INICIO, COL_G_CUAR_FIN,
                COL_G_FECHA_RESUELTO, COL_G_MODIFICADO])
        g.to_excel(writer, sheet_name=HOJA_GARANTIAS, index=False)

        # Devoluciones
        d = datos.get("devoluciones")
        if d is not None:
            d = d.copy().drop(columns=["MES ETIQUETA", "TIPO"], errors="ignore")
        else:
            d = pd.DataFrame(columns=[
                COL_D_FECHA_REPORTE, COL_D_FOLIO, COL_D_FECHA, COL_D_PROVEEDOR,
                COL_D_TOTAL, COL_D_PENDIENTE, COL_D_APLICADO_MXN,
                COL_D_APLICADO_EST, COL_D_RESOLUCION, COL_D_TIPO_CLIENTE,
                COL_D_EJECUTIVO, COL_D_COMPRADOR, COL_D_ESTADO, COL_D_NOTAS])
        d.to_excel(writer, sheet_name=HOJA_DEVOLUCIONES, index=False)

        # NC
        nc = datos.get("nc")
        if nc is not None:
            nc = nc.copy().drop(columns=["MES ETIQUETA"], errors="ignore")
            nc = nc.rename(columns={
                COL_NC_OBSERVACIONES:
                "OBSERVACIONES , QUE SE ESTA REALIZANDO PARA QUE NOS EMITAN LA NC"})
        else:
            nc = pd.DataFrame(columns=[
                COL_NC_ENTRADA, COL_NC_FECHA_FACTURA, COL_NC_FECHA_REPORTE,
                COL_NC_FISCAL, COL_NC_PROVEEDOR, COL_NC_NUM_FACTURA,
                COL_NC_IMP_FACTURA, COL_NC_IMP_PENDIENTE, COL_NC_FECHA_NC,
                COL_NC_FOLIO_NC,
                "OBSERVACIONES , QUE SE ESTA REALIZANDO PARA QUE NOS EMITAN LA NC",
                COL_NC_EJECUTIVA, COL_NC_COMPRADOR, COL_NC_ESTADO,
                COL_NC_ETAPA])
        nc.to_excel(writer, sheet_name=HOJA_NC, index=False)


# =============================================================================
# 4. LÓGICA DE ESTADO Y SEMÁFOROS
# =============================================================================

def dias_transcurridos_garantia(fila: pd.Series) -> int | None:
    """Días transcurridos desde la fecha de recepción del reporte."""
    inicio = _a_fecha(fila.get(COL_G_FECHA_RECEPCION))
    if inicio is None:
        return None
    return (hoy_mx() - inicio).days


def dias_restantes_garantia(fila: pd.Series) -> int | None:
    """Días que faltan para vencer el plazo de 90 días."""
    d = dias_transcurridos_garantia(fila)
    if d is None:
        return None
    return DIAS_PLAZO_TOTAL - d


def fecha_vencimiento_garantia(fila: pd.Series):
    inicio = _a_fecha(fila.get(COL_G_FECHA_RECEPCION))
    if inicio is None:
        return None
    return inicio + timedelta(days=DIAS_PLAZO_TOTAL)


def semaforo_garantia(fila: pd.Series) -> tuple[str, str]:
    """Devuelve (icono, texto) del semáforo de un folio de garantía.

    - Resuelto/Cancelado tienen su propio icono.
    - En cuarentena: azul con días para salir.
    - Activo: 🟢 en tiempo · 🟡 por vencerse (≤15d) · 🔴 vencido.
    """
    estado = str(fila.get(COL_G_ESTADO, "")).strip()
    if estado == ESTADO_RESUELTO:
        return "✅", "Resuelto"
    if estado == ESTADO_CANCELADO:
        return "⚫", "Cancelado"
    if estado == ESTADO_CUARENTENA:
        fin = _a_fecha(fila.get(COL_G_CUAR_FIN))
        if fin:
            faltan = (fin - hoy_mx()).days
            return "🧊", f"En cuarentena · sale en {faltan} día(s)"
        return "🧊", "En cuarentena"
    # Activo
    restantes = dias_restantes_garantia(fila)
    if restantes is None:
        return "—", "Sin fecha de recepción"
    if restantes < 0:
        return "🔴", f"VENCIDO hace {abs(restantes)} día(s)"
    if restantes <= DIAS_POR_VENCER:
        return "🟡", f"Por vencerse: {restantes} día(s)"
    return "🟢", f"En tiempo: {restantes} día(s)"


def esta_vencido_garantia(fila: pd.Series) -> bool:
    """True si el folio está activo y pasó los 90 días sin resolverse."""
    if str(fila.get(COL_G_ESTADO, "")).strip() != ESTADO_ACTIVO:
        return False
    r = dias_restantes_garantia(fila)
    return r is not None and r < 0


def dias_transcurridos_nc(fila: pd.Series) -> int | None:
    inicio = _a_fecha(fila.get(COL_NC_FECHA_REPORTE))
    if inicio is None:
        return None
    return (hoy_mx() - inicio).days


def dias_restantes_nc(fila: pd.Series) -> int | None:
    d = dias_transcurridos_nc(fila)
    if d is None:
        return None
    return DIAS_PLAZO_NC - d


def semaforo_nc(fila: pd.Series) -> tuple[str, str]:
    estado = str(fila.get(COL_NC_ESTADO, "")).strip()
    if estado == "Resuelto":
        return "✅", "Resuelto"
    if estado == "Cancelado":
        return "⚫", "Cancelado"
    restantes = dias_restantes_nc(fila)
    if restantes is None:
        return "—", "Sin fecha"
    if restantes < 0:
        return "🔴", f"VENCIDA hace {abs(restantes)} día(s)"
    if restantes <= DIAS_POR_VENCER_NC:
        return "🟡", f"Por vencerse: {restantes} día(s)"
    return "🟢", f"En tiempo: {restantes} día(s)"


def esta_vencida_nc(fila: pd.Series) -> bool:
    if str(fila.get(COL_NC_ESTADO, "")).strip() != "Pendiente":
        return False
    r = dias_restantes_nc(fila)
    return r is not None and r < 0


# =============================================================================
# 5. AUTENTICACIÓN
# =============================================================================

def verificar_acceso() -> bool:
    if st.session_state.get("autenticado", False):
        return True
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        st.markdown("## 🔐 Seguimiento a devoluciones")
        st.caption("Acceso restringido. Ingresa la contraseña para continuar.")
        with st.form("form_login"):
            password = st.text_input("Contraseña", type="password",
                                     placeholder="••••••••")
            enviar = st.form_submit_button("Ingresar", use_container_width=True)
        if enviar:
            correcta = st.secrets.get("DASHBOARD_PASSWORD")
            if correcta is None:
                st.error("⚠️ Falta 'DASHBOARD_PASSWORD' en los secretos de la app.")
            elif password == correcta:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
    return False


# =============================================================================
# 6. GUARDADO + SINCRONIZACIÓN CON GITHUB
# =============================================================================

def subir_a_github(mensaje_commit: str) -> tuple[bool, str]:
    """Hace commit y push del Excel usando GitPython y st.secrets['GITHUB_TOKEN']."""
    try:
        from git import Repo
    except ImportError:
        return False, "GitPython no está instalado (revisa requirements.txt)."
    token = st.secrets.get("GITHUB_TOKEN")
    if not token:
        return False, ("Falta 'GITHUB_TOKEN' en los secretos. Los cambios se "
                       "guardaron localmente pero no se subieron al repositorio.")
    try:
        repo = Repo(RUTA_BASE, search_parent_directories=True)
        with repo.config_writer() as cw:
            cw.set_value("user", "name", "Dashboard Devoluciones")
            cw.set_value("user", "email", "dashboard@devoluciones.app")
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
        return True, "Cambios subidos al repositorio ✅"
    except Exception as e:
        detalle = str(e)
        if "403" in detalle or "denied" in detalle.lower():
            ayuda = " — El token no tiene permiso de escritura."
        elif "could not read Password" in detalle or "Authentication" in detalle:
            ayuda = " — El token es inválido o expiró; genera uno nuevo."
        else:
            ayuda = ""
        return False, f"Error al subir a GitHub: {detalle}{ayuda}"


def persistir(datos: dict, mensaje_commit: str, mensaje_ok: str) -> None:
    """Guarda a Excel, sube a GitHub, recarga en el ORDEN correcto y re-renderiza."""
    try:
        guardar_excel(datos, RUTA_EXCEL)
    except Exception as e:
        st.error(f"No se pudo escribir el Excel: {e}")
        return
    with st.spinner("Guardando y sincronizando…"):
        exito, msg = subir_a_github(mensaje_commit)
    cargar_datos.clear()
    st.session_state["version_datos"] += 1
    st.session_state["datos"] = cargar_datos(RUTA_EXCEL,
                                              st.session_state["version_datos"])
    if exito:
        st.session_state["flash"] = ("success", f"✅ {mensaje_ok} · {msg}")
    else:
        st.session_state["flash"] = ("warning",
                                     f"💾 {mensaje_ok} · Guardado local, pero: {msg}")
    st.rerun()


# =============================================================================
# 7. FILTROS COMPARTIDOS
# =============================================================================
#
#   Todas las pestañas usan el mismo conjunto de filtros: periodo (mes o rango),
#   proveedor, comprador y búsqueda por folio. Cada pestaña pasa el DataFrame
#   correspondiente (garantías o NC) y recibe el mismo DataFrame ya filtrado.


def _orden_mes(etiqueta: str) -> str:
    """Devuelve una clave AAAA-MM ordenable a partir de 'Marzo 2026'."""
    if etiqueta == "Sin fecha":
        return "0000-00"
    try:
        nombre, año = etiqueta.split()
        m = list(MESES_ES.values()).index(nombre) + 1
        return f"{año}-{m:02d}"
    except Exception:
        return "0000-00"


def _rango_meses_placeholder(df: pd.DataFrame) -> str:
    """Devuelve el rango de meses disponibles como 'Mes Año — Mes Año'.

    Sirve como texto guía en los multiselects de mes: en vez del genérico
    "Todos los meses", muestra los extremos reales del periodo cargado.
    """
    if df.empty or "MES ETIQUETA" not in df.columns:
        return "Todos los meses"
    meses = [m for m in df["MES ETIQUETA"].dropna().unique()
             if m != "Sin fecha"]
    if not meses:
        return "Todos los meses"
    meses = sorted(meses, key=_orden_mes)
    if len(meses) == 1:
        return meses[0]
    return f"{meses[0]} — {meses[-1]}"


def _filtros_barra(df: pd.DataFrame, prefijo: str,
                    col_fecha: str, col_folio: str,
                    col_proveedor: str, col_comprador: str,
                    etiqueta_folio: str = "Buscar folio") -> tuple:
    """Barra de filtros común: periodo, proveedor, comprador, folio.

    Devuelve (df_filtrado, etiqueta_periodo). Las claves de widget llevan
    `prefijo` para no chocar entre pestañas.
    """
    # ---- FILA 1: periodo ----
    cA, cB = st.columns([1, 3])
    modo = cA.radio("Filtrar periodo por", ["Mes", "Rango de fechas"],
                     horizontal=True, key=f"{prefijo}_modo")

    hoy = hoy_mx()
    if modo == "Mes":
        meses_disp = sorted(
            [m for m in df["MES ETIQUETA"].dropna().unique()
             if m != "Sin fecha"], key=_orden_mes)
        placeholder_meses = _rango_meses_placeholder(df)
        sel_meses = cB.multiselect(
            "Mes(es)", options=meses_disp,
            placeholder=placeholder_meses,
            key=f"{prefijo}_meses")
        if sel_meses:
            df_p = df[df["MES ETIQUETA"].isin(sel_meses)]
            etiqueta_periodo = ", ".join(sel_meses)
        else:
            df_p = df
            etiqueta_periodo = placeholder_meses
    else:
        # Rango de fechas
        fmin_val = df[col_fecha].min()
        fmin = _a_fecha(fmin_val) if pd.notna(fmin_val) else hoy - timedelta(days=180)
        default = (fmin, hoy)
        rango = cB.date_input("Rango de fechas", value=default,
                                format="DD/MM/YYYY", key=f"{prefijo}_rango")
        if isinstance(rango, tuple) and len(rango) == 2:
            desde, hasta = rango
            df_p = df[(df[col_fecha].dt.date >= desde) &
                       (df[col_fecha].dt.date <= hasta)]
            etiqueta_periodo = f"{desde:%d/%m/%Y} — {hasta:%d/%m/%Y}"
        else:
            df_p = df
            etiqueta_periodo = "Todo el periodo"

    # ---- FILA 2: proveedor, comprador, folio ----
    c1, c2, c3 = st.columns(3)
    proveedores = sorted([p for p in df_p[col_proveedor].dropna().unique()
                          if str(p).strip()])
    sel_prov = c1.multiselect("Proveedor", options=proveedores,
                                placeholder="Todos", key=f"{prefijo}_prov")

    compradores = sorted([c for c in df_p[col_comprador].dropna().unique()
                           if str(c).strip()]) if col_comprador in df_p.columns else []
    sel_comp = c2.multiselect("Comprador", options=compradores,
                                placeholder="Todos", key=f"{prefijo}_comp")

    txt_folio = c3.text_input(etiqueta_folio, placeholder="Ej. DC-MZ017",
                                key=f"{prefijo}_folio")

    df_f = df_p.copy()
    if sel_prov:
        df_f = df_f[df_f[col_proveedor].isin(sel_prov)]
    if sel_comp and col_comprador in df_f.columns:
        df_f = df_f[df_f[col_comprador].isin(sel_comp)]
    if txt_folio.strip():
        df_f = df_f[df_f[col_folio].astype(str).str.contains(
            txt_folio.strip(), case=False, na=False)]

    n_activos = sum([bool(sel_meses if modo == "Mes" else False),
                     bool(sel_prov), bool(sel_comp), bool(txt_folio.strip())])
    if n_activos > 0:
        c1, c2 = st.columns([4, 1])
        c1.caption(f"**{len(df_f)}** de **{len(df)}** registros · "
                    f"{n_activos} filtro(s) activo(s)")
        if c2.button("🧹 Limpiar filtros", key=f"{prefijo}_limpiar",
                       use_container_width=True):
            for k in [f"{prefijo}_meses", f"{prefijo}_prov",
                       f"{prefijo}_comp", f"{prefijo}_folio"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()
    else:
        st.caption(f"**{len(df_f)}** registros")

    return df_f, etiqueta_periodo


# =============================================================================
# 8. TABLERO DIRECTIVO
# =============================================================================
#
#   Vista de presentación para juntas. Puede mostrar SOLO Garantías, SOLO NC o
#   AMBOS. Cada sección con sus KPIs y sus gráficas. Los filtros son opcionales
#   arriba (periodo).


def _tarjeta_kpi(col, icono: str, titulo: str, valor: str,
                  subtitulo: str = "", color: str = "#1f4e79") -> None:
    col.markdown(
        f"""<div style='padding:1rem 1.2rem;border-left:5px solid {color};
                       background:#f8f9fa;border-radius:6px;
                       box-shadow:0 1px 3px rgba(0,0,0,0.08);height:100%'>
              <div style='color:#666;font-size:0.85rem;
                          text-transform:uppercase;letter-spacing:0.5px'>
                {icono} {titulo}
              </div>
              <div style='font-size:1.9rem;font-weight:700;color:{color};
                          margin-top:0.35rem;line-height:1.1'>
                {valor}
              </div>
              <div style='color:#888;font-size:0.82rem;margin-top:0.15rem'>
                {subtitulo}
              </div>
            </div>""",
        unsafe_allow_html=True,
    )


def _filtro_periodo_tablero(df: pd.DataFrame, col_fecha: str,
                              prefijo: str) -> tuple:
    """Solo filtro de periodo, para el tablero directivo."""
    cA, cB = st.columns([1, 3])
    modo = cA.radio("Periodo", ["Mes", "Rango de fechas"],
                     horizontal=True, key=f"{prefijo}_t_modo")
    hoy = hoy_mx()
    if modo == "Mes":
        meses_disp = sorted(
            [m for m in df["MES ETIQUETA"].dropna().unique()
             if m != "Sin fecha"], key=_orden_mes)
        placeholder_meses = _rango_meses_placeholder(df)
        sel = cB.multiselect("Mes(es)", options=meses_disp,
                               placeholder=placeholder_meses,
                               key=f"{prefijo}_t_meses")
        if sel:
            return df[df["MES ETIQUETA"].isin(sel)], ", ".join(sel)
        return df, placeholder_meses
    else:
        fmin_val = df[col_fecha].min()
        fmin = _a_fecha(fmin_val) if pd.notna(fmin_val) else hoy - timedelta(days=180)
        rango = cB.date_input("Rango de fechas", value=(fmin, hoy),
                                format="DD/MM/YYYY", key=f"{prefijo}_t_rango")
        if isinstance(rango, tuple) and len(rango) == 2:
            d, h = rango
            df_f = df[(df[col_fecha].dt.date >= d) &
                       (df[col_fecha].dt.date <= h)]
            return df_f, f"{d:%d/%m/%Y} — {h:%d/%m/%Y}"
        return df, "Todo el periodo"


def vista_tablero(datos: dict) -> None:
    st.markdown("#### 📊 Tablero Directivo")
    st.caption("Panorama general para presentación en juntas.")

    seccion = st.radio(
        "Sección a mostrar",
        ["Solo Garantías", "Solo NC pendientes", "Ambos"],
        horizontal=True, key="tab_seccion")

    st.divider()

    if seccion in ("Solo Garantías", "Ambos"):
        _tablero_seccion_garantias(datos)

    if seccion == "Ambos":
        st.divider()

    if seccion in ("Solo NC pendientes", "Ambos"):
        _tablero_seccion_nc(datos)


def _tablero_seccion_garantias(datos: dict) -> None:
    df = datos.get("garantias")
    if df is None or df.empty:
        st.info("Sin datos de garantías.")
        return

    st.markdown("## 📋 Folios de Garantía")
    df_f, etiq = _filtro_periodo_tablero(df, COL_G_FECHA_RECEPCION, "gar")
    st.caption(f"📅 Periodo: **{etiq}** · Los tableros, gráficas e indicadores "
                "excluyen folios en **cuarentena** (aún no se gestionan). "
                "La tabla de detalle mensual al final los conserva para el total.")

    # Excluir CUARENTENA para tableros, gráficas e indicadores de cumplimiento
    df_g = df_f[df_f[COL_G_ESTADO] != ESTADO_CUARENTENA].copy()

    # ---- KPIs ----
    total = len(df_g)
    activos = (df_g[COL_G_ESTADO] == ESTADO_ACTIVO).sum()
    n_cuarentena_excluidos = (df_f[COL_G_ESTADO] == ESTADO_CUARENTENA).sum()
    resueltos = (df_g[COL_G_ESTADO] == ESTADO_RESUELTO).sum()
    cancelados = (df_g[COL_G_ESTADO] == ESTADO_CANCELADO).sum()
    monto_total = df_g[COL_G_IMPORTE].sum()
    monto_activo = df_g.loc[df_g[COL_G_ESTADO] == ESTADO_ACTIVO,
                              COL_G_IMPORTE].sum()
    valido = total - cancelados
    pct_res = (resueltos / valido * 100) if valido > 0 else 0.0
    vencidos = int(df_g.apply(esta_vencido_garantia, axis=1).sum()) if not df_g.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    _tarjeta_kpi(c1, "📁", "Folios en gestión", f"{total:,}",
                 f"{activos} activos · {resueltos} resueltos"
                 + (f" · ({n_cuarentena_excluidos} en cuarentena excluidos)"
                    if n_cuarentena_excluidos else ""),
                 color="#1f4e79")
    _tarjeta_kpi(c2, "💰", "Monto en gestión",
                 _fmt_mxn(monto_total),
                 f"{_fmt_mxn(monto_activo)} por cobrar",
                 color="#0f766e")
    _tarjeta_kpi(c3, "✅", "% Resueltos",
                 f"{pct_res:.1f}%",
                 f"{resueltos:,} de {valido:,} folios",
                 color="#059669")
    _tarjeta_kpi(c4, "🚨", "Vencidos (90 días)",
                 f"{vencidos:,}",
                 MSG_VENCIDO if vencidos > 0 else "Sin vencimientos",
                 color="#dc2626" if vencidos > 0 else "#94a3b8")

    # ---- INDICADOR DE CUMPLIMIENTO por mes ----
    st.markdown("---")
    st.markdown("#### 🎯 Indicador de cumplimiento — Folios de Garantía")
    st.caption(f"**Objetivo:** {META_CUMPLIMIENTO_GARANTIA:.0f}% de folios cerrados "
                f"dentro de los {DIAS_PLAZO_TOTAL} días · "
                f"🔴 Alerta si algún folio abierto rebasa **{DIAS_ALERTA_ROJA} días** · "
                "⚠️ No pueden pasar **3 meses consecutivos** sin cerrar folios.")
    _tabla_cumplimiento_garantias(df_g)

    # ---- Gráficas (2 columnas x 2 filas para aprovechar el ancho) ----
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🏭 Top 10 proveedores por monto")
        _grafica_top_prov_garantias(df_g)
        st.markdown("#### 👥 Distribución por comprador")
        _grafica_por_comprador(df_g)
    with c2:
        st.markdown("#### 📅 Folios por mes y estado")
        _grafica_evolucion_mensual(df_g)
        st.markdown("#### 🚨 Vencidos por mes de recepción")
        _grafica_vencidos_por_mes(df_g)

    # ---- Tabla mensual detallada (SIN filtros — usa df completo, incluye no gestionados) ----
    st.markdown("---")
    st.markdown("#### 📋 Detalle mensual (base completa, sin filtros)")
    st.caption("Folios y monto por mes, agrupados por estado. Incluye TODOS "
                "los folios (aun los no gestionados). La fila **TOTAL** acumula "
                "todas las columnas.")
    _tabla_detalle_mensual(datos["garantias"])


def _tabla_cumplimiento_garantias(df_g: pd.DataFrame) -> None:
    """Tabla de cumplimiento mensual para folios de garantía.

    Por cada mes de recepción calcula:
      - Total de folios (en gestión)
      - Resueltos dentro de 90 días
      - % cumplimiento
      - Semáforo vs META_CUMPLIMIENTO_GARANTIA
      - Folios abiertos > DIAS_ALERTA_ROJA días (penalización)
    """
    if df_g is None or df_g.empty:
        st.info("Sin folios en gestión para calcular cumplimiento.")
        return
    d = df_g[df_g["MES ETIQUETA"] != "Sin fecha"].copy()
    if d.empty:
        st.info("Sin folios con mes válido.")
        return

    filas = []
    for mes, grupo in d.groupby("MES ETIQUETA"):
        total = len(grupo)
        # Resueltos dentro del plazo: FECHA RESUELTO - FECHA RECEPCION <= 90
        resueltos = grupo[grupo[COL_G_ESTADO] == ESTADO_RESUELTO].copy()
        cerrados_a_tiempo = 0
        for _, r in resueltos.iterrows():
            ini = _a_fecha(r.get(COL_G_FECHA_RECEPCION))
            fin = _a_fecha(r.get(COL_G_FECHA_RESUELTO))
            if ini and fin and (fin - ini).days <= DIAS_PLAZO_TOTAL:
                cerrados_a_tiempo += 1
        pct = (cerrados_a_tiempo / total * 100) if total > 0 else 0.0
        # Folios que rebasan la penalización
        penalizados = 0
        for _, r in grupo.iterrows():
            if r[COL_G_ESTADO] in (ESTADO_ACTIVO, ESTADO_CUARENTENA):
                d_transc = dias_transcurridos_garantia(r)
                if d_transc is not None and d_transc > DIAS_ALERTA_ROJA:
                    penalizados += 1
        estado_semaforo = ("🟢" if pct >= META_CUMPLIMIENTO_GARANTIA
                           else "🟡" if pct >= META_CUMPLIMIENTO_GARANTIA - 10
                           else "🔴")
        if penalizados > 0:
            estado_semaforo = "🔴"
        filas.append({
            "Mes": mes,
            "_orden": _orden_mes(mes),
            "Estado": estado_semaforo,
            "Folios totales": total,
            "Cerrados a tiempo (≤90 d)": cerrados_a_tiempo,
            "% Cumplimiento": round(pct, 1),
            f"Penalizados (>{DIAS_ALERTA_ROJA} d)": penalizados,
        })
    tabla = pd.DataFrame(filas).sort_values("_orden").drop(columns="_orden")

    # Detectar 3 meses consecutivos sin cerrar
    meses_ord = tabla["Mes"].tolist()
    cerrados_por_mes = dict(zip(tabla["Mes"], tabla["Cerrados a tiempo (≤90 d)"]))
    for i in range(len(meses_ord) - 2):
        tres = [cerrados_por_mes[meses_ord[i + k]] for k in range(3)]
        if all(v == 0 for v in tres):
            st.error(f"⚠️ **3 meses consecutivos sin cerrar folios**: "
                     f"{meses_ord[i]}, {meses_ord[i+1]}, {meses_ord[i+2]}.")

    # Totales
    total_folios = tabla["Folios totales"].sum()
    total_cerrados = tabla["Cerrados a tiempo (≤90 d)"].sum()
    total_penal = tabla[f"Penalizados (>{DIAS_ALERTA_ROJA} d)"].sum()
    pct_global = (total_cerrados / total_folios * 100) if total_folios > 0 else 0.0

    fila_total = pd.DataFrame([{
        "Mes": "TOTAL",
        "Estado": ("🟢" if pct_global >= META_CUMPLIMIENTO_GARANTIA
                    else "🟡" if pct_global >= META_CUMPLIMIENTO_GARANTIA - 10
                    else "🔴"),
        "Folios totales": total_folios,
        "Cerrados a tiempo (≤90 d)": total_cerrados,
        "% Cumplimiento": round(pct_global, 1),
        f"Penalizados (>{DIAS_ALERTA_ROJA} d)": total_penal,
    }])
    tabla_f = pd.concat([tabla, fila_total], ignore_index=True).set_index("Mes")

    st.dataframe(
        tabla_f, use_container_width=True,
        column_config={
            "Estado": st.column_config.TextColumn("🚦", width="small"),
            "Folios totales": st.column_config.NumberColumn(format="%d"),
            "Cerrados a tiempo (≤90 d)": st.column_config.NumberColumn(format="%d"),
            "% Cumplimiento": st.column_config.NumberColumn(
                "% Cumplimiento", format="%.1f%%",
                help=f"Meta: {META_CUMPLIMIENTO_GARANTIA:.0f}%"),
            f"Penalizados (>{DIAS_ALERTA_ROJA} d)":
                st.column_config.NumberColumn(format="%d"),
        },
    )


def _tabla_detalle_mensual(df_completo: pd.DataFrame) -> None:
    """Tabla mensual con folios y monto por estado. Ignora filtros."""
    if df_completo is None or df_completo.empty:
        st.info("Sin datos.")
        return
    d = df_completo.copy()
    d = d[d["MES ETIQUETA"] != "Sin fecha"]
    if d.empty:
        st.info("Sin datos con fecha de recepción.")
        return

    # Construir por mes
    filas = []
    for mes, grupo in d.groupby("MES ETIQUETA"):
        act = grupo[grupo[COL_G_ESTADO] == ESTADO_ACTIVO]
        cua = grupo[grupo[COL_G_ESTADO] == ESTADO_CUARENTENA]
        res = grupo[grupo[COL_G_ESTADO] == ESTADO_RESUELTO]
        f_act, f_cua, f_res = len(act), len(cua), len(res)
        m_act = act[COL_G_IMPORTE].sum()
        m_cua = cua[COL_G_IMPORTE].sum()
        m_res = res[COL_G_IMPORTE].sum()
        filas.append({
            "Mes": mes,
            "_orden": _orden_mes(mes),
            "Folios activos": f_act,
            "Folios cuarentena": f_cua,
            "Folios resueltos": f_res,
            "Folios TOTAL": f_act + f_cua + f_res,
            "Monto activo": m_act,
            "Monto cuarentena": m_cua,
            "Monto resuelto": m_res,
            "Monto TOTAL": m_act + m_cua + m_res,
        })
    tabla = pd.DataFrame(filas).sort_values("_orden").drop(columns="_orden")
    tabla = tabla.set_index("Mes")

    # Fila TOTAL al final
    total = tabla.sum(numeric_only=True)
    total.name = "TOTAL"
    tabla = pd.concat([tabla, total.to_frame().T])

    st.dataframe(
        tabla, use_container_width=True,
        column_config={
            "Folios activos": st.column_config.NumberColumn(
                "Folios activos", format="%d"),
            "Folios cuarentena": st.column_config.NumberColumn(
                "Folios en cuarentena", format="%d"),
            "Folios resueltos": st.column_config.NumberColumn(
                "Folios resueltos", format="%d"),
            "Folios TOTAL": st.column_config.NumberColumn(
                "Folios TOTAL", format="%d"),
            "Monto activo": st.column_config.NumberColumn(
                "Monto activo (por cobrar)", format="$%.2f"),
            "Monto cuarentena": st.column_config.NumberColumn(
                "Monto en cuarentena", format="$%.2f"),
            "Monto resuelto": st.column_config.NumberColumn(
                "Monto resuelto (cobrado)", format="$%.2f"),
            "Monto TOTAL": st.column_config.NumberColumn(
                "Monto TOTAL", format="$%.2f"),
        },
    )


def _tablero_seccion_nc(datos: dict) -> None:
    df = datos.get("nc")
    if df is None or df.empty:
        st.info("Sin datos de notas de crédito.")
        return

    st.markdown("## 💳 Notas de Crédito Pendientes")
    df_f, etiq = _filtro_periodo_tablero(df, COL_NC_FECHA_REPORTE, "nc")
    st.caption(f"📅 Periodo: **{etiq}**")

    df_nc = df_f.copy()

    pend = df_nc[df_nc[COL_NC_ESTADO] == "Pendiente"]
    res = df_nc[df_nc[COL_NC_ESTADO] == "Resuelto"]

    n_pend = len(pend)
    monto_pend = pend[COL_NC_IMP_PENDIENTE].sum()
    n_res = len(res)
    monto_res = res[COL_NC_IMP_PENDIENTE].sum()
    n_venc = int(pend.apply(esta_vencida_nc, axis=1).sum()) if not pend.empty else 0
    prov_pend = pend[COL_NC_PROVEEDOR].nunique()

    c1, c2, c3, c4 = st.columns(4)
    _tarjeta_kpi(c1, "⏳", "NC pendientes", f"{n_pend:,}",
                 f"{_fmt_mxn(monto_pend)} por conseguir",
                 color="#ea580c")
    _tarjeta_kpi(c2, "✅", "NC ya conseguidas", f"{n_res:,}",
                 f"{_fmt_mxn(monto_res)} entregadas",
                 color="#059669")
    _tarjeta_kpi(c3, "🚨", f"Vencidas (>{DIAS_PLAZO_NC} días)",
                 f"{n_venc:,}",
                 "Requieren atención inmediata" if n_venc > 0 else "Sin vencidas",
                 color="#dc2626" if n_venc > 0 else "#94a3b8")
    _tarjeta_kpi(c4, "🏭", "Proveedores con NC pend.",
                 f"{prov_pend:,}",
                 f"con {n_pend} nota(s) por conseguir",
                 color="#1f4e79")

    # ---- INDICADOR DE CUMPLIMIENTO NC por mes ----
    st.markdown("---")
    st.markdown("#### 🎯 Indicador de cumplimiento — Notas de Crédito")
    st.caption(f"**Objetivo:** {META_CUMPLIMIENTO_NC:.0f}% de NC cerradas dentro "
                f"de los {DIAS_PLAZO_NC} días desde la fecha de reporte · "
                f"🔴 Alerta si alguna NC abierta rebasa **{DIAS_ALERTA_NC} días**.")
    _tabla_cumplimiento_nc(df_nc)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🏭 Top 10 proveedores por monto pendiente")
        _grafica_top_prov_nc(pend)
    with c2:
        st.markdown("#### 📅 NC por mes de recepción")
        _grafica_nc_por_mes(df_nc)

    # ---- Tabla mensual detallada de NC (SIN filtros — incluye no gestionadas) ----
    st.markdown("---")
    st.markdown("#### 📋 Detalle mensual de NC (base completa)")
    st.caption("NC por mes agrupadas por estado. Incluye TODAS las NC (aun las "
                "no gestionadas). La fila **TOTAL** acumula todas las columnas.")
    _tabla_detalle_mensual_nc(datos["nc"])


def _tabla_cumplimiento_nc(df_nc: pd.DataFrame) -> None:
    """Tabla de cumplimiento mensual para notas de crédito."""
    if df_nc is None or df_nc.empty:
        st.info("Sin NC en gestión para calcular cumplimiento.")
        return
    d = df_nc[df_nc["MES ETIQUETA"] != "Sin fecha"].copy()
    if d.empty:
        st.info("Sin NC con mes válido.")
        return

    filas = []
    for mes, grupo in d.groupby("MES ETIQUETA"):
        total = len(grupo)
        # Cerradas a tiempo: FECHA NC - FECHA REPORTE <= 20 días
        resueltas = grupo[grupo[COL_NC_ESTADO] == "Resuelto"]
        cerradas_a_tiempo = 0
        for _, r in resueltas.iterrows():
            ini = _a_fecha(r.get(COL_NC_FECHA_REPORTE))
            fin = _a_fecha(r.get(COL_NC_FECHA_NC))
            if ini and fin and (fin - ini).days <= DIAS_PLAZO_NC:
                cerradas_a_tiempo += 1
        pct = (cerradas_a_tiempo / total * 100) if total > 0 else 0.0
        # NC que rebasan la penalización
        penalizadas = 0
        for _, r in grupo.iterrows():
            if r[COL_NC_ESTADO] == "Pendiente":
                d_transc = dias_transcurridos_nc(r)
                if d_transc is not None and d_transc > DIAS_ALERTA_NC:
                    penalizadas += 1
        semaforo = ("🟢" if pct >= META_CUMPLIMIENTO_NC
                     else "🟡" if pct >= META_CUMPLIMIENTO_NC - 10
                     else "🔴")
        if penalizadas > 0:
            semaforo = "🔴"
        filas.append({
            "Mes": mes,
            "_orden": _orden_mes(mes),
            "Estado": semaforo,
            "NC totales": total,
            "Cerradas a tiempo (≤20 d)": cerradas_a_tiempo,
            "% Cumplimiento": round(pct, 1),
            f"Penalizadas (>{DIAS_ALERTA_NC} d)": penalizadas,
        })
    tabla = pd.DataFrame(filas).sort_values("_orden").drop(columns="_orden")

    # Totales
    total_nc = tabla["NC totales"].sum()
    total_cerr = tabla["Cerradas a tiempo (≤20 d)"].sum()
    total_pen = tabla[f"Penalizadas (>{DIAS_ALERTA_NC} d)"].sum()
    pct_g = (total_cerr / total_nc * 100) if total_nc > 0 else 0.0
    fila_total = pd.DataFrame([{
        "Mes": "TOTAL",
        "Estado": ("🟢" if pct_g >= META_CUMPLIMIENTO_NC
                    else "🟡" if pct_g >= META_CUMPLIMIENTO_NC - 10
                    else "🔴"),
        "NC totales": total_nc,
        "Cerradas a tiempo (≤20 d)": total_cerr,
        "% Cumplimiento": round(pct_g, 1),
        f"Penalizadas (>{DIAS_ALERTA_NC} d)": total_pen,
    }])
    tabla_f = pd.concat([tabla, fila_total], ignore_index=True).set_index("Mes")

    st.dataframe(
        tabla_f, use_container_width=True,
        column_config={
            "Estado": st.column_config.TextColumn("🚦", width="small"),
            "NC totales": st.column_config.NumberColumn(format="%d"),
            "Cerradas a tiempo (≤20 d)": st.column_config.NumberColumn(format="%d"),
            "% Cumplimiento": st.column_config.NumberColumn(
                "% Cumplimiento", format="%.1f%%",
                help=f"Meta: {META_CUMPLIMIENTO_NC:.0f}%"),
            f"Penalizadas (>{DIAS_ALERTA_NC} d)":
                st.column_config.NumberColumn(format="%d"),
        },
    )


def _tabla_detalle_mensual_nc(df_completo: pd.DataFrame) -> None:
    """Tabla mensual NC con conteo y monto por estado. Ignora filtros."""
    if df_completo is None or df_completo.empty:
        st.info("Sin datos.")
        return
    d = df_completo.copy()
    d = d[d["MES ETIQUETA"] != "Sin fecha"]
    if d.empty:
        st.info("Sin NC con fecha de reporte.")
        return
    filas = []
    for mes, grupo in d.groupby("MES ETIQUETA"):
        pend = grupo[grupo[COL_NC_ESTADO] == "Pendiente"]
        res = grupo[grupo[COL_NC_ESTADO] == "Resuelto"]
        n_p, n_r = len(pend), len(res)
        m_p = pend[COL_NC_IMP_PENDIENTE].sum()
        m_r = res[COL_NC_IMP_PENDIENTE].sum()
        filas.append({
            "Mes": mes,
            "_orden": _orden_mes(mes),
            "NC pendientes": n_p,
            "NC resueltas": n_r,
            "NC TOTAL": n_p + n_r,
            "Monto pendiente": m_p,
            "Monto resuelto": m_r,
            "Monto TOTAL": m_p + m_r,
        })
    tabla = pd.DataFrame(filas).sort_values("_orden").drop(columns="_orden")
    tabla = tabla.set_index("Mes")
    total = tabla.sum(numeric_only=True)
    total.name = "TOTAL"
    tabla = pd.concat([tabla, total.to_frame().T])
    st.dataframe(
        tabla, use_container_width=True,
        column_config={
            "NC pendientes": st.column_config.NumberColumn(format="%d"),
            "NC resueltas": st.column_config.NumberColumn(format="%d"),
            "NC TOTAL": st.column_config.NumberColumn(format="%d"),
            "Monto pendiente": st.column_config.NumberColumn(
                "Monto pendiente (por cobrar)", format="$%.2f"),
            "Monto resuelto": st.column_config.NumberColumn(
                "Monto resuelto (cobrado)", format="$%.2f"),
            "Monto TOTAL": st.column_config.NumberColumn(format="$%.2f"),
        },
    )


def _grafica_top_prov_garantias(df: pd.DataFrame) -> None:
    """Top 10 proveedores por monto (folios activos + cuarentena)."""
    d = df[df[COL_G_ESTADO].isin([ESTADO_ACTIVO, ESTADO_CUARENTENA])]
    if d.empty:
        st.info("Sin folios vigentes para graficar.")
        return
    top = (d.groupby(COL_G_PROVEEDOR)
           .agg(Monto=(COL_G_IMPORTE, "sum"),
                Folios=(COL_G_FOLIO, "count"))
           .sort_values("Monto", ascending=False).head(10))
    _dibujar_barras_prov(top, "#1f4e79")


def _grafica_top_prov_nc(df: pd.DataFrame) -> None:
    if df.empty:
        st.info("Sin NC pendientes para graficar.")
        return
    top = (df.groupby(COL_NC_PROVEEDOR)
           .agg(Monto=(COL_NC_IMP_PENDIENTE, "sum"),
                Folios=(COL_NC_ENTRADA, "count"))
           .sort_values("Monto", ascending=False).head(10))
    _dibujar_barras_prov(top, "#ea580c")


def _dibujar_barras_prov(top: pd.DataFrame, color: str) -> None:
    if not ALTAIR_OK:
        st.dataframe(top, use_container_width=True)
        return
    d = top.reset_index()
    d.columns = ["Proveedor", "Monto", "Folios"]
    d["_txt"] = d.apply(lambda r: f"${r['Monto']:,.2f} ({int(r['Folios'])})",
                          axis=1)
    base = alt.Chart(d)
    barras = base.mark_bar(color=color).encode(
        y=alt.Y("Proveedor:N", title=None, sort="-x"),
        x=alt.X("Monto:Q", title="Monto (MXN)",
                axis=alt.Axis(format="$,.0f")),
        tooltip=[
            alt.Tooltip("Proveedor:N"),
            alt.Tooltip("Monto:Q", format="$,.2f"),
            alt.Tooltip("Folios:Q", format=",d"),
        ],
    )
    etq = base.mark_text(align="left", dx=4, fontSize=11, fontWeight="bold",
                          color="#333").encode(
        y=alt.Y("Proveedor:N", sort="-x"),
        x=alt.X("Monto:Q"),
        text=alt.Text("_txt:N"))
    st.altair_chart((barras + etq).properties(height=max(280, 30 * len(d))),
                     use_container_width=True)


def _grafica_evolucion_mensual(df: pd.DataFrame) -> None:
    d = df[df["MES ETIQUETA"] != "Sin fecha"]
    if d.empty:
        st.info("Sin datos por mes.")
        return
    tabla = (d.groupby(["MES ETIQUETA", COL_G_ESTADO])[COL_G_FOLIO]
             .count().unstack(fill_value=0))
    tabla = tabla.reindex(sorted(tabla.index, key=_orden_mes))
    if not ALTAIR_OK:
        st.bar_chart(tabla, height=320)
        return
    largo = tabla.reset_index().melt(id_vars="MES ETIQUETA",
                                      var_name="Estado", value_name="Folios")
    largo = largo[largo["Folios"] > 0]
    orden = [ESTADO_ACTIVO, ESTADO_CUARENTENA, ESTADO_RESUELTO,
             ESTADO_CANCELADO]
    colores = ["#f59e0b", "#3b82f6", "#10b981", "#94a3b8"]
    g = alt.Chart(largo).mark_bar().encode(
        x=alt.X("MES ETIQUETA:N", title="Mes", sort=list(tabla.index),
                axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("Folios:Q", title="Folios"),
        color=alt.Color("Estado:N", sort=orden,
                         scale=alt.Scale(domain=orden, range=colores),
                         legend=alt.Legend(orient="bottom")),
        tooltip=["MES ETIQUETA", "Estado", "Folios"],
    ).properties(height=320)
    st.altair_chart(g, use_container_width=True)


def _grafica_por_comprador(df: pd.DataFrame) -> None:
    d = df[df[COL_G_ESTADO].isin([ESTADO_ACTIVO, ESTADO_CUARENTENA])]
    if d.empty:
        st.info("Sin folios vigentes para graficar por comprador.")
        return
    d = d.copy()
    d["_comp"] = d[COL_G_COMPRADOR].replace("", "Sin comprador")
    agr = (d.groupby("_comp")
           .agg(Monto=(COL_G_IMPORTE, "sum"),
                Folios=(COL_G_FOLIO, "count"))
           .sort_values("Monto", ascending=False))
    if not ALTAIR_OK:
        st.dataframe(agr, use_container_width=True)
        return
    dd = agr.reset_index()
    dd.columns = ["Comprador", "Monto", "Folios"]
    dd["_txt"] = dd.apply(lambda r: f"${r['Monto']:,.2f} ({int(r['Folios'])})",
                            axis=1)
    base = alt.Chart(dd)
    barras = base.mark_bar(color="#7c3aed").encode(
        y=alt.Y("Comprador:N", title=None, sort="-x"),
        x=alt.X("Monto:Q", title="Monto (MXN)",
                axis=alt.Axis(format="$,.0f")),
        tooltip=[alt.Tooltip("Comprador:N"),
                 alt.Tooltip("Monto:Q", format="$,.2f"),
                 alt.Tooltip("Folios:Q", format=",d")],
    )
    etq = base.mark_text(align="left", dx=4, fontSize=11, fontWeight="bold",
                          color="#333").encode(
        y=alt.Y("Comprador:N", sort="-x"),
        x=alt.X("Monto:Q"),
        text=alt.Text("_txt:N"))
    st.altair_chart((barras + etq).properties(height=max(280, 30 * len(dd))),
                     use_container_width=True)


def _grafica_vencidos_por_mes(df: pd.DataFrame) -> None:
    """Distribución mensual de folios vencidos (>90 días desde recepción)."""
    if df.empty:
        st.info("Sin datos.")
        return
    d = df.copy()
    d["_venc"] = d.apply(esta_vencido_garantia, axis=1)
    venc = d[d["_venc"]]
    if venc.empty:
        st.success("✅ No hay folios vencidos en este periodo.")
        return
    venc = venc[venc["MES ETIQUETA"] != "Sin fecha"]
    if venc.empty:
        return
    agr = (venc.groupby("MES ETIQUETA")
           .agg(Folios=(COL_G_FOLIO, "count"),
                Monto=(COL_G_IMPORTE, "sum")))
    agr = agr.reindex(sorted(agr.index, key=_orden_mes))
    if not ALTAIR_OK:
        st.bar_chart(agr[["Folios"]], height=280)
        return
    d2 = agr.reset_index()
    d2["_txt"] = d2.apply(
        lambda r: f"{int(r['Folios'])} · ${r['Monto']:,.0f}", axis=1)
    base = alt.Chart(d2)
    barras = base.mark_bar(color="#dc2626").encode(
        x=alt.X("MES ETIQUETA:N", title="Mes de recepción",
                sort=list(agr.index),
                axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("Folios:Q", title="Folios vencidos"),
        tooltip=[
            alt.Tooltip("MES ETIQUETA:N", title="Mes"),
            alt.Tooltip("Folios:Q", format=",d"),
            alt.Tooltip("Monto:Q", title="Monto", format="$,.2f"),
        ],
    )
    etq = base.mark_text(align="center", dy=-8, fontSize=11,
                          fontWeight="bold", color="#dc2626").encode(
        x=alt.X("MES ETIQUETA:N", sort=list(agr.index)),
        y=alt.Y("Folios:Q"),
        text=alt.Text("_txt:N"))
    st.altair_chart((barras + etq).properties(height=280),
                     use_container_width=True)


def _grafica_nc_por_mes(df: pd.DataFrame) -> None:
    d = df[df["MES ETIQUETA"] != "Sin fecha"]
    if d.empty:
        st.info("Sin datos por mes.")
        return
    tabla = (d.groupby(["MES ETIQUETA", COL_NC_ESTADO])[COL_NC_ENTRADA]
             .count().unstack(fill_value=0))
    tabla = tabla.reindex(sorted(tabla.index, key=_orden_mes))
    if not ALTAIR_OK:
        st.bar_chart(tabla, height=320)
        return
    largo = tabla.reset_index().melt(id_vars="MES ETIQUETA",
                                      var_name="Estado", value_name="NC")
    largo = largo[largo["NC"] > 0]
    orden = ["Pendiente", "Resuelto", "Cancelado"]
    colores = ["#ea580c", "#10b981", "#94a3b8"]
    g = alt.Chart(largo).mark_bar().encode(
        x=alt.X("MES ETIQUETA:N", title="Mes", sort=list(tabla.index),
                axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("NC:Q", title="Notas de crédito"),
        color=alt.Color("Estado:N", sort=orden,
                         scale=alt.Scale(domain=orden, range=colores),
                         legend=alt.Legend(orient="bottom")),
        tooltip=["MES ETIQUETA", "Estado", "NC"],
    ).properties(height=320)
    st.altair_chart(g, use_container_width=True)


# =============================================================================
# 9. FOLIOS DE GARANTÍA
# =============================================================================


def vista_garantias(datos: dict) -> None:
    st.markdown("#### 📋 Folios de Garantía")
    st.caption("Lista completa de folios. Plazo total: **90 días** desde la "
                "fecha de recepción del reporte. Un folio queda RESUELTO al "
                "capturar su nota de crédito.")

    df = datos.get("garantias")
    if df is None or df.empty:
        st.info("No hay folios de garantía cargados.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    # ---- Filtros adicionales: por estado y por etapa del flujo ----
    c_e1, c_e2 = st.columns(2)
    estados_disp = ["Todos", ESTADO_ACTIVO, ESTADO_CUARENTENA,
                    ESTADO_RESUELTO, ESTADO_CANCELADO]
    f_estado = c_e1.selectbox("Estado", estados_disp, key="g_estado")

    etapas_disp = ["Todas"] + ETAPAS_GARANTIA
    f_etapa = c_e2.selectbox("Etapa del flujo", etapas_disp, key="g_etapa",
                              help="Filtra por el estatus intermedio del proceso.")

    df_e = df if f_estado == "Todos" else df[df[COL_G_ESTADO] == f_estado]
    if f_etapa != "Todas":
        df_e = df_e[df_e[COL_G_ETAPA] == f_etapa]

    # ---- Filtros comunes ----
    df_f, etiq = _filtros_barra(
        df_e, prefijo="g",
        col_fecha=COL_G_FECHA_RECEPCION,
        col_folio=COL_G_FOLIO,
        col_proveedor=COL_G_PROVEEDOR,
        col_comprador=COL_G_COMPRADOR,
        etiqueta_folio="Buscar folio reporte")

    # ---- Alerta al inicio ----
    if not df_f.empty:
        venc = df_f[df_f.apply(esta_vencido_garantia, axis=1)]
        if not venc.empty:
            st.error(f"🚨 **{len(venc)}** folio(s) vencido(s): {MSG_VENCIDO}")

    if df_f.empty:
        st.info("No hay folios con los filtros actuales.")
        return

    # ---- Tabla ----
    vista = df_f.copy()
    vista["🚦"] = vista.apply(lambda f: semaforo_garantia(f)[0], axis=1)
    vista["Días transc."] = vista.apply(dias_transcurridos_garantia, axis=1)
    vista["Días restantes"] = vista.apply(dias_restantes_garantia, axis=1)
    vista["Vence"] = vista.apply(fecha_vencimiento_garantia, axis=1)

    cols = ["🚦", COL_G_FOLIO, COL_G_PROVEEDOR, COL_G_COMPRADOR,
            COL_G_IMPORTE, COL_G_FECHA_RECEPCION, "Vence",
            "Días transc.", "Días restantes", COL_G_ESTADO, COL_G_ETAPA,
            COL_G_FOLIO_DEV, COL_G_FOLIO_AJUSTE, COL_G_NOTA_CREDITO]
    st.dataframe(
        vista[cols], use_container_width=True, hide_index=True,
        column_config={
            COL_G_IMPORTE: st.column_config.NumberColumn(
                "Importe", format="$%.2f"),
            COL_G_FECHA_RECEPCION: st.column_config.DateColumn(
                "Recepción", format="DD/MM/YYYY"),
            "Vence": st.column_config.DateColumn(format="DD/MM/YYYY"),
            COL_G_ETAPA: "Etapa del flujo",
            COL_G_FOLIO_DEV: "F. Devolución",
            COL_G_FOLIO_AJUSTE: "F. Ajuste",
            COL_G_NOTA_CREDITO: "Nota Crédito",
        },
    )

    st.divider()
    st.markdown("#### ✏️ Editar folio")

    df_sel = df_f.copy()
    df_sel["etiqueta"] = (df_sel[COL_G_FOLIO] + " · "
                           + df_sel[COL_G_PROVEEDOR].str.slice(0, 40)
                           + " · [" + df_sel[COL_G_ESTADO] + "]")
    sel = st.selectbox("Selecciona un folio",
                        options=df_sel["etiqueta"].tolist(),
                        key="g_edit_sel")
    fila = df_sel[df_sel["etiqueta"] == sel].iloc[0]
    _editor_garantia(fila, datos)


def _editor_garantia(fila: pd.Series, datos: dict) -> None:
    icono, texto_sem = semaforo_garantia(fila)
    st.markdown(
        f"""<div style='display:flex;gap:1.2rem;flex-wrap:wrap;
                        padding:0.5rem 0.8rem;margin:0.3rem 0 0.8rem 0;
                        border:1px solid #e6e6e6;border-radius:6px;
                        font-size:0.85rem;background:#fafafa'>
          <span><b>Folio:</b> {fila[COL_G_FOLIO]}</span>
          <span><b>Proveedor:</b> {fila[COL_G_PROVEEDOR]}</span>
          <span><b>Importe:</b> {_fmt_mxn(fila[COL_G_IMPORTE])}</span>
          <span><b>Comprador:</b> {fila[COL_G_COMPRADOR] or '—'}</span>
          <span><b>Recepción:</b> {_fmt_fecha(fila[COL_G_FECHA_RECEPCION])}</span>
          <span><b>Estado:</b> {icono} {texto_sem}</span>
        </div>""",
        unsafe_allow_html=True,
    )

    if esta_vencido_garantia(fila):
        st.error(f"🚨 **{MSG_VENCIDO}**")

    estado = str(fila[COL_G_ESTADO]).strip()

    if estado == ESTADO_CUARENTENA:
        st.info("Este folio está en cuarentena. Gestionalo desde la pestaña "
                 "**🧊 Cuarentena**.")
        return

    if estado == ESTADO_CANCELADO:
        st.warning("Este folio fue cancelado y no se contabiliza en indicadores.")
        if st.button("↩️ Reactivar folio", key=f"react_{fila[COL_G_FOLIO]}"):
            _actualizar_garantia(fila[COL_G_FOLIO], datos,
                                  {COL_G_ESTADO: ESTADO_ACTIVO},
                                  "Folio reactivado")
        return

    with st.form(f"form_g_{fila[COL_G_FOLIO]}"):
        c1, c2, c3 = st.columns(3)
        folio_dev = c1.text_input(
            "Folio de devolución",
            value=str(fila.get(COL_G_FOLIO_DEV, "") or ""))
        folio_aj = c2.text_input(
            "Folio de ajuste",
            value=str(fila.get(COL_G_FOLIO_AJUSTE, "") or ""))
        nota_cred = c3.text_input(
            "Nota de crédito",
            value=str(fila.get(COL_G_NOTA_CREDITO, "") or ""),
            help="Al capturar la nota de crédito, el folio queda RESUELTO.")

        opciones_resp = ["", "Recolección", "Destrucción", "Sin respuesta"]
        actual_resp = str(fila.get(COL_G_RESPUESTA, "") or "")
        idx_resp = opciones_resp.index(actual_resp) if actual_resp in opciones_resp else 0

        c_r1, c_r2 = st.columns(2)
        respuesta = c_r1.selectbox(
            "Respuesta del proveedor", options=opciones_resp,
            index=idx_resp)

        # Selector de ETAPA (estatus intermedio del flujo)
        actual_etapa = str(fila.get(COL_G_ETAPA, "") or "")
        idx_etapa = (ETAPAS_GARANTIA.index(actual_etapa)
                     if actual_etapa in ETAPAS_GARANTIA else 0)
        etapa = c_r2.selectbox(
            "Etapa del flujo", options=ETAPAS_GARANTIA,
            index=idx_etapa,
            help="Marca en qué punto del proceso va este folio.")

        notas = st.text_area("Notas / acciones",
                              value=str(fila.get(COL_G_NOTAS, "") or ""),
                              height=100)

        c1, c2, c3 = st.columns([2, 1, 1])
        guardar = c1.form_submit_button("💾 Guardar cambios",
                                          use_container_width=True,
                                          type="primary")
        cancelar = c2.form_submit_button("⛔ Cancelar folio",
                                           use_container_width=True)
        recibir = c3.form_submit_button(
            "✅ Marcar como resuelto",
            use_container_width=True,
            help="Solo si ya se capturó la nota de crédito.")

    if guardar:
        cambios = {
            COL_G_FOLIO_DEV: folio_dev.strip(),
            COL_G_FOLIO_AJUSTE: folio_aj.strip(),
            COL_G_NOTA_CREDITO: nota_cred.strip(),
            COL_G_RESPUESTA: respuesta,
            COL_G_ETAPA: etapa,
            COL_G_NOTAS: notas.strip(),
        }
        if nota_cred.strip():
            cambios[COL_G_ESTADO] = ESTADO_RESUELTO
            cambios[COL_G_FECHA_RESUELTO] = pd.Timestamp(hoy_mx())
            msg = "Folio marcado como RESUELTO (con NC capturada)"
        else:
            msg = "Cambios guardados"
        _actualizar_garantia(fila[COL_G_FOLIO], datos, cambios, msg)

    if cancelar:
        _actualizar_garantia(fila[COL_G_FOLIO], datos,
                              {COL_G_ESTADO: ESTADO_CANCELADO},
                              "Folio cancelado")

    if recibir:
        if not nota_cred.strip():
            st.warning("⚠️ Captura primero el número de nota de crédito.")
        else:
            _actualizar_garantia(fila[COL_G_FOLIO], datos, {
                COL_G_NOTA_CREDITO: nota_cred.strip(),
                COL_G_ESTADO: ESTADO_RESUELTO,
                COL_G_FECHA_RESUELTO: pd.Timestamp(hoy_mx()),
            }, "Folio marcado como RESUELTO")


def _actualizar_garantia(folio: str, datos: dict, cambios: dict,
                          mensaje: str) -> None:
    df = datos["garantias"]
    m = df[COL_G_FOLIO] == folio
    for col, val in cambios.items():
        df.loc[m, col] = val
    df.loc[m, COL_G_MODIFICADO] = f"{ahora_mx():%d/%m/%Y %H:%M}"
    persistir(datos, f"Garantía {folio}: {mensaje}", mensaje)


# =============================================================================
# 10. CUARENTENA
# =============================================================================
#
#   Pestaña dedicada. Los folios ≤ $300 entran automáticamente y se acumulan
#   por proveedor. Al vencer el plazo se liberan solos y ahí arrancan los 90
#   días. Se puede liberar manualmente sin clave.


def vista_cuarentena(datos: dict) -> None:
    st.markdown("#### 🧊 Cuarentena")
    st.caption(f"Folios con importe ≤ **${UMBRAL_CUARENTENA:,.0f}** en espera "
                f"de acumular monto por proveedor. Plazo: {DIAS_CUARENTENA} "
                "días. Al vencer se liberan solos y ahí comienzan los 90 días.")

    df = datos.get("garantias")
    if df is None:
        st.info("No hay folios cargados.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    en_cuar = df[df[COL_G_ESTADO] == ESTADO_CUARENTENA].copy()

    if en_cuar.empty:
        st.success("✅ No hay folios en cuarentena en este momento.")
        return

    # ---- Filtros ----
    df_f, etiq = _filtros_barra(
        en_cuar, prefijo="q",
        col_fecha=COL_G_CUAR_INICIO,
        col_folio=COL_G_FOLIO,
        col_proveedor=COL_G_PROVEEDOR,
        col_comprador=COL_G_COMPRADOR,
        etiqueta_folio="Buscar folio")

    # ---- KPIs ----
    c1, c2, c3 = st.columns(3)
    _tarjeta_kpi(c1, "🧊", "Folios en cuarentena",
                 f"{len(df_f):,}",
                 f"de {len(en_cuar):,} totales",
                 color="#3b82f6")
    _tarjeta_kpi(c2, "💰", "Monto acumulado",
                 _fmt_mxn(df_f[COL_G_IMPORTE].sum()),
                 f"promedio {_fmt_mxn(df_f[COL_G_IMPORTE].mean() if len(df_f) else 0)}",
                 color="#0f766e")
    prov_listos = 0
    if not df_f.empty:
        acum = df_f.groupby(COL_G_PROVEEDOR)[COL_G_IMPORTE].sum()
        prov_listos = int((acum > UMBRAL_CUARENTENA).sum())
    _tarjeta_kpi(c3, "✅", "Proveedores listos",
                 f"{prov_listos:,}",
                 f"acumulado > ${UMBRAL_CUARENTENA:,.0f}",
                 color="#059669")

    
    # ---- Acumulado por proveedor ----
    st.markdown("#### 🏭 Acumulado por proveedor")
    resumen = (df_f.groupby(COL_G_PROVEEDOR)
               .agg(Folios=(COL_G_FOLIO, "count"),
                    Acumulado=(COL_G_IMPORTE, "sum"))
               .sort_values("Acumulado", ascending=False))
    resumen["¿Supera umbral?"] = resumen["Acumulado"].apply(
        lambda v: "✅ Sí" if v > UMBRAL_CUARENTENA else "—")
    st.dataframe(
        resumen, use_container_width=True,
        column_config={
            "Folios": st.column_config.NumberColumn(format="%d"),
            "Acumulado": st.column_config.NumberColumn(format="$%.2f"),
        })

    st.divider()

    # ---- Administrar por proveedor ----
    st.markdown("#### ⚙️ Administrar cuarentena por proveedor")
    st.caption("Selecciona un proveedor para ver el detalle y liberar sus "
                "folios al flujo normal (sin clave).")

    prov_sel = st.selectbox("Proveedor", options=resumen.index.tolist(),
                              key="q_prov_sel")
    grupo = df_f[df_f[COL_G_PROVEEDOR] == prov_sel].copy()
    grupo["Días restantes cuar."] = grupo[COL_G_CUAR_FIN].apply(
        lambda f: (_a_fecha(f) - hoy_mx()).days if pd.notna(f) else None)

    st.dataframe(
        grupo[[COL_G_FOLIO, COL_G_IMPORTE, COL_G_CUAR_INICIO,
               COL_G_CUAR_FIN, "Días restantes cuar."]],
        use_container_width=True, hide_index=True,
        column_config={
            COL_G_IMPORTE: st.column_config.NumberColumn("Importe",
                                                          format="$%.2f"),
            COL_G_CUAR_INICIO: st.column_config.DateColumn("Inicio",
                                                            format="DD/MM/YYYY"),
            COL_G_CUAR_FIN: st.column_config.DateColumn("Sale",
                                                        format="DD/MM/YYYY"),
            "Días restantes cuar.": st.column_config.NumberColumn(format="%d"),
        },
    )

    c1, c2 = st.columns(2)
    if c1.button(f"🚀 Liberar TODOS los folios de {prov_sel}",
                  key=f"lib_all_{prov_sel}", type="primary",
                  use_container_width=True):
        dfx = datos["garantias"]
        m = ((dfx[COL_G_ESTADO] == ESTADO_CUARENTENA) &
             (dfx[COL_G_PROVEEDOR] == prov_sel))
        n = int(m.sum())
        dfx.loc[m, COL_G_ESTADO] = ESTADO_ACTIVO
        dfx.loc[m, COL_G_FECHA_RECEPCION] = pd.Timestamp(hoy_mx())
        dfx.loc[m, COL_G_MODIFICADO] = f"{ahora_mx():%d/%m/%Y %H:%M}"
        persistir(datos,
                    f"Cuarentena: {n} folio(s) de {prov_sel} liberados",
                    f"{n} folio(s) liberados al flujo normal.")

    if c2.button(f"⛔ Cancelar TODOS los folios de {prov_sel}",
                  key=f"can_all_{prov_sel}", use_container_width=True):
        dfx = datos["garantias"]
        m = ((dfx[COL_G_ESTADO] == ESTADO_CUARENTENA) &
             (dfx[COL_G_PROVEEDOR] == prov_sel))
        n = int(m.sum())
        dfx.loc[m, COL_G_ESTADO] = ESTADO_CANCELADO
        dfx.loc[m, COL_G_MODIFICADO] = f"{ahora_mx():%d/%m/%Y %H:%M}"
        persistir(datos,
                    f"Cuarentena: {n} folio(s) de {prov_sel} cancelados",
                    f"{n} folio(s) cancelados.")


# =============================================================================
# 11. NOTAS DE CRÉDITO PENDIENTES
# =============================================================================


def vista_nc_pendientes(datos: dict) -> None:
    st.markdown("#### 💳 Notas de Crédito Pendientes")
    st.caption(f"Notas por conceptos varios que aún no llegan a administración. "
                f"Plazo de **{DIAS_PLAZO_NC} días** desde la fecha de recepción "
                "del reporte.")

    df = datos.get("nc")
    if df is None or df.empty:
        st.info("No hay notas de crédito cargadas.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    # ---- Alerta inicial ----
    df_pend = df[df[COL_NC_ESTADO] == "Pendiente"]
    if not df_pend.empty:
        venc = df_pend[df_pend.apply(esta_vencida_nc, axis=1)]
        if not venc.empty:
            st.error(f"🚨 **{len(venc)}** nota(s) de crédito VENCIDA(S) "
                     f"(más de {DIAS_PLAZO_NC} días sin resolver).")

    # ---- KPIs generales ----
    df_res = df[df[COL_NC_ESTADO] == "Resuelto"]
    c1, c2, c3, c4 = st.columns(4)
    _tarjeta_kpi(c1, "⏳", "NC pendientes", f"{len(df_pend):,}",
                 f"{_fmt_mxn(df_pend[COL_NC_IMP_PENDIENTE].sum())} por conseguir",
                 color="#ea580c")
    _tarjeta_kpi(c2, "✅", "NC ya conseguidas", f"{len(df_res):,}",
                 f"{_fmt_mxn(df_res[COL_NC_IMP_PENDIENTE].sum())} entregadas",
                 color="#059669")
    n_venc = int(df_pend.apply(esta_vencida_nc, axis=1).sum()) if not df_pend.empty else 0
    _tarjeta_kpi(c3, "🚨", f"Vencidas (>{DIAS_PLAZO_NC} días)",
                 f"{n_venc:,}",
                 "Requieren atención" if n_venc > 0 else "Sin vencidas",
                 color="#dc2626" if n_venc > 0 else "#94a3b8")
    _tarjeta_kpi(c4, "🏭", "Proveedores", f"{df_pend[COL_NC_PROVEEDOR].nunique():,}",
                 "con NC pendientes", color="#1f4e79")

    
    # ---- Filtros adicionales ----
    c_a1, c_a2 = st.columns([1, 2])
    ver_resueltas = c_a1.checkbox("Mostrar también resueltas", value=False,
                                    key="nc_ver_res")
    etapas_disp = ["Todas"] + ETAPAS_NC
    f_etapa_nc = c_a2.selectbox("Etapa del flujo", etapas_disp,
                                  key="nc_etapa",
                                  help="Filtra por el estatus intermedio de la NC.")

    df_base = df if ver_resueltas else df[df[COL_NC_ESTADO] != "Resuelto"]
    df_base = df_base[df_base[COL_NC_ESTADO] != "Cancelado"]
    if f_etapa_nc != "Todas":
        df_base = df_base[df_base[COL_NC_ETAPA] == f_etapa_nc]

    # ---- Filtros comunes ----
    df_f, etiq = _filtros_barra(
        df_base, prefijo="nc",
        col_fecha=COL_NC_FECHA_REPORTE,
        col_folio=COL_NC_ENTRADA,
        col_proveedor=COL_NC_PROVEEDOR,
        col_comprador=COL_NC_COMPRADOR,
        etiqueta_folio="Buscar folio de entrada")

    if df_f.empty:
        st.info("No hay notas de crédito con los filtros actuales.")
        return

    # ---- Tabla ----
    vista = df_f.copy()
    vista["🚦"] = vista.apply(lambda f: semaforo_nc(f)[0], axis=1)
    vista["Días transc."] = vista.apply(dias_transcurridos_nc, axis=1)
    vista["Días restantes"] = vista.apply(dias_restantes_nc, axis=1)

    cols = ["🚦", COL_NC_ENTRADA, COL_NC_PROVEEDOR, COL_NC_NUM_FACTURA,
            COL_NC_IMP_FACTURA, COL_NC_IMP_PENDIENTE,
            COL_NC_FECHA_REPORTE, "Días transc.", "Días restantes",
            COL_NC_FECHA_NC, COL_NC_FOLIO_NC, COL_NC_ESTADO, COL_NC_ETAPA,
            COL_NC_COMPRADOR, COL_NC_OBSERVACIONES]
    cols = [c for c in cols if c in vista.columns]
    st.dataframe(
        vista[cols], use_container_width=True, hide_index=True,
        column_config={
            COL_NC_IMP_FACTURA: st.column_config.NumberColumn(
                "Importe factura", format="$%.2f"),
            COL_NC_IMP_PENDIENTE: st.column_config.NumberColumn(
                "Importe NC pendiente", format="$%.2f"),
            COL_NC_FECHA_REPORTE: st.column_config.DateColumn(
                "Fecha recepción", format="DD/MM/YYYY"),
            COL_NC_FECHA_NC: st.column_config.DateColumn(
                "Fecha NC", format="DD/MM/YYYY"),
            COL_NC_ETAPA: "Etapa del flujo",
            COL_NC_FOLIO_NC: "Folio NC",
        },
    )

    st.divider()
    st.markdown("#### ✏️ Actualizar nota de crédito")
    df_sel = df_f.copy()
    df_sel["etiqueta"] = (df_sel[COL_NC_ENTRADA].astype(str) + " · "
                           + df_sel[COL_NC_PROVEEDOR].str.slice(0, 40)
                           + " · " + df_sel[COL_NC_IMP_PENDIENTE].apply(_fmt_mxn))
    sel = st.selectbox("Selecciona una nota",
                        options=df_sel["etiqueta"].tolist(),
                        key="nc_edit_sel")
    fila = df_sel[df_sel["etiqueta"] == sel].iloc[0]
    _editor_nc(fila, datos)


def _editor_nc(fila: pd.Series, datos: dict) -> None:
    icono, texto = semaforo_nc(fila)
    st.markdown(
        f"""<div style='display:flex;gap:1.2rem;flex-wrap:wrap;
                        padding:0.5rem 0.8rem;margin:0.3rem 0 0.8rem 0;
                        border:1px solid #e6e6e6;border-radius:6px;
                        font-size:0.85rem;background:#fafafa'>
          <span><b>Folio entrada:</b> {fila[COL_NC_ENTRADA]}</span>
          <span><b>Proveedor:</b> {fila[COL_NC_PROVEEDOR]}</span>
          <span><b>Importe pendiente:</b> {_fmt_mxn(fila[COL_NC_IMP_PENDIENTE])}</span>
          <span><b>Recepción reporte:</b> {_fmt_fecha(fila[COL_NC_FECHA_REPORTE])}</span>
          <span><b>Estado:</b> {icono} {texto}</span>
        </div>""",
        unsafe_allow_html=True,
    )

    with st.form(f"form_nc_{fila[COL_NC_ENTRADA]}"):
        c1, c2 = st.columns(2)
        fecha_rep_val = _a_fecha(fila.get(COL_NC_FECHA_REPORTE)) or hoy_mx()
        fecha_recep = c1.date_input(
            "Fecha de recepción del reporte",
            value=fecha_rep_val, format="DD/MM/YYYY",
            help="Desde esta fecha corren los 20 días.")

        fecha_nc_actual = _a_fecha(fila.get(COL_NC_FECHA_NC))
        fecha_nc = c2.date_input(
            "Fecha de la nota de crédito (si ya llegó)",
            value=fecha_nc_actual or hoy_mx(), format="DD/MM/YYYY")
        capturar_nc = st.checkbox(
            "Marcar que la NC ya se envió a administración",
            value=fecha_nc_actual is not None,
            key=f"cap_nc_{fila[COL_NC_ENTRADA]}")

        c_nc1, c_nc2 = st.columns(2)
        folio_nc = c_nc1.text_input(
            "Folio de la nota de crédito",
            value=str(fila.get(COL_NC_FOLIO_NC, "") or ""))

        # Selector de ETAPA
        actual_etapa_nc = str(fila.get(COL_NC_ETAPA, "") or "")
        idx_etapa_nc = (ETAPAS_NC.index(actual_etapa_nc)
                        if actual_etapa_nc in ETAPAS_NC else 0)
        etapa_nc = c_nc2.selectbox(
            "Etapa del flujo", options=ETAPAS_NC, index=idx_etapa_nc,
            help="Marca en qué punto va la gestión de esta NC.")

        observ = st.text_area(
            "Observaciones (qué se está haciendo para conseguirla)",
            value=str(fila.get(COL_NC_OBSERVACIONES, "") or ""),
            height=100)

        c1, c2 = st.columns([2, 1])
        guardar = c1.form_submit_button("💾 Guardar cambios",
                                          use_container_width=True,
                                          type="primary")
        cancelar = c2.form_submit_button("⛔ Cancelar seguimiento",
                                           use_container_width=True)

    if guardar:
        cambios = {
            COL_NC_FECHA_REPORTE: pd.Timestamp(fecha_recep),
            COL_NC_FOLIO_NC: folio_nc.strip(),
            COL_NC_ETAPA: etapa_nc,
            COL_NC_OBSERVACIONES: observ.strip(),
        }
        if capturar_nc:
            cambios[COL_NC_FECHA_NC] = pd.Timestamp(fecha_nc)
            cambios[COL_NC_ESTADO] = "Resuelto"
            msg = "NC marcada como RESUELTA (enviada a administración)"
        else:
            cambios[COL_NC_FECHA_NC] = pd.NaT
            cambios[COL_NC_ESTADO] = "Pendiente"
            msg = "NC actualizada"
        _actualizar_nc(fila[COL_NC_ENTRADA], datos, cambios, msg)

    if cancelar:
        _actualizar_nc(fila[COL_NC_ENTRADA], datos,
                        {COL_NC_ESTADO: "Cancelado"},
                        "Seguimiento de NC cancelado")


def _actualizar_nc(folio_entrada, datos: dict, cambios: dict,
                    mensaje: str) -> None:
    df = datos["nc"]
    m = df[COL_NC_ENTRADA] == folio_entrada
    for col, val in cambios.items():
        df.loc[m, col] = val
    persistir(datos, f"NC {folio_entrada}: {mensaje}", mensaje)


# =============================================================================
# 12. FOLIOS DE DEVOLUCIÓN HISTÓRICOS
# =============================================================================


def vista_devoluciones_sueltas(datos: dict) -> None:
    st.markdown("#### 📦 Folios de devolución históricos")
    st.caption("Folios de devolución que no se ligan a un folio reporte actual "
                "(en su mayoría de 2023-2024). Puedes activarlos, bloquearlos o "
                "cancelarlos cuando ya no procedan.")

    df = datos.get("devoluciones")
    if df is None or df.empty:
        st.info("No hay folios de devolución cargados.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    # Filtro por estado
    f_estado = st.selectbox("Estado",
                              ["Todos", "Activo", "Bloqueado", "Cancelado"],
                              key="d_estado")
    df_e = df if f_estado == "Todos" else df[df[COL_D_ESTADO] == f_estado]

    df_f, etiq = _filtros_barra(
        df_e, prefijo="dev",
        col_fecha=COL_D_FECHA,
        col_folio=COL_D_FOLIO,
        col_proveedor=COL_D_PROVEEDOR,
        col_comprador=COL_D_COMPRADOR,
        etiqueta_folio="Buscar folio")

    if df_f.empty:
        st.info("No hay folios con los filtros actuales.")
        return

    st.caption(f"Total pendiente: **{_fmt_mxn(df_f[COL_D_PENDIENTE].sum())}**")

    cols = [COL_D_FOLIO, COL_D_FECHA, COL_D_PROVEEDOR, COL_D_TOTAL,
            COL_D_PENDIENTE, COL_D_APLICADO_MXN, "TIPO", COL_D_RESOLUCION,
            COL_D_COMPRADOR, COL_D_ESTADO]
    cols = [c for c in cols if c in df_f.columns]
    st.dataframe(
        df_f[cols], use_container_width=True, hide_index=True,
        column_config={
            COL_D_TOTAL: st.column_config.NumberColumn("Total", format="$%.2f"),
            COL_D_PENDIENTE: st.column_config.NumberColumn("Pendiente",
                                                            format="$%.2f"),
            COL_D_APLICADO_MXN: st.column_config.NumberColumn("Aplicado",
                                                               format="$%.2f"),
            COL_D_FECHA: st.column_config.DateColumn(format="DD/MM/YYYY"),
        },
    )

    st.divider()
    st.markdown("#### ✏️ Actualizar folio")
    df_sel = df_f.copy()
    df_sel["etiqueta"] = (df_sel[COL_D_FOLIO].astype(str) + " · "
                           + df_sel[COL_D_PROVEEDOR].str.slice(0, 40)
                           + " · " + df_sel[COL_D_PENDIENTE].apply(_fmt_mxn))
    sel = st.selectbox("Selecciona un folio",
                        options=df_sel["etiqueta"].tolist(),
                        key="d_edit_sel")
    fila = df_sel[df_sel["etiqueta"] == sel].iloc[0]

    st.markdown(
        f"""<div style='padding:0.5rem 0.8rem;margin:0.3rem 0 0.8rem 0;
                        border:1px solid #e6e6e6;border-radius:6px;
                        font-size:0.85rem;background:#fafafa'>
          <b>Folio:</b> {fila[COL_D_FOLIO]} · <b>{fila[COL_D_PROVEEDOR]}</b> · 
          Total: {_fmt_mxn(fila[COL_D_TOTAL])} · 
          Pendiente: {_fmt_mxn(fila[COL_D_PENDIENTE])} · 
          Estado: <b>{fila[COL_D_ESTADO]}</b>
        </div>""",
        unsafe_allow_html=True,
    )

    with st.form(f"form_d_{fila[COL_D_FOLIO]}"):
        notas = st.text_area("Notas",
                              value=str(fila.get(COL_D_NOTAS, "") or ""))
        c1, c2, c3 = st.columns(3)
        act = c1.form_submit_button("↩️ Activar", use_container_width=True)
        blq = c2.form_submit_button("🔒 Bloquear", use_container_width=True)
        can = c3.form_submit_button("⛔ Cancelar", use_container_width=True)

    def _act_dev(nuevo, msg):
        df2 = datos["devoluciones"]
        m = df2[COL_D_FOLIO] == fila[COL_D_FOLIO]
        df2.loc[m, COL_D_ESTADO] = nuevo
        df2.loc[m, COL_D_NOTAS] = notas.strip()
        persistir(datos, f"Devolución {fila[COL_D_FOLIO]}: {msg}", msg)

    if act:
        _act_dev("Activo", "Folio activado")
    if blq:
        _act_dev("Bloqueado", "Folio bloqueado")
    if can:
        _act_dev("Cancelado", "Folio cancelado")


# =============================================================================
# 13. BASE DE DATOS Y GUÍA
# =============================================================================


def _excel_en_memoria(datos: dict) -> bytes:
    buffer = io.BytesIO()
    guardar_excel(datos, "/tmp/_desc.xlsx")
    with open("/tmp/_desc.xlsx", "rb") as fh:
        buffer.write(fh.read())
    return buffer.getvalue()


def vista_base_datos(datos: dict) -> None:
    st.markdown("#### 🗄️ Base de datos")
    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    hay_datos = (datos.get("garantias") is not None and
                 not datos["garantias"].empty)

    if hay_datos:
        st.markdown("#### ⬇️ Descargar base actual")
        st.caption("Descarga el Excel con las tres hojas tal como están hoy.")
        hoy_str = ahora_mx().strftime("%Y%m%d_%H%M")
        try:
            excel_bytes = _excel_en_memoria(datos)
            st.download_button(
                "⬇️ Descargar datos.xlsx",
                data=excel_bytes,
                file_name=f"datos_{hoy_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.warning(f"No se pudo generar el Excel para descarga: {e}")
    else:
        st.info("📥 **Primera carga:** No hay datos válidos aún. Sube abajo el "
                "archivo con las tres hojas requeridas.")
        try:
            plantilla_bytes = _excel_en_memoria({"garantias": None,
                                                  "devoluciones": None,
                                                  "nc": None})
            st.download_button(
                "📄 Descargar plantilla vacía",
                data=plantilla_bytes,
                file_name="plantilla_datos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Estructura de las tres hojas requeridas, sin datos.",
            )
        except Exception:
            pass

    st.divider()

    st.markdown("#### ⬆️ Cargar nueva base")
    st.caption(f"Sube un Excel con las tres hojas: **{HOJA_GARANTIAS}**, "
                f"**{HOJA_DEVOLUCIONES}** y **{HOJA_NC}**. Reemplazará por "
                "completo la base actual.")
    archivo = st.file_uploader("Selecciona el archivo", type=["xlsx"])
    if archivo is None:
        return

    try:
        xl = pd.ExcelFile(archivo)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    faltantes = [h for h in [HOJA_GARANTIAS, HOJA_DEVOLUCIONES, HOJA_NC]
                 if h not in xl.sheet_names]
    if faltantes:
        st.error(f"Faltan hojas: {', '.join(faltantes)}")
        st.caption(f"Hojas encontradas: {', '.join(xl.sheet_names)}")
        return

    c1, c2, c3 = st.columns(3)
    n_g = len(pd.read_excel(xl, sheet_name=HOJA_GARANTIAS))
    n_d = len(pd.read_excel(xl, sheet_name=HOJA_DEVOLUCIONES))
    n_nc = len(pd.read_excel(xl, sheet_name=HOJA_NC))
    c1.metric("Folios garantía", n_g)
    c2.metric("Folios devolución", n_d)
    c3.metric("NC pendientes", n_nc)

    confirmar = st.checkbox(
        "Entiendo que se reemplazará por completo la base actual",
        key="confirmar_carga")
    if st.button("⬆️ Cargar y sincronizar", type="primary",
                  use_container_width=True, disabled=not confirmar):
        try:
            with open(RUTA_EXCEL, "wb") as fh:
                archivo.seek(0)
                fh.write(archivo.read())
        except Exception as e:
            st.error(f"No se pudo escribir el archivo: {e}")
            return
        cargar_datos.clear()
        st.session_state["version_datos"] += 1
        st.session_state["datos"] = cargar_datos(RUTA_EXCEL,
                                                  st.session_state["version_datos"])
        with st.spinner("Subiendo al repositorio…"):
            exito, msg = subir_a_github(
                f"Carga masiva ({n_g}+{n_d}+{n_nc} registros) "
                f"{ahora_mx():%d/%m/%Y %H:%M}")
        if exito:
            st.success(f"✅ Base cargada. {msg}")
        else:
            st.warning(f"💾 Guardado local, pero: {msg}")
        st.rerun()


def vista_guia() -> None:
    st.markdown("#### 📖 Guía del sistema")
    st.markdown(f"""
Esta aplicación da seguimiento a devoluciones con proveedores. Toda la gestión
la lleva **una sola persona** con acceso mediante contraseña única.

### 🔑 Concepto general
- Cada folio de garantía se mide de **extremo a extremo**: desde su fecha de
  recepción hasta que se captura la **nota de crédito**.
- El plazo total es de **{DIAS_PLAZO_TOTAL} días**.
- Un folio queda **RESUELTO** en el momento en que se captura la nota de crédito.

### 📊 Tablero Directivo
Es la pestaña para juntas. Puedes elegir mostrar **Solo Garantías**, **Solo NC
pendientes** o **Ambos**. Cada sección con sus propias tarjetas de indicadores
y sus propias gráficas: KPIs, top proveedores, evolución mensual y distribución
por comprador (para garantías).

### 📋 Folios de Garantía
Lista completa con filtros por estado, periodo, proveedor, comprador y folio.
Editor con los tres campos: Folio Devolución, Folio Ajuste y Nota de Crédito.
Al capturar la nota de crédito, el folio pasa automáticamente a **Resuelto**.

### 🧊 Cuarentena
Pestaña dedicada. Los folios con importe ≤ **${UMBRAL_CUARENTENA:,.0f}** entran
automáticamente por **{DIAS_CUARENTENA} días** para acumular monto con otros
del mismo proveedor. Al vencer el plazo se liberan solos y ahí empiezan los
{DIAS_PLAZO_TOTAL} días. También se pueden liberar manualmente sin clave.

### 💳 Notas de Crédito Pendientes
Pestaña aparte para NC de conceptos varios (faltantes, rebates, descuentos,
fletes, etc.). Plazo: **{DIAS_PLAZO_NC} días** desde la fecha de recepción del
reporte. Al capturar la fecha de la NC pasa a **Resuelto**.

### 📦 Folios de devolución históricos
Folios de la hoja 2 que no se ligan a un folio reporte actual. Se pueden
activar, bloquear o cancelar.

### 🚦 Semáforo
- 🟢 En tiempo · 🟡 Por vencerse · 🔴 Vencido · ✅ Resuelto · 🧊 Cuarentena · ⚫ Cancelado

### ✏️ Modificaciones
Sin claves para editar, cancelar o reactivar. Todos los cambios se guardan
automáticamente y se sincronizan con el repositorio.
""")


# =============================================================================
# 13. TABLERO DE FLUJO (KANBAN)
# =============================================================================
#
#   Vista de flujo tipo Kanban. Cada columna es una etapa del proceso. Las
#   tarjetas son folios (o NC). Se puede arrastrar entre columnas si está
#   instalado el paquete `streamlit-sortables`. Si no, se muestra el mismo
#   layout con un menú "Mover a" en cada tarjeta.


def vista_kanban(datos: dict) -> None:
    """Tablero de flujo tipo Kanban VISUAL: columnas por etapa con tarjetas.

    En cada tarjeta hay un selector "Mover a" para cambiarla de etapa. No
    requiere paquetes externos y no genera commits al rozar la pantalla.
    """
    st.markdown("#### 🔄 Tablero de flujo")
    st.caption("Cada columna es una etapa del proceso. En cada tarjeta puedes "
                "usar el selector para moverla a otra etapa. Solo se muestran "
                "folios/NC en proceso (no resueltos ni cancelados).")

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    tipo = st.radio(
        "Tipo de flujo",
        ["📋 Folios de Garantía", "💳 NC Pendientes"],
        horizontal=True, key="kanban_tipo")

    st.divider()

    if tipo == "📋 Folios de Garantía":
        _kanban_garantias(datos)
    else:
        _kanban_nc(datos)


def _kanban_garantias(datos: dict) -> None:
    df = datos.get("garantias")
    if df is None:
        st.info("No hay folios de garantía cargados.")
        return

    df_act = df[df[COL_G_ESTADO] == ESTADO_ACTIVO].copy()
    if df_act.empty:
        st.info("No hay folios activos para mostrar en el flujo.")
        return

    # Filtro por proveedor
    proveedores = sorted([p for p in df_act[COL_G_PROVEEDOR].dropna().unique()
                          if str(p).strip()])
    f_prov = st.multiselect("Filtrar por proveedor", options=proveedores,
                              placeholder="Todos", key="kg_prov")
    if f_prov:
        df_act = df_act[df_act[COL_G_PROVEEDOR].isin(f_prov)]

    # Normalizar la etapa
    df_act[COL_G_ETAPA] = df_act[COL_G_ETAPA].fillna(ETAPA_GARANTIA_INICIAL)
    df_act.loc[~df_act[COL_G_ETAPA].isin(ETAPAS_GARANTIA), COL_G_ETAPA] = \
        ETAPA_GARANTIA_INICIAL

    st.caption(f"**{len(df_act)}** folio(s) activo(s) en el flujo · "
                f"Monto total: **{_fmt_mxn(df_act[COL_G_IMPORTE].sum())}**")

    
    # Layout: una columna por etapa
    cols = st.columns(len(ETAPAS_GARANTIA))
    for i, etapa in enumerate(ETAPAS_GARANTIA):
        with cols[i]:
            grupo = df_act[df_act[COL_G_ETAPA] == etapa]
            # Encabezado de columna
            st.markdown(
                f"<div style='background:#1f4e79;color:white;padding:0.5rem;"
                f"border-radius:6px 6px 0 0;text-align:center;font-weight:600;"
                f"font-size:0.72rem;line-height:1.2'>"
                f"{etapa}<br>"
                f"<span style='color:#cbd5e1;font-size:0.7rem;font-weight:400'>"
                f"{len(grupo)} folio(s) · {_fmt_mxn(grupo[COL_G_IMPORTE].sum())}"
                f"</span></div>",
                unsafe_allow_html=True)
            _render_tarjetas_garantia(grupo, etapa, datos)


def _render_tarjetas_garantia(grupo: pd.DataFrame, etapa_actual: str,
                                 datos: dict) -> None:
    """Renderiza las tarjetas de una columna del kanban de garantías."""
    if grupo.empty:
        st.markdown(
            "<div style='color:#94a3b8;text-align:center;padding:1rem;"
            "font-size:0.75rem;font-style:italic'>Vacío</div>",
            unsafe_allow_html=True)
        return

    # Ordenar por días transcurridos (más urgentes primero)
    grupo = grupo.copy()
    grupo["_dias"] = grupo.apply(dias_transcurridos_garantia, axis=1)
    grupo = grupo.sort_values("_dias", ascending=False)

    # Mostrar hasta 30 tarjetas por columna (evita saturar la pantalla)
    LIMITE = 30
    for _, r in grupo.head(LIMITE).iterrows():
        icono, _ = semaforo_garantia(r)
        restantes = dias_restantes_garantia(r) or 0
        color_dias = ("#dc2626" if restantes < 0
                       else "#ea580c" if restantes <= 15
                       else "#059669")
        etiqueta_dias = (f"{abs(restantes)}d vencido" if restantes < 0
                          else f"{restantes}d restantes")

        st.markdown(
            f"<div style='background:white;padding:0.5rem;"
            f"border:1px solid #e5e7eb;border-radius:4px;"
            f"margin-bottom:0.4rem;font-size:0.72rem;line-height:1.3'>"
            f"<div style='display:flex;justify-content:space-between;"
            f"align-items:center'>"
            f"<b>{icono} {r[COL_G_FOLIO]}</b>"
            f"<span style='color:{color_dias};font-size:0.68rem;"
            f"font-weight:600'>{etiqueta_dias}</span>"
            f"</div>"
            f"<div style='color:#4b5563;margin-top:0.2rem'>"
            f"{str(r[COL_G_PROVEEDOR])[:28]}</div>"
            f"<div style='color:#0f766e;font-weight:600;margin-top:0.15rem'>"
            f"{_fmt_mxn(r[COL_G_IMPORTE])}</div>"
            f"</div>",
            unsafe_allow_html=True)

        # Selector "Mover a" (solo si hay otras etapas donde moverla)
        opciones = ["— Mover a…"] + [e for e in ETAPAS_GARANTIA
                                        if e != etapa_actual]
        key_sel = f"mv_g_{r[COL_G_FOLIO]}"
        seleccion = st.selectbox("", options=opciones, key=key_sel,
                                    label_visibility="collapsed")
        if seleccion != "— Mover a…":
            df = datos["garantias"]
            m = df[COL_G_FOLIO] == r[COL_G_FOLIO]
            df.loc[m, COL_G_ETAPA] = seleccion
            df.loc[m, COL_G_MODIFICADO] = f"{ahora_mx():%d/%m/%Y %H:%M}"
            persistir(datos,
                        f"Kanban: folio {r[COL_G_FOLIO]} → {seleccion}",
                        f"Folio {r[COL_G_FOLIO]} movido a: {seleccion}")

    if len(grupo) > LIMITE:
        st.caption(f"+ {len(grupo) - LIMITE} más (usa filtro por proveedor)")


def _kanban_nc(datos: dict) -> None:
    df = datos.get("nc")
    if df is None:
        st.info("No hay notas de crédito cargadas.")
        return

    df_pend = df[df[COL_NC_ESTADO] == "Pendiente"].copy()
    if df_pend.empty:
        st.info("No hay NC pendientes para mostrar en el flujo.")
        return

    proveedores = sorted([p for p in df_pend[COL_NC_PROVEEDOR].dropna().unique()
                          if str(p).strip()])
    f_prov = st.multiselect("Filtrar por proveedor", options=proveedores,
                              placeholder="Todos", key="knc_prov")
    if f_prov:
        df_pend = df_pend[df_pend[COL_NC_PROVEEDOR].isin(f_prov)]

    df_pend[COL_NC_ETAPA] = df_pend[COL_NC_ETAPA].fillna(ETAPA_NC_INICIAL)
    df_pend.loc[~df_pend[COL_NC_ETAPA].isin(ETAPAS_NC), COL_NC_ETAPA] = \
        ETAPA_NC_INICIAL

    st.caption(f"**{len(df_pend)}** NC pendiente(s) · "
                f"Monto: **{_fmt_mxn(df_pend[COL_NC_IMP_PENDIENTE].sum())}**")

    
    cols = st.columns(len(ETAPAS_NC))
    for i, etapa in enumerate(ETAPAS_NC):
        with cols[i]:
            grupo = df_pend[df_pend[COL_NC_ETAPA] == etapa]
            st.markdown(
                f"<div style='background:#ea580c;color:white;padding:0.5rem;"
                f"border-radius:6px 6px 0 0;text-align:center;font-weight:600;"
                f"font-size:0.72rem;line-height:1.2'>"
                f"{etapa}<br>"
                f"<span style='color:#fed7aa;font-size:0.7rem;font-weight:400'>"
                f"{len(grupo)} NC · "
                f"{_fmt_mxn(grupo[COL_NC_IMP_PENDIENTE].sum())}"
                f"</span></div>",
                unsafe_allow_html=True)
            _render_tarjetas_nc(grupo, etapa, datos)


def _render_tarjetas_nc(grupo: pd.DataFrame, etapa_actual: str,
                          datos: dict) -> None:
    if grupo.empty:
        st.markdown(
            "<div style='color:#94a3b8;text-align:center;padding:1rem;"
            "font-size:0.75rem;font-style:italic'>Vacío</div>",
            unsafe_allow_html=True)
        return

    grupo = grupo.copy()
    grupo["_dias"] = grupo.apply(dias_transcurridos_nc, axis=1)
    grupo = grupo.sort_values("_dias", ascending=False)

    LIMITE = 30
    for _, r in grupo.head(LIMITE).iterrows():
        icono, _ = semaforo_nc(r)
        restantes = dias_restantes_nc(r) or 0
        color_dias = ("#dc2626" if restantes < 0
                       else "#ea580c" if restantes <= 5
                       else "#059669")
        etiqueta_dias = (f"{abs(restantes)}d vencida" if restantes < 0
                          else f"{restantes}d restantes")

        st.markdown(
            f"<div style='background:white;padding:0.5rem;"
            f"border:1px solid #e5e7eb;border-radius:4px;"
            f"margin-bottom:0.4rem;font-size:0.72rem;line-height:1.3'>"
            f"<div style='display:flex;justify-content:space-between;"
            f"align-items:center'>"
            f"<b>{icono} {r[COL_NC_ENTRADA]}</b>"
            f"<span style='color:{color_dias};font-size:0.68rem;"
            f"font-weight:600'>{etiqueta_dias}</span>"
            f"</div>"
            f"<div style='color:#4b5563;margin-top:0.2rem'>"
            f"{str(r[COL_NC_PROVEEDOR])[:28]}</div>"
            f"<div style='color:#ea580c;font-weight:600;margin-top:0.15rem'>"
            f"{_fmt_mxn(r[COL_NC_IMP_PENDIENTE])}</div>"
            f"</div>",
            unsafe_allow_html=True)

        opciones = ["— Mover a…"] + [e for e in ETAPAS_NC
                                        if e != etapa_actual]
        key_sel = f"mv_nc_{r[COL_NC_ENTRADA]}"
        seleccion = st.selectbox("", options=opciones, key=key_sel,
                                    label_visibility="collapsed")
        if seleccion != "— Mover a…":
            df = datos["nc"]
            m = df[COL_NC_ENTRADA] == r[COL_NC_ENTRADA]
            df.loc[m, COL_NC_ETAPA] = seleccion
            persistir(datos,
                        f"Kanban NC: {r[COL_NC_ENTRADA]} → {seleccion}",
                        f"NC {r[COL_NC_ENTRADA]} movida a: {seleccion}")

    if len(grupo) > LIMITE:
        st.caption(f"+ {len(grupo) - LIMITE} más (usa filtro por proveedor)")


def main() -> None:
    if not verificar_acceso():
        st.stop()

    if "version_datos" not in st.session_state:
        st.session_state["version_datos"] = 0
    if "datos" not in st.session_state:
        st.session_state["datos"] = cargar_datos(
            RUTA_EXCEL, st.session_state["version_datos"])
    datos = st.session_state["datos"]

    # Barra lateral
    st.sidebar.markdown("### ⚙️ Panel")
    if st.sidebar.button("🔄 Recargar datos", use_container_width=True):
        cargar_datos.clear()
        st.session_state["version_datos"] += 1
        st.session_state["datos"] = cargar_datos(
            RUTA_EXCEL, st.session_state["version_datos"])
        st.rerun()
    if st.sidebar.button("🚪 Salir", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    # Encabezado compacto (una sola línea)
    st.markdown(
        "<div style='margin:0 0 0.3rem 0;padding:0'>"
        "<span style='font-size:1.15rem;font-weight:700;color:#1f4e79'>"
        "📋 Seguimiento a Devoluciones</span>"
        "<span style='color:#666;font-size:0.82rem'> · MASYFERR / SANVER FORTE</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    if datos.get("error"):
        st.error(datos["error"])
        st.info("Sube un archivo válido en la pestaña **Base de datos** para "
                "comenzar.")
        vista_base_datos(datos)
        return

    tabs = st.tabs([
        "📊 Tablero directivo",
        "🔄 Flujo (Kanban)",
        "📋 Folios de garantía",
        "🧊 Cuarentena",
        "💳 NC pendientes",
        "📦 Devoluciones históricas",
        "🗄️ Base de datos",
        "📖 Guía",
    ])
    with tabs[0]:
        vista_tablero(datos)
    with tabs[1]:
        vista_kanban(datos)
    with tabs[2]:
        vista_garantias(datos)
    with tabs[3]:
        vista_cuarentena(datos)
    with tabs[4]:
        vista_nc_pendientes(datos)
    with tabs[5]:
        vista_devoluciones_sueltas(datos)
    with tabs[6]:
        vista_base_datos(datos)
    with tabs[7]:
        vista_guia()


if __name__ == "__main__":
    main()
