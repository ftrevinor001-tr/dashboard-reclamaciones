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
    initial_sidebar_state="expanded",
)

RUTA_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_EXCEL = os.path.join(RUTA_BASE, "datos.xlsx")

# --- Hojas del Excel ---
HOJA_GARANTIAS = "datos- FOLIO DE GARANTIA"
HOJA_DEVOLUCIONES = "FOLIOS DE DEVOLUCION-GARANTIA"
HOJA_NC = "NC PENDIENTES"

# --- Umbrales del proceso ---
DIAS_PLAZO_TOTAL = 90     # plazo total de un folio de garantía
DIAS_PLAZO_NC = 20        # plazo para gestionar una nota de crédito pendiente
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
COL_G_CUAR_INICIO = "CUARENTENA INICIO"     # fecha entró
COL_G_CUAR_FIN = "CUARENTENA FIN"           # fecha calculada de liberación
COL_G_FECHA_RESUELTO = "FECHA RESUELTO"     # cuándo se capturó la NC
COL_G_MODIFICADO = "ULTIMA MODIFICACION"

# Estados posibles de un folio de garantía
ESTADO_ACTIVO = "Activo"
ESTADO_CUARENTENA = "Cuarentena"
ESTADO_RESUELTO = "Resuelto"
ESTADO_CANCELADO = "Cancelado"

# --- Columnas de la hoja 2: FOLIOS DE DEVOLUCIÓN ---
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
COL_NC_ENTRADA = "FOLIO DE ENTRADA"
COL_NC_FECHA_FACTURA = "FECHA FACTURA"
COL_NC_FECHA_REPORTE = "FECHA RECEPCION REPORTE"  # NUEVA — arranca el reloj
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
        for col, default in [(COL_D_ESTADO, "Activo"), (COL_D_NOTAS, "")]:
            if col not in d.columns:
                d[col] = default
        # Tipos
        d[COL_D_FECHA] = pd.to_datetime(d[COL_D_FECHA], errors="coerce").dt.normalize()
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
        for col, default in [
            (COL_NC_FECHA_REPORTE, pd.NaT),
            (COL_NC_ESTADO, ""),
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
        # Si no hay fecha de recepción del reporte, usar la fecha de la factura
        # como aproximación inicial.
        m_sin_rep = nc[COL_NC_FECHA_REPORTE].isna()
        nc.loc[m_sin_rep, COL_NC_FECHA_REPORTE] = nc.loc[m_sin_rep, COL_NC_FECHA_FACTURA]
        # Estado por defecto: si tiene FECHA DE LA NC → Resuelto, si no → Pendiente
        vacio_e = nc[COL_NC_ESTADO] == ""
        con_nc_e = nc[COL_NC_FECHA_NC].notna()
        nc.loc[vacio_e & con_nc_e, COL_NC_ESTADO] = "Resuelto"
        nc.loc[vacio_e & ~con_nc_e, COL_NC_ESTADO] = "Pendiente"
        nc["MES ETIQUETA"] = nc[COL_NC_FECHA_REPORTE].apply(etiqueta_mes)
        resultado["nc"] = nc

    return resultado


def guardar_excel(datos: dict, ruta: str) -> None:
    """Sobrescribe el Excel con las tres hojas."""
    with pd.ExcelWriter(ruta, engine="openpyxl",
                        datetime_format="DD/MM/YYYY",
                        date_format="DD/MM/YYYY") as writer:
        if datos.get("garantias") is not None:
            g = datos["garantias"].copy()
            g = g.drop(columns=["MES ETIQUETA"], errors="ignore")
            g.to_excel(writer, sheet_name=HOJA_GARANTIAS, index=False)
        if datos.get("devoluciones") is not None:
            d = datos["devoluciones"].copy()
            d = d.drop(columns=["MES ETIQUETA", "TIPO"], errors="ignore")
            d.to_excel(writer, sheet_name=HOJA_DEVOLUCIONES, index=False)
        if datos.get("nc") is not None:
            nc = datos["nc"].copy()
            nc = nc.drop(columns=["MES ETIQUETA"], errors="ignore")
            # Restaurar el nombre largo de la observación
            nc = nc.rename(columns={
                COL_NC_OBSERVACIONES:
                "OBSERVACIONES , QUE SE ESTA REALIZANDO PARA QUE NOS EMITAN LA NC"})
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
# 7. TABLERO DIRECTIVO
# =============================================================================

def _filtro_periodo(df_g: pd.DataFrame, df_nc: pd.DataFrame,
                     prefijo: str = "dir") -> tuple:
    """Filtro por rango de fechas o por mes. Aplica a ambos DataFrames."""
    c1, c2, c3 = st.columns([1, 1.5, 1.5])
    modo = c1.radio("Filtrar por", ["Mes", "Rango de fechas"],
                     horizontal=True, key=f"{prefijo}_modo")

    hoy = hoy_mx()
    meses_disp = sorted(
        set(df_g["MES ETIQUETA"].dropna().unique()) |
        set(df_nc["MES ETIQUETA"].dropna().unique()) - {"Sin fecha"},
        key=lambda x: pd.to_datetime(
            f"{x.split()[1]}-{list(MESES_ES).__getitem__(list(MESES_ES.values()).index(x.split()[0]))}-01"
            if x != "Sin fecha" else "1900-01-01"),
    )

    if modo == "Mes":
        opciones = meses_disp if meses_disp else ["Sin fecha"]
        # Preseleccionar el mes más reciente
        sel = c2.multiselect("Mes(es)", options=opciones,
                              default=[opciones[-1]] if opciones else [],
                              key=f"{prefijo}_meses")
        c3.empty()
        if sel:
            g_filt = df_g[df_g["MES ETIQUETA"].isin(sel)]
            nc_filt = df_nc[df_nc["MES ETIQUETA"].isin(sel)]
        else:
            g_filt, nc_filt = df_g, df_nc
        etiqueta = ", ".join(sel) if sel else "Todo el periodo"
    else:
        # Rango
        fmin_g = df_g[COL_G_FECHA_RECEPCION].min()
        fmin_nc = df_nc[COL_NC_FECHA_REPORTE].min()
        fmin = min([f for f in [fmin_g, fmin_nc] if pd.notna(f)],
                   default=pd.Timestamp(hoy - timedelta(days=180)))
        rango = c2.date_input(
            "Rango de fechas", value=(fmin.date(), hoy),
            format="DD/MM/YYYY", key=f"{prefijo}_rango")
        c3.empty()
        if isinstance(rango, tuple) and len(rango) == 2:
            desde, hasta = rango
            g_filt = df_g[
                (df_g[COL_G_FECHA_RECEPCION].dt.date >= desde) &
                (df_g[COL_G_FECHA_RECEPCION].dt.date <= hasta)]
            nc_filt = df_nc[
                (df_nc[COL_NC_FECHA_REPORTE].dt.date >= desde) &
                (df_nc[COL_NC_FECHA_REPORTE].dt.date <= hasta)]
            etiqueta = f"{desde:%d/%m/%Y} — {hasta:%d/%m/%Y}"
        else:
            g_filt, nc_filt = df_g, df_nc
            etiqueta = "Todo el periodo"

    return g_filt, nc_filt, etiqueta


def _tarjeta_kpi(col, icono: str, titulo: str, valor: str,
                  subtitulo: str = "", color: str = "#1f4e79") -> None:
    """Tarjeta grande de KPI para el tablero directivo."""
    col.markdown(
        f"""<div style='padding:1rem 1.2rem;border-left:5px solid {color};
                       background:#f8f9fa;border-radius:6px;
                       box-shadow:0 1px 3px rgba(0,0,0,0.08);height:100%;'>
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


def vista_tablero(datos: dict) -> None:
    """Pestaña 1: Tablero Directivo — visión general para juntas."""
    st.markdown("### 📊 Tablero Directivo")
    st.caption("Panorama general para presentación en juntas. "
               "Filtra por mes o rango de fechas.")

    df_g = datos["garantias"]
    df_nc = datos["nc"]
    if df_g is None or df_nc is None:
        st.error("No se pudieron cargar los datos.")
        return

    g_filt, nc_filt, etiqueta_periodo = _filtro_periodo(df_g, df_nc, "dir")

    st.markdown(f"<div style='color:#666;font-size:0.85rem;margin-bottom:0.8rem'>"
                f"📅 Periodo: <b>{etiqueta_periodo}</b></div>",
                unsafe_allow_html=True)

    # ================ TARJETAS KPI ================
    total_folios = len(g_filt)
    monto_total = g_filt[COL_G_IMPORTE].sum()

    activos_mask = g_filt[COL_G_ESTADO] == ESTADO_ACTIVO
    cuar_mask = g_filt[COL_G_ESTADO] == ESTADO_CUARENTENA
    resueltos_mask = g_filt[COL_G_ESTADO] == ESTADO_RESUELTO
    cancelados_mask = g_filt[COL_G_ESTADO] == ESTADO_CANCELADO

    n_activos = activos_mask.sum()
    n_cuar = cuar_mask.sum()
    n_resueltos = resueltos_mask.sum()
    n_cancelados = cancelados_mask.sum()

    # % resueltos = resueltos / (total - cancelados)
    total_valido = total_folios - n_cancelados
    pct_resueltos = (n_resueltos / total_valido * 100) if total_valido > 0 else 0.0

    # Vencidos (activos con >90 días)
    if activos_mask.any():
        n_vencidos = int(g_filt[activos_mask].apply(esta_vencido_garantia,
                                                     axis=1).sum())
    else:
        n_vencidos = 0

    # Notas de crédito conseguidas (folios resueltos)
    monto_resuelto = g_filt.loc[resueltos_mask, COL_G_IMPORTE].sum()
    n_nc_conseguidas = n_resueltos

    # NC pendientes
    nc_pend = nc_filt[nc_filt[COL_NC_ESTADO] == "Pendiente"]
    n_nc_pend = len(nc_pend)
    monto_nc_pend = nc_pend[COL_NC_IMP_PENDIENTE].sum()

    # --- FILA 1: 4 KPIs principales ---
    c1, c2, c3, c4 = st.columns(4)
    _tarjeta_kpi(c1, "📁", "Folios totales",
                 f"{total_folios:,}",
                 f"{n_activos} activos · {n_cuar} en cuarentena",
                 color="#1f4e79")
    _tarjeta_kpi(c2, "💰", "Monto total",
                 _fmt_mxn(monto_total),
                 f"{_fmt_mxn(monto_resuelto)} ya resueltos",
                 color="#0f766e")
    _tarjeta_kpi(c3, "✅", "% Resuelto",
                 f"{pct_resueltos:.1f}%",
                 f"{n_resueltos:,} de {total_valido:,} folios (sin cancelados)",
                 color="#059669")
    _tarjeta_kpi(c4, "📝", "NC conseguidas",
                 f"{n_nc_conseguidas:,}",
                 f"{_fmt_mxn(monto_resuelto)} en notas de crédito",
                 color="#7c3aed")

    # --- FILA 2: Alertas ---
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    _tarjeta_kpi(c1, "🚨", "Folios vencidos (90 días)",
                 f"{n_vencidos:,}",
                 f"{MSG_VENCIDO}" if n_vencidos > 0 else "Sin vencimientos",
                 color="#dc2626" if n_vencidos > 0 else "#94a3b8")
    _tarjeta_kpi(c2, "⏳", "NC pendientes (concepto varios)",
                 f"{n_nc_pend:,}",
                 f"{_fmt_mxn(monto_nc_pend)} por conseguir",
                 color="#ea580c" if n_nc_pend > 0 else "#94a3b8")

    st.divider()

    # ================ GRÁFICAS ================
    st.markdown("#### 📈 Análisis por proveedor")

    # Top 10 proveedores por importe (activos + cuarentena, sin cancelados)
    g_vig = g_filt[g_filt[COL_G_ESTADO].isin([ESTADO_ACTIVO, ESTADO_CUARENTENA])]
    if not g_vig.empty:
        top = (g_vig.groupby(COL_G_PROVEEDOR)
               .agg(Monto=(COL_G_IMPORTE, "sum"),
                    Folios=(COL_G_FOLIO, "count"))
               .sort_values("Monto", ascending=False).head(10))
        _grafica_top_proveedores(top)
    else:
        st.info("No hay folios vigentes para graficar.")

    st.divider()
    st.markdown("#### 📅 Evolución mensual")
    _grafica_evolucion_mensual(df_g)


def _grafica_top_proveedores(top: pd.DataFrame) -> None:
    """Barras horizontales de Top 10 proveedores por monto."""
    if not ALTAIR_OK or top.empty:
        st.dataframe(top, use_container_width=True)
        return
    datos = top.reset_index().rename(columns={COL_G_PROVEEDOR: "Proveedor"})
    datos["_txt"] = datos.apply(
        lambda r: f"${r['Monto']:,.2f} ({int(r['Folios'])})", axis=1)
    base = alt.Chart(datos)
    barras = base.mark_bar(color="#1f4e79").encode(
        y=alt.Y("Proveedor:N", title=None, sort="-x"),
        x=alt.X("Monto:Q", title="Monto (MXN)",
                axis=alt.Axis(format="$,.0f")),
        tooltip=[
            alt.Tooltip("Proveedor:N"),
            alt.Tooltip("Monto:Q", title="Monto", format="$,.2f"),
            alt.Tooltip("Folios:Q", title="Folios", format=",d"),
        ],
    )
    etiquetas = base.mark_text(align="left", dx=4, fontSize=11,
                                 fontWeight="bold", color="#333").encode(
        y=alt.Y("Proveedor:N", sort="-x"),
        x=alt.X("Monto:Q"),
        text=alt.Text("_txt:N"),
    )
    st.altair_chart(
        (barras + etiquetas).properties(
            height=max(280, 30 * len(datos)),
            title="Top 10 proveedores por monto (folios vigentes)"),
        use_container_width=True)


def _grafica_evolucion_mensual(df_g: pd.DataFrame) -> None:
    """Barras apiladas por mes: folios activos, resueltos, cancelados."""
    if df_g.empty:
        st.info("Sin datos para graficar.")
        return
    d = df_g.copy()
    # Solo meses con datos válidos
    d = d[d["MES ETIQUETA"] != "Sin fecha"]
    if d.empty:
        st.info("Sin datos por mes.")
        return
    tabla = (d.groupby(["MES ETIQUETA", COL_G_ESTADO])[COL_G_FOLIO]
             .count().unstack(fill_value=0))
    # Orden cronológico
    def _mes_key(x):
        try:
            nombre, año = x.split()
            m = list(MESES_ES.values()).index(nombre) + 1
            return f"{año}-{m:02d}"
        except Exception:
            return "0000-00"
    tabla = tabla.reindex(sorted(tabla.index, key=_mes_key))
    if not ALTAIR_OK:
        st.bar_chart(tabla, height=320)
        return
    largo = tabla.reset_index().melt(id_vars="MES ETIQUETA",
                                      var_name="Estado", value_name="Folios")
    largo = largo[largo["Folios"] > 0]
    orden_estados = [ESTADO_ACTIVO, ESTADO_CUARENTENA, ESTADO_RESUELTO,
                     ESTADO_CANCELADO]
    colores = ["#f59e0b", "#3b82f6", "#10b981", "#94a3b8"]
    grafica = alt.Chart(largo).mark_bar().encode(
        x=alt.X("MES ETIQUETA:N", title="Mes", sort=list(tabla.index),
                axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("Folios:Q", title="Cantidad de folios"),
        color=alt.Color("Estado:N", sort=orden_estados,
                         scale=alt.Scale(domain=orden_estados, range=colores),
                         legend=alt.Legend(orient="bottom")),
        tooltip=["MES ETIQUETA", "Estado", "Folios"],
    ).properties(height=320,
                 title="Folios de garantía por mes y estado")
    st.altair_chart(grafica, use_container_width=True)


# =============================================================================
# 8. FOLIOS DE GARANTÍA (Pestaña 2)
# =============================================================================

def vista_garantias(datos: dict) -> None:
    st.markdown("### 📋 Folios de Garantía")
    st.caption("Lista completa de folios. Selecciona uno para editarlo. "
               "Plazo total: **90 días** desde la fecha de recepción.")

    df = datos["garantias"]
    if df is None:
        st.error("No hay datos de garantías.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    # ---- Filtros ----
    c1, c2, c3, c4 = st.columns(4)
    estados_disp = ["Todos", ESTADO_ACTIVO, ESTADO_CUARENTENA, ESTADO_RESUELTO,
                    ESTADO_CANCELADO]
    f_estado = c1.selectbox("Estado", estados_disp, key="g_f_estado")
    proveedores = sorted(df[COL_G_PROVEEDOR].dropna().unique().tolist())
    f_prov = c2.multiselect("Proveedor", proveedores, placeholder="Todos",
                             key="g_f_prov")
    compradores = sorted(df[COL_G_COMPRADOR].dropna().unique().tolist())
    f_comp = c3.multiselect("Comprador", compradores, placeholder="Todos",
                             key="g_f_comp")
    f_folio = c4.text_input("Buscar folio", placeholder="Ej. DC-MZ017",
                             key="g_f_folio")

    d = df.copy()
    if f_estado != "Todos":
        d = d[d[COL_G_ESTADO] == f_estado]
    if f_prov:
        d = d[d[COL_G_PROVEEDOR].isin(f_prov)]
    if f_comp:
        d = d[d[COL_G_COMPRADOR].isin(f_comp)]
    if f_folio.strip():
        d = d[d[COL_G_FOLIO].str.contains(f_folio.strip(), case=False, na=False)]

    st.caption(f"**{len(d)}** de **{len(df)}** folios")

    # Alerta al inicio
    vencidos = d[d.apply(esta_vencido_garantia, axis=1)]
    if not vencidos.empty:
        st.error(f"🚨 **{len(vencidos)}** folio(s) vencido(s) sin resolverse "
                 f"({MSG_VENCIDO}).")

    # ---- Tabla resumen ----
    if d.empty:
        st.info("No hay folios con los filtros actuales.")
        return

    vista = d.copy()
    vista["🚦"] = vista.apply(lambda f: semaforo_garantia(f)[0], axis=1)
    vista["Días transcurridos"] = vista.apply(dias_transcurridos_garantia, axis=1)
    vista["Días restantes"] = vista.apply(dias_restantes_garantia, axis=1)
    vista["Vence"] = vista.apply(fecha_vencimiento_garantia, axis=1)

    cols_vista = ["🚦", COL_G_FOLIO, COL_G_PROVEEDOR, COL_G_COMPRADOR,
                  COL_G_IMPORTE, COL_G_FECHA_RECEPCION, "Vence",
                  "Días transcurridos", "Días restantes", COL_G_ESTADO,
                  COL_G_FOLIO_DEV, COL_G_FOLIO_AJUSTE, COL_G_NOTA_CREDITO]
    st.dataframe(
        vista[cols_vista], use_container_width=True, hide_index=True,
        column_config={
            COL_G_IMPORTE: st.column_config.NumberColumn("Importe",
                                                          format="$%.2f"),
            COL_G_FECHA_RECEPCION: st.column_config.DateColumn(
                "Recepción", format="DD/MM/YYYY"),
            "Vence": st.column_config.DateColumn(format="DD/MM/YYYY"),
            COL_G_FOLIO_DEV: "F. Devolución",
            COL_G_FOLIO_AJUSTE: "F. Ajuste",
            COL_G_NOTA_CREDITO: "Nota Crédito",
        },
    )

    st.divider()

    # ---- Editor de folio individual ----
    st.markdown("#### ✏️ Editar folio")
    d_sel = d.copy()
    d_sel["etiqueta"] = (d_sel[COL_G_FOLIO] + " · "
                          + d_sel[COL_G_PROVEEDOR].str.slice(0, 40)
                          + " · [" + d_sel[COL_G_ESTADO] + "]")
    etiqueta_sel = st.selectbox("Selecciona un folio",
                                 options=d_sel["etiqueta"].tolist())
    fila = d_sel[d_sel["etiqueta"] == etiqueta_sel].iloc[0]
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
        _panel_cuarentena(fila, datos)
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

        respuesta = st.selectbox(
            "Respuesta del proveedor",
            options=["", "Recolección", "Destrucción", "Sin respuesta"],
            index=(["", "Recolección", "Destrucción", "Sin respuesta"]
                   .index(str(fila.get(COL_G_RESPUESTA, "") or ""))
                   if str(fila.get(COL_G_RESPUESTA, "") or "")
                   in ["", "Recolección", "Destrucción", "Sin respuesta"]
                   else 0),
        )

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
            COL_G_NOTAS: notas.strip(),
        }
        # Si tiene NC → automáticamente Resuelto
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


def _panel_cuarentena(fila: pd.Series, datos: dict) -> None:
    """Panel especial para folios en cuarentena."""
    inicio = _a_fecha(fila.get(COL_G_CUAR_INICIO))
    fin = _a_fecha(fila.get(COL_G_CUAR_FIN))
    if fin:
        faltan = (fin - hoy_mx()).days
        if faltan > 0:
            st.info(f"🧊 En cuarentena desde {inicio:%d/%m/%Y} — sale en "
                     f"**{faltan} día(s)** ({fin:%d/%m/%Y}).")
        else:
            st.warning(f"🧊 Cuarentena vencida el {fin:%d/%m/%Y}. Al recargar "
                        "se activará automáticamente.")
    st.caption(f"Los folios con importe ≤ ${UMBRAL_CUARENTENA:,.0f} entran a "
                "cuarentena para acumular monto por proveedor. Se liberan solos "
                "al vencer el plazo.")
    c1, c2 = st.columns(2)
    if c1.button("🚀 Liberar ahora", key=f"lib_{fila[COL_G_FOLIO]}",
                  use_container_width=True, type="primary"):
        _actualizar_garantia(fila[COL_G_FOLIO], datos, {
            COL_G_ESTADO: ESTADO_ACTIVO,
            COL_G_FECHA_RECEPCION: pd.Timestamp(hoy_mx()),
        }, "Folio liberado de cuarentena")
    if c2.button("⛔ Cancelar folio", key=f"can_{fila[COL_G_FOLIO]}",
                  use_container_width=True):
        _actualizar_garantia(fila[COL_G_FOLIO], datos,
                              {COL_G_ESTADO: ESTADO_CANCELADO},
                              "Folio cancelado desde cuarentena")


def _actualizar_garantia(folio: str, datos: dict, cambios: dict,
                          mensaje: str) -> None:
    df = datos["garantias"]
    m = df[COL_G_FOLIO] == folio
    for col, val in cambios.items():
        df.loc[m, col] = val
    df.loc[m, COL_G_MODIFICADO] = f"{ahora_mx():%d/%m/%Y %H:%M}"
    persistir(datos, f"Garantía {folio}: {mensaje}", mensaje)


# =============================================================================
# 9. NOTAS DE CRÉDITO PENDIENTES (Pestaña 3)
# =============================================================================

def vista_nc_pendientes(datos: dict) -> None:
    st.markdown("### 💳 Notas de Crédito Pendientes")
    st.caption(f"Notas por conceptos varios que aún no llegan a administración. "
               f"Plazo de **{DIAS_PLAZO_NC} días** desde la fecha de recepción "
               "del reporte.")

    df = datos["nc"]
    if df is None:
        st.error("No hay datos de notas de crédito.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    # Solo pendientes por defecto
    df_pend = df[df[COL_NC_ESTADO] == "Pendiente"]
    df_resueltas = df[df[COL_NC_ESTADO] == "Resuelto"]

    # Alerta al inicio
    if not df_pend.empty:
        vencidas = df_pend[df_pend.apply(esta_vencida_nc, axis=1)]
        if not vencidas.empty:
            st.error(f"🚨 **{len(vencidas)}** nota(s) de crédito VENCIDA(S) "
                     f"(más de {DIAS_PLAZO_NC} días sin resolver).")

    c1, c2, c3, c4 = st.columns(4)
    _tarjeta_kpi(c1, "⏳", "NC pendientes",
                 f"{len(df_pend):,}",
                 f"{_fmt_mxn(df_pend[COL_NC_IMP_PENDIENTE].sum())} por conseguir",
                 color="#ea580c")
    _tarjeta_kpi(c2, "✅", "NC resueltas",
                 f"{len(df_resueltas):,}",
                 f"{_fmt_mxn(df_resueltas[COL_NC_IMP_PENDIENTE].sum())} entregadas",
                 color="#059669")
    n_venc = int(df_pend.apply(esta_vencida_nc, axis=1).sum()) if not df_pend.empty else 0
    _tarjeta_kpi(c3, "🚨", "Vencidas (>20 días)",
                 f"{n_venc:,}",
                 "Requieren atención inmediata" if n_venc > 0 else "Sin vencidas",
                 color="#dc2626" if n_venc > 0 else "#94a3b8")
    proveedores_pend = df_pend[COL_NC_PROVEEDOR].nunique()
    _tarjeta_kpi(c4, "🏭", "Proveedores",
                 f"{proveedores_pend:,}",
                 "con NC pendientes",
                 color="#1f4e79")

    st.divider()

    # Filtros
    c1, c2, c3 = st.columns(3)
    ver_resueltas = c1.checkbox("Mostrar también resueltas", value=False,
                                 key="nc_ver_res")
    proveedores = sorted(df[COL_NC_PROVEEDOR].dropna().unique().tolist())
    f_prov = c2.multiselect("Proveedor", proveedores, placeholder="Todos",
                             key="nc_f_prov")
    f_folio = c3.text_input("Buscar folio de entrada",
                             placeholder="Ej. 176054", key="nc_f_folio")

    d = df.copy()
    if not ver_resueltas:
        d = d[d[COL_NC_ESTADO] != "Resuelto"]
    d = d[d[COL_NC_ESTADO] != "Cancelado"]
    if f_prov:
        d = d[d[COL_NC_PROVEEDOR].isin(f_prov)]
    if f_folio.strip():
        d = d[d[COL_NC_ENTRADA].astype(str).str.contains(
            f_folio.strip(), case=False, na=False)]

    if d.empty:
        st.info("No hay notas de crédito con los filtros actuales.")
        return

    # Tabla
    vista = d.copy()
    vista["🚦"] = vista.apply(lambda f: semaforo_nc(f)[0], axis=1)
    vista["Días transcurridos"] = vista.apply(dias_transcurridos_nc, axis=1)
    vista["Días restantes"] = vista.apply(dias_restantes_nc, axis=1)

    cols_vista = ["🚦", COL_NC_ENTRADA, COL_NC_PROVEEDOR, COL_NC_NUM_FACTURA,
                  COL_NC_IMP_FACTURA, COL_NC_IMP_PENDIENTE,
                  COL_NC_FECHA_REPORTE, "Días transcurridos", "Días restantes",
                  COL_NC_FECHA_NC, COL_NC_FOLIO_NC, COL_NC_ESTADO,
                  COL_NC_COMPRADOR, COL_NC_OBSERVACIONES]
    cols_vista = [c for c in cols_vista if c in vista.columns]
    st.dataframe(
        vista[cols_vista], use_container_width=True, hide_index=True,
        column_config={
            COL_NC_IMP_FACTURA: st.column_config.NumberColumn(
                "Importe factura", format="$%.2f"),
            COL_NC_IMP_PENDIENTE: st.column_config.NumberColumn(
                "Importe NC pendiente", format="$%.2f"),
            COL_NC_FECHA_REPORTE: st.column_config.DateColumn(
                "Fecha recepción", format="DD/MM/YYYY"),
            COL_NC_FECHA_NC: st.column_config.DateColumn(
                "Fecha NC", format="DD/MM/YYYY"),
            COL_NC_FOLIO_NC: "Folio NC",
        },
    )

    st.divider()

    # Editor de NC individual
    st.markdown("#### ✏️ Actualizar nota de crédito")
    d_sel = d.copy()
    d_sel["etiqueta"] = (d_sel[COL_NC_ENTRADA].astype(str) + " · "
                          + d_sel[COL_NC_PROVEEDOR].str.slice(0, 40)
                          + " · " + d_sel[COL_NC_IMP_PENDIENTE].apply(_fmt_mxn))
    etiqueta_sel = st.selectbox("Selecciona una nota",
                                 options=d_sel["etiqueta"].tolist(),
                                 key="nc_edit_sel")
    fila = d_sel[d_sel["etiqueta"] == etiqueta_sel].iloc[0]
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

        folio_nc = st.text_input(
            "Folio de la nota de crédito",
            value=str(fila.get(COL_NC_FOLIO_NC, "") or ""))

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
# 10. FOLIOS DE DEVOLUCIÓN SUELTOS (dentro de garantías, vista aparte)
# =============================================================================

def vista_devoluciones_sueltas(datos: dict) -> None:
    st.markdown("### 📦 Folios de devolución históricos")
    st.caption("Folios de devolución que no se ligan a un folio reporte actual "
               "(en su mayoría de 2023-2024). Aquí puedes revisarlos, "
               "cancelarlos o bloquearlos cuando ya no procedan.")

    df = datos["devoluciones"]
    if df is None or df.empty:
        st.info("No hay folios de devolución cargados.")
        return

    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    # Filtros
    c1, c2, c3 = st.columns(3)
    f_estado = c1.selectbox("Estado", ["Todos", "Activo", "Bloqueado", "Cancelado"],
                             key="d_f_estado")
    f_tipo = c2.selectbox("Tipo",
                           ["Todos", "Garantía (cliente)",
                            "Incidente proveedor", "Otro"], key="d_f_tipo")
    f_prov = c3.text_input("Buscar proveedor", key="d_f_prov")

    d = df.copy()
    if f_estado != "Todos":
        d = d[d[COL_D_ESTADO] == f_estado]
    if f_tipo != "Todos":
        d = d[d["TIPO"] == f_tipo]
    if f_prov.strip():
        d = d[d[COL_D_PROVEEDOR].str.contains(f_prov.strip(), case=False, na=False)]

    st.caption(f"**{len(d)}** de **{len(df)}** folios · "
                f"Total pendiente: **{_fmt_mxn(d[COL_D_PENDIENTE].sum())}**")

    if d.empty:
        st.info("No hay folios con los filtros actuales.")
        return

    cols = [COL_D_FOLIO, COL_D_FECHA, COL_D_PROVEEDOR, COL_D_TOTAL,
            COL_D_PENDIENTE, COL_D_APLICADO_MXN, "TIPO", COL_D_RESOLUCION,
            COL_D_COMPRADOR, COL_D_ESTADO]
    cols = [c for c in cols if c in d.columns]
    st.dataframe(
        d[cols], use_container_width=True, hide_index=True,
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

    # Editor
    st.markdown("#### ✏️ Actualizar folio")
    d_sel = d.copy()
    d_sel["etiqueta"] = (d_sel[COL_D_FOLIO].astype(str) + " · "
                          + d_sel[COL_D_PROVEEDOR].str.slice(0, 40)
                          + " · " + d_sel[COL_D_PENDIENTE].apply(_fmt_mxn))
    sel = st.selectbox("Selecciona un folio", options=d_sel["etiqueta"].tolist())
    fila = d_sel[d_sel["etiqueta"] == sel].iloc[0]

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

    def _act_dev(nuevo_estado, msg):
        df2 = datos["devoluciones"]
        m = df2[COL_D_FOLIO] == fila[COL_D_FOLIO]
        df2.loc[m, COL_D_ESTADO] = nuevo_estado
        df2.loc[m, COL_D_NOTAS] = notas.strip()
        persistir(datos, f"Devolución {fila[COL_D_FOLIO]}: {msg}", msg)

    if act:
        _act_dev("Activo", "Folio activado")
    if blq:
        _act_dev("Bloqueado", "Folio bloqueado")
    if can:
        _act_dev("Cancelado", "Folio cancelado")


# =============================================================================
# 11. CARGA / DESCARGA DE BASE DE DATOS
# =============================================================================

def _excel_en_memoria(datos: dict) -> bytes:
    buffer = io.BytesIO()
    guardar_excel(datos, "/tmp/_desc.xlsx")
    with open("/tmp/_desc.xlsx", "rb") as fh:
        buffer.write(fh.read())
    return buffer.getvalue()


def vista_base_datos(datos: dict) -> None:
    st.markdown("### 🗄️ Base de datos")
    if "flash" in st.session_state:
        tipo, msg = st.session_state.pop("flash")
        (st.success if tipo == "success" else st.warning)(msg)

    st.markdown("#### ⬇️ Descargar base actual")
    st.caption("Descarga el Excel con las tres hojas tal como están hoy. "
                "Puedes editarlo, agregar nuevas filas y volver a cargarlo.")
    hoy_str = ahora_mx().strftime("%Y%m%d_%H%M")
    st.download_button(
        "⬇️ Descargar datos.xlsx",
        data=_excel_en_memoria(datos),
        file_name=f"datos_{hoy_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    st.markdown("#### ⬆️ Cargar nueva base")
    st.caption("Sube un Excel con las tres hojas: "
                f"**{HOJA_GARANTIAS}**, **{HOJA_DEVOLUCIONES}** y **{HOJA_NC}**. "
                "Reemplazará por completo la base actual.")
    archivo = st.file_uploader("Selecciona el archivo", type=["xlsx"])
    if archivo is None:
        return

    try:
        xl = pd.ExcelFile(archivo)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    hojas_encontradas = xl.sheet_names
    faltantes = [h for h in [HOJA_GARANTIAS, HOJA_DEVOLUCIONES, HOJA_NC]
                 if h not in hojas_encontradas]
    if faltantes:
        st.error(f"Faltan hojas: {', '.join(faltantes)}")
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


# =============================================================================
# 12. GUÍA
# =============================================================================

def vista_guia() -> None:
    st.markdown("### 📖 Guía del sistema")
    st.markdown(f"""
Esta aplicación da seguimiento a devoluciones con proveedores. Toda la gestión
la lleva **una sola persona** con acceso mediante contraseña única.

### 🔑 Concepto general
- Cada folio de garantía se mide de **extremo a extremo**: desde su fecha de
  recepción hasta que se captura la **nota de crédito**.
- El plazo total es de **{DIAS_PLAZO_TOTAL} días**.
- Un folio queda **RESUELTO** en el momento en que se captura la nota de crédito.

### 🧊 Cuarentena
- Los folios con importe menor o igual a **${UMBRAL_CUARENTENA:,.0f}** entran
  automáticamente a **cuarentena** por **{DIAS_CUARENTENA} días** para acumular
  monto con otros folios del mismo proveedor.
- Al vencer el plazo, se liberan solos y **ahí empiezan a contar los
  {DIAS_PLAZO_TOTAL} días**.
- También puedes liberarlos manualmente en cualquier momento (sin clave).

### 💳 Notas de Crédito Pendientes
- Pestaña aparte para NC de conceptos varios (faltantes, rebates, descuentos,
  fletes, etc.).
- Plazo: **{DIAS_PLAZO_NC} días** desde la fecha de recepción del reporte hasta
  que se envía la NC a administración.
- Cuando se captura la fecha de la NC, la NC pasa a estado **Resuelto** y sale
  del listado de pendientes.

### 🚦 Semáforo
- 🟢 En tiempo · 🟡 Por vencerse · 🔴 Vencido · ✅ Resuelto · 🧊 Cuarentena · ⚫ Cancelado

### 📊 Tablero Directivo
Es la pestaña que se presenta en juntas. Muestra:
- Número de folios y monto total
- % de folios resueltos
- Notas de crédito conseguidas
- Alertas de folios vencidos
- Top 10 proveedores por importe
- Evolución mensual

Todo con **filtro por mes o rango de fechas**.

### ✏️ Modificaciones
- Ya no se requieren claves para editar, cancelar o reactivar folios.
- Todos los cambios se guardan automáticamente y se sincronizan con el
  repositorio.

### 🗄️ Base de datos
- Se puede **descargar** el Excel actual en cualquier momento.
- Se puede **cargar** una versión actualizada (reemplaza toda la base).
- El archivo debe tener las tres hojas:
  `{HOJA_GARANTIAS}`, `{HOJA_DEVOLUCIONES}` y `{HOJA_NC}`.
""")


# =============================================================================
# 13. MAIN
# =============================================================================

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

    # Encabezado
    st.markdown(
        "<div style='margin-bottom:0.4rem'>"
        "<span style='font-size:1.3rem;font-weight:700;color:#1f4e79'>"
        "📋 Seguimiento a Devoluciones</span>"
        "<span style='color:#666;font-size:0.9rem'> · MASYFERR / SANVER FORTE</span>"
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
        "📋 Folios de garantía",
        "💳 NC pendientes",
        "📦 Devoluciones históricas",
        "🗄️ Base de datos",
        "📖 Guía",
    ])
    with tabs[0]:
        vista_tablero(datos)
    with tabs[1]:
        vista_garantias(datos)
    with tabs[2]:
        vista_nc_pendientes(datos)
    with tabs[3]:
        vista_devoluciones_sueltas(datos)
    with tabs[4]:
        vista_base_datos(datos)
    with tabs[5]:
        vista_guia()


if __name__ == "__main__":
    main()
