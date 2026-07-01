# =============================================================================
#  DASHBOARD DE DEVOLUCIONES Y RECLAMACIONES — Seguimiento por etapas
#  -----------------------------------------------------------------------------
#  Aplicación Streamlit que gestiona el ciclo de vida completo de una
#  reclamación a través de 4 etapas encadenadas (máquina de estados):
#
#     1) Reporte de reclamo   (7 días)
#     2) Gestión              (30 días)
#     3) Cuentas por pagar    (53 días)
#     4) Destino final        (recolección 20 días / destrucción)
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

import pandas as pd
import streamlit as st

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

# Clave para reactivar (desbloquear) etapas ya cerradas.
CLAVE_REACTIVACION = "devoluciones2026"

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
ETAPA_4 = "Destino final"
ETAPA_FINAL = "FINALIZADO"

# Duraciones (días) de cada etapa.
DIAS_ETAPA_1 = 7    # desde FECHA CORTE
DIAS_ETAPA_2 = 30   # desde el vencimiento de la etapa 1
DIAS_ETAPA_3 = 53   # desde el vencimiento de la etapa 2
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
    ("negociacion", "Negociación con proveedor", "E3 FECHA NEGOCIACION", "E3 USUARIO NEGOCIACION"),
    ("seguimiento", "Seguimiento", "E3 FECHA SEGUIMIENTO", "E3 USUARIO SEGUIMIENTO"),
    ("recepcion_nc", "Recepción de nota de crédito", "E3 FECHA RECEPCION NC", "E3 USUARIO RECEPCION NC"),
    ("aplicacion", "Aplicación de pago", "E3 FECHA APLICACION PAGO", "E3 USUARIO APLICACION"),
]
PASOS_E4_DESTRUCCION = [
    ("informe", "Informe a proveedor", "E4 FECHA INFORME PROVEEDOR", "E4 USUARIO INFORME"),
    ("definicion", "Definición de destino final", "E4 FECHA DEFINICION DESTINO", "E4 USUARIO DEFINICION"),
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
    hoy = date.today()
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
    # Las columnas de fecha a datetime (para poder asignar Timestamps).
    for col in cols_fecha:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Etapa inicial por defecto
    if COL_ETAPA in df.columns:
        df[COL_ETAPA] = df[COL_ETAPA].replace("", pd.NA).fillna(ETAPA_1)

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


def guardar_excel(df: pd.DataFrame, ruta: str) -> None:
    """Sobrescribe el Excel quitando columnas auxiliares."""
    df.drop(columns=["CLAVE", COL_MES_ETIQUETA], errors="ignore").to_excel(
        ruta, sheet_name=NOMBRE_HOJA, index=False
    )


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

    E1: FECHA CORTE + 7
    E2: límite E1 + 30
    E3: límite E2 + 53
    Devuelve un dict con date o None por etapa.
    """
    corte = _a_fecha(fila.get(COL_FECHA_CORTE))
    limites = {"E1": None, "E2": None, "E3": None}
    if corte:
        limites["E1"] = corte + timedelta(days=DIAS_ETAPA_1)
        limites["E2"] = limites["E1"] + timedelta(days=DIAS_ETAPA_2)
        limites["E3"] = limites["E2"] + timedelta(days=DIAS_ETAPA_3)
    return limites


def estado_alarma(fecha_limite, terminada: bool) -> tuple[str, str]:
    """Devuelve (nivel, mensaje) de alarma para una etapa activa.

    nivel: 'ok' | 'warn' | 'danger' | 'done' | 'none'
    """
    if terminada:
        return "done", "Etapa terminada."
    f = _a_fecha(fecha_limite)
    if f is None:
        return "none", "Sin fecha límite definida."
    hoy = date.today()
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
        limites = calcular_fechas_limite(fila)
        # Determinar la fecha límite y bandera de terminada de la etapa activa
        if etapa == ETAPA_1:
            lim, term = limites["E1"], es_verdadero(fila.get(COL_E1_TERM))
        elif etapa == ETAPA_2:
            lim, term = limites["E2"], es_verdadero(fila.get(COL_E2_TERM))
        elif etapa == ETAPA_3:
            lim, term = limites["E3"], es_verdadero(fila.get(COL_E3_TERM))
        elif etapa == ETAPA_4:
            lim, term = _a_fecha(fila.get(COL_E4_LIMITE_REC)), es_verdadero(fila.get(COL_E4_TERM))
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
    # Vencidas primero
    avisos.sort(key=lambda a: 0 if a["nivel"] == "danger" else 1)
    return avisos


def aplicar_guardado(df: pd.DataFrame, clave: str, cambios: dict,
                     usuario: str, mensaje_log: str) -> str:
    """Aplica un diccionario de cambios a la fila y registra auditoría.

    Devuelve el mensaje de commit.
    """
    mascara = df["CLAVE"] == clave
    ahora = datetime.now()
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
    """Botón protegido por clave para reabrir una etapa terminada."""
    with st.expander("🔓 Reactivar esta etapa (requiere clave)"):
        st.caption(
            "Solo personal autorizado. Reactivar reabre la etapa para corregir "
            "información; las etapas posteriores se recalculan al avanzar de nuevo."
        )
        c1, c2 = st.columns([2, 1])
        clave_in = c1.text_input(
            "Clave de reactivación", type="password",
            key=f"react_{etapa_cod}_{clave_registro}",
        )
        if c2.button("Reactivar", key=f"btn_react_{etapa_cod}_{clave_registro}",
                     use_container_width=True):
            if clave_in != CLAVE_REACTIVACION:
                st.error("Clave incorrecta.")
                return
            df = st.session_state["df"]
            cambios = {COL_ETAPA: ETAPA_NOMBRE_POR_COD[etapa_cod]}
            # Quitar la marca de terminada de esta etapa
            cambios[COL_TERM_POR_COD[etapa_cod]] = "NO"
            mensaje = aplicar_guardado(
                df, clave_registro, cambios, "Reactivación",
                f"Etapa '{ETAPA_NOMBRE_POR_COD[etapa_cod]}' reactivada",
            )
            persistir_y_sincronizar(df, mensaje, "Etapa reactivada.")


# Mapas auxiliares para reactivación
ETAPA_NOMBRE_POR_COD = {"E1": ETAPA_1, "E2": ETAPA_2, "E3": ETAPA_3, "E4": ETAPA_4}
COL_TERM_POR_COD = {"E1": COL_E1_TERM, "E2": COL_E2_TERM, "E3": COL_E3_TERM, "E4": COL_E4_TERM}


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
    hoy = date.today()
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
            st.metric("Días que tardó la etapa", dias)
        if str(fila.get(COL_E1_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E1_OBS)}")
        _boton_reactivar("E1", clave)
        return

    with st.form(f"form_e1_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, PASOS_E1, "e1", solo_lectura=False)
        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E1_OBS, "") or ""),
                           placeholder="Notas sobre la recepción/revisión/envío…")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e1_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

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
            cambios[COL_E1_DIAS] = (date.today() - corte).days
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
                "Fecha compromiso de recolección", value=date.today(),
                format="DD/MM/YYYY", key=f"comp_{clave}",
            )
            st.info("Al cerrar, se activará la pestaña **Destino final** (modalidad recolección).")
        elif respuesta == RESP_DESTRUCCION:
            st.warning("Destrucción: se reportará a Devoluciones y se informará a CxP de la "
                       "nota de crédito. Al cerrar se activará **Cuentas por pagar**.")
        else:
            st.warning("Sin respuesta: al vencer el plazo se notifica a CxP para negociación. "
                       "Al cerrar se activará **Cuentas por pagar**.")

        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E2_OBS, "") or ""))
        cerrar = st.checkbox("Marcar 'Respuesta de proveedor' como definitiva y cerrar la etapa",
                             key=f"cerrar_e2_{clave}")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e2_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

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

    mensaje_log = "Avance en Gestión"
    if cerrar:
        # Cerrar etapa 2 y bifurcar
        cambios[PASOS_E2[-1][2]] = pd.Timestamp(date.today())  # marca respuesta
        for cu in [p[3] for p in PASOS_E2]:
            if not str(fila.get(cu, "")).strip() and cu not in cambios:
                cambios[cu] = usuario.strip()
        lim_e2 = limites["E2"]
        if lim_e2:
            cambios[COL_E2_DIAS] = (date.today() - lim_e2).days + DIAS_ETAPA_2
        cambios[COL_E2_TERM] = "SÍ"
        if respuesta == RESP_RECOLECCION:
            cambios[COL_ETAPA] = ETAPA_4
            cambios[COL_E4_MODALIDAD] = RESP_RECOLECCION
            cambios[COL_E4_LIMITE_REC] = pd.Timestamp(date.today() + timedelta(days=DIAS_RECOLECCION))
            mensaje_log = "Etapa 2 TERMINADA · Recolección → pasa a Destino final"
        else:
            cambios[COL_ETAPA] = ETAPA_3
            cambios[COL_E4_MODALIDAD] = RESP_DESTRUCCION  # destino final será destrucción
            mensaje_log = f"Etapa 2 TERMINADA · {respuesta} → pasa a Cuentas por pagar"

    mensaje = aplicar_guardado(df, clave, cambios, usuario.strip(), mensaje_log)
    persistir_y_sincronizar(df, mensaje, "Avance guardado.", celebrar=bool(cerrar))


def pestania_etapa3(fila: pd.Series) -> None:
    """Etapa 3 — Cuentas por pagar (53 días desde el vencimiento de la etapa 2)."""
    clave = fila["CLAVE"]
    etapa_actual = str(fila.get(COL_ETAPA, "")).strip()
    # Solo aplica si el flujo fue destrucción o sin respuesta (no recolección)
    es_recoleccion = str(fila.get(COL_RESPUESTA_TIPO, "")).strip() == RESP_RECOLECCION
    disponible = etapa_actual in (ETAPA_3, ETAPA_4, ETAPA_FINAL) and not es_recoleccion
    terminada = _etapa_bloqueada(fila, "E3")
    limites = calcular_fechas_limite(fila)

    st.markdown(f"#### 3️⃣ {ETAPA_3}")
    st.caption("Pasos: Negociación → Seguimiento → Recepción de nota de crédito → "
               "Aplicación de pago. Plazo de 53 días desde el vencimiento de Gestión.")

    if es_recoleccion:
        st.info("Esta reclamación fue de **recolección**, por lo que no pasa por "
                "Cuentas por pagar. Continúa en **Destino final**.")
        return
    if not disponible:
        st.warning("🔒 Esta etapa se activa cuando cierres **Gestión** con respuesta "
                   "de destrucción o sin respuesta.")
        return

    _mostrar_alarma(limites["E3"], terminada)

    # Aviso por vencimiento: si venció sin avanzar, alarma de destrucción por vencimiento
    nivel, _ = estado_alarma(limites["E3"], terminada)
    if nivel == "danger" and not terminada:
        st.error("⚠️ Vencido sin respuesta del proveedor: informar al proveedor de la "
                 "**destrucción por vencimiento**. Estatus sugerido: 'Aplicación a factura'. "
                 "Notificar a Devoluciones para proceder con la destrucción.")

    if terminada:
        st.success("Etapa terminada. Continúa en **Destino final**.")
        _resumen_pasos(fila, PASOS_E3)
        if str(fila.get(COL_E3_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E3_OBS)}")
        _boton_reactivar("E3", clave)
        return

    with st.form(f"form_e3_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, PASOS_E3, "e3", solo_lectura=False)
        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E3_OBS, "") or ""))
        st.caption("Al registrar 'Aplicación de pago' se da por terminada la etapa y se "
                   "activa **Destino final**.")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e3_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

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
            cambios[COL_E3_DIAS] = (date.today() - lim_e3).days + DIAS_ETAPA_3
        cambios[COL_E3_TERM] = "SÍ"
        cambios[COL_ETAPA] = ETAPA_4
        cambios[COL_E4_MODALIDAD] = RESP_DESTRUCCION
        mensaje_log = "Etapa 3 (Cuentas por pagar) TERMINADA → pasa a Destino final"

    mensaje = aplicar_guardado(df, clave, cambios, usuario.strip(), mensaje_log)
    persistir_y_sincronizar(df, mensaje, "Avance guardado.", celebrar=aplicacion_hecha)


def pestania_etapa4(fila: pd.Series) -> None:
    """Etapa 4 — Destino final (recolección 20 días / destrucción)."""
    clave = fila["CLAVE"]
    etapa_actual = str(fila.get(COL_ETAPA, "")).strip()
    modalidad = str(fila.get(COL_E4_MODALIDAD, "")).strip() or RESP_DESTRUCCION
    disponible = etapa_actual in (ETAPA_4, ETAPA_FINAL)
    terminada = _etapa_bloqueada(fila, "E4")

    st.markdown(f"#### 4️⃣ {ETAPA_4}")
    st.caption("Cierre del proceso. La ruta depende de la respuesta del proveedor: "
               "recolección o destrucción.")

    if not disponible:
        st.warning("🔒 Esta etapa se activa al cerrar la etapa previa "
                   "(Gestión con recolección, o Cuentas por pagar).")
        return

    st.info(f"Modalidad: **{modalidad}**")

    # ----- RECOLECCIÓN -----
    if modalidad == RESP_RECOLECCION:
        lim_rec = _a_fecha(fila.get(COL_E4_LIMITE_REC))
        _mostrar_alarma(lim_rec, terminada)
        nivel, _ = estado_alarma(lim_rec, terminada)
        if nivel == "danger" and not terminada:
            st.error("⚠️ Vencieron los 20 días de recolección: proceder a enviarlo a "
                     "**destrucción**. Usa el botón de abajo para cambiar a destrucción.")
            if st.button("➡️ Cambiar a destrucción por vencimiento", key=f"to_destr_{clave}"):
                df = st.session_state["df"]
                cambios = {COL_E4_MODALIDAD: RESP_DESTRUCCION,
                           COL_E4_ESTATUS: "Definición de destino final"}
                mensaje = aplicar_guardado(df, clave, cambios, "Sistema",
                                           "Recolección vencida → cambia a destrucción")
                persistir_y_sincronizar(df, mensaje, "Cambiado a destrucción.")

        pasos = PASOS_E4_RECOLECCION
        ultimo_etiqueta = "Recepción de folio de devolución"
    # ----- DESTRUCCIÓN -----
    else:
        st.caption("Destrucción por vencimiento: Informe a proveedor → Definición de "
                   "destino final → Reporte al almacén → Recepción de folio de ajuste.")
        pasos = PASOS_E4_DESTRUCCION
        ultimo_etiqueta = "Recepción de folio de ajuste"

    if terminada:
        st.success("🎉 Proceso TERMINADO. Todas las etapas y fechas quedaron registradas.")
        _resumen_pasos(fila, pasos)
        if str(fila.get(COL_E4_OBS, "")).strip():
            st.info(f"📝 Observaciones: {fila.get(COL_E4_OBS)}")
        _boton_reactivar("E4", clave)
        return

    with st.form(f"form_e4_{clave}"):
        cambios, cols_usuario = _registrar_pasos_form(fila, pasos, "e4", solo_lectura=False)
        obs = st.text_area("Observaciones de la etapa", value=str(fila.get(COL_E4_OBS, "") or ""))
        st.caption(f"Al registrar '{ultimo_etiqueta}' se da por terminado TODO el proceso.")
        usuario = st.text_input("Tu nombre / usuario", key=f"u_e4_{clave}")
        guardar = st.form_submit_button("💾 Guardar avance", use_container_width=True)

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
    mensaje_log = "Avance en Destino final"
    if ultimo_hecho:
        cambios[COL_E4_TERM] = "SÍ"
        cambios[COL_ETAPA] = ETAPA_FINAL
        mensaje_log = "Etapa 4 (Destino final) TERMINADA · PROCESO FINALIZADO"

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

    # Ficha del registro
    with st.container(border=True):
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"**Folio:** {fila[COL_FOLIO]}")
        c1.markdown(f"**Proveedor:** {fila[COL_PROVEEDOR]}")
        imp = pd.to_numeric(pd.Series([fila.get(COL_IMPORTE)]), errors="coerce").iloc[0]
        c2.markdown(f"**Importe:** ${imp:,.2f}" if pd.notna(imp) else "**Importe:** —")
        c2.markdown(f"**Comprador:** {fila[COL_COMPRADOR]}")
        c3.markdown(f"**Fecha corte:** {_fmt_fecha(fila.get(COL_FECHA_CORTE))}")
        c3.markdown(f"**Carta firmada:** {fila.get(COL_CARTA_FIRMADA, '—')}")
        c4.markdown(f"**Etapa actual:** `{fila.get(COL_ETAPA, '—')}`")
        c4.markdown(f"**Respuesta prov.:** {fila.get(COL_RESPUESTA_TIPO, '') or '—'}")

    t1, t2, t3, t4 = st.tabs([
        f"1️⃣ {ETAPA_1}", f"2️⃣ {ETAPA_2}", f"3️⃣ {ETAPA_3}", f"4️⃣ {ETAPA_4}",
    ])
    with t1:
        pestania_etapa1(fila)
    with t2:
        pestania_etapa2(fila)
    with t3:
        pestania_etapa3(fila)
    with t4:
        pestania_etapa4(fila)


# =============================================================================
# 10. TABLA, FILTROS, KPIs Y NOTIFICACIONES
# =============================================================================

def construir_filtros(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("🔎 Filtros")
    dff = df.copy()

    # Mes de reclamación (cascada inicial)
    if COL_MES_ETIQUETA in dff.columns:
        orden = (dff[[COL_MES_ETIQUETA, COL_MES]].drop_duplicates()
                 .sort_values(COL_MES)[COL_MES_ETIQUETA].tolist())
        sel = st.sidebar.multiselect("Mes de reclamación", options=orden,
                                     placeholder="Todos los meses")
        if sel:
            dff = dff[dff[COL_MES_ETIQUETA].isin(sel)]

    # Comprador
    compradores = sorted(x for x in dff[COL_COMPRADOR].unique() if x)
    selc = st.sidebar.multiselect("Comprador", options=compradores,
                                  placeholder="Todos los compradores")
    if selc:
        dff = dff[dff[COL_COMPRADOR].isin(selc)]

    # Etapa
    etapas_pres = [e for e in (ETAPA_1, ETAPA_2, ETAPA_3, ETAPA_4, ETAPA_FINAL)
                   if e in set(dff[COL_ETAPA])]
    sele = st.sidebar.multiselect("Etapa actual", options=etapas_pres,
                                  placeholder="Todas las etapas")
    if sele:
        dff = dff[dff[COL_ETAPA].isin(sele)]

    st.sidebar.caption(f"Mostrando **{len(dff)}** de **{len(df)}** registros.")
    return dff


def mostrar_kpis(df: pd.DataFrame) -> None:
    total = len(df)
    importe = pd.to_numeric(df.get(COL_IMPORTE), errors="coerce").fillna(0).sum()
    finalizados = int((df[COL_ETAPA] == ETAPA_FINAL).sum())
    en_proceso = total - finalizados
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Reclamaciones", f"{total:,}")
    c2.metric("💰 Importe total (MXN)", f"${importe:,.2f}")
    c3.metric("⏳ En proceso", en_proceso)
    c4.metric("✅ Finalizadas", finalizados)


def mostrar_notificaciones(df: pd.DataFrame) -> None:
    avisos = construir_notificaciones(df)
    if not avisos:
        st.success("✅ No hay reclamaciones por vencerse o vencidas con los filtros actuales.")
        return
    vencidas = [a for a in avisos if a["nivel"] == "danger"]
    por_vencer = [a for a in avisos if a["nivel"] == "warn"]
    c1, c2 = st.columns(2)
    c1.metric("🔴 Vencidas", len(vencidas))
    c2.metric("🟡 Por vencerse", len(por_vencer))
    for a in avisos:
        icono = "🔴" if a["nivel"] == "danger" else "🟡"
        st.markdown(f"{icono} **Folio {a['folio']}** · {a['proveedor']} · "
                    f"_{a['etapa']}_ — {a['mensaje']} (Comprador: {a['comprador']})")


def mostrar_tabla(df: pd.DataFrame) -> None:
    columnas = [COL_FOLIO, COL_PROVEEDOR, COL_COMPRADOR, COL_IMPORTE, COL_FECHA_CORTE,
                COL_ETAPA, COL_RESPUESTA_TIPO, COL_E1_TERM, COL_E2_TERM, COL_E3_TERM,
                COL_E4_TERM, COL_MODIFICADO_POR, COL_FECHA_MODIFICACION]
    columnas = [c for c in columnas if c in df.columns]
    st.dataframe(
        df[columnas], use_container_width=True, hide_index=True,
        column_config={
            COL_IMPORTE: st.column_config.NumberColumn(format="$%.2f"),
            COL_FECHA_CORTE: st.column_config.DateColumn(format="DD/MM/YYYY"),
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
Aquí se captura la **respuesta del proveedor**:
- **Recolección** → se pide *fecha compromiso* y se activa **Destino final**.
- **Destrucción** → se reporta a Devoluciones, se informa a CxP de la nota de
  crédito y se activa **Cuentas por pagar**.
- **Sin respuesta** (al vencer) → se notifica a CxP para negociación y se activa
  **Cuentas por pagar**.

### 3️⃣ {ETAPA_3} · plazo {DIAS_ETAPA_3} días
Pasos: **Negociación → Seguimiento → Recepción de nota de crédito → Aplicación de pago**.
Si vence sin respuesta del proveedor, se lanza la alarma de **destrucción por
vencimiento** (estatus *Aplicación a factura*) y se notifica a Devoluciones.
Al registrar *Aplicación de pago* se activa **Destino final**.

### 4️⃣ {ETAPA_4}
- **Recolección**: **Programación → Recolección → Recepción de folio de devolución**,
  con un contador de **{DIAS_RECOLECCION} días**. Si vence, se manda a destrucción.
- **Destrucción**: **Informe a proveedor → Definición de destino final → Reporte al
  almacén → Recepción de folio de ajuste**.
Al registrar el último paso, **todo el proceso se da por terminado**.

### 🔓 Reactivar etapas
Cada etapa cerrada tiene un botón **Reactivar** protegido con clave. Solo el
personal autorizado puede reabrir una etapa para corregir información.

### 🔔 Alarmas
En la pestaña **Resumen** se listan las reclamaciones *por vencerse* (≤ 15 días) y
*vencidas*. La estructura de avisos está preparada para enviarse por correo en el
futuro.

### ✍️ Registro
Cada paso guarda **fecha y usuario**; las observaciones quedan en la base de datos
y la bitácora completa se acumula en *Acciones / Notas* para su posterior análisis.
        """
    )


# =============================================================================
# 11b. GESTIÓN DE LA BASE DE DATOS (descargar / cargar)
# =============================================================================

def _excel_en_memoria(df: pd.DataFrame) -> bytes:
    """Genera el Excel completo (sin columnas auxiliares) como bytes descargables."""
    import io
    buffer = io.BytesIO()
    df_export = df.drop(columns=["CLAVE", COL_MES_ETIQUETA], errors="ignore")
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
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
    hoy = datetime.now().strftime("%Y%m%d_%H%M")
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
                          f"({n_nuevos} registros, {datetime.now():%d/%m/%Y %H:%M})")
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

    st.markdown("### 📋 Dashboard de Devoluciones y Reclamaciones")
    st.markdown("#### Seguimiento a devoluciones")

    st.sidebar.title("⚙️ Panel de control")
    if st.sidebar.button("🔄 Recargar datos", use_container_width=True):
        cargar_datos.clear()
        st.session_state["version_datos"] += 1
        st.session_state["df"] = cargar_datos(RUTA_EXCEL, st.session_state["version_datos"])
        st.rerun()
    if st.sidebar.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state.clear()
        st.rerun()
    st.sidebar.divider()

    df_filtrado = construir_filtros(df)

    mostrar_kpis(df_filtrado)
    st.divider()

    tab_editar, tab_tabla, tab_resumen, tab_bd, tab_guia = st.tabs(
        ["✏️ Editar reclamación", "📊 Tabla", "🔔 Resumen y alarmas",
         "🗄️ Base de datos", "📖 Guía"]
    )
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
