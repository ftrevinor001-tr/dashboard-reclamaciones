# =============================================================================
#  DASHBOARD DE DEVOLUCIONES Y RECLAMACIONES
#  -----------------------------------------------------------------------------
#  Aplicación Streamlit que gestiona un registro de reclamaciones almacenado
#  en un archivo Excel (datos.xlsx) dentro del mismo repositorio de GitHub.
#  Los cambios se guardan en el Excel y se suben automáticamente a GitHub
#  mediante GitPython usando un token seguro (st.secrets["GITHUB_TOKEN"]).
#
#  Estructura del código (modular, dentro de un solo archivo):
#    1. Configuración general y constantes
#    2. Módulo de autenticación (login)
#    3. Módulo de datos (carga, normalización, semáforo de fechas)
#    4. Módulo de sincronización con GitHub
#    5. Componentes de interfaz (KPIs, filtros, tabla, formulario de edición)
#    6. Función principal (main)
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
    page_title="Dashboard de Reclamaciones",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Ruta del Excel (vive en la raíz del repositorio, junto a app.py)
RUTA_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_EXCEL = os.path.join(RUTA_BASE, "datos.xlsx")
NOMBRE_HOJA = "datos"

# Opciones FIJAS de estatus (requerimiento estricto)
OPCIONES_ESTATUS = [
    "Enviada",
    "En revisión",
    "En proceso",
    "Aceptado",
    "Devuelto",
    "Pagado",
    "Rechazado",
    "Terminado",
]

# Nombres canónicos de columnas (los encabezados reales del Excel traen
# saltos de línea y espacios dobles, por eso se normalizan al cargar).
COL_MES = "MES DE DEVOLUCION"
COL_ID = "ID"
COL_PROVEEDOR = "PROVEEDOR"
COL_FOLIO = "FOLIO REPORTE"
COL_CARTA_FIRMADA = "CARTA FIRMADA"
COL_IMPORTE = "IMPORTE (MXN)"
COL_COMPRADOR = "COMPRADOR"
COL_FECHA_CORTE = "FECHA CORTE"
COL_LIMITE_NOTIFICAR = "LÍMITE NOTIFICAR (+7 DÍAS)"
COL_SEGUIMIENTO = "FECHA CON COMPRADOR SEGUIMIENTO (+1 MES)"
COL_AVISO_CXP = "AVISO CXP (+2 MESES)"
COL_VENCIMIENTO = "FECHA VENCIMIENTO (+90 DÍAS)"
COL_DIAS_TRANS = "DÍAS TRANSCURRIDOS"
COL_DIAS_REST = "DÍAS RESTANTES"
COL_CARTA_ENVIADA = "¿CARTA ENVIADA?"
COL_FECHA_ENVIO = "FECHA ENVÍO CARTA"
COL_RESPUESTA = "¿RESPUESTA PROVEEDOR?"
COL_FECHA_RESPUESTA = "FECHA RESPUESTA"
COL_AVISO_ENVIADO = "¿AVISO CXP ENVIADO?"
COL_ESTATUS = "ESTATUS ACTUAL"
COL_NOTAS = "ACCIONES / NOTAS"
# Columnas de auditoría (se crean si no existen)
COL_MODIFICADO_POR = "MODIFICADO POR"
COL_FECHA_MODIFICACION = "FECHA MODIFICACIÓN"

# Columnas de fecha que se muestran como SEMÁFORO (estatus, no fecha)
COLUMNAS_SEMAFORO = [
    COL_LIMITE_NOTIFICAR,
    COL_SEGUIMIENTO,
    COL_AVISO_CXP,
    COL_VENCIMIENTO,
]

# Valores posibles del semáforo
EN_TIEMPO = "🟢 EN TIEMPO"
POR_VENCERSE = "🟡 POR VENCERSE"
VENCIDO = "🔴 VENCIDO"
SIN_FECHA = "—"


# =============================================================================
# 2. MÓDULO DE AUTENTICACIÓN (LOGIN)
# =============================================================================

def verificar_acceso() -> bool:
    """Pantalla de inicio de sesión simple con contraseña maestra.

    La contraseña se valida contra st.secrets["DASHBOARD_PASSWORD"].
    Devuelve True solo si el usuario ya está autenticado.
    """
    if st.session_state.get("autenticado", False):
        return True

    # --- Diseño de la pantalla de login ---
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.2, 1])
    with col_centro:
        st.markdown("## 🔐 Dashboard de Reclamaciones")
        st.caption("Acceso restringido. Ingresa la contraseña para continuar.")

        with st.form("form_login"):
            password = st.text_input(
                "Contraseña", type="password", placeholder="••••••••"
            )
            enviar = st.form_submit_button("Ingresar", use_container_width=True)

        if enviar:
            password_correcta = st.secrets.get("DASHBOARD_PASSWORD")
            if password_correcta is None:
                st.error(
                    "⚠️ No se encontró 'DASHBOARD_PASSWORD' en los secretos "
                    "de Streamlit. Revisa la guía de configuración."
                )
            elif password == password_correcta:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta. Inténtalo de nuevo.")
    return False


def boton_cerrar_sesion() -> None:
    """Botón para cerrar sesión desde la barra lateral."""
    if st.sidebar.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state.clear()
        st.rerun()


# =============================================================================
# 3. MÓDULO DE DATOS
# =============================================================================

def _normalizar_encabezado(nombre: str) -> str:
    """Limpia saltos de línea y espacios múltiples de un encabezado."""
    return " ".join(str(nombre).split()).upper()


def _a_fecha(valor):
    """Convierte un valor (datetime, serial de Excel o texto) a date.

    Devuelve None si no es una fecha válida.
    """
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    # Ya es fecha / datetime
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.date() if hasattr(valor, "date") else valor
    if isinstance(valor, date):
        return valor
    # Serial de Excel (número de días desde 1899-12-30)
    if isinstance(valor, (int, float)):
        try:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(valor))).date()
        except Exception:
            return None
    # Texto
    try:
        return pd.to_datetime(str(valor), dayfirst=True, errors="raise").date()
    except Exception:
        return None


def calcular_semaforo(fecha_limite) -> str:
    """Aplica la regla de semáforo solicitada.

    - Más de 15 días antes del vencimiento  -> EN TIEMPO
    - 15 días o menos (incluido el día de hoy) -> POR VENCERSE
    - Fecha de vencimiento ya pasada         -> VENCIDO
    """
    fecha = _a_fecha(fecha_limite)
    if fecha is None:
        return SIN_FECHA

    hoy = date.today()
    if fecha < hoy:
        return VENCIDO
    if fecha <= hoy + timedelta(days=15):
        return POR_VENCERSE
    return EN_TIEMPO


@st.cache_data(show_spinner="Cargando base de datos…")
def cargar_datos(ruta: str, _version: int) -> pd.DataFrame:
    """Carga el Excel y normaliza encabezados y tipos.

    El parámetro _version permite invalidar la caché tras cada guardado.
    """
    df = pd.read_excel(ruta, sheet_name=NOMBRE_HOJA)
    df.columns = [_normalizar_encabezado(c) for c in df.columns]

    # Asegurar columnas de auditoría
    for col in (COL_MODIFICADO_POR, COL_FECHA_MODIFICACION):
        if col not in df.columns:
            df[col] = ""

    # Normalizar columnas de texto clave
    for col in (COL_PROVEEDOR, COL_COMPRADOR, COL_ESTATUS, COL_FOLIO):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Clave única de reclamación: el ID de proveedor se repite entre meses,
    # por lo que la combinación ID + FOLIO identifica cada registro.
    df["CLAVE"] = df[COL_ID].astype(str) + " | " + df[COL_FOLIO]
    return df


def agregar_columnas_semaforo(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve una copia con las 4 columnas de fecha convertidas a semáforo."""
    vista = df.copy()
    for col in COLUMNAS_SEMAFORO:
        if col in vista.columns:
            vista[col] = vista[col].apply(calcular_semaforo)
    return vista


def guardar_excel(df: pd.DataFrame, ruta: str) -> None:
    """Sobrescribe el archivo Excel con el DataFrame actualizado."""
    df_guardar = df.drop(columns=["CLAVE"], errors="ignore")
    df_guardar.to_excel(ruta, sheet_name=NOMBRE_HOJA, index=False)


# =============================================================================
# 4. MÓDULO DE SINCRONIZACIÓN CON GITHUB
# =============================================================================

def subir_a_github(mensaje_commit: str) -> tuple[bool, str]:
    """Hace commit y push del Excel al repositorio usando GitPython.

    Usa st.secrets["GITHUB_TOKEN"] para autenticarse. Devuelve una tupla
    (exito, mensaje) para mostrar retroalimentación al usuario.
    """
    try:
        from git import Repo
    except ImportError:
        return False, "GitPython no está instalado (revisa requirements.txt)."

    token = st.secrets.get("GITHUB_TOKEN")
    if not token:
        return False, (
            "No se encontró 'GITHUB_TOKEN' en los secretos de Streamlit. "
            "El cambio se guardó en el Excel local, pero NO se subió a GitHub."
        )

    try:
        repo = Repo(RUTA_BASE, search_parent_directories=True)

        # Identidad del commit (necesaria en entornos efímeros como
        # Streamlit Cloud, donde git no tiene usuario configurado).
        with repo.config_writer() as cw:
            cw.set_value("user", "name", "Dashboard Reclamaciones")
            cw.set_value("user", "email", "dashboard@reclamaciones.app")

        repo.index.add([RUTA_EXCEL])

        # Si no hay cambios reales, no hacemos commit vacío.
        if not repo.index.diff("HEAD"):
            return True, "No había cambios nuevos que subir."

        repo.index.commit(mensaje_commit)

        # Inyectar el token en la URL del remoto de forma temporal y segura.
        origen = repo.remote(name="origin")
        url_original = origen.url
        url_limpia = url_original
        if url_limpia.startswith("git@github.com:"):
            url_limpia = url_limpia.replace("git@github.com:", "https://github.com/")
        if "@" in url_limpia and url_limpia.startswith("https://"):
            # Quitar credenciales previas si las hubiera
            url_limpia = "https://" + url_limpia.split("@", 1)[1]
        url_con_token = url_limpia.replace("https://", f"https://{token}@")

        try:
            origen.set_url(url_con_token)
            origen.push()
        finally:
            # Restaurar la URL sin token para no dejarlo expuesto en .git/config
            origen.set_url(url_original)

        return True, "Cambios subidos a GitHub correctamente. ✅"
    except Exception as e:
        return False, f"Error al subir a GitHub: {e}"


# =============================================================================
# 5. COMPONENTES DE INTERFAZ
# =============================================================================

def mostrar_kpis(df: pd.DataFrame) -> None:
    """Tarjetas de métricas en la parte superior del dashboard."""
    total = len(df)
    importe = pd.to_numeric(df.get(COL_IMPORTE), errors="coerce").fillna(0).sum()
    vista = agregar_columnas_semaforo(df)
    vencidos = int((vista[COL_VENCIMIENTO] == VENCIDO).sum())
    por_vencer = int((vista[COL_VENCIMIENTO] == POR_VENCERSE).sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Reclamaciones", f"{total:,}")
    c2.metric("💰 Importe total (MXN)", f"${importe:,.2f}")
    c3.metric("🟡 Por vencerse", por_vencer)
    c4.metric("🔴 Vencidas", vencidos)


def construir_filtros(df: pd.DataFrame) -> pd.DataFrame:
    """Barra lateral con filtros en cascada: Comprador → Estatus → Semáforo.

    Cada filtro reduce las opciones disponibles del siguiente.
    """
    st.sidebar.header("🔎 Filtros")

    df_filtrado = df.copy()

    # --- Filtro por COMPRADOR ---
    compradores = sorted(x for x in df_filtrado[COL_COMPRADOR].unique() if x)
    sel_compradores = st.sidebar.multiselect(
        "Comprador", options=compradores, placeholder="Todos los compradores"
    )
    if sel_compradores:
        df_filtrado = df_filtrado[df_filtrado[COL_COMPRADOR].isin(sel_compradores)]

    # --- Filtro por ESTATUS (en cascada: solo estatus presentes tras el
    #     filtro anterior, manteniendo las opciones fijas como base) ---
    estatus_presentes = set(df_filtrado[COL_ESTATUS].unique())
    opciones_estatus = [e for e in OPCIONES_ESTATUS if e in estatus_presentes]
    # Incluir también estatus históricos del archivo que no estén en la lista fija
    extras = sorted(e for e in estatus_presentes if e and e not in OPCIONES_ESTATUS)
    opciones_estatus += extras

    sel_estatus = st.sidebar.multiselect(
        "Estatus actual", options=opciones_estatus, placeholder="Todos los estatus"
    )
    if sel_estatus:
        df_filtrado = df_filtrado[df_filtrado[COL_ESTATUS].isin(sel_estatus)]

    # --- Filtro por SEMÁFORO de vencimiento (cascada final) ---
    vista = agregar_columnas_semaforo(df_filtrado)
    semaforos_presentes = [
        s for s in (EN_TIEMPO, POR_VENCERSE, VENCIDO, SIN_FECHA)
        if s in set(vista[COL_VENCIMIENTO])
    ]
    sel_semaforo = st.sidebar.multiselect(
        "Semáforo de vencimiento (+90 días)",
        options=semaforos_presentes,
        placeholder="Todos",
    )
    if sel_semaforo:
        df_filtrado = df_filtrado[vista[COL_VENCIMIENTO].isin(sel_semaforo)]

    st.sidebar.caption(f"Mostrando **{len(df_filtrado)}** de **{len(df)}** registros.")
    return df_filtrado


def mostrar_tabla(df: pd.DataFrame) -> None:
    """Tabla principal: las columnas de fecha clave se muestran como semáforo."""
    vista = agregar_columnas_semaforo(df)

    columnas_visibles = [
        COL_ID, COL_PROVEEDOR, COL_FOLIO, COL_CARTA_FIRMADA, COL_IMPORTE,
        COL_COMPRADOR, COL_FECHA_CORTE,
        COL_LIMITE_NOTIFICAR, COL_SEGUIMIENTO, COL_AVISO_CXP, COL_VENCIMIENTO,
        COL_CARTA_ENVIADA, COL_FECHA_ENVIO, COL_RESPUESTA, COL_FECHA_RESPUESTA,
        COL_AVISO_ENVIADO, COL_ESTATUS, COL_NOTAS,
        COL_MODIFICADO_POR, COL_FECHA_MODIFICACION,
    ]
    columnas_visibles = [c for c in columnas_visibles if c in vista.columns]

    st.dataframe(
        vista[columnas_visibles],
        use_container_width=True,
        hide_index=True,
        column_config={
            COL_IMPORTE: st.column_config.NumberColumn(format="$%.2f"),
            COL_FECHA_CORTE: st.column_config.DateColumn(format="DD/MM/YYYY"),
            COL_FECHA_ENVIO: st.column_config.DateColumn(format="DD/MM/YYYY"),
            COL_FECHA_RESPUESTA: st.column_config.DateColumn(format="DD/MM/YYYY"),
        },
    )


def derivar_estatus_por_fase(carta: bool, respuesta: bool, aviso: bool) -> str:
    """Deriva el ESTATUS ACTUAL según la fase más avanzada activa.

    Fases del proceso de reclamación:
      1. Carta enviada           -> "Enviada"
      2. Respuesta del proveedor -> "En revisión"
      3. Aviso a CxP enviado     -> "En proceso"
    Los estatus finales (Aceptado, Devuelto, Pagado, Rechazado, Terminado)
    se asignan manualmente con el selector.
    """
    if aviso:
        return "En proceso"
    if respuesta:
        return "En revisión"
    if carta:
        return "Enviada"
    return ""


def formulario_edicion(df: pd.DataFrame) -> None:
    """Formulario para editar un registro, firmarlo y subirlo a GitHub."""
    st.subheader("✏️ Editar reclamación")

    if df.empty:
        st.info("No hay registros con los filtros actuales.")
        return

    # --- Selección del registro por clave única (ID + FOLIO) ---
    df_sel = df.copy()
    df_sel["ETIQUETA"] = (
        "ID " + df_sel[COL_ID].astype(str)
        + " · " + df_sel[COL_FOLIO]
        + " · " + df_sel[COL_PROVEEDOR].str.slice(0, 40)
    )
    etiqueta = st.selectbox(
        "Selecciona la reclamación (ID · Folio · Proveedor)",
        options=df_sel["ETIQUETA"].tolist(),
    )
    fila = df_sel[df_sel["ETIQUETA"] == etiqueta].iloc[0]
    clave = fila["CLAVE"]

    # --- Resumen del registro seleccionado ---
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"**Proveedor:** {fila[COL_PROVEEDOR]}")
        importe = pd.to_numeric(pd.Series([fila.get(COL_IMPORTE)]), errors="coerce").iloc[0]
        c2.markdown(f"**Importe:** ${importe:,.2f} MXN" if pd.notna(importe) else "**Importe:** —")
        c3.markdown(f"**Comprador:** {fila[COL_COMPRADOR]}")
        c1.markdown(f"**Estatus actual:** `{fila[COL_ESTATUS] or '—'}`")
        c2.markdown(f"**Vencimiento:** {calcular_semaforo(fila.get(COL_VENCIMIENTO))}")
        c3.markdown(f"**Folio:** {fila[COL_FOLIO]}")

    # --- Formulario de edición ---
    with st.form("form_edicion", clear_on_submit=False):
        st.markdown("##### Fases del proceso (marca las completadas)")
        col_a, col_b, col_c = st.columns(3)
        chk_carta = col_a.checkbox(
            "📤 ¿Carta enviada?",
            value=str(fila.get(COL_CARTA_ENVIADA, "")).strip().upper() in ("SÍ", "SI"),
        )
        chk_respuesta = col_b.checkbox(
            "📨 ¿Respuesta del proveedor?",
            value=str(fila.get(COL_RESPUESTA, "")).strip().upper() in ("SÍ", "SI"),
        )
        chk_aviso = col_c.checkbox(
            "📣 ¿Aviso a CxP enviado?",
            value=str(fila.get(COL_AVISO_ENVIADO, "")).strip().upper() in ("SÍ", "SI"),
        )

        st.markdown("##### Estatus final")
        estatus_sugerido = derivar_estatus_por_fase(chk_carta, chk_respuesta, chk_aviso)
        indice_defecto = (
            OPCIONES_ESTATUS.index(fila[COL_ESTATUS])
            if fila[COL_ESTATUS] in OPCIONES_ESTATUS
            else (OPCIONES_ESTATUS.index(estatus_sugerido) if estatus_sugerido in OPCIONES_ESTATUS else 0)
        )
        usar_automatico = st.toggle(
            "Asignar estatus automáticamente según las fases marcadas",
            value=True,
            help="Carta → Enviada · Respuesta → En revisión · Aviso CxP → En proceso. "
                 "Desactívalo para elegir un estatus final manualmente.",
        )
        estatus_manual = st.selectbox(
            "Estatus (manual)", options=OPCIONES_ESTATUS, index=indice_defecto
        )

        st.markdown("##### Comentarios y firma")
        comentario = st.text_area(
            "Comentarios / acciones sobre este estatus",
            placeholder="Describe la acción realizada, acuerdos, fechas comprometidas…",
        )
        usuario = st.text_input("Tu Nombre / Usuario", placeholder="Ej. María López")

        guardar = st.form_submit_button("💾 Guardar y subir a GitHub", type="primary",
                                        use_container_width=True)

    if not guardar:
        return

    # --- Validaciones ---
    if not usuario.strip():
        st.error("✍️ Debes firmar con tu nombre antes de guardar.")
        return

    # --- Aplicar cambios al DataFrame completo (en sesión) ---
    df_completo = st.session_state["df"]
    mascara = df_completo["CLAVE"] == clave
    if not mascara.any():
        st.error("No se encontró el registro seleccionado. Recarga la página.")
        return

    ahora = datetime.now()
    hoy = date.today()
    valor_anterior = df_completo.loc[mascara, COL_ESTATUS].iloc[0]

    # Fases (SÍ/NO) y fechas asociadas automáticas
    df_completo.loc[mascara, COL_CARTA_ENVIADA] = "SÍ" if chk_carta else "NO"
    df_completo.loc[mascara, COL_RESPUESTA] = "SÍ" if chk_respuesta else "NO"
    df_completo.loc[mascara, COL_AVISO_ENVIADO] = "SÍ" if chk_aviso else "NO"
    if chk_carta and _a_fecha(df_completo.loc[mascara, COL_FECHA_ENVIO].iloc[0]) is None:
        df_completo.loc[mascara, COL_FECHA_ENVIO] = pd.Timestamp(hoy)
    if chk_respuesta and _a_fecha(df_completo.loc[mascara, COL_FECHA_RESPUESTA].iloc[0]) is None:
        df_completo.loc[mascara, COL_FECHA_RESPUESTA] = pd.Timestamp(hoy)

    # Estatus: automático por fase o manual
    estatus_final = (
        estatus_sugerido if (usar_automatico and estatus_sugerido) else estatus_manual
    )
    df_completo.loc[mascara, COL_ESTATUS] = estatus_final

    # Notas: se anexa el comentario con sello de fecha/hora y firma
    if comentario.strip():
        nota_previa = str(df_completo.loc[mascara, COL_NOTAS].iloc[0] or "").strip()
        if nota_previa.lower() in ("nan", "none"):
            nota_previa = ""
        nueva_nota = f"[{ahora:%d/%m/%Y %H:%M} · {usuario.strip()}] {comentario.strip()}"
        df_completo.loc[mascara, COL_NOTAS] = (
            f"{nota_previa} | {nueva_nota}" if nota_previa else nueva_nota
        )

    # Auditoría
    df_completo.loc[mascara, COL_MODIFICADO_POR] = usuario.strip()
    df_completo.loc[mascara, COL_FECHA_MODIFICACION] = f"{ahora:%d/%m/%Y %H:%M:%S}"

    # --- Guardar Excel y subir a GitHub ---
    try:
        guardar_excel(df_completo, RUTA_EXCEL)
    except Exception as e:
        st.error(f"No se pudo escribir el archivo Excel: {e}")
        return

    mensaje_commit = (
        f"Actualización reclamación {clave}: "
        f"'{valor_anterior or '—'}' → '{estatus_final}' por {usuario.strip()} "
        f"({ahora:%d/%m/%Y %H:%M})"
    )
    with st.spinner("Subiendo cambios a GitHub…"):
        exito, mensaje = subir_a_github(mensaje_commit)

    # Invalidar caché y refrescar datos en sesión
    st.session_state["version_datos"] += 1
    st.session_state["df"] = df_completo

    if exito:
        st.success(f"✅ Registro actualizado. {mensaje}")
        st.toast("Cambios guardados", icon="💾")
    else:
        st.warning(f"💾 El Excel se guardó localmente, pero: {mensaje}")


# =============================================================================
# 6. FUNCIÓN PRINCIPAL
# =============================================================================

def main() -> None:
    if not verificar_acceso():
        st.stop()

    # --- Estado inicial ---
    if "version_datos" not in st.session_state:
        st.session_state["version_datos"] = 0
    if "df" not in st.session_state:
        st.session_state["df"] = cargar_datos(
            RUTA_EXCEL, st.session_state["version_datos"]
        )

    df = st.session_state["df"]

    # --- Encabezado ---
    st.title("📋 Dashboard de Devoluciones y Reclamaciones")
    st.caption(
        "Gestiona el registro de reclamaciones, actualiza estatus y sincroniza "
        "automáticamente con GitHub."
    )

    # --- Barra lateral ---
    st.sidebar.title("⚙️ Panel de control")
    if st.sidebar.button("🔄 Recargar datos", use_container_width=True):
        st.session_state["version_datos"] += 1
        st.session_state["df"] = cargar_datos(
            RUTA_EXCEL, st.session_state["version_datos"]
        )
        st.rerun()
    boton_cerrar_sesion()
    st.sidebar.divider()

    df_filtrado = construir_filtros(df)

    # --- Cuerpo principal ---
    mostrar_kpis(df_filtrado)
    st.divider()

    tab_tabla, tab_editar = st.tabs(["📊 Tabla de reclamaciones", "✏️ Editar registro"])
    with tab_tabla:
        st.markdown(
            f"Leyenda de fechas clave: {EN_TIEMPO} · {POR_VENCERSE} · {VENCIDO}"
        )
        mostrar_tabla(df_filtrado)
    with tab_editar:
        formulario_edicion(df_filtrado)


if __name__ == "__main__":
    main()
