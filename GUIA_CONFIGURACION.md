# Guía — Dashboard de Devoluciones y Reclamaciones (seguimiento por etapas)

## Archivos del repositorio

En la raíz del repo `ftrevinor001-tr/dashboard-reclamaciones`:

```
app.py             ← aplicación (máquina de estados de 4 etapas)
requirements.txt   ← dependencias
datos.xlsx         ← base de datos (hoja "datos") — usa la versión NUEVA con las columnas de etapas
```

> ⚠️ Importante: sube el `datos.xlsx` nuevo que se entrega junto a esta versión. Conserva tus 89 registros y añade todas las columnas que las etapas necesitan (estatus, fechas, usuarios y observaciones por etapa). Si subes el Excel viejo, la app igual creará las columnas faltantes, pero es más limpio partir del nuevo.

## Secretos en Streamlit Cloud

En **⋮ → Settings → Secrets** (formato TOML):

```toml
GITHUB_TOKEN = "github_pat_XXXXXXXXXXXX"
DASHBOARD_PASSWORD = "TuContraseñaSegura"
```

- `GITHUB_TOKEN`: token fine-grained con acceso al repositorio y permiso **Contents: Read and write**.
- `DASHBOARD_PASSWORD`: contraseña de acceso al dashboard.

## Clave de reactivación

Para reabrir una etapa ya cerrada se usa la clave **`devoluciones2026`**, escrita en el código (constante `CLAVE_REACTIVACION` en `app.py`). Si quieres cambiarla, edita esa línea. Compártela solo con el personal autorizado.

## Cómo funciona el proceso

El identificador de cada reclamación es el **Folio de reporte**. Cada reclamación recorre 4 etapas; al terminar la última fase de una etapa, esa etapa se bloquea y se activa la siguiente:

1. **Reporte de reclamo** (7 días desde la fecha de corte): Recepción de Folio → Revisión de folio → Envío a proveedores.
2. **Gestión** (30 días): Enviado a proveedor → Seguimiento → Respuesta de proveedor. Se bifurca según la respuesta: recolección (→ Destino final), destrucción o sin respuesta (→ Cuentas por pagar).
3. **Cuentas por pagar** (53 días): Negociación → Seguimiento → Recepción de nota de crédito → Aplicación de pago.
4. **Destino final**: recolección (Programación → Recolección → Recepción de folio de devolución, límite 20 días) o destrucción (Informe a proveedor → Definición de destino final → Reporte al almacén → Recepción de folio de ajuste).

Cada paso guarda fecha y usuario; las observaciones por etapa se guardan en columnas propias y la bitácora completa se acumula en *Acciones / Notas*. La pestaña **Resumen y alarmas** lista las reclamaciones por vencerse y vencidas (estructura lista para correo más adelante). La pestaña **Guía** explica el proceso dentro de la app.
