# Handoff — Contabilidad RM2

Documento para retomar el trabajo de esta app desde otra cuenta (otro equipo, otra sesión de Claude Code, u otro desarrollador).

---

## 1. Qué es la app

App web en **Streamlit** para el equipo de contabilidad de **Fondo Raíz RM2**. Un solo dashboard ("Contabilidad RM2") con **dos módulos**:

| Módulo | Qué hace | Entrada | Salida |
|---|---|---|---|
| **Conciliación** | Verifica que los movimientos contables cuadren contra el listado de propietarios (IVA TC14 = 19% de comisión TC20, retefuente TC23 = 11%, etc.) | Archivo de propietarios + archivo contable (+ opcional catálogo TC) | Excel multi-hoja con verificaciones y semáforos |
| **Actas Fondo — Contai** | Genera las planillas de asientos contables (ACTA y COMISIONES) a partir del movimiento bancario | `Movimiento Cta 28150501 RM2 -Fondo.xls` | Dos Excel: Plantilla ACTA + Plantilla COMISIONES |

Ambos módulos viven en **un único `app.py`** (~2059 líneas). La navegación es por `st.session_state["modulo"]` (`None` = landing, `"conciliacion"`, `"contai"`).

---

## 2. Repositorio y despliegue

- **GitHub**: https://github.com/jlloalejo-art/contabilidad-rm2.git (cuenta `jlloalejo-art`)
- **Rama de producción**: `main`
- **Hosting**: Streamlit Community Cloud — **auto-despliega al hacer push a `main`** (tarda 1-2 min).
- **Repo público** (decisión tomada porque el plan Community solo permite 1 app privada por workspace; los públicos son ilimitados).
- **Carpeta local**: `/Users/alejandro/renta-metro-cuadrado/Contabilidad RM2/`

### Archivos del repo
```
app.py                    ← toda la lógica + UI (ambos módulos)
requirements.txt          ← streamlit, pandas, openpyxl, xlrd
.streamlit/config.toml    ← tema claro fijo (navy #1a4a7a)
HANDOFF.md                ← este documento
```

### Para retomar desde otra cuenta
1. Clonar: `git clone https://github.com/jlloalejo-art/contabilidad-rm2.git`
2. (Opcional) Correr local: `pip install -r requirements.txt && streamlit run app.py`
3. Editar `app.py`, commit, `git push` → Streamlit Cloud redespliega solo.
4. Si vas a desplegar en **otra** cuenta de Streamlit Cloud: apunta un nuevo deploy al repo (o a un fork), archivo principal `app.py`, rama `main`.

> **Nota**: existe una app "hermana" antigua de solo-conciliación en otra URL de Streamlit (`conciliacion-rm2-...streamlit.app`). Ya está integrada aquí; esa app vieja se puede ignorar/retirar.

---

## 3. Mapa de `app.py`

| Líneas aprox. | Sección |
|---|---|
| 9–100 | **Constantes** (NIT, cuentas, sets de TipoCausa, catálogo TC) |
| 102–222 | Helpers generales + helpers de estilo Excel (`_hdr`, `_row`, `_color`, `_widths`…) |
| 224–521 | **`run_conciliacion()`** — lógica del módulo Conciliación |
| 523–1028 | **`build_excel()`** — genera el Excel de conciliación (multi-hoja) |
| 1030–1117 | Helpers del módulo Contai (`parse_valor`, `es_reversion`, `format_ad_ref`, `get_cuenta`, `tipo_mov`, `make_row`) |
| 1119–1188 | **`procesar_acta()`** |
| 1190–1264 | **`procesar_comisiones()`** |
| 1266–1345 | `generar_excel()` + `resumen_cuentas()` (Contai) |
| 1347–1623 | `set_page_config` + **CSS global** |
| 1624–1647 | Navegación (`_nav_bar`, session_state) |
| 1649+ | Landing + render de cada módulo |

---

## 4. Reglas de negocio críticas (Contai) — NO cambiar sin verificar contra archivos de referencia

Estas reglas se determinaron **empíricamente** comparando contra planillas hechas a mano. Son contraintuitivas; documentadas aquí para no revertirlas por error.

### Constantes clave
```python
NIT_FONDO     = 901155077
CENTRO_COSTO  = "0501"
CUENTA_CONTRA = 22050501           # contrapartida auto-generada en COMISIONES
TC_ACTA       = {1,2,3,4,7,11,30,32,33,37}
TC_COMISIONES = {20,14,23}
TC_EXCLUIR    = {0,16,40}
```

### Dirección del Tipo (columna H)
- Fuente **CRÉDITO → `2`** en la planilla
- Fuente **DÉBITO → `1`**
- (`tipo_mov(debito, credito) = 2 if credito>0 else 1`)

### Gravado vs No Gravado (Canon, TC=1) — **invertido respecto a la descripción verbal**
- **Forward** (no reversión): CON IVA → **No Gravado** (`41550501`); SIN IVA → **Gravado** (`41550502`)
- **Reversión**: CON IVA → **Gravado reversal** (`41750501`); SIN IVA → **No Gravado reversal** (`41750502`)
- Un canon "tiene IVA" si el mismo `No.Transacción` tiene también un TC=3 (IVA COMERCIAL).

### Reversiones
- Solo las reversiones de **comisiones** (incluye IVA comisiones TC14 y retefuente comisiones TC23) son comprobante **`DV`** y llevan el **número de transacción de la cuenta original**.
- Reversiones de canon/IVA/retefte en el ACTA son **`CP`** (no DV).

### Formato Doc Ref de reversión (`format_ad_ref`)
`2026030005` → `AD2630005` (año corto `26`, mes sin cero `3`, secuencia `0005`). Doc Ref de comisiones normales: `FE{trans_int}`.

### ACTA — otras reglas
- Todos los renglones: comprobante `CP`, misma **fecha_acta** (input del usuario, no la fecha por fila), mismo documento.
- Orden: grupo TC={1,3,4} primero, luego el resto; dentro de cada grupo por `No.Trans` ASC, `TipoCausa` ASC.

### COMISIONES — otras reglas
- Agrupar por `No.Transacción`. Consecutivos separados para CP (`cons_cp`) y DV (`cons_dv`).
- TC=20: columna Base va **vacía** (no tomar el Valor Base de la fuente, que trae el canon).
- Contrapartida `22050501` por grupo: `valor = com + iva - ret`, tipo=2 para CP, tipo=1 para DV.

### Decisión pendiente conocida
- **AUTORRENTA** (`23657102` + `13551590`, ~1.1% del canon): el usuario decidió **NO** auto-generarlas.

---

## 5. Estado actual / último trabajo

- **Últimas correcciones (julio 2026)**:
  1. **`IllegalCharacterError` también en el módulo Contai**: `generar_excel()` escribía el texto crudo del `.xls` (columna Detalle con caracteres de control) sin pasar por `_sanitize()` — fallaba la descarga de las plantillas ACTA y COMISIONES. Se aplicó `_sanitize()` a cada celda en `generar_excel()` (línea ~1334). Esto cierra el punto 1 de "Cosas a vigilar": ahora *ambos* módulos sanitizan.
  2. **Conciliación — filtros borraban el reporte**: el resultado vivía dentro de `if st.button("▶ Conciliar")`; al usar un filtro (rerun con botón en `False`) desaparecía todo. Ahora el resultado se guarda en `st.session_state["concil_result"]` y se renderiza fuera del `if` (línea ~1768).
  3. **Propietarios — lectura frágil**: `prop[prop[12] == 'False']` solo aceptaba el texto `'False'` (fallo silencioso si el export traía booleano) y columnas por posición fija. Ahora hay `_find_prop_col()` (busca por encabezado, respaldo posicional) y `_es_activo()` (acepta texto o booleano) — línea ~224.
  4. **`requirements.txt`**: se mantiene `pandas>=2.0.0`. Ojo: el código usa `include_groups=False` (pandas ≥ 2.2), pero Cloud ya instala una versión ≥2.2 en la práctica. Subir el piso a `>=2.2.0` invalidó el entorno cacheado de Streamlit Cloud (Python 3.14) y tumbó la app ("Oh no. Error running app"), por lo que se revirtió. **No tocar el piso de pandas sin verificar el redeploy.**
- **Corrección previa (commit `2ad914c`)**: `IllegalCharacterError` en el módulo Conciliación. Se agregó `_sanitize()` (línea ~523) aplicado a `src1`, `src2`, `nombre_cont`, `detalle`.
- **UI**: tema claro fijo, paleta navy (`#1a4a7a` / `#0f2744`), fuente Inter, landing con tarjetas-botón por módulo, KPI cards.
- **Validación de exactitud**: ACTA 595 filas = 595 de referencia; COMISIONES 528 = 528. Archivos de mayo (CANON/OTROS) también confirmados 100%.

  5. **Contai — lectura de columnas robusta**: `df_src["TipoCausa"]` (y demás) tumbaba la app con `KeyError` si el export traía encabezados con variaciones. Ahora, tras leer el archivo, se normalizan los encabezados por nombre (`_col_contai`, con variantes de acentos/espacios) y si falta una columna esencial se muestra un `st.error` claro con las columnas encontradas en vez de crashear (línea ~1968).

### Tareas conocidas pendientes (revisión de código, aún sin abordar)
- **Parseo de dinero** (`parse_money`/`parse_valor`): asume formato US (coma miles, punto decimal); con formato colombiano `1.234.567,89` devuelve 0 en silencio.
- Menores: título del Excel de conciliación hardcodeado "CORTE ABRIL 2026" (línea ~539); `No. Inmueble` se muestra como flotante (`101.0`).

---

## 6. Cosas a vigilar si retomas

1. **Si aparece de nuevo un error de caracteres al descargar**: revisa que `_sanitize()` se aplique en *toda* escritura de celda con texto proveniente del `.xls` fuente (no solo las cubiertas hoy).
2. **No inviertas la lógica Gravado/No Gravado ni la dirección del Tipo** — parecen "al revés" pero son correctas contra la referencia real. Verifica siempre contra una planilla hecha a mano antes de tocar `get_cuenta` o `tipo_mov`.
3. **`xlrd`** es necesario para leer los `.xls` viejos (formato Excel 97-2003). No lo quites de `requirements.txt`.
4. **Streamlit Cloud** cachea; si un cambio no aparece, fuerza "Reboot app" desde Manage app.
5. **`requirements.txt` es un lock COMPLETO** (las 45 versiones fijadas con `==`, tomadas del entorno que funcionaba en Cloud con Python 3.14.6). Build 100% reproducible. Motivo: un rebuild subió `pyarrow` 24→25 y `websockets` 16.0→16.1 y la app empezó a caer con `Segmentation fault` → "Oh no. Error running app". La pista de un segfault es la línea `run-streamlit.sh: ... Segmentation fault`, no un traceback de Python. (pandas 3.0.3 NO es el problema; corría bien.)
   - **Para actualizar dependencias a propósito**: cambia la versión, push, y **verifica el log de build** (Manage app → logs). Si aparece un segfault, revierte.
   - **Riesgo del lock**: si Streamlit Cloud cambia de versión de Python, alguna versión pineada podría no tener wheel. En ese caso, actualiza los paquetes afectados (empezando por `pyarrow`/`numpy`/`pandas`) verificando el log.
